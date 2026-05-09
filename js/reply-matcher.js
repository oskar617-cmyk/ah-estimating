// js/reply-matcher.js
// Match an inbound message to an open RFQ supplier entry.
//
// Three tiers:
//   1. In-Reply-To header: most replies. The message has a header
//      "In-Reply-To: <internetMessageId>" pointing at our original RFQ.
//   2. PDF content match: accounting-app emails (Xero/MYOB) come from
//      generic addresses with no In-Reply-To. We extract attached-PDF text
//      and look for the job's address / job code / trade name.
//   3. Manual: surface to user, they pick the matching open RFQ.
//
// Across all tiers we walk every job's rfq-tracker.json. To keep this
// efficient on poll runs we cache per-job-folder trackers in
// state.jobTrackerCache for the duration of one inbox-processing pass.

import { CONFIG } from './config.js';
import { state } from './state.js';
import {
  graphFetch, getAhSiteId, encodeUriPath,
  listMessageAttachments, getAttachmentBytes
} from './graph.js';
import { readTracker } from './audit.js';
import { extractPdfText } from './pdf-tools.js';

// Helper: list all job folders in AH Site Documents that match the pattern.
// Returns array of { folderName, jobCode, jobName, address }.
async function listAllJobFolders() {
  const siteId = await getAhSiteId();
  const r = await graphFetch(`/sites/${siteId}/drive/root/children?$top=200&$select=id,name,folder`);
  const items = r.value || [];
  return items
    .filter(it => it.folder && CONFIG.jobFolderPattern.test(it.name))
    .map(it => {
      const m = it.name.match(CONFIG.jobFolderPattern);
      return { folderName: it.name, jobCode: m[1].trim(), jobName: m[2].trim(), address: m[3].trim() };
    });
}

// Read (and cache for this pass) a job's rfq-tracker.json.
async function loadTrackerCached(jobFolderName) {
  if (state.jobTrackerCache.has(jobFolderName)) {
    return state.jobTrackerCache.get(jobFolderName);
  }
  const tracker = await readTracker(jobFolderName);
  state.jobTrackerCache.set(jobFolderName, tracker);
  return tracker;
}

// Clear the per-pass tracker cache (called at end of poll cycle).
export function clearTrackerCache() {
  state.jobTrackerCache.clear();
}

// Build the lookup tables used by Tier 1. Returns:
//   { byMessageId: Map<msgId, { jobFolder, rfqId, supplierId, supplierEmail }>,
//     byConversationId: Map<convId, { ... }> }
async function buildSentMessageIndex() {
  const byMessageId = new Map();
  const byConversationId = new Map();
  const jobs = await listAllJobFolders();
  for (const job of jobs) {
    const tracker = await loadTrackerCached(job.folderName);
    if (!tracker || !Array.isArray(tracker.rfqs)) continue;
    for (const rfq of tracker.rfqs) {
      for (const sup of (rfq.suppliers || [])) {
        const ref = {
          jobFolder: job.folderName,
          jobCode: job.jobCode,
          jobName: job.jobName,
          rfqId: rfq.id,
          rfqCategory: rfq.category,
          supplierId: sup.id,
          supplierEmail: (sup.email || '').toLowerCase(),
          supplierCompany: sup.companyName,
          supplierContact: sup.contactName,
          budgetRowNo: rfq.budgetRowNo
        };
        if (sup.internetMessageId) byMessageId.set(sup.internetMessageId, ref);
        if (sup.conversationId) byConversationId.set(sup.conversationId, ref);
      }
    }
  }
  return { byMessageId, byConversationId, jobs };
}

// ---------- Tier 1: In-Reply-To header ----------

function getHeader(message, name) {
  const headers = message.internetMessageHeaders || [];
  const lower = name.toLowerCase();
  for (const h of headers) {
    if ((h.name || '').toLowerCase() === lower) return h.value || '';
  }
  return '';
}

function tier1Match(message, index) {
  // Microsoft Graph stores headers in internetMessageHeaders, but the
  // In-Reply-To and References values come back wrapped in <...> brackets.
  const inReplyTo = getHeader(message, 'In-Reply-To').trim();
  const references = getHeader(message, 'References').trim();
  // Try In-Reply-To first
  if (inReplyTo) {
    const ids = parseMessageIds(inReplyTo);
    for (const id of ids) {
      if (index.byMessageId.has(id)) return { tier: 1, ref: index.byMessageId.get(id) };
    }
  }
  // Fall back to scanning References (chains of message-ids)
  if (references) {
    const ids = parseMessageIds(references);
    for (const id of ids) {
      if (index.byMessageId.has(id)) return { tier: 1, ref: index.byMessageId.get(id) };
    }
  }
  // Last resort within Tier 1: conversationId match (Graph's threading)
  if (message.conversationId && index.byConversationId.has(message.conversationId)) {
    return { tier: 1, ref: index.byConversationId.get(message.conversationId) };
  }
  return null;
}

function parseMessageIds(headerValue) {
  // Headers like "<id1@host>" or "<id1@host> <id2@host>"
  const matches = headerValue.match(/<[^>]+>/g) || [];
  return matches.map(m => m.replace(/[<>]/g, ''));
}

// ---------- Tier 2: PDF content match ----------

// For each PDF attachment, extract text and search for job code (e.g. "2304"),
// job name (e.g. "GV"), or address fragment. If a match is unique, use it.
async function tier2Match(message, index) {
  if (!message.hasAttachments) return null;
  let attachments;
  try {
    attachments = await listMessageAttachments(message.id);
  } catch (e) { console.warn('Tier 2 attachment list failed:', e); return null; }
  const pdfs = attachments.filter(a => /\.pdf$/i.test(a.name) || a.contentType === 'application/pdf');
  for (const att of pdfs) {
    let text = '';
    try {
      const bytes = await getAttachmentBytes(message.id, att.id);
      text = await extractPdfText(bytes);
    } catch (e) {
      console.warn(`PDF text extraction failed for ${att.name}:`, e);
      continue;
    }
    const ref = scanPdfTextForMatch(text, index);
    if (ref) return { tier: 2, ref, evidence: att.name };
  }
  return null;
}

function scanPdfTextForMatch(text, index) {
  if (!text) return null;
  const lower = text.toLowerCase();
  // Score each job by how many distinctive tokens appear in the PDF text.
  // The most-distinctive token is jobCode (4 digits) — but it's risky if the
  // PDF has other 4-digit numbers (postcodes, ABNs, etc). We require the
  // jobCode to appear together with the jobName or address fragment.
  let best = null;
  let bestScore = 0;
  for (const job of index.jobs) {
    let score = 0;
    if (job.jobCode && lower.includes(job.jobCode)) score += 2;
    if (job.jobName && new RegExp(`\\b${escapeRegex(job.jobName.toLowerCase())}\\b`).test(lower)) score += 2;
    // Address fragment: take the first 2 words of the address (e.g. "31 Langford")
    const addrFrag = (job.address || '').split(/\s+/).slice(0, 2).join(' ').toLowerCase();
    if (addrFrag && addrFrag.length >= 4 && lower.includes(addrFrag)) score += 3;
    if (score > bestScore) { bestScore = score; best = job; }
  }
  // Require minimum score of 4 (e.g. jobCode + jobName, or addrFrag + jobCode)
  if (!best || bestScore < 4) return null;
  // We matched to a job. Now find the most likely RFQ within that job
  // by scanning text for trade-category keywords. If multiple match,
  // pick the one with strongest hit.
  const tracker = state.jobTrackerCache.get(best.folderName);
  if (!tracker || !Array.isArray(tracker.rfqs)) return null;
  let bestRfq = null;
  let bestRfqScore = 0;
  for (const rfq of tracker.rfqs) {
    const cat = (rfq.category || '').toLowerCase();
    if (!cat) continue;
    const reCat = new RegExp(`\\b${escapeRegex(cat)}\\b`, 'i');
    if (reCat.test(text)) {
      // Multi-supplier RFQ — Tier 2 can't tell which supplier within the
      // RFQ this PDF is for. Most accounting-app emails have one supplier
      // anyway. We return the RFQ-level match without a specific supplier
      // and let the manual review flow disambiguate if needed.
      const score = cat.length; // longer match = more specific
      if (score > bestRfqScore) {
        bestRfqScore = score;
        bestRfq = rfq;
      }
    }
  }
  if (!bestRfq) return null;
  // Try to pick the supplier by sender email if present in the index map
  // — but we don't have message context here, so caller will fill in if able.
  return {
    jobFolder: best.folderName,
    jobCode: best.jobCode,
    jobName: best.jobName,
    rfqId: bestRfq.id,
    rfqCategory: bestRfq.category,
    supplierId: null,         // unknown — caller may supply if it can
    supplierEmail: null,
    supplierCompany: null,
    supplierContact: null,
    budgetRowNo: bestRfq.budgetRowNo
  };
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

// ---------- Public API ----------

// Match a single message. Returns:
//   { tier: 1|2|3, ref: { jobFolder, rfqId, supplierId, ... } }  on success
//   { tier: 3, ref: null }                                        on no match (manual)
export async function matchMessageToRfq(message, indexCache) {
  const index = indexCache || (await buildSentMessageIndex());
  // Tier 1
  const t1 = tier1Match(message, index);
  if (t1) return t1;
  // Tier 2
  const t2 = await tier2Match(message, index);
  if (t2) return t2;
  // Tier 3: needs manual matching
  return { tier: 3, ref: null };
}

// Build the index once and return it so the caller can re-use it across
// many messages in a single poll cycle (avoids re-reading every tracker
// for every message).
export async function buildIndex() {
  return buildSentMessageIndex();
}
