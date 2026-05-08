// js/inbox.js
// Inbox poller: every 30s, scan est@'s inbox for new messages, match each
// to an open RFQ, classify with AI, extract amount, save attachments to
// SharePoint, write quote to budget Excel, file the email, and surface a
// notification.

import { CONFIG } from './config.js';
import { state } from './state.js';
import {
  graphFetch, getAhSiteId, encodeUriPath, readJson, uploadJson,
  listInboxSince, getMessage, listMessageAttachments, getAttachmentBytes,
  arrayBufferToBase64, readXlsx, markMessageRead
} from './graph.js';
import { getToken } from './auth.js';
import { buildIndex, matchMessageToRfq, clearTrackerCache } from './reply-matcher.js';
import { extractPdfText, emailBodyToPdfBytes } from './pdf-tools.js';
import { classifyEmail, extractQuoteAmount, summarizeFilename } from './classification.js';
import { fileMessageToJob, buildOfficeFolderName, clearFolderIdCache } from './mail-filer.js';
import { addNotification } from './notifications.js';
import { logAudit } from './audit.js';

// ----- Public boot/teardown -----

export function startInboxPoller() {
  // Only sender's inbox is polled. If app is opened by oskar (admin) but
  // the signed-in account is est@, polling runs.
  if (!state.currentUserEmail) return;
  // Only poll for the sender mailbox (per project decision: only est@auhs
  // sends and receives RFQs in v1).
  if (state.currentUserEmail !== CONFIG.senderEmail) {
    console.log(`Inbox poller skipped — signed in as ${state.currentUserEmail}, not ${CONFIG.senderEmail}`);
    return;
  }
  if (state.inboxPollerHandle) return;  // already running
  // Initialise the lookback window
  if (!state.inboxLastPolledAt) {
    state.inboxLastPolledAt = new Date(
      Date.now() - CONFIG.inboxLookbackMinutesOnFirstPoll * 60 * 1000
    ).toISOString();
  }
  // First run immediately, then on interval
  pollInbox().catch(err => console.warn('Initial inbox poll failed:', err));
  state.inboxPollerHandle = setInterval(() => {
    pollInbox().catch(err => console.warn('Inbox poll failed:', err));
  }, CONFIG.inboxPollIntervalMs);
}

export function stopInboxPoller() {
  if (state.inboxPollerHandle) {
    clearInterval(state.inboxPollerHandle);
    state.inboxPollerHandle = null;
  }
}

// ----- Per-pass orchestration -----

let pollInProgress = false;

async function pollInbox() {
  if (pollInProgress) return;
  pollInProgress = true;
  try {
    const since = state.inboxLastPolledAt;
    const messages = await listInboxSince(since);
    if (!messages.length) {
      state.inboxLastPolledAt = new Date().toISOString();
      return;
    }
    // Build the sent-message index ONCE per poll
    const index = await buildIndex();
    for (const msg of messages) {
      // Dedupe: skip if we've already processed this id this session
      if (state.processedMessageIds.has(msg.id)) continue;
      try {
        await processMessage(msg, index);
      } catch (err) {
        console.error(`Process message ${msg.id} failed:`, err);
      }
    }
    state.inboxLastPolledAt = new Date().toISOString();
  } finally {
    pollInProgress = false;
    clearTrackerCache();
    clearFolderIdCache();
  }
}

// ----- Processing pipeline for one message -----

async function processMessage(msg, index) {
  // Skip messages we sent ourselves (Graph filter on Inbox should already
  // exclude these, but defensive check).
  if (((msg.from || {}).emailAddress || {}).address &&
      msg.from.emailAddress.address.toLowerCase() === CONFIG.senderEmail.toLowerCase()) {
    state.processedMessageIds.add(msg.id);
    return;
  }

  // Tier 1/2 matching
  const match = await matchMessageToRfq(msg, index);
  // If Tier 3 (no match), surface as Pending Match notification and stop —
  // user will trigger processNotificationManualMatch later.
  if (match.tier === 3) {
    await emitPendingMatch(msg);
    state.processedMessageIds.add(msg.id);
    return;
  }

  // Fully process
  await fullyProcessMatched(msg, match);
  state.processedMessageIds.add(msg.id);
}

// Pending Match notification — user must manually pick the RFQ.
async function emitPendingMatch(msg) {
  // Need full body for the manual-match modal to show context
  const full = await getMessage(msg.id);
  addNotification({
    id: msg.id,
    createdAt: new Date().toISOString(),
    read: false,
    classification: 'Pending Match',
    subject: msg.subject,
    fromName: ((msg.from || {}).emailAddress || {}).name || '',
    fromEmail: ((msg.from || {}).emailAddress || {}).address || '',
    receivedAt: msg.receivedDateTime,
    bodyPreview: msg.bodyPreview,
    bodyHtml: full.body && full.body.contentType === 'html' ? full.body.content : escapeAsPre(full.body && full.body.content),
    job: null, rfqId: null, rfqCategory: null, supplier: null,
    budgetRowNo: null, extractedAmount: null, savedAttachments: [],
    tier: 3
  });
}

function escapeAsPre(s) {
  if (!s) return '';
  return `<pre style="white-space:pre-wrap;font-family:inherit;">${s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>`;
}

// Process a message that we've matched to an RFQ.
async function fullyProcessMatched(msg, match) {
  const ref = match.ref;
  // Fetch full message body for classification + storage
  const full = await getMessage(msg.id);
  const bodyHtml = full.body && full.body.contentType === 'html'
    ? full.body.content
    : escapeAsPre(full.body && full.body.content);
  const bodyText = htmlToPlainText(bodyHtml);

  // Classify
  let classification = 'Question', confidence = 0;
  try {
    const c = await classifyEmail({
      subject: msg.subject,
      fromName: ((msg.from || {}).emailAddress || {}).name || '',
      fromEmail: ((msg.from || {}).emailAddress || {}).address || '',
      bodyText
    });
    classification = c.classification || 'Question';
    confidence = c.confidence || 0;
  } catch (err) {
    console.warn('Classification failed:', err);
  }

  // For OOO and Decline and Suspicious, no quote / no Excel write — just
  // surface the notification and stop. (Per project rules, OOO doesn't
  // bump bell badge — handled by notifications.js.)
  if (classification === 'Out-of-Office' || classification === 'Decline' || classification === 'Suspicious' || classification === 'Unrelated') {
    addNotification({
      id: msg.id,
      createdAt: new Date().toISOString(),
      read: false,
      classification,
      subject: msg.subject,
      fromName: ((msg.from || {}).emailAddress || {}).name || '',
      fromEmail: ((msg.from || {}).emailAddress || {}).address || '',
      receivedAt: msg.receivedDateTime,
      bodyPreview: msg.bodyPreview,
      bodyHtml,
      job: { folderName: ref.jobFolder, code: ref.jobCode, name: ref.jobName, address: '' },
      rfqId: ref.rfqId,
      rfqCategory: ref.rfqCategory,
      supplier: ref.supplierEmail ? { id: ref.supplierId, companyName: ref.supplierCompany, contactName: ref.supplierContact, email: ref.supplierEmail } : null,
      budgetRowNo: ref.budgetRowNo,
      tier: match.tier,
      extractedAmount: null,
      savedAttachments: []
    });
    // Mark read in Outlook (so unread count there matches our processed view)
    try { await markMessageRead(msg.id, true); } catch (e) { /* non-fatal */ }
    // Don't move OOO/Suspicious/Decline/Unrelated out of inbox — user wants
    // those to stay so they can manually deal with them.
    return;
  }

  // For Question or Quote, proceed with attachment save + (Quote) Excel write
  const savedAttachments = await saveAttachments(msg.id, ref);

  // For Quote, extract amount and write to Excel
  let extractedAmount = null;
  let extractedCurrency = 'AUD';
  let pendingEntry = null;
  if (classification === 'Quote') {
    let attachmentText = '';
    // Use the first PDF's text (already extracted during save) if we kept it
    if (savedAttachments.length && savedAttachments[0].text) {
      attachmentText = savedAttachments[0].text;
    }
    try {
      const ext = await extractQuoteAmount({ subject: msg.subject, bodyText, attachmentText });
      if (ext && ext.amount != null) {
        extractedAmount = Number(ext.amount);
        extractedCurrency = ext.currency || 'AUD';
      }
    } catch (err) { console.warn('Amount extraction failed:', err); }

    // Write to budget Excel and create a pendingReview entry
    try {
      pendingEntry = await writeQuoteToBudget(ref, savedAttachments, extractedAmount);
    } catch (err) {
      console.warn('Budget write failed:', err);
    }
  }

  // Update tracker: append a reply entry on the supplier
  try {
    await appendReplyToTracker(ref, {
      messageId: msg.id,
      internetMessageId: full.internetMessageId,
      receivedAt: msg.receivedDateTime,
      classification,
      confidence,
      amount: extractedAmount,
      attachments: savedAttachments.map(a => ({ name: a.savedName, sharePointPath: a.sharePointPath })),
      pendingReviewId: pendingEntry ? pendingEntry.id : null
    }, pendingEntry);
  } catch (err) {
    console.warn('Tracker update failed:', err);
  }

  // File message into job folder
  try {
    const officeFolder = buildOfficeFolderName(ref.jobCode, ref.jobName, await deriveAddress(ref.jobFolder));
    await fileMessageToJob(msg.id, officeFolder);
  } catch (err) {
    console.warn('Mail filing failed:', err);
  }

  // Add notification
  addNotification({
    id: msg.id,
    createdAt: new Date().toISOString(),
    read: false,
    classification,
    subject: msg.subject,
    fromName: ((msg.from || {}).emailAddress || {}).name || '',
    fromEmail: ((msg.from || {}).emailAddress || {}).address || '',
    receivedAt: msg.receivedDateTime,
    bodyPreview: msg.bodyPreview,
    bodyHtml,
    job: { folderName: ref.jobFolder, code: ref.jobCode, name: ref.jobName, address: '' },
    rfqId: ref.rfqId,
    rfqCategory: ref.rfqCategory,
    supplier: ref.supplierEmail ? { id: ref.supplierId, companyName: ref.supplierCompany, contactName: ref.supplierContact, email: ref.supplierEmail } : null,
    budgetRowNo: ref.budgetRowNo,
    extractedAmount, extractedCurrency,
    savedAttachments: savedAttachments.map(a => ({ name: a.savedName, path: a.sharePointPath })),
    tier: match.tier,
    warning: !ref.supplierEmail && match.tier === 2 ? 'Tier 2 match — supplier identity inferred from PDF content. Review carefully.' : null
  });

  await logAudit('REPLY_PROCESSED', `${ref.jobCode} ${ref.jobName} / ${ref.rfqCategory}`, {
    classification, amount: extractedAmount, tier: match.tier
  });
}

// ----- Attachment saving -----

async function saveAttachments(messageId, ref) {
  const saved = [];
  let attachments;
  try {
    attachments = await listMessageAttachments(messageId);
  } catch (e) { console.warn('Attachment list failed:', e); return []; }

  // Determine version: count existing files in Quote/ that match the
  // [trade] - [company] vN pattern.
  const version = await nextVersionForSupplier(ref);

  for (const att of attachments) {
    // Only save PDFs (Word/Excel quotes are rare and we don't have a parser
    // for them in v1)
    if (!/\.pdf$/i.test(att.name) && att.contentType !== 'application/pdf') continue;
    try {
      const bytes = await getAttachmentBytes(messageId, att.id);
      // Optionally extract text + summarize for filename
      let pdfText = '';
      try { pdfText = await extractPdfText(bytes); } catch (e) { /* extraction can fail */ }
      let summary = '';
      try {
        const s = await summarizeFilename({ originalName: att.name, attachmentText: pdfText });
        summary = (s && s.summary) || '';
      } catch (e) { /* fallback to original */ }
      const savedName = buildAttachmentFilename({
        trade: ref.rfqCategory,
        company: ref.supplierCompany,
        version,
        amount: null,  // amount unknown at save time; filename gets renamed after extraction below
        summary: summary || stripExt(att.name)
      });
      const sharePointPath = await uploadAttachment(ref.jobFolder, savedName, bytes);
      saved.push({ originalName: att.name, savedName, sharePointPath, text: pdfText, bytes });
    } catch (err) {
      console.warn(`Attachment ${att.name} save failed:`, err);
    }
  }

  // If no PDFs attached, generate a PDF from the email body
  if (saved.length === 0) {
    try {
      const full = await getMessage(messageId);
      const bodyHtml = full.body && full.body.contentType === 'html'
        ? full.body.content : (full.body && full.body.content) || '';
      const bytes = await emailBodyToPdfBytes({
        subject: full.subject,
        from: ((full.from || {}).emailAddress || {}).address || '',
        to: (full.toRecipients || []).map(r => (r.emailAddress || {}).address).join(', '),
        receivedAt: full.receivedDateTime,
        bodyHtml
      });
      const savedName = buildAttachmentFilename({
        trade: ref.rfqCategory,
        company: ref.supplierCompany,
        version,
        amount: null,
        summary: 'EmailBody'
      });
      const sharePointPath = await uploadAttachment(ref.jobFolder, savedName, bytes);
      saved.push({ originalName: 'email-body.pdf', savedName, sharePointPath, text: '', bytes, generated: true });
    } catch (err) {
      console.warn('Email body PDF generation failed:', err);
    }
  }
  return saved;
}

function buildAttachmentFilename({ trade, company, version, amount, summary }) {
  const tradeSafe = sanitize(trade || 'unknown');
  const companySafe = sanitize(company || 'unknown');
  const summarySafe = sanitize(summary || '').slice(0, 30);
  const amountStr = amount == null ? 'TBC' : Number(amount).toLocaleString('en-AU', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  const parts = [`${tradeSafe} - ${companySafe} v${version} - ${amountStr}`];
  if (summarySafe) parts.push(summarySafe);
  return `${parts.join(' - ')}.pdf`;
}

function sanitize(s) {
  // Remove characters illegal in SharePoint filenames; leave commas alone.
  return String(s || '').replace(/[\\/:*?"<>|#&%]/g, '').replace(/\s+/g, ' ').trim();
}

function stripExt(filename) {
  return filename.replace(/\.[^.]+$/, '');
}

async function nextVersionForSupplier(ref) {
  // Scan Quote/ folder for files matching this trade+company pattern
  try {
    const siteId = await getAhSiteId();
    const path = `${ref.jobFolder}/Quote`;
    const r = await graphFetch(
      `/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/children?$top=500&$select=name,file`
    );
    const files = (r.value || []).filter(it => it.file).map(it => it.name);
    const re = new RegExp(`^${escapeRegex(sanitize(ref.rfqCategory || ''))}\\s*-\\s*${escapeRegex(sanitize(ref.supplierCompany || ''))}\\s*v(\\d+)\\b`, 'i');
    let max = 0;
    for (const f of files) {
      const m = f.match(re);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
    return max + 1;
  } catch (e) {
    return 1;
  }
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

async function uploadAttachment(jobFolder, filename, arrayBuffer) {
  const siteId = await getAhSiteId();
  const path = `${jobFolder}/Quote/${filename}`;
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
    {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/octet-stream'
      },
      body: arrayBuffer
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Attachment upload failed: ${res.status} ${text}`);
  }
  return path;
}

// ----- Budget Excel write -----

async function writeQuoteToBudget(ref, savedAttachments, amount) {
  if (!ref.budgetRowNo) {
    return null;  // can't write without mapping
  }
  const siteId = await getAhSiteId();
  const m = ref.jobFolder.match(CONFIG.jobFolderPattern);
  const jobName = m ? m[2].trim() : '';
  const filename = `0 Budget Control ${jobName}.xlsx`;
  const buf = await readXlsx(siteId, `${ref.jobFolder}/Quote`, filename);
  const wb = XLSX.read(buf, { type: 'array', cellStyles: true, cellFormula: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref']);

  // Find Name N / Quote N columns
  const slots = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const headerCell = sheet[XLSX.utils.encode_cell({ c, r: 0 })];
    const v = headerCell ? String(headerCell.v || '').trim() : '';
    let m1 = v.match(/^Name\s*(\d+)$/i);
    if (m1) slots.push({ idx: parseInt(m1[1], 10), nameCol: c });
    let m2 = v.match(/^Quote\s*(\d+)$/i);
    if (m2) {
      const slot = slots.find(s => s.idx === parseInt(m2[1], 10));
      if (slot) slot.quoteCol = c;
    }
  }
  slots.sort((a, b) => a.idx - b.idx);

  // Find target row by No. column (col B = idx 1), strip nbsp + whitespace
  const targetNo = String(ref.budgetRowNo).replace(/[\s\xa0]+/g, '').trim();
  let targetRow = -1;
  for (let r = 1; r <= range.e.r; r++) {
    const noCell = sheet[XLSX.utils.encode_cell({ c: 1, r })];
    const noVal = noCell ? String(noCell.v || '').replace(/[\s\xa0]+/g, '').trim() : '';
    if (noVal === targetNo) { targetRow = r; break; }
  }
  if (targetRow < 0) {
    console.warn(`Budget row ${targetNo} not found in ${filename}`);
    return null;
  }

  // Find first empty Name N slot
  let chosenSlot = null;
  for (const s of slots) {
    if (!s.quoteCol) continue;
    const nameCell = sheet[XLSX.utils.encode_cell({ c: s.nameCol, r: targetRow })];
    if (!nameCell || nameCell.v == null || String(nameCell.v).trim() === '') { chosenSlot = s; break; }
  }
  if (!chosenSlot) {
    console.warn(`No empty Name slot on row ${targetNo} for ${ref.supplierCompany}`);
    return null;
  }

  // Write
  sheet[XLSX.utils.encode_cell({ c: chosenSlot.nameCol, r: targetRow })] = { t: 's', v: ref.supplierCompany || '' };
  if (amount != null) {
    sheet[XLSX.utils.encode_cell({ c: chosenSlot.quoteCol, r: targetRow })] = { t: 'n', v: amount };
  }
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const blob = new Uint8Array(out);
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(`${ref.jobFolder}/Quote/${filename}`)}:/content`,
    {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body: blob
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Budget write failed: ${res.status} ${text}`);
  }

  return {
    id: 'pr-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8),
    createdAt: new Date().toISOString(),
    supplierId: ref.supplierId,
    supplierCompany: ref.supplierCompany,
    supplierEmail: ref.supplierEmail,
    rfqId: ref.rfqId,
    rfqCategory: ref.rfqCategory,
    budgetRowNo: ref.budgetRowNo,
    budgetSlotIndex: chosenSlot.idx,
    amount,
    currency: 'AUD',
    pdfFilename: savedAttachments[0] ? savedAttachments[0].savedName : null,
    emailSubject: ref.subject,
    status: 'pending'
  };
}

// ----- Tracker update -----

async function appendReplyToTracker(ref, replyEntry, pendingReview) {
  const siteId = await getAhSiteId();
  const path = `${ref.jobFolder}/Quote`;
  const tracker = await readJson(siteId, path, 'rfq-tracker.json');
  if (!tracker) return;
  if (!Array.isArray(tracker.pendingReviews)) tracker.pendingReviews = [];
  if (pendingReview) tracker.pendingReviews.push(pendingReview);
  const rfq = (tracker.rfqs || []).find(r => r.id === ref.rfqId);
  if (rfq) {
    const sup = (rfq.suppliers || []).find(s => s.id === ref.supplierId);
    if (sup) {
      if (!Array.isArray(sup.replies)) sup.replies = [];
      sup.replies.push(replyEntry);
    }
  }
  await uploadJson(siteId, path, 'rfq-tracker.json', tracker);
}

// Lookup the address from the AH Site folder name (parses the folder pattern)
async function deriveAddress(jobFolderName) {
  const m = jobFolderName.match(CONFIG.jobFolderPattern);
  return m ? m[3].trim() : '';
}

// ----- Manual matching path -----

// List all open RFQs across all jobs (used by the Pending Match modal).
// Returns [{ key, label }] where key is encoded jobFolder|rfqId|supplierId.
export async function listAllOpenRfqs() {
  const index = await buildIndex();
  const out = [];
  for (const ref of index.byMessageId.values()) {
    out.push({
      key: `${ref.jobFolder}|${ref.rfqId}|${ref.supplierId}`,
      label: `${ref.jobCode} ${ref.jobName} · ${ref.rfqCategory} · ${ref.supplierCompany}`
    });
  }
  // Add entries for suppliers we don't have message ids for (older sends)
  // — these are scanned from trackers directly.
  // Already included above; dedupe by key.
  const seen = new Set();
  return out.filter(o => seen.has(o.key) ? false : (seen.add(o.key), true));
}

// User clicked an item in the Pending Match modal and picked an RFQ.
// We re-process the message as if Tier 1/2 had matched it.
export async function processNotificationManualMatch(messageId, key) {
  const [jobFolder, rfqId, supplierId] = key.split('|');
  // Load tracker and find the supplier
  const siteId = await getAhSiteId();
  const tracker = await readJson(siteId, `${jobFolder}/Quote`, 'rfq-tracker.json');
  if (!tracker) throw new Error('Tracker not found');
  const rfq = (tracker.rfqs || []).find(r => r.id === rfqId);
  if (!rfq) throw new Error('RFQ not found');
  const sup = (rfq.suppliers || []).find(s => s.id === supplierId);
  if (!sup) throw new Error('Supplier not found');
  const m = jobFolder.match(CONFIG.jobFolderPattern);
  const ref = {
    jobFolder, jobCode: m ? m[1].trim() : '', jobName: m ? m[2].trim() : '',
    rfqId, rfqCategory: rfq.category,
    supplierId: sup.id, supplierEmail: sup.email,
    supplierCompany: sup.companyName, supplierContact: sup.contactName,
    budgetRowNo: rfq.budgetRowNo
  };
  // Re-fetch message metadata since the notification only had a preview
  const msg = await getMessage(messageId);
  await fullyProcessMatched(msg, { tier: 3, ref });
  // Update the existing notification (addNotification dedupes by id)
}

// ----- Local helpers -----

function htmlToPlainText(html) {
  if (!html) return '';
  let s = String(html);
  s = s.replace(/<script[\s\S]*?<\/script>/gi, '');
  s = s.replace(/<style[\s\S]*?<\/style>/gi, '');
  s = s.replace(/<\/(p|div|li|tr|h[1-6])>/gi, '\n');
  s = s.replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = s.replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'");
  return s.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
}
