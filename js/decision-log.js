// js/decision-log.js
// Logs every Gemini classification/extraction decision and every subsequent
// user reaction (Confirmed / Edited / Rejected) so we can later export the
// data to a separate "AH Est Email Classifier" tuning chat.
//
// Storage: SharePoint AAA Quote Common Docs / _app /
//   classifier-decisions-YYYY-MM.json   (rotates monthly to keep size bounded)
//   classifier-decisions-cursor.json    (last-export pointer + reset history)
//
// Decision entry shape (one record per Gemini call):
//   {
//     id,                  // unique
//     createdAt,           // ISO when decision was made
//     task,                // 'classify' | 'extractAmount' | 'summarizeFilename'
//     input: {...},        // exact payload sent to Gemini
//     output: {...},       // exact result Gemini returned
//     // Tuning context (filled in once user reacts):
//     userReaction: null | 'confirmed' | 'edited' | 'rejected',
//     userCorrection: null | {...},   // for 'edited': what they changed it to
//     userWhy: null | 'free text',    // optional why they corrected
//     reactedAt: null | ISO,
//     // Misc:
//     isTest: false,                  // user can flag entries as test data
//     // Linkage:
//     messageId: null,                // Graph message id (if applicable)
//     jobFolder: null,                // matched job
//     rfqId: null, supplierId: null,
//     // Auto-classification source notification id (so reactions can be
//     // matched back to the originating decision)
//     notificationId: null
//   }

import { CONFIG } from './config.js';
import { state } from './state.js';
import { getAhSiteId, encodeUriPath, readJson, uploadJson, graphFetch } from './graph.js';

// In-memory queue used while a poll cycle is running. Flushed at end.
const pendingAppends = [];
let flushScheduled = false;

// ---------- Public API ----------

// Record a new Gemini decision. Returns the entry's id (used to link the
// user-reaction record later).
export function logDecision({ task, input, output, messageId, jobFolder, rfqId, supplierId, notificationId, isTest }) {
  const id = 'd-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8);
  const entry = {
    id,
    createdAt: new Date().toISOString(),
    task,
    input: redactInput(input),
    output,
    userReaction: null,
    userCorrection: null,
    userWhy: null,
    reactedAt: null,
    isTest: !!isTest,
    messageId: messageId || null,
    jobFolder: jobFolder || null,
    rfqId: rfqId || null,
    supplierId: supplierId || null,
    notificationId: notificationId || null
  };
  pendingAppends.push(entry);
  scheduleFlush();
  return id;
}

// Mark a previously-logged decision with the user's reaction. The decision
// id is whatever logDecision() returned. Re-fetches the month's log file,
// patches the matching entry, and writes back.
//
// reaction: 'confirmed' | 'edited' | 'rejected'
// correction: optional (for 'edited' — what the user changed it to)
// why: optional free-text reason
export async function recordReaction(decisionId, reaction, correction, why) {
  if (!decisionId) return;
  // Find which monthly file holds this id. Easiest: look at the entry's
  // createdAt — but caller doesn't pass that. Walk current and previous
  // month's logs (decisions almost always reacted to within ~weeks).
  const months = monthsToTry();
  for (const ym of months) {
    const filename = monthlyFilename(ym);
    const log = await readMonthlyLog(filename);
    if (!log) continue;
    const idx = (log.entries || []).findIndex(e => e.id === decisionId);
    if (idx >= 0) {
      log.entries[idx].userReaction = reaction;
      if (correction !== undefined) log.entries[idx].userCorrection = correction;
      if (why !== undefined) log.entries[idx].userWhy = why || null;
      log.entries[idx].reactedAt = new Date().toISOString();
      await writeMonthlyLog(filename, log);
      return true;
    }
  }
  // If the decision is still queued (not flushed yet), patch in-memory:
  const inFlight = pendingAppends.find(e => e.id === decisionId);
  if (inFlight) {
    inFlight.userReaction = reaction;
    if (correction !== undefined) inFlight.userCorrection = correction;
    if (why !== undefined) inFlight.userWhy = why || null;
    inFlight.reactedAt = new Date().toISOString();
    return true;
  }
  console.warn('recordReaction: decision id not found:', decisionId);
  return false;
}

// Mark a decision as test data (so it's excluded from exports). Caller
// passes the decision id; we patch in-place.
export async function markDecisionAsTest(decisionId, isTest) {
  if (!decisionId) return;
  const months = monthsToTry();
  for (const ym of months) {
    const filename = monthlyFilename(ym);
    const log = await readMonthlyLog(filename);
    if (!log) continue;
    const idx = (log.entries || []).findIndex(e => e.id === decisionId);
    if (idx >= 0) {
      log.entries[idx].isTest = !!isTest;
      await writeMonthlyLog(filename, log);
      return true;
    }
  }
  const inFlight = pendingAppends.find(e => e.id === decisionId);
  if (inFlight) { inFlight.isTest = !!isTest; return true; }
  return false;
}

// Read the cursor file (or create a default).
export async function readCursor() {
  const siteId = await getAhSiteId();
  const cursor = await readJson(siteId, CONFIG.decisionLogPath, CONFIG.decisionCursorFilename);
  return cursor || { lastExportAt: null, exportHistory: [] };
}

// Write the cursor file (after a successful export).
export async function writeCursor(cursor) {
  const siteId = await getAhSiteId();
  await uploadJson(siteId, CONFIG.decisionLogPath, CONFIG.decisionCursorFilename, cursor);
}

// Reset the cursor (allows re-exporting from a chosen point).
export async function resetCursorTo(iso) {
  const cursor = await readCursor();
  cursor.exportHistory = (cursor.exportHistory || []);
  cursor.exportHistory.push({ resetTo: iso, resetAt: new Date().toISOString(), kind: 'reset' });
  cursor.lastExportAt = iso;
  await writeCursor(cursor);
  return cursor;
}

// Read all decision entries since `sinceIso` (exclusive). Walks however
// many monthly files are needed. Returns sorted oldest-first.
export async function readDecisionsSince(sinceIso, maxEntries) {
  const since = sinceIso ? new Date(sinceIso).getTime() : 0;
  const out = [];
  // Determine which months to read: from the month containing sinceIso
  // through the current month.
  const months = monthsBetween(sinceIso, new Date().toISOString());
  for (const ym of months) {
    const filename = monthlyFilename(ym);
    const log = await readMonthlyLog(filename);
    if (!log || !Array.isArray(log.entries)) continue;
    for (const e of log.entries) {
      if (new Date(e.createdAt).getTime() > since) out.push(e);
      if (maxEntries && out.length >= maxEntries) break;
    }
    if (maxEntries && out.length >= maxEntries) break;
  }
  out.sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  return out;
}

// ---------- Internal: monthly file rotation ----------

function ymOf(iso) {
  const d = iso ? new Date(iso) : new Date();
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

function monthlyFilename(ym) {
  return `${CONFIG.decisionLogPrefix}${ym}.json`;
}

function monthsToTry() {
  // Current and previous 2 months — covers most reactions.
  const out = [];
  const now = new Date();
  for (let i = 0; i < 3; i++) {
    const d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - i, 1));
    out.push(`${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`);
  }
  return out;
}

function monthsBetween(startIso, endIso) {
  const out = [];
  const start = startIso ? new Date(startIso) : new Date(0);
  const end = endIso ? new Date(endIso) : new Date();
  // Move start to first of its month
  const cur = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), 1));
  const stop = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), 1));
  while (cur <= stop) {
    out.push(`${cur.getUTCFullYear()}-${String(cur.getUTCMonth() + 1).padStart(2, '0')}`);
    cur.setUTCMonth(cur.getUTCMonth() + 1);
  }
  return out;
}

async function readMonthlyLog(filename) {
  const siteId = await getAhSiteId();
  return readJson(siteId, CONFIG.decisionLogPath, filename);
}

async function writeMonthlyLog(filename, log) {
  const siteId = await getAhSiteId();
  await uploadJson(siteId, CONFIG.decisionLogPath, filename, log);
}

// ---------- Internal: batched flush ----------

function scheduleFlush() {
  if (flushScheduled) return;
  flushScheduled = true;
  // Flush after 1.5s of no further appends — enough to batch a poll cycle.
  setTimeout(() => { flushPending().catch(err => console.warn('Decision log flush failed:', err)); }, 1500);
}

async function flushPending() {
  flushScheduled = false;
  if (pendingAppends.length === 0) return;
  // Group by month so we minimise reads/writes
  const grouped = new Map();  // ym -> entries[]
  for (const e of pendingAppends) {
    const ym = ymOf(e.createdAt);
    if (!grouped.has(ym)) grouped.set(ym, []);
    grouped.get(ym).push(e);
  }
  pendingAppends.length = 0;
  for (const [ym, entries] of grouped.entries()) {
    const filename = monthlyFilename(ym);
    let log = await readMonthlyLog(filename);
    if (!log) log = { version: 1, ym, entries: [] };
    log.entries.push(...entries);
    try { await writeMonthlyLog(filename, log); }
    catch (err) {
      console.error(`Decision log write failed for ${filename}:`, err);
      // Re-queue — caller should retry on next flush
      pendingAppends.push(...entries);
    }
  }
}

// ---------- Internal: input redaction ----------
// Don't store the entire body verbatim — keep first ~2000 chars so the
// export is human-readable but not absurdly large. Strip nothing else;
// classifier-relevant fields must be preserved.
function redactInput(input) {
  if (!input || typeof input !== 'object') return input;
  const out = { ...input };
  for (const k of Object.keys(out)) {
    if (typeof out[k] === 'string' && out[k].length > 2000) {
      out[k] = out[k].slice(0, 2000) + '\n…[truncated]';
    }
  }
  return out;
}
