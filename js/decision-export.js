// js/decision-export.js
// Generates the markdown export consumed by the AH Est Email Classifier
// tuning chat. Plus a "copy the starter prompt" helper that the user
// pastes into the new chat alongside the file.

import { CONFIG } from './config.js';
import { readDecisionsSince, readCursor, writeCursor } from './decision-log.js';

// ---------- Markdown export ----------

// Build the markdown export text. Returns { markdown, filename, summary }
// where summary is { total, classifyConfirmed, classifyEdited, ... } so
// the caller can show the user what they're about to download.
export async function buildExport({ excludeTest = true } = {}) {
  const cursor = await readCursor();
  const since = cursor.lastExportAt;
  let entries = await readDecisionsSince(since, CONFIG.decisionExportMaxEntries + 1);
  // Cap at max + 1 so we know if we hit the limit
  const truncated = entries.length > CONFIG.decisionExportMaxEntries;
  if (truncated) entries = entries.slice(0, CONFIG.decisionExportMaxEntries);
  // Filter out test entries
  if (excludeTest) entries = entries.filter(e => !e.isTest);

  const summary = summarise(entries);
  const filename = makeFilename();
  const markdown = renderMarkdown({
    entries,
    summary,
    since,
    truncated,
    excludeTest
  });
  return { markdown, filename, summary, since, truncated, count: entries.length };
}

// Mark the export as taken — moves the cursor to "now".
export async function commitExport(asOfIso) {
  const cursor = await readCursor();
  cursor.exportHistory = cursor.exportHistory || [];
  cursor.exportHistory.push({
    exportAt: asOfIso || new Date().toISOString(),
    kind: 'export'
  });
  cursor.lastExportAt = asOfIso || new Date().toISOString();
  await writeCursor(cursor);
}

// ---------- Helpers ----------

function makeFilename() {
  const d = new Date();
  const ymd = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  return `gemini-decisions-${ymd}.md`;
}

function summarise(entries) {
  const s = {
    total: entries.length,
    byTask: {},
    classify: {
      confirmed: 0, edited: 0, rejected: 0, unreacted: 0,
      byClassification: {}
    },
    extractAmount: { confirmed: 0, edited: 0, rejected: 0, unreacted: 0 },
    summarizeFilename: { confirmed: 0, edited: 0, unreacted: 0 }
  };
  for (const e of entries) {
    s.byTask[e.task] = (s.byTask[e.task] || 0) + 1;
    if (e.task === 'classify') {
      const c = (e.output && e.output.classification) || '?';
      s.classify.byClassification[c] = (s.classify.byClassification[c] || 0) + 1;
      const r = e.userReaction || 'unreacted';
      if (s.classify[r] != null) s.classify[r]++;
    } else if (e.task === 'extractAmount') {
      const r = e.userReaction || 'unreacted';
      if (s.extractAmount[r] != null) s.extractAmount[r]++;
    } else if (e.task === 'summarizeFilename') {
      const r = e.userReaction || 'unreacted';
      if (s.summarizeFilename[r] != null) s.summarizeFilename[r]++;
    }
  }
  return s;
}

function renderMarkdown({ entries, summary, since, truncated, excludeTest }) {
  const lines = [];
  lines.push(`# Gemini Decisions Export`);
  lines.push('');
  lines.push(`Generated: ${new Date().toISOString()}`);
  lines.push(`Range: ${since ? '> ' + since : 'all time'} → now`);
  lines.push(`Entries included: **${summary.total}**${truncated ? ' (capped at limit — re-export to get more)' : ''}`);
  lines.push(`Test entries excluded: ${excludeTest ? 'yes' : 'no'}`);
  lines.push('');
  lines.push('## Summary');
  lines.push('');
  // Group by task with classification breakdown
  if (summary.byTask.classify) {
    lines.push(`**classify**: ${summary.byTask.classify}`);
    lines.push(`- confirmed: ${summary.classify.confirmed}, edited: ${summary.classify.edited}, rejected: ${summary.classify.rejected}, no reaction yet: ${summary.classify.unreacted}`);
    const byCls = Object.entries(summary.classify.byClassification).sort((a, b) => b[1] - a[1]);
    if (byCls.length) {
      lines.push('- by classification:');
      for (const [k, v] of byCls) lines.push(`  - ${k}: ${v}`);
    }
    lines.push('');
  }
  if (summary.byTask.extractAmount) {
    lines.push(`**extractAmount**: ${summary.byTask.extractAmount}`);
    lines.push(`- confirmed: ${summary.extractAmount.confirmed}, edited: ${summary.extractAmount.edited}, rejected: ${summary.extractAmount.rejected}, no reaction yet: ${summary.extractAmount.unreacted}`);
    lines.push('');
  }
  if (summary.byTask.summarizeFilename) {
    lines.push(`**summarizeFilename**: ${summary.byTask.summarizeFilename}`);
    lines.push(`- confirmed: ${summary.summarizeFilename.confirmed}, edited: ${summary.summarizeFilename.edited}, no reaction yet: ${summary.summarizeFilename.unreacted}`);
    lines.push('');
  }

  // Group entries by task → reaction for readable analysis
  lines.push('## Entries');
  lines.push('');

  const buckets = bucketEntries(entries);
  for (const [bucketKey, bucketEntries] of buckets) {
    if (bucketEntries.length === 0) continue;
    lines.push(`### ${bucketKey}  (${bucketEntries.length})`);
    lines.push('');
    for (const e of bucketEntries) {
      lines.push(formatEntry(e));
      lines.push('');
    }
  }

  return lines.join('\n');
}

function bucketEntries(entries) {
  // Order matters — most useful for tuning analysis is "edits" and
  // "rejects" since those are where Gemini was wrong.
  const order = [
    ['classify · edited',      e => e.task === 'classify' && e.userReaction === 'edited'],
    ['classify · rejected',    e => e.task === 'classify' && e.userReaction === 'rejected'],
    ['classify · confirmed',   e => e.task === 'classify' && e.userReaction === 'confirmed'],
    ['classify · no reaction', e => e.task === 'classify' && !e.userReaction],
    ['extractAmount · edited',      e => e.task === 'extractAmount' && e.userReaction === 'edited'],
    ['extractAmount · rejected',    e => e.task === 'extractAmount' && e.userReaction === 'rejected'],
    ['extractAmount · confirmed',   e => e.task === 'extractAmount' && e.userReaction === 'confirmed'],
    ['extractAmount · no reaction', e => e.task === 'extractAmount' && !e.userReaction],
    ['summarizeFilename · edited',      e => e.task === 'summarizeFilename' && e.userReaction === 'edited'],
    ['summarizeFilename · confirmed',   e => e.task === 'summarizeFilename' && e.userReaction === 'confirmed'],
    ['summarizeFilename · no reaction', e => e.task === 'summarizeFilename' && !e.userReaction]
  ];
  const out = new Map();
  for (const [key] of order) out.set(key, []);
  for (const e of entries) {
    for (const [key, pred] of order) {
      if (pred(e)) { out.get(key).push(e); break; }
    }
  }
  return out;
}

function formatEntry(e) {
  const lines = [];
  lines.push(`#### ${e.id}`);
  lines.push(`- **createdAt**: ${e.createdAt}`);
  if (e.jobFolder) lines.push(`- **job**: ${e.jobFolder}`);
  if (e.userReaction) {
    lines.push(`- **userReaction**: \`${e.userReaction}\` at ${e.reactedAt}`);
    if (e.userCorrection) lines.push(`- **userCorrection**: \`${JSON.stringify(e.userCorrection)}\``);
    if (e.userWhy) lines.push(`- **userWhy**: ${e.userWhy}`);
  }
  lines.push('');
  lines.push('**Input:**');
  lines.push('```json');
  lines.push(JSON.stringify(e.input, null, 2));
  lines.push('```');
  lines.push('');
  lines.push('**Gemini output:**');
  lines.push('```json');
  lines.push(JSON.stringify(e.output, null, 2));
  lines.push('```');
  return lines.join('\n');
}

// ---------- Browser-side download trigger ----------

export function downloadAsFile(filename, text) {
  const blob = new Blob([text], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ---------- The starter prompt for the AH Est Email Classifier chat ----------
// Static text. User clicks "Copy Starter Prompt" → paste into the
// classifier chat alongside the markdown file attachment.

export const STARTER_PROMPT_FOR_CLASSIFIER_CHAT =
`This is the AH Est Email Classifier chat (per its own pinned system prompt).

I'm sending you an export of real Gemini decisions and my reactions to them.
Filename will be \`gemini-decisions-YYYY-MM-DD.md\` — attached/pasted alongside this message.

Workflow you should follow:

1. Analyse the export. Look for clusters where Gemini was systematically wrong:
   - "edited" entries with consistent direction (e.g. classify said Question, user changed to Quote)
   - "rejected" entries (Gemini hallucinated the wrong amount, supplier didn't actually quote)
   - confidence patterns (low confidence + correct vs high confidence + wrong)
2. Tell me your findings in plain language BEFORE producing code.
3. Wait for me to confirm.
4. Then produce a full updated worker.js as a zip named
   \`ah-est-email-classifier-tune-[short-summary].zip\` ready for me to drop
   into the GitHub repo on the main branch.
5. Tell me what changed in the prompt(s) so I have a record.

Don't change the Worker's contract (three tasks, six classifications). Don't add env variables. Stay on Gemini 2.5 Flash unless I approve a switch.

If the export doesn't show enough signal to confidently improve the prompt, say so plainly. A bad prompt edit is worse than no edit.

— Export below —`;
