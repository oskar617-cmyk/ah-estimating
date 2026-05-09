// js/jobs.js
// Jobs list, job detail screen, and the existing-quotes migration flow.

import { CONFIG } from './config.js';
import { state } from './state.js';
import { navigate, setOnReturnToJobs } from './nav.js';
import { graphFetch, getAhSiteId, encodeUriPath, readJson, uploadJson } from './graph.js';
import { logAudit, readTracker, writeTracker } from './audit.js';
import { showToast, escapeHtml, confirmModal } from './ui.js';
import { renderPendingReviewSection, attachPendingReviewHandlers } from './pending-review.js';

export async function loadJobs() {
  const container = document.getElementById('jobs-content');
  container.innerHTML = '<div class="loading"><div class="spinner"></div><div>Loading Jobs...</div></div>';
  try {
    const siteId = await getAhSiteId();
    const result = await graphFetch(`/sites/${siteId}/drive/root/children?$top=200&$select=id,name,folder,webUrl`);
    const items = result.value || [];
    const jobs = items
      .filter(it => it.folder && CONFIG.jobFolderPattern.test(it.name))
      .map(it => {
        const m = it.name.match(CONFIG.jobFolderPattern);
        return { id: it.id, folderName: it.name, jobCode: m[1].trim(), jobName: m[2].trim(), address: m[3].trim(), webUrl: it.webUrl };
      })
      .sort((a, b) => b.jobCode.localeCompare(a.jobCode));
    renderJobs(jobs);
  } catch (err) {
    console.error('Load jobs error:', err);
    container.innerHTML = `<div class="empty-state"><div style="color: var(--red); margin-bottom: 8px;">Failed To Load Jobs</div><div class="text-small">${escapeHtml(err.message)}</div><button class="btn-secondary mt-16" onclick="loadJobs()">Try Again</button></div>`;
  }
}

function renderJobs(jobs) {
  const container = document.getElementById('jobs-content');
  if (jobs.length === 0) {
    container.innerHTML = `<div class="empty-state"><div>No Jobs Found</div><div class="text-small mt-8">Tap the + button to create a new job.</div></div>`;
    return;
  }
  const html = jobs.map(j => `
    <div class="job-card" data-id="${escapeHtml(j.id)}">
      <div class="job-code">${escapeHtml(j.jobCode)} ${escapeHtml(j.jobName)}</div>
      <div class="job-address">${escapeHtml(j.address)}</div>
    </div>
  `).join('');
  container.innerHTML = `<div class="jobs-list">${html}</div>`;
  container.querySelectorAll('.job-card').forEach((card, i) => card.addEventListener('click', () => openJob(jobs[i])));
}

function openJob(job) {
  state.currentJob = job;
  navigate('job-detail-screen', { jobId: job.id });
  loadJobDetail();
}

export function openCurrentJobInSharePoint() {
  if (state.currentJob && state.currentJob.webUrl) window.open(state.currentJob.webUrl, '_blank');
}

// Display "Oskar, Est" instead of full emails
function formatTeamEmails(emails) {
  if (!emails || !emails.length) return 'Not Set';
  return emails.map(e => {
    const local = (e.split('@')[0] || '').toLowerCase();
    return local.charAt(0).toUpperCase() + local.slice(1);
  }).join(', ');
}

export async function loadJobDetail() {
  const container = document.getElementById('jd-content');
  document.getElementById('jd-title').textContent = `${state.currentJob.jobCode} ${state.currentJob.jobName}`;
  container.innerHTML = '<div class="loading"><div class="spinner"></div><div>Loading Job...</div></div>';
  try {
    const siteId = await getAhSiteId();
    const quotePath = `${state.currentJob.folderName}/Quote`;
    let tracker = await readTracker(state.currentJob.folderName);
    if (!tracker) {
      const quoteFolder = await graphFetch(
        `/sites/${siteId}/drive/root:/${encodeUriPath(quotePath)}:/children?$top=200&$select=id,name,file`
      ).catch(err => err.status === 404 ? { value: [] } : Promise.reject(err));
      const existingPdfs = (quoteFolder.value || []).filter(it => it.file && /\.pdf$/i.test(it.name));
      if (existingPdfs.length > 0) { renderMigrationPrompt(existingPdfs.length); return; }
      else { tracker = createEmptyTracker(); await writeTracker(state.currentJob.folderName, tracker); }
    }
    renderJobDetail(tracker);
  } catch (err) {
    console.error('Load job detail error:', err);
    container.innerHTML = `<div class="empty-state"><div style="color: var(--red); margin-bottom: 8px;">Failed To Load Job</div><div class="text-small">${escapeHtml(err.message)}</div><button class="btn-secondary mt-16" onclick="loadJobDetail()">Try Again</button></div>`;
  }
}

function createEmptyTracker() {
  return {
    version: 1,
    jobCode: state.currentJob.jobCode,
    jobName: state.currentJob.jobName,
    address: state.currentJob.address,
    projectTeamEmails: [],
    rfqs: [],
    createdAt: new Date().toISOString(),
    createdBy: state.currentUserEmail
  };
}

function renderMigrationPrompt(count) {
  const container = document.getElementById('jd-content');
  container.innerHTML = `
    <div class="info-card">
      <div style="font-size: 17px; font-weight: 600; margin-bottom: 8px;">Existing Quotes Detected</div>
      <div class="text-muted text-small" style="margin-bottom: 16px;">
        This job's <strong>Quote</strong> folder already has <strong>${count}</strong> PDF file${count === 1 ? '' : 's'}, but no tracker has been set up yet. What would you like to do?
      </div>
      <div class="btn-row">
        <button class="btn-primary" onclick="migrateExistingQuotes()">Migrate Existing Files</button>
        <button class="btn-secondary" onclick="skipMigration()">Skip - Start Fresh</button>
      </div>
      <div class="text-muted text-small mt-16">
        <strong>Migrate</strong>: scans filenames for trade, company, version and amount, creates tracker entries you can review later.<br>
        <strong>Skip</strong>: leaves existing files alone; the tracker tracks only NEW RFQs sent through this app.
      </div>
    </div>`;
}

export async function skipMigration() {
  try {
    const siteId = await getAhSiteId();
    const quotePath = `${state.currentJob.folderName}/Quote`;
    const tracker = createEmptyTracker(); tracker.migrationSkipped = true;
    await writeTracker(state.currentJob.folderName, tracker);
    await logAudit('JOB_TRACKER_INITIALISED', state.currentJob.folderName, { migrated: false });
    renderJobDetail(tracker);
    showToast('Tracker Initialised', 'success');
  } catch (err) { console.error(err); showToast('Failed To Initialise Tracker', 'error'); }
}

export async function migrateExistingQuotes() {
  const container = document.getElementById('jd-content');
  container.innerHTML = '<div class="loading"><div class="spinner"></div><div>Scanning Existing Quotes...</div></div>';
  try {
    const siteId = await getAhSiteId();
    const quotePath = `${state.currentJob.folderName}/Quote`;
    const result = await graphFetch(`/sites/${siteId}/drive/root:/${encodeUriPath(quotePath)}:/children?$top=500&$select=id,name,file,webUrl,size,lastModifiedDateTime`);
    const pdfs = (result.value || []).filter(it => it.file && /\.pdf$/i.test(it.name));
    const tracker = createEmptyTracker();
    tracker.migrated = true;
    const parsed = []; const needsReview = [];
    for (const pdf of pdfs) {
      const stem = pdf.name.replace(/\.pdf$/i, '');
      const m = stem.match(/^(.+?)\s*-\s*(.+?)\s+v(\d+)(?:\s*-\s*([\d,]+(?:\.\d+)?))?$/i);
      if (m) {
        parsed.push({ fileName: pdf.name, fileId: pdf.id, webUrl: pdf.webUrl, trade: m[1].trim(), company: m[2].trim(), version: parseInt(m[3], 10), amount: m[4] ? parseFloat(m[4].replace(/,/g, '')) : null, receivedAt: pdf.lastModifiedDateTime });
      } else {
        needsReview.push({ fileName: pdf.name, fileId: pdf.id, webUrl: pdf.webUrl, receivedAt: pdf.lastModifiedDateTime });
      }
    }
    tracker.migratedQuotes = parsed; tracker.needsReview = needsReview;
    await writeTracker(state.currentJob.folderName, tracker);
    await logAudit('JOB_TRACKER_INITIALISED', state.currentJob.folderName, { migrated: true, parsedCount: parsed.length, needsReviewCount: needsReview.length });
    renderJobDetail(tracker);
    showToast(`Migrated ${parsed.length} Quote${parsed.length === 1 ? '' : 's'}`, 'success');
  } catch (err) {
    console.error(err); showToast('Migration Failed', 'error');
    container.innerHTML = `<div class="empty-state"><div style="color: var(--red); margin-bottom: 8px;">Migration Failed</div><div class="text-small">${escapeHtml(err.message)}</div><button class="btn-secondary mt-16" onclick="loadJobDetail()">Back</button></div>`;
  }
}

function renderJobDetail(tracker) {
  const container = document.getElementById('jd-content');
  const teamDisplay = formatTeamEmails(tracker.projectTeamEmails);
  const rfqs = Array.isArray(tracker.rfqs) ? tracker.rfqs : [];
  const migratedCount = (tracker.migratedQuotes || []).length;
  const needsReviewCount = (tracker.needsReview || []).length;

  // Group RFQs by trade category. Same trade may have multiple RFQ batches
  // (e.g. re-sent later). We display each batch under its trade group.
  const groups = new Map();
  for (const r of rfqs) {
    if (!groups.has(r.category)) groups.set(r.category, []);
    groups.get(r.category).push(r);
  }

  const senderAllowed = state.currentUserEmail === 'est@auhs.com.au';

  let migratedSection = '';
  if (migratedCount > 0 || needsReviewCount > 0) {
    const items = [];
    if (migratedCount > 0) items.push(`<div class="text-small">${migratedCount} Parsed Quote${migratedCount === 1 ? '' : 's'} From Existing Files</div>`);
    if (needsReviewCount > 0) items.push(`<div class="text-small mt-4" style="color: var(--amber);">${needsReviewCount} File${needsReviewCount === 1 ? ' Needs' : 's Need'} Manual Review</div>`);
    migratedSection = `<div class="section-title">Migrated From Existing Files</div><div class="info-card">${items.join('')}</div>`;
  }

  let rfqsHtml;
  if (rfqs.length === 0) {
    rfqsHtml = '<div class="info-card"><div class="text-muted text-small">No RFQs Sent Yet</div></div>';
  } else {
    const groupKeys = Array.from(groups.keys()).sort();
    rfqsHtml = groupKeys.map(category => {
      const batches = groups.get(category);
      const counts = aggregateTradeCounts(batches);
      const rollupHtml = renderTradeRollup(counts);
      const batchHtml = batches
        .slice()
        .sort((a, b) => (b.sentAt || '').localeCompare(a.sentAt || ''))
        .map(b => renderRfqBatchActivity(b, senderAllowed))
        .join('');
      return `
        <div class="rfq-group">
          <div class="rfq-group-header">
            <div class="rfq-group-title">${escapeHtml(category)}</div>
            ${rollupHtml}
          </div>
          <div class="rfq-group-body">${batchHtml}</div>
        </div>`;
    }).join('');
  }

  const sendBtn = senderAllowed
    ? `<button class="btn-primary mt-16" onclick="openSendRfq()">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
        Send New RFQ
      </button>`
    : `<button class="btn-primary mt-16" disabled title="Only est@auhs.com.au can send RFQs">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
        Send New RFQ (Est Only)
      </button>`;

  const pendingReviewHtml = renderPendingReviewSection(tracker);

  container.innerHTML = `
    <div class="job-header">
      <div class="job-header-code">${escapeHtml(state.currentJob.jobCode)} ${escapeHtml(state.currentJob.jobName)}</div>
      <div class="job-header-address">${escapeHtml(state.currentJob.address)}</div>
      <div class="job-header-meta">
        <div><strong>Project Team:</strong> ${escapeHtml(teamDisplay)}</div>
      </div>
    </div>
    ${pendingReviewHtml}
    <div class="section-title">RFQs <span class="text-muted text-small" style="font-weight: normal; text-transform: none; letter-spacing: 0;">(${rfqs.length} Total)</span></div>
    ${rfqsHtml}
    ${sendBtn}
    ${migratedSection}
    <div class="section-title">Folder</div>
    <div class="info-card">
      <a href="${escapeHtml(state.currentJob.webUrl)}" target="_blank" style="color: var(--blue); text-decoration: none; font-size: 14px;">Open ${escapeHtml(state.currentJob.folderName)} In SharePoint</a>
    </div>
  `;
  // Wire up Pending Review action buttons (Confirm/Edit/Reject)
  if (pendingReviewHtml) {
    attachPendingReviewHandlers(tracker, state.currentJob.folderName, () => loadJobDetail());
  }
}

// Aggregate counts across all batches in a trade. Used at the trade-level
// rollup at the top of each trade group.
//   { rfqCount, suppliers, sent, replied, quoted, questioned, suspicious,
//     givenUp, picked }
function aggregateTradeCounts(batches) {
  const c = {
    rfqCount: batches.length,
    suppliers: 0, sent: 0, replied: 0, quoted: 0,
    questioned: 0, suspicious: 0, givenUp: 0, picked: 0
  };
  for (const b of batches) {
    if (b.status === 'picked') c.picked++;
    if (b.status === 'given_up') c.givenUp++;
    for (const s of b.suppliers || []) {
      c.suppliers++;
      const replies = s.replies || [];
      if (replies.length > 0) c.replied++;
      if (replies.some(r => r.classification === 'Quote')) c.quoted++;
      if (replies.some(r => r.classification === 'Question')) c.questioned++;
      if (replies.some(r => r.classification === 'Suspicious')) c.suspicious++;
    }
    // Suppliers with no replies and not given-up count as 'sent'
    for (const s of b.suppliers || []) {
      if (!(s.replies && s.replies.length > 0) && s.status !== 'given_up') c.sent++;
    }
  }
  return c;
}

// Render the trade-level rollup as plain-language counts.
// Example: "3 RFQs · 6 sent · 1 quoted · 1 question"
function renderTradeRollup(counts) {
  const parts = [];
  parts.push(`${counts.rfqCount} RFQ${counts.rfqCount === 1 ? '' : 's'}`);
  if (counts.sent > 0)        parts.push(`<span class="rollup-pill rollup-sent">${counts.sent} sent</span>`);
  if (counts.quoted > 0)      parts.push(`<span class="rollup-pill rollup-quoted">${counts.quoted} quoted</span>`);
  if (counts.questioned > 0)  parts.push(`<span class="rollup-pill rollup-question">${counts.questioned} question</span>`);
  if (counts.suspicious > 0)  parts.push(`<span class="rollup-pill rollup-suspicious">${counts.suspicious} suspicious</span>`);
  if (counts.givenUp > 0)     parts.push(`<span class="rollup-pill rollup-given-up">${counts.givenUp} old</span>`);
  if (counts.picked > 0)      parts.push(`<span class="rollup-pill rollup-picked">${counts.picked} picked</span>`);
  return `<div class="trade-rollup">${parts.join('<span class="rollup-sep">·</span>')}</div>`;
}

// Per-RFQ mixed status: counts each supplier within the RFQ.
function computeRfqMixedStatus(batch) {
  const out = { sent: 0, replied: 0, quoted: 0, questioned: 0, suspicious: 0, givenUp: 0 };
  const suppliers = batch.suppliers || [];
  for (const s of suppliers) {
    const replies = s.replies || [];
    if (s.status === 'given_up') { out.givenUp++; continue; }
    if (replies.some(r => r.classification === 'Quote')) { out.quoted++; continue; }
    if (replies.some(r => r.classification === 'Suspicious')) { out.suspicious++; continue; }
    if (replies.some(r => r.classification === 'Question')) { out.questioned++; continue; }
    if (replies.length > 0) { out.replied++; continue; }
    out.sent++;
  }
  out.total = suppliers.length;
  return out;
}

function renderRfqMixedBadge(s) {
  const parts = [];
  if (s.quoted > 0)      parts.push(`<span class="badge badge-quoted">🟢 ${s.quoted} Quoted</span>`);
  if (s.questioned > 0)  parts.push(`<span class="badge badge-question">❓ ${s.questioned} Question</span>`);
  if (s.suspicious > 0)  parts.push(`<span class="badge badge-suspicious">⚠ ${s.suspicious} Suspicious</span>`);
  if (s.replied > 0)     parts.push(`<span class="badge badge-replied">🔵 ${s.replied} Replied</span>`);
  if (s.sent > 0)        parts.push(`<span class="badge badge-sent">🟡 ${s.sent} Sent</span>`);
  if (s.givenUp > 0)     parts.push(`<span class="badge badge-given-up">⚫ ${s.givenUp} Old</span>`);
  if (parts.length === 0) parts.push('<span class="badge badge-not-sent">⚪ Empty</span>');
  return `<div class="rfq-mixed-badges">${parts.join('')}</div>`;
}

// Aggregate the highest-priority status across all batches under one trade.
// Priority (high to low): picked > suspicious > question > quoted > replied > sent > given_up > none.
function aggregateTradeStatus(batches) {
  let totalSuppliers = 0, replied = 0, quoted = 0;
  let hasPicked = false, hasSuspicious = false, hasQuestion = false;
  let allGivenUp = batches.length > 0;
  for (const b of batches) {
    if (b.status === 'picked') hasPicked = true;
    if (b.status !== 'given_up') allGivenUp = false;
    for (const s of b.suppliers || []) {
      totalSuppliers++;
      const replies = s.replies || [];
      if (replies.length > 0) replied++;
      // Quote priority: any reply classified 'quote' counts
      if (replies.some(r => r.classification === 'Quote')) quoted++;
      if (replies.some(r => r.classification === 'Suspicious')) hasSuspicious = true;
      if (replies.some(r => r.classification === 'Question')) hasQuestion = true;
    }
  }
  if (hasPicked) return { kind: 'picked' };
  if (hasSuspicious) return { kind: 'suspicious' };
  if (hasQuestion && quoted === 0) return { kind: 'question' };
  if (quoted > 0) return { kind: 'quoted', n: quoted };
  if (replied > 0) return { kind: 'replied', n: replied, m: totalSuppliers };
  if (allGivenUp) return { kind: 'given-up' };
  if (totalSuppliers > 0) return { kind: 'sent', n: totalSuppliers };
  return { kind: 'not-sent' };
}

function renderStatusBadge(status) {
  switch (status.kind) {
    case 'not-sent':   return '<span class="badge badge-not-sent">⚪ Not Sent</span>';
    case 'sent':       return `<span class="badge badge-sent">🟡 ${status.n} Sent</span>`;
    case 'replied':    return `<span class="badge badge-replied">🔵 ${status.n}/${status.m} Replied</span>`;
    case 'quoted':     return `<span class="badge badge-quoted">🟢 ${status.n} Quoted</span>`;
    case 'question':   return '<span class="badge badge-question">❓ Question</span>';
    case 'suspicious': return '<span class="badge badge-suspicious">⚠ Suspicious</span>';
    case 'picked':     return '<span class="badge badge-selected">✅ Trade Selected</span>';
    case 'given-up':   return '<span class="badge badge-given-up">⚫ Given Up</span>';
    default:           return '';
  }
}

// Render one RFQ batch as a plain-language activity entry.
function renderRfqBatchActivity(batch, isAdmin) {
  const senderLocal = (batch.sentBy || '').split('@')[0] || '?';
  const sentDate = formatDateTimeShort(batch.sentAt);
  const supplierNames = (batch.suppliers || []).map(s => escapeHtml(s.companyName)).join(', ');
  const supplierCount = (batch.suppliers || []).length;
  const mixedStatus = computeRfqMixedStatus(batch);
  const mixedBadgeHtml = renderRfqMixedBadge(mixedStatus);
  const deleteBtn = isAdmin
    ? `<button class="rfq-delete-btn" title="Delete RFQ" onclick="window.deleteRfq('${escapeHtml(batch.id)}', event)">
         <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="14" height="14">
           <polyline points="3 6 5 6 21 6"/>
           <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
           <path d="M10 11v6"/><path d="M14 11v6"/>
           <path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
         </svg>
       </button>`
    : '';
  const lines = [];
  lines.push(`
    <div class="activity-item">
      <div class="activity-icon">✉</div>
      <div class="activity-text">
        <div class="activity-title-row">
          <div><strong>${escapeHtml(senderLocal)}</strong> sent RFQ to <strong>${supplierNames}</strong></div>
          ${mixedBadgeHtml}
        </div>
        <div class="activity-meta">${escapeHtml(sentDate)} · Reply by ${escapeHtml(batch.respondByDate || '?')}${batch.budgetRowNo ? ' · Maps to row ' + escapeHtml(batch.budgetRowNo) : ''}${batch.sowAttached ? '' : ' · <span style="color:var(--amber);">No SOW attached</span>'}</div>
      </div>
      ${deleteBtn}
    </div>
  `);
  // Each supplier's replies (none yet for v1; placeholder for Phase 4d inbox)
  for (const s of batch.suppliers || []) {
    const replies = s.replies || [];
    for (const r of replies) {
      lines.push(renderReplyActivity(s, r));
    }
    if (s.followupCount > 0) {
      lines.push(`
        <div class="activity-item activity-system">
          <div class="activity-icon">↩</div>
          <div class="activity-text">
            <div>Sent ${s.followupCount} follow-up${s.followupCount === 1 ? '' : 's'} to ${escapeHtml(s.companyName)}</div>
            <div class="activity-meta">Latest: ${escapeHtml(formatDateTimeShort(s.lastFollowupAt))}</div>
          </div>
        </div>
      `);
    }
  }
  return `<div class="rfq-batch">${lines.join('')}</div>`;
}

function renderReplyActivity(supplier, reply) {
  const icon = reply.classification === 'quote' ? '📥'
    : reply.classification === 'question' ? '❓'
    : reply.classification === 'suspicious' ? '⚠'
    : reply.classification === 'decline' ? '❌'
    : '↻';
  let line;
  if (reply.classification === 'quote') {
    const amount = reply.amount != null ? `$${reply.amount.toLocaleString()}` : '(amount tbc)';
    line = `<strong>${escapeHtml(supplier.companyName)}</strong> replied with quote ${escapeHtml(amount)}`;
  } else if (reply.classification === 'question') {
    line = `<strong>${escapeHtml(supplier.companyName)}</strong> asked: "${escapeHtml(reply.summary || '')}"`;
  } else if (reply.classification === 'suspicious') {
    line = `<strong>${escapeHtml(supplier.companyName)}</strong> reply flagged suspicious`;
  } else if (reply.classification === 'decline') {
    line = `<strong>${escapeHtml(supplier.companyName)}</strong> declined to quote`;
  } else {
    line = `<strong>${escapeHtml(supplier.companyName)}</strong> replied`;
  }
  return `
    <div class="activity-item">
      <div class="activity-icon">${icon}</div>
      <div class="activity-text">
        <div>${line}</div>
        <div class="activity-meta">${escapeHtml(formatDateTimeShort(reply.receivedAt))}</div>
      </div>
    </div>`;
}

function formatDateTimeShort(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  let h = d.getHours();
  const m = String(d.getMinutes()).padStart(2, '0');
  const ampm = h >= 12 ? 'pm' : 'am';
  h = h % 12; if (h === 0) h = 12;
  return `${dayNames[d.getDay()]} ${d.getDate()} ${monthNames[d.getMonth()]}, ${h}:${m}${ampm}`;
}

// Wire up the "refresh on return" callback once at boot
setOnReturnToJobs(loadJobs);

// Hard-delete an RFQ entry from the tracker. Triggered by the trash icon
// on each RFQ row (admin-only). Confirms before deleting; never deletes
// the supplier emails themselves.
async function deleteRfq(rfqId, ev) {
  if (ev) ev.stopPropagation();
  if (state.currentUserEmail !== 'est@auhs.com.au') {
    showToast('Only Est Can Delete', 'error');
    return;
  }
  const tracker = await readTracker(state.currentJob.folderName);
  if (!tracker || !Array.isArray(tracker.rfqs)) return;
  const rfq = tracker.rfqs.find(r => r.id === rfqId);
  if (!rfq) { showToast('RFQ Not Found', 'error'); return; }
  const supplierList = (rfq.suppliers || []).map(s => escapeHtml(s.companyName)).join(', ');
  const proceed = await confirmModal(
    'Delete This RFQ?',
    `Permanently delete the <strong>${escapeHtml(rfq.category)}</strong> RFQ to <strong>${supplierList}</strong>?<br><br>The emails already sent will not be unsent. This only removes the entry from the tracker.`,
    'Delete', 'Cancel'
  );
  if (!proceed) return;
  tracker.rfqs = tracker.rfqs.filter(r => r.id !== rfqId);
  try {
    await writeTracker(state.currentJob.folderName, tracker);
    await logAudit('RFQ_DELETED', `${rfq.category} (${state.currentJob.jobCode} ${state.currentJob.jobName})`, {
      rfqId,
      suppliers: (rfq.suppliers || []).map(s => s.email),
      sentAt: rfq.sentAt
    });
    showToast('Deleted', 'success');
    loadJobDetail();
  } catch (err) {
    console.error(err);
    showToast('Delete Failed', 'error');
  }
}

// Inline-onclick exposure
window.loadJobs = loadJobs;
window.loadJobDetail = loadJobDetail;
window.skipMigration = skipMigration;
window.migrateExistingQuotes = migrateExistingQuotes;
window.openCurrentJobInSharePoint = openCurrentJobInSharePoint;
window.deleteRfq = deleteRfq;
