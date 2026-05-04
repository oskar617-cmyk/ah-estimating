// js/jobs.js
// Jobs list, job detail screen, and the existing-quotes migration flow.

import { CONFIG } from './config.js';
import { state } from './state.js';
import { navigate, setOnReturnToJobs } from './nav.js';
import { graphFetch, getAhSiteId, encodeUriPath, readJson, uploadJson } from './graph.js';
import { logAudit } from './audit.js';
import { showToast, escapeHtml } from './ui.js';

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
    let tracker = await readJson(siteId, quotePath, 'rfq-tracker.json');
    if (!tracker) {
      const quoteFolder = await graphFetch(
        `/sites/${siteId}/drive/root:/${encodeUriPath(quotePath)}:/children?$top=200&$select=id,name,file`
      ).catch(err => err.status === 404 ? { value: [] } : Promise.reject(err));
      const existingPdfs = (quoteFolder.value || []).filter(it => it.file && /\.pdf$/i.test(it.name));
      if (existingPdfs.length > 0) { renderMigrationPrompt(existingPdfs.length); return; }
      else { tracker = createEmptyTracker(); await uploadJson(siteId, quotePath, 'rfq-tracker.json', tracker); }
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
    await uploadJson(siteId, quotePath, 'rfq-tracker.json', tracker);
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
    await uploadJson(siteId, quotePath, 'rfq-tracker.json', tracker);
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
  const rfqCount = (tracker.rfqs || []).length;
  const migratedCount = (tracker.migratedQuotes || []).length;
  const needsReviewCount = (tracker.needsReview || []).length;
  let migratedSection = '';
  if (migratedCount > 0 || needsReviewCount > 0) {
    const items = [];
    if (migratedCount > 0) items.push(`<div class="text-small">${migratedCount} Parsed Quote${migratedCount === 1 ? '' : 's'} From Existing Files</div>`);
    if (needsReviewCount > 0) items.push(`<div class="text-small mt-4" style="color: var(--amber);">${needsReviewCount} File${needsReviewCount === 1 ? ' Needs' : 's Need'} Manual Review</div>`);
    migratedSection = `<div class="section-title">Migrated From Existing Files</div><div class="info-card">${items.join('')}</div>`;
  }
  container.innerHTML = `
    <div class="job-header">
      <div class="job-header-code">${escapeHtml(state.currentJob.jobCode)} ${escapeHtml(state.currentJob.jobName)}</div>
      <div class="job-header-address">${escapeHtml(state.currentJob.address)}</div>
      <div class="job-header-meta">
        <div><strong>Project Team:</strong> ${escapeHtml(teamDisplay)}</div>
      </div>
    </div>
    <div class="section-title">RFQs <span class="text-muted text-small" style="font-weight: normal; text-transform: none; letter-spacing: 0;">(${rfqCount} Total)</span></div>
    <div class="info-card">
      ${rfqCount === 0 ? '<div class="text-muted text-small">No RFQs Sent Yet</div>' : '<div class="text-muted text-small">RFQ List Display Coming In Next Phase</div>'}
      <button class="btn-primary mt-16" disabled title="Coming In Next Phase">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
        Send New RFQ
      </button>
    </div>
    ${migratedSection}
    <div class="section-title">Folder</div>
    <div class="info-card">
      <a href="${escapeHtml(state.currentJob.webUrl)}" target="_blank" style="color: var(--blue); text-decoration: none; font-size: 14px;">Open ${escapeHtml(state.currentJob.folderName)} In SharePoint</a>
    </div>
  `;
}

// Wire up the "refresh on return" callback once at boot
setOnReturnToJobs(loadJobs);

// Inline-onclick exposure
window.loadJobs = loadJobs;
window.loadJobDetail = loadJobDetail;
window.skipMigration = skipMigration;
window.migrateExistingQuotes = migrateExistingQuotes;
window.openCurrentJobInSharePoint = openCurrentJobInSharePoint;
