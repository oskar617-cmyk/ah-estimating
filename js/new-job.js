// js/new-job.js
// New Job creation flow: form, validation, project team chips with email
// search, and the multi-step folder/file creation runner.

import { CONFIG } from './config.js';
import { state } from './state.js';
import { navigate, showScreen } from './nav.js';
import {
  graphFetch, getAhSiteId, getAhOfficeId, encodeUriPath,
  ensureFolder, copyFile, uploadJson
} from './graph.js';
import { logAudit } from './audit.js';
import { showToast, showModal, closeModal, confirmModal, escapeHtml } from './ui.js';
import { loadJobs } from './jobs.js';

export function openNewJob() {
  document.getElementById('nj-code').value = '';
  document.getElementById('nj-name').value = '';
  document.getElementById('nj-address').value = '';
  ['nj-code-err', 'nj-name-err', 'nj-address-err'].forEach(id => document.getElementById(id).classList.add('hidden'));
  document.getElementById('nj-progress').classList.add('hidden');
  document.getElementById('nj-submit').disabled = false;
  state.projectTeamEmails = [
    { email: 'oskar@auhs.com.au', name: 'Oskar', locked: true },
    { email: 'est@auhs.com.au', name: 'Est', locked: true }
  ];
  renderTeamChips();
  navigate('new-job-screen', {});
}

function renderTeamChips() {
  const container = document.getElementById('nj-team-chips');
  const chipsHtml = state.projectTeamEmails.map((c, i) => `
    <span class="email-chip${c.locked ? ' locked' : ''}" title="${escapeHtml(c.email)}">
      <span class="email-chip-text">${escapeHtml(c.name || c.email)}</span>
      ${c.locked ? '' : `
        <button class="email-chip-remove" data-idx="${i}" title="Remove">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
        </button>`}
    </span>
  `).join('');
  container.innerHTML = chipsHtml + `
    <button class="email-chip-add" onclick="openEmailSearch()">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
      Add
    </button>`;
  container.querySelectorAll('.email-chip-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const idx = parseInt(btn.dataset.idx, 10);
      state.projectTeamEmails.splice(idx, 1);
      renderTeamChips();
    });
  });
}

export function openEmailSearch() {
  showModal(`
    <div class="modal-title">Add Project Team Member</div>
    <div class="modal-body">Search by name or email. Suggestions come from your Auzzie Homes directory.</div>
    <div class="email-search-wrap">
      <input id="email-search-input" type="text" placeholder="Type a name or email..." autocomplete="off" />
      <div id="email-suggestions" class="email-suggestions hidden"></div>
    </div>
    <div class="modal-actions"><button class="btn-secondary" onclick="closeModal()">Cancel</button></div>`);
  setTimeout(() => {
    const input = document.getElementById('email-search-input');
    if (input) {
      input.focus();
      input.addEventListener('input', onEmailSearchInput);
      input.addEventListener('keydown', onEmailSearchKeydown);
    }
  }, 50);
}

function onEmailSearchInput(e) {
  const q = e.target.value.trim();
  const sug = document.getElementById('email-suggestions');
  clearTimeout(state.emailSearchTimer);
  if (q.length < 2) { sug.classList.add('hidden'); sug.innerHTML = ''; return; }
  sug.classList.remove('hidden');
  sug.innerHTML = '<div class="email-suggestion-loading"><div class="spinner-sm"></div> Searching...</div>';
  state.emailSearchTimer = setTimeout(() => searchTenantUsers(q), 300);
}

function onEmailSearchKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    const v = e.target.value.trim();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) addTeamMember({ email: v, name: v.split('@')[0] });
  } else if (e.key === 'Escape') closeModal();
}

async function searchTenantUsers(query) {
  const sug = document.getElementById('email-suggestions');
  if (!sug) return;
  try {
    const q = query.replace(/'/g, "''");
    const filter = `startswith(displayName,'${q}') or startswith(mail,'${q}') or startswith(userPrincipalName,'${q}')`;
    const result = await graphFetch(`/users?$filter=${encodeURIComponent(filter)}&$top=10&$select=displayName,mail,userPrincipalName`);
    const users = (result.value || []).filter(u => u.mail || u.userPrincipalName);
    if (users.length === 0) { sug.innerHTML = '<div class="email-suggestion-empty">No matches. Press Enter to add as a custom email.</div>'; return; }
    sug.innerHTML = users.map((u, i) => `
      <div class="email-suggestion" data-idx="${i}">
        <div class="email-suggestion-name">${escapeHtml(u.displayName || u.mail)}</div>
        <div class="email-suggestion-mail">${escapeHtml(u.mail || u.userPrincipalName)}</div>
      </div>`).join('');
    sug.querySelectorAll('.email-suggestion').forEach((el, i) => {
      el.addEventListener('click', () => { const u = users[i]; addTeamMember({ email: (u.mail || u.userPrincipalName).toLowerCase(), name: u.displayName }); });
    });
  } catch (err) { console.error(err); sug.innerHTML = '<div class="email-suggestion-empty">Search failed. Type the full email and press Enter.</div>'; }
}

function addTeamMember(member) {
  const lower = member.email.toLowerCase();
  if (state.projectTeamEmails.some(c => c.email.toLowerCase() === lower)) { showToast('Already Added', 'error'); return; }
  state.projectTeamEmails.push({ email: lower, name: member.name, locked: false });
  closeModal(); renderTeamChips();
}

export async function submitNewJob() {
  const codeEl = document.getElementById('nj-code'); const nameEl = document.getElementById('nj-name'); const addressEl = document.getElementById('nj-address');
  const code = codeEl.value.trim(); const name = nameEl.value.trim().toUpperCase(); const address = addressEl.value.trim();
  ['nj-code-err', 'nj-name-err', 'nj-address-err'].forEach(id => document.getElementById(id).classList.add('hidden'));
  [codeEl, nameEl, addressEl].forEach(el => el.classList.remove('invalid'));
  let valid = true;
  if (!CONFIG.codeRegex.test(code)) { showFieldError('nj-code', 'Must Be Exactly 4 Digits'); valid = false; }
  if (!CONFIG.nameRegex.test(name)) { showFieldError('nj-name', 'Must Start With A Letter, 2-5 Capitals/Digits'); valid = false; }
  if (address.length < 5) { showFieldError('nj-address', 'Address Required'); valid = false; }
  if (/[\\\/:*?"<>|]/.test(address)) { showFieldError('nj-address', 'Address Contains Illegal Characters'); valid = false; }
  if (!valid) return;
  try {
    const siteId = await getAhSiteId();
    const result = await graphFetch(`/sites/${siteId}/drive/root/children?$top=200&$select=name,folder`);
    const items = result.value || [];
    let exactDup = false; let nameDupExisting = null;
    for (const it of items) {
      if (!it.folder) continue;
      const m = it.name.match(CONFIG.jobFolderPattern); if (!m) continue;
      const eCode = m[1]; const eName = m[2].toUpperCase();
      if (eCode === code && eName === name) { exactDup = true; break; }
      if (eName === name) nameDupExisting = m;
    }
    if (exactDup) { showFieldError('nj-code', `Job ${code} ${name} Already Exists`); return; }
    if (nameDupExisting) {
      const proceed = await confirmModal('Duplicate Job Name', `Another job named <strong>${escapeHtml(name)}</strong> already exists: <strong>${escapeHtml(nameDupExisting[1])} ${escapeHtml(nameDupExisting[2])}</strong>.<br><br>Continue creating this new job anyway?`, 'Continue', 'Cancel');
      if (!proceed) return;
    }
  } catch (err) { console.error(err); showToast('Could Not Verify Duplicates', 'error'); return; }
  document.getElementById('nj-submit').disabled = true;
  await runJobCreation(code, name, address);
}

function showFieldError(fieldId, msg) {
  const errEl = document.getElementById(fieldId + '-err');
  errEl.textContent = msg; errEl.classList.remove('hidden');
  document.getElementById(fieldId).classList.add('invalid');
}

async function runJobCreation(code, name, address) {
  const ahSiteFolder = `${code} ${name} Site Docs - ${address}`;
  const ahOfficeFolder = `${code} ${name} - ${address}`;
  const tradiesSubfolder = `AAA Docs for Tradies ${name}`;
  const budgetTargetName = `0 Budget Control ${name}.xlsx`;
  const steps = [];
  steps.push({ id: 's-ahsite-root', label: `Create AH Site folder: ${ahSiteFolder}`, run: async () => { const siteId = await getAhSiteId(); await ensureFolder(siteId, '', ahSiteFolder); }});
  for (const sub of CONFIG.ahSiteSubfolders) {
    const finalName = sub === 'AAA Docs for Tradies' ? tradiesSubfolder : sub;
    steps.push({ id: `s-ahsite-${sub.replace(/\s+/g, '-').toLowerCase()}`, label: `Create subfolder: ${finalName}`, run: async () => { const siteId = await getAhSiteId(); await ensureFolder(siteId, ahSiteFolder, finalName); }});
  }
  steps.push({ id: 's-budget', label: `Copy budget template: ${budgetTargetName}`, run: async () => {
    const siteId = await getAhSiteId();
    const existing = await graphFetch(`/sites/${siteId}/drive/root:/${encodeUriPath(`${ahSiteFolder}/Quote/${budgetTargetName}`)}`).catch(err => err.status === 404 ? null : Promise.reject(err));
    if (existing) return;
    await copyFile(siteId, `${CONFIG.commonDocsPath}/${CONFIG.budgetTemplateName}`, `${ahSiteFolder}/Quote`, budgetTargetName);
  }});
  steps.push({ id: 's-tracker', label: 'Initialise RFQ tracker', run: async () => {
    const siteId = await getAhSiteId();
    const tracker = { version: 1, jobCode: code, jobName: name, address, projectTeamEmails: state.projectTeamEmails.map(c => c.email), rfqs: [], createdAt: new Date().toISOString(), createdBy: state.currentUserEmail };
    await uploadJson(siteId, `${ahSiteFolder}/Quote`, 'rfq-tracker.json', tracker);
  }});
  steps.push({ id: 's-ahoffice-root', label: `Create AH Office folder: ${ahOfficeFolder}`, run: async () => { const officeId = await getAhOfficeId(); await ensureFolder(officeId, '', ahOfficeFolder); }});
  for (const sub of CONFIG.ahOfficeSubfolders) {
    steps.push({ id: `s-ahoffice-${sub.replace(/\s+/g, '-').toLowerCase()}`, label: `Create AH Office subfolder: ${sub}`, run: async () => { const officeId = await getAhOfficeId(); await ensureFolder(officeId, ahOfficeFolder, sub); }});
  }
  steps.push({ id: 's-audit', label: 'Log to audit trail', run: async () => { await logAudit('JOB_CREATED', `${code} ${name}`, { ahSiteFolder, ahOfficeFolder, projectTeamEmails: state.projectTeamEmails.map(c => c.email) }); }});

  const progressEl = document.getElementById('nj-progress');
  const progressList = document.getElementById('nj-progress-list');
  progressEl.classList.remove('hidden');
  progressList.innerHTML = steps.map(s => `<div class="progress-item pending" id="prog-${s.id}"><div class="progress-icon">·</div><div>${escapeHtml(s.label)}</div></div>`).join('');

  let failedStep = null;
  for (const step of steps) {
    const el = document.getElementById('prog-' + step.id);
    el.classList.remove('pending'); el.classList.add('active');
    el.querySelector('.progress-icon').innerHTML = '<div class="spinner-sm"></div>';
    try {
      await step.run();
      el.classList.remove('active'); el.classList.add('done');
      el.querySelector('.progress-icon').innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>`;
    } catch (err) {
      console.error(`Step ${step.id} failed:`, err);
      el.classList.remove('active'); el.classList.add('failed');
      el.querySelector('.progress-icon').innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`;
      failedStep = { step, err }; break;
    }
  }
  const oldRetry = document.getElementById('nj-retry-block'); if (oldRetry) oldRetry.remove();
  if (failedStep) {
    const errMsg = failedStep.err.message || String(failedStep.err);
    const retryBlock = document.createElement('div');
    retryBlock.id = 'nj-retry-block';
    retryBlock.innerHTML = `<div class="form-error mt-12" style="display: block;">${escapeHtml(errMsg)}</div><div class="btn-row mt-12"><button class="btn-primary" id="nj-retry">Retry</button><button class="btn-secondary" onclick="goBack()">Back To Jobs</button></div>`;
    progressList.parentElement.appendChild(retryBlock);
    document.getElementById('nj-retry').addEventListener('click', () => runJobCreation(code, name, address));
    document.getElementById('nj-submit').disabled = false;
  } else {
    showToast('Job Created', 'success');
    setTimeout(() => { state.navStack.pop(); showScreen('jobs-screen'); loadJobs(); }, 800);
  }
}

// Inline-onclick exposure
window.openNewJob = openNewJob;
window.submitNewJob = submitNewJob;
window.openEmailSearch = openEmailSearch;
window.closeModal = closeModal;
