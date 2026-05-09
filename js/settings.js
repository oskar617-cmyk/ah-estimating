// js/settings.js
// Settings screen orchestration: opens the screen, switches tabs,
// handles the Signature tab. Catalog tab logic lives in catalog.js.

import { state } from './state.js';
import { navigate } from './nav.js';
import { loadAppConfig, saveAppConfig, logAudit } from './audit.js';
import { showToast, confirmModal } from './ui.js';
import { loadCatalogTab } from './catalog.js';
import { loadEmailTemplatesTab } from './email-templates.js';
import {
  buildExport, commitExport, downloadAsFile, STARTER_PROMPT_FOR_CLASSIFIER_CHAT
} from './decision-export.js';
import { readCursor, resetCursorTo } from './decision-log.js';

export async function openSettings() {
  navigate('settings-screen', {});
  // Force a fresh SOW filename scan each time Settings opens, so newly
  // uploaded SOW Word docs are reflected in the indicators.
  state.sowFilenames = null;
  document.querySelectorAll('.settings-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.settings-section').forEach(s => s.classList.remove('active'));
  document.querySelector('.settings-tab[data-tab="catalog"]').classList.add('active');
  document.getElementById('settings-catalog').classList.add('active');
  document.querySelectorAll('.settings-tab').forEach(tab => {
    tab.onclick = () => {
      document.querySelectorAll('.settings-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      const which = tab.dataset.tab;
      document.querySelectorAll('.settings-section').forEach(s => s.classList.remove('active'));
      document.getElementById('settings-' + which).classList.add('active');
      if (which === 'signature') loadSignatureTab();
      if (which === 'catalog') loadCatalogTab();
      if (which === 'email-templates') loadEmailTemplatesTab();
      if (which === 'ai-tuning') loadAiTuningTab();
    };
  });
  await loadCatalogTab();
}

export async function loadSignatureTab() {
  await loadAppConfig();
  document.getElementById('sig-title').value = state.appConfig.signature.title || '';
  document.getElementById('sig-body').value = state.appConfig.signature.body || '';
}

export async function saveSignature() {
  state.appConfig.signature = {
    title: document.getElementById('sig-title').value.trim(),
    body: document.getElementById('sig-body').value
  };
  try {
    await saveAppConfig();
    await logAudit('SIGNATURE_UPDATED', state.currentUserEmail, null);
    showToast('Signature Saved', 'success');
  } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
}

// ----- AI Tuning tab -----

async function loadAiTuningTab() {
  const statusEl = document.getElementById('ai-tuning-status');
  statusEl.innerHTML = '<div class="loading"><div class="spinner"></div><div>Checking decision log...</div></div>';
  try {
    const cursor = await readCursor();
    const lastTxt = cursor.lastExportAt
      ? `Last export: ${new Date(cursor.lastExportAt).toLocaleString()}`
      : 'No export yet — first export will include all decisions on record.';
    statusEl.innerHTML = `
      <div style="font-weight:600;">${lastTxt}</div>
      <div class="text-muted text-small mt-4">Exports cap at 500 entries. Re-export to fetch the next batch if more remain.</div>
    `;
  } catch (err) {
    statusEl.innerHTML = `<div style="color:var(--red);">Failed to read cursor: ${err.message}</div>`;
  }
  // Wire up handlers (re-bind on every tab open to be safe)
  const exportBtn = document.getElementById('btn-export-ai-tuning');
  const copyBtn = document.getElementById('btn-copy-starter-prompt');
  const resetBtn = document.getElementById('btn-reset-export-cursor');
  exportBtn.onclick = handleExport;
  copyBtn.onclick = handleCopyStarterPrompt;
  resetBtn.onclick = handleResetCursor;
}

async function handleExport() {
  const btn = document.getElementById('btn-export-ai-tuning');
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Building export...';
  try {
    const built = await buildExport({ excludeTest: true });
    if (built.count === 0) {
      showToast('Nothing New To Export', 'info');
      return;
    }
    downloadAsFile(built.filename, built.markdown);
    await commitExport();
    const truncatedTxt = built.truncated ? ' (capped — re-export later for the rest)' : '';
    showToast(`Exported ${built.count} entries${truncatedTxt}`, 'success');
    await logAudit('AI_TUNING_EXPORTED', `${built.count} entries`, {
      since: built.since,
      truncated: built.truncated
    });
    // Refresh the status display
    loadAiTuningTab();
  } catch (err) {
    console.error(err);
    showToast('Export Failed: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
    btn.textContent = original;
  }
}

async function handleCopyStarterPrompt() {
  try {
    await navigator.clipboard.writeText(STARTER_PROMPT_FOR_CLASSIFIER_CHAT);
    showToast('Starter Prompt Copied', 'success');
  } catch (err) {
    // Fallback: show in a modal so user can copy manually
    showToast('Clipboard blocked — copy from the modal', 'info');
    const m = await import('./ui.js');
    m.showModal(`
      <div class="modal-title">Starter Prompt</div>
      <div class="text-small text-muted mb-12">Copy this and paste it as the first message in your AH Est Email Classifier chat.</div>
      <textarea readonly rows="12" style="width:100%;font-size:12px;font-family:monospace;">${STARTER_PROMPT_FOR_CLASSIFIER_CHAT.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</textarea>
      <div class="modal-actions"><button class="btn-secondary" onclick="closeModal()">Close</button></div>
    `);
  }
}

async function handleResetCursor() {
  const proceed = await confirmModal(
    'Reset Export Cursor?',
    'This rewinds the export pointer so the next export will include older entries again. Useful if you want to re-export a previous period. Choose how far back to go in the next prompt.',
    'Reset…', 'Cancel'
  );
  if (!proceed) return;
  // Simple reset choices
  const m = await import('./ui.js');
  m.showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">Reset Export Cursor To…</div>
    <div class="form-group">
      <label class="form-label">Pick a starting point</label>
      <div class="btn-row" style="flex-direction:column;gap:6px;">
        <button class="btn-secondary" onclick="window._resetCursorChoice('all')">All time (re-export everything)</button>
        <button class="btn-secondary" onclick="window._resetCursorChoice('30d')">Last 30 days</button>
        <button class="btn-secondary" onclick="window._resetCursorChoice('7d')">Last 7 days</button>
      </div>
    </div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Cancel</button>
    </div>
  `);
  window._resetCursorChoice = async (which) => {
    const ui = await import('./ui.js');
    let iso = null;
    if (which === '30d') iso = new Date(Date.now() - 30 * 86400000).toISOString();
    else if (which === '7d') iso = new Date(Date.now() - 7 * 86400000).toISOString();
    // 'all' → null (treated as "since beginning of time")
    try {
      await resetCursorTo(iso);
      ui.closeModal();
      ui.showToast('Cursor Reset', 'success');
      loadAiTuningTab();
    } catch (err) {
      ui.showToast('Reset Failed: ' + err.message, 'error');
    }
  };
}

// Inline-onclick exposure
window.openSettings = openSettings;
window.saveSignature = saveSignature;
