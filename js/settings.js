// js/settings.js
// Settings screen orchestration: opens the screen, switches tabs,
// handles the Signature tab. Catalog tab logic lives in catalog.js.

import { state } from './state.js';
import { navigate } from './nav.js';
import { loadAppConfig, saveAppConfig, logAudit } from './audit.js';
import { showToast } from './ui.js';
import { loadCatalogTab } from './catalog.js';
import { loadEmailTemplatesTab } from './email-templates.js';

export async function openSettings() {
  navigate('settings-screen', {});
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

// Inline-onclick exposure
window.openSettings = openSettings;
window.saveSignature = saveSignature;
