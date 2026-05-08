// js/email-templates.js
// Email Templates tab in Settings.
//
// Architecture: one global default template + per-trade overrides.
// Templates live in state.appConfig.emailTemplates = { default: '...', byCategory: { 'Concrete': '...' } }
// Subject line uses a single global subject template (default: '{streetAddress} - RFQ - {tradeName}').
// All persisted in estimating-config.json.

import { state } from './state.js';
import { loadAppConfig, saveAppConfig, logAudit } from './audit.js';
import { showToast, showModal, closeModal, escapeHtml } from './ui.js';

// ----- Defaults -----

export const DEFAULT_EMAIL_BODY =
`Hi {firstName},
Please find attached the SOW for {fullAddress}.
Document folder: {tradiesLink}
Please respond by {respondByDate}.
Below is list of current file snapshot:
{filesList}

{signature}`;

export const DEFAULT_SUBJECT_TEMPLATE = '{streetAddress} - RFQ - {tradeName}';

// ----- Placeholders catalog (used in editor sidebar + sample preview) -----

const PLACEHOLDERS = [
  { key: 'firstName',     label: 'Recipient first name',          sample: 'Bob' },
  { key: 'companyName',   label: 'Recipient company name',        sample: "Bob's Concreting" },
  { key: 'jobCode',       label: 'Job code',                      sample: '2310' },
  { key: 'jobName',       label: 'Job name',                      sample: 'LF' },
  { key: 'streetAddress', label: 'Street number + name',          sample: '31 Langford' },
  { key: 'fullAddress',   label: 'Full address',                  sample: '31 Langford St Reservoir 3073' },
  { key: 'tradeName',     label: 'Trade / Supplier name',         sample: 'Roof Plumbing' },
  { key: 'requirements',  label: 'Job-specific requirements',     sample: 'Roof to be Colorbond Monument. Include all flashings, downpipes and sumps. Allow 6 days from start.' },
  { key: 'tradiesLink',   label: 'SharePoint drawings folder link', sample: '<a href="https://example">Tradies Folder Link</a>' },
  { key: 'filesList',     label: 'File list table (auto-built)',  sample: '<i>(File list table appears here at send time)</i>' },
  { key: 'respondByDate', label: 'Reply-by date',                 sample: 'Mon 12 May 2026' },
  { key: 'signature',     label: 'Signature block',               sample: 'Oskar Lue<br>Estimator<br>oskar@auhs.com.au<br><br>Auzzie Homes Pty Ltd' }
];

// ----- Migration / seeding -----

function ensureTemplatesShape() {
  if (!state.appConfig.emailTemplates) {
    state.appConfig.emailTemplates = {
      default: DEFAULT_EMAIL_BODY,
      subjectTemplate: DEFAULT_SUBJECT_TEMPLATE,
      byCategory: {}
    };
  }
  if (!state.appConfig.emailTemplates.subjectTemplate) {
    state.appConfig.emailTemplates.subjectTemplate = DEFAULT_SUBJECT_TEMPLATE;
  }
  if (!state.appConfig.emailTemplates.byCategory) {
    state.appConfig.emailTemplates.byCategory = {};
  }
}

// ----- Public lookups (used by future Phase 4c-iii Send RFQ flow) -----

// Resolve the body template for a category, falling back to the default.
export function getTemplateForCategory(category) {
  ensureTemplatesShape();
  const overrides = state.appConfig.emailTemplates.byCategory;
  return (category && overrides[category] != null && overrides[category] !== '')
    ? overrides[category]
    : state.appConfig.emailTemplates.default;
}

export function getSubjectTemplate() {
  ensureTemplatesShape();
  return state.appConfig.emailTemplates.subjectTemplate;
}

// Render a template string by substituting {placeholders} from a values object.
// IMPORTANT: this version expects HTML-safe values where appropriate. The
// caller (send-rfq for real emails, preview for samples) is responsible for
// escaping plain-text values like firstName before passing them in.
//
// Two-pass approach to keep injected HTML pristine:
//   1. Convert template's plain-text linebreaks to <br> first
//   2. Substitute placeholders (including HTML-bearing ones like
//      {filesList} and {tradiesLink}) into the now-HTML template
// This way HTML inside placeholder values isn't broken up by an over-eager
// \n → <br> pass.
export function renderTemplate(templateStr, values) {
  if (!templateStr) return '';
  let out = templateStr;
  for (const key of Object.keys(values || {})) {
    const re = new RegExp('\\{' + key + '\\}', 'g');
    out = out.replace(re, values[key] != null ? String(values[key]) : '');
  }
  return out;
}

// Convert plain-text linebreaks in the template to <br> BEFORE the caller
// substitutes HTML-bearing placeholders. Use this on the raw template string,
// then call renderTemplate() on the result with HTML-formatted values.
export function templateToHtml(templateStr) {
  if (!templateStr) return '';
  return templateStr.replace(/\n/g, '<br>');
}

// ----- Tab loader -----

export async function loadEmailTemplatesTab() {
  await loadAppConfig();
  ensureTemplatesShape();
  state.activeTemplateCategory = state.activeTemplateCategory || '__default__';
  renderEmailTemplatesTab();
}

function renderEmailTemplatesTab() {
  const trades = (state.appConfig.trades || []).slice().sort((a, b) => a.category.localeCompare(b.category));
  const overrides = state.appConfig.emailTemplates.byCategory;

  // Sidebar: Default + each trade. Show indicator on trades that have an override.
  const sidebarItems = [
    { key: '__default__', label: 'Default Template', isDefault: true, hasOverride: false },
    ...trades.map(t => ({
      key: t.category,
      label: t.category,
      isDefault: false,
      hasOverride: overrides[t.category] != null && overrides[t.category] !== ''
    }))
  ];

  const sidebarHtml = sidebarItems.map(item => `
    <div class="tpl-sidebar-item${state.activeTemplateCategory === item.key ? ' active' : ''}" data-key="${escapeHtml(item.key)}">
      <span class="tpl-sidebar-label">${escapeHtml(item.label)}</span>
      ${item.isDefault
        ? '<span class="tpl-sidebar-tag tpl-sidebar-tag-default">Base</span>'
        : (item.hasOverride
            ? '<span class="tpl-sidebar-tag tpl-sidebar-tag-override">Override</span>'
            : '<span class="tpl-sidebar-tag tpl-sidebar-tag-inherit">Default</span>')}
    </div>
  `).join('');

  document.getElementById('settings-email-templates').innerHTML = `
    <div class="info-card mb-12">
      <div style="font-weight: 600;">Email Templates</div>
      <div class="text-muted text-small mt-4">
        Edit the <strong>Default</strong> template that applies to every trade. Override individual trades only when the wording needs to differ.
        Placeholders like <code>{firstName}</code> are replaced when the email is sent.
      </div>
    </div>

    <div class="form-group">
      <label class="form-label">Subject Line Template</label>
      <input id="tpl-subject" type="text" value="${escapeHtml(state.appConfig.emailTemplates.subjectTemplate)}" />
      <div class="form-hint">Applies to every RFQ. Default: <code>{streetAddress} - RFQ - {tradeName}</code></div>
    </div>

    <div class="tpl-layout">
      <aside class="tpl-sidebar">
        <div class="text-muted text-small" style="font-weight: 600; padding: 6px 8px; text-transform: uppercase; letter-spacing: 0.5px;">Trades / Suppliers</div>
        <div class="tpl-sidebar-list">${sidebarHtml}</div>
      </aside>
      <section class="tpl-editor">
        <div id="tpl-editor-pane"></div>
      </section>
      <aside class="tpl-placeholders">
        <div class="text-muted text-small" style="font-weight: 600; padding: 6px 8px; text-transform: uppercase; letter-spacing: 0.5px;">Placeholders</div>
        <div class="text-muted text-small" style="padding: 0 8px 8px;">Click to copy to clipboard.</div>
        <div class="tpl-placeholder-list">
          ${PLACEHOLDERS.map(p => `
            <button class="tpl-placeholder-btn" data-ph="${escapeHtml(p.key)}" title="${escapeHtml(p.label)}">
              <code>{${escapeHtml(p.key)}}</code>
              <span class="tpl-placeholder-label">${escapeHtml(p.label)}</span>
            </button>
          `).join('')}
        </div>
      </aside>
    </div>
  `;

  // Sidebar click handler
  document.querySelectorAll('.tpl-sidebar-item').forEach(el => {
    el.addEventListener('click', () => {
      state.activeTemplateCategory = el.dataset.key;
      renderEmailTemplatesTab();
    });
  });

  // Placeholder buttons → copy to clipboard
  document.querySelectorAll('.tpl-placeholder-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const text = '{' + btn.dataset.ph + '}';
      try {
        await navigator.clipboard.writeText(text);
        showToast(`Copied ${text}`, 'success');
      } catch (err) {
        // Fallback for older browsers / non-secure contexts
        const ta = document.createElement('textarea');
        ta.value = text; document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); showToast(`Copied ${text}`, 'success'); }
        catch (e) { showToast('Copy Failed', 'error'); }
        document.body.removeChild(ta);
      }
    });
  });

  // Subject line save on blur
  document.getElementById('tpl-subject').addEventListener('blur', async (e) => {
    const v = e.target.value.trim() || DEFAULT_SUBJECT_TEMPLATE;
    if (v !== state.appConfig.emailTemplates.subjectTemplate) {
      state.appConfig.emailTemplates.subjectTemplate = v;
      try {
        await saveAppConfig();
        await logAudit('SUBJECT_TEMPLATE_UPDATED', '__global__', { value: v });
        showToast('Subject Saved', 'success');
      } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
    }
  });

  renderEditorPane();
}

function renderEditorPane() {
  const cat = state.activeTemplateCategory;
  const isDefault = cat === '__default__';
  const overrides = state.appConfig.emailTemplates.byCategory;
  const hasOverride = !isDefault && overrides[cat] != null && overrides[cat] !== '';
  const value = isDefault
    ? state.appConfig.emailTemplates.default
    : (hasOverride ? overrides[cat] : state.appConfig.emailTemplates.default);
  const editing = isDefault || hasOverride;

  const headerLabel = isDefault ? 'Default Template' : escapeHtml(cat);
  const statusLine = isDefault
    ? '<span class="text-muted text-small">This template applies to every trade unless overridden.</span>'
    : (hasOverride
        ? '<span class="text-small" style="color: var(--green);">Custom override active for this trade.</span>'
        : '<span class="text-muted text-small">Currently inheriting from Default. Click <em>Override</em> to customise.</span>');

  document.getElementById('tpl-editor-pane').innerHTML = `
    <div class="tpl-editor-header">
      <div>
        <div style="font-weight: 600; font-size: 16px;">${headerLabel}</div>
        <div class="mt-4">${statusLine}</div>
      </div>
      <div class="btn-row">
        <button class="btn-secondary small" id="tpl-preview-btn">Preview</button>
        ${isDefault
          ? `<button class="btn-secondary small" id="tpl-reset-default-btn" title="Restore the original built-in default">Reset</button>`
          : (hasOverride
              ? `<button class="btn-danger small" id="tpl-remove-override-btn">Remove Override</button>`
              : `<button class="btn-primary small" id="tpl-add-override-btn">Override</button>`)}
        ${editing ? `<button class="btn-primary small" id="tpl-save-btn">Save</button>` : ''}
      </div>
    </div>
    <textarea id="tpl-body" rows="16" ${editing ? '' : 'disabled style="opacity:0.6;"'}>${escapeHtml(value)}</textarea>
  `;

  // Preview always available
  document.getElementById('tpl-preview-btn').addEventListener('click', () => openPreview(cat));

  if (isDefault) {
    document.getElementById('tpl-reset-default-btn').addEventListener('click', async () => {
      state.appConfig.emailTemplates.default = DEFAULT_EMAIL_BODY;
      try {
        await saveAppConfig();
        await logAudit('TEMPLATE_DEFAULT_RESET', '__default__', null);
        renderEmailTemplatesTab();
        showToast('Default Reset', 'success');
      } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
    });
  } else if (hasOverride) {
    document.getElementById('tpl-remove-override-btn').addEventListener('click', async () => {
      delete overrides[cat];
      try {
        await saveAppConfig();
        await logAudit('TEMPLATE_OVERRIDE_REMOVED', cat, null);
        renderEmailTemplatesTab();
        showToast('Override Removed', 'success');
      } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
    });
  } else {
    document.getElementById('tpl-add-override-btn').addEventListener('click', async () => {
      // Seed override with current default so user has a starting point
      overrides[cat] = state.appConfig.emailTemplates.default;
      try {
        await saveAppConfig();
        await logAudit('TEMPLATE_OVERRIDE_ADDED', cat, null);
        renderEmailTemplatesTab();
        showToast('Override Created', 'success');
      } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
    });
  }

  if (editing) {
    document.getElementById('tpl-save-btn').addEventListener('click', async () => {
      const newValue = document.getElementById('tpl-body').value;
      if (isDefault) {
        state.appConfig.emailTemplates.default = newValue;
      } else {
        overrides[cat] = newValue;
      }
      try {
        await saveAppConfig();
        await logAudit('TEMPLATE_SAVED', cat, { isDefault });
        // Re-render sidebar (override status may change) without losing focus jumpily
        renderEmailTemplatesTab();
        showToast('Template Saved', 'success');
      } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
    });
  }
}

// ----- Preview modal -----

function openPreview(category) {
  const isDefault = category === '__default__';
  const body = isDefault
    ? state.appConfig.emailTemplates.default
    : getTemplateForCategory(category);
  const subjectTemplate = state.appConfig.emailTemplates.subjectTemplate;

  // Build sample values from PLACEHOLDERS samples
  const sample = {};
  for (const p of PLACEHOLDERS) sample[p.key] = p.sample;
  // If we're previewing a specific trade, use its name in tradeName sample
  if (!isDefault) sample.tradeName = category;

  const renderedSubject = renderTemplate(subjectTemplate, sample);
  // New order: convert linebreaks first, THEN substitute. Sample values
  // (including filesList table HTML) are injected into already-HTML template.
  const bodyAsHtml = templateToHtml(body);
  const renderedBodyHtml = renderTemplate(bodyAsHtml, sample);

  showModal(`
    <div class="modal-title">Preview${isDefault ? '' : ' — ' + escapeHtml(category)}</div>
    <div class="modal-body" style="margin-bottom: 12px;">
      Sample values used (placeholders shown filled in):
    </div>
    <div class="email-preview">
      <div class="email-preview-headers">
        <div class="email-preview-row"><span class="email-preview-label">From:</span> <span>${escapeHtml(sample.firstName ? 'Estimator' : '')} &lt;est@auhs.com.au&gt;</span></div>
        <div class="email-preview-row"><span class="email-preview-label">To:</span> <span>${escapeHtml(sample.firstName)} &lt;bob@example.com&gt;</span></div>
        <div class="email-preview-row"><span class="email-preview-label">Subject:</span> <strong>${escapeHtml(renderedSubject)}</strong></div>
      </div>
      <div class="email-preview-body">${renderedBodyHtml}</div>
    </div>
    <div class="modal-actions">
      <button class="btn-primary" onclick="closeModal()">Close</button>
    </div>
  `);
}

window.loadEmailTemplatesTab = loadEmailTemplatesTab;
