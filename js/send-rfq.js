// js/send-rfq.js
// Send RFQ wizard: pick trade → pick suppliers → fill requirements + dates
// → confirm budget row → preview → batch send → summary.
//
// Persists the batch into rfq-tracker.json. Sends via /me/sendMail; replies
// route back to whoever signed in to send (for v1, only est@auhs.com.au).

import { CONFIG } from './config.js';
import { state } from './state.js';
import { navigate, showScreen } from './nav.js';
import {
  graphFetch, getAhSiteId, encodeUriPath,
  readJson, uploadJson, listFiles, readBinary, fileExists,
  createAnonymousReadLink, sendMail, arrayBufferToBase64
} from './graph.js';
import { loadAppConfig, loadSuppliers, logAudit } from './audit.js';
import { showToast, showModal, closeModal, confirmModal, escapeHtml } from './ui.js';
import { getTemplateForCategory, getSubjectTemplate, renderTemplate, templateToHtml } from './email-templates.js';
import { openCompanyEditor } from './companies.js';

// In-memory wizard state. Reset each time the wizard opens.
const wizard = {
  step: 1,                  // 1..5
  trade: null,              // selected catalog item
  selectedSupplierIds: new Set(),
  requirements: '',
  daysToRespond: CONFIG.defaultDaysToRespond,
  daysToFollowup: CONFIG.defaultDaysToFollowup,
  budgetRowNo: null,
  // Snapshot data captured at confirm-time, used for send + tracker write
  snapshot: null
};

export async function openSendRfq() {
  if (state.currentUserEmail !== CONFIG.senderEmail) {
    showToast(`Only ${CONFIG.senderEmail} Can Send RFQs`, 'error');
    return;
  }
  await loadAppConfig();
  await loadSuppliers();
  // Reset wizard state
  wizard.step = 1;
  wizard.trade = null;
  wizard.selectedSupplierIds = new Set();
  wizard.requirements = '';
  wizard.daysToRespond = CONFIG.defaultDaysToRespond;
  wizard.daysToFollowup = CONFIG.defaultDaysToFollowup;
  wizard.budgetRowNo = null;
  wizard.snapshot = null;
  navigate('send-rfq-screen', {});
  renderWizard();
}

function renderWizard() {
  document.getElementById('rfq-title').textContent =
    `Send RFQ — Step ${wizard.step} of 5`;
  const root = document.getElementById('rfq-content');
  if (wizard.step === 1) renderStep1Trade(root);
  else if (wizard.step === 2) renderStep2Suppliers(root);
  else if (wizard.step === 3) renderStep3Details(root);
  else if (wizard.step === 4) renderStep4Preview(root);
  else if (wizard.step === 5) renderStep5Summary(root);
}

// --------- Step 1: Pick a trade ---------
function renderStep1Trade(root) {
  const trades = (state.appConfig.trades || []).slice().sort((a, b) => a.category.localeCompare(b.category));
  root.innerHTML = `
    <div class="filter-bar">
      <input id="rfq-trade-search" type="text" placeholder="Search trades..." />
    </div>
    <div id="rfq-trade-grid" class="rfq-tile-grid"></div>
    <div class="btn-row mt-16">
      <button class="btn-secondary" onclick="goBack()">Cancel</button>
    </div>
  `;
  const gridEl = document.getElementById('rfq-trade-grid');
  function renderGrid(filter) {
    const f = (filter || '').toLowerCase().trim();
    const items = trades.filter(t => !f || t.category.toLowerCase().includes(f));
    if (items.length === 0) {
      gridEl.innerHTML = '<div class="empty-state" style="grid-column:1/-1;"><div>No Matches</div></div>';
      return;
    }
    gridEl.innerHTML = items.map(t => {
      const supplierCount = (state.suppliersData.suppliers || []).filter(s => (s.trades || []).includes(t.category) && s.active !== false).length;
      const sowExists = (state.sowFilenames || []).includes(`${t.category}.docx`);
      const budgetOK = !!t.budgetRowNo;
      const peopleSvg = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="11" height="11" style="vertical-align:-1px;"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>';
      const supTitle = `${supplierCount} active supplier${supplierCount === 1 ? '' : 's'}`;
      const budgetTitle = budgetOK ? `Budget row ${t.budgetRowNo}` : 'No budget row mapped';
      const sowTitle = sowExists ? `${t.category}.docx found in SOW Templates` : `${t.category}.docx missing from SOW Templates`;
      // Tile: trade name + three small status pills (suppliers / budget / SOW)
      return `
        <button class="rfq-tile" data-cat="${escapeHtml(t.category)}" type="button">
          <div class="rfq-tile-name">${escapeHtml(t.category)}</div>
          <div class="rfq-tile-meta">
            <span class="rfq-tile-dot ${supplierCount > 0 ? 'ok' : 'warn'}" title="${escapeHtml(supTitle)}">${supplierCount} ${peopleSvg}</span>
            <span class="rfq-tile-dot ${budgetOK ? 'ok' : 'warn'}" title="${escapeHtml(budgetTitle)}">${budgetOK ? escapeHtml(t.budgetRowNo) : 'no row'}</span>
            <span class="rfq-tile-dot ${sowExists ? 'ok' : 'warn'}" title="${escapeHtml(sowTitle)}">${sowExists ? 'SOW' : 'no SOW'}</span>
          </div>
        </button>`;
    }).join('');
    gridEl.querySelectorAll('.rfq-tile').forEach(el => {
      el.addEventListener('click', () => {
        const cat = el.dataset.cat;
        wizard.trade = state.appConfig.trades.find(t => t.category === cat);
        wizard.budgetRowNo = wizard.trade.budgetRowNo;
        wizard.daysToRespond = wizard.trade.daysToRespond || CONFIG.defaultDaysToRespond;
        wizard.daysToFollowup = wizard.trade.daysToFollowup || CONFIG.defaultDaysToFollowup;
        wizard.step = 2;
        renderWizard();
      });
    });
  }
  renderGrid('');
  document.getElementById('rfq-trade-search').addEventListener('input', (e) => renderGrid(e.target.value));
}

// --------- Step 2: Pick suppliers ---------
function renderStep2Suppliers(root) {
  const cat = wizard.trade.category;
  const suppliers = (state.suppliersData.suppliers || [])
    .filter(s => (s.trades || []).includes(cat) && s.active !== false)
    .sort((a, b) => (a.companyName || '').localeCompare(b.companyName || ''));
  const empty = suppliers.length === 0;
  root.innerHTML = `
    <div class="info-card mb-12" style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;">
      <div style="flex:1;min-width:200px;">
        <div style="font-weight:600;">Select Suppliers For ${escapeHtml(cat)}</div>
        <div class="text-muted text-small mt-4">Each selected supplier receives a separate email with the same body.</div>
      </div>
      <button class="btn-secondary small" id="rfq-add-company">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
        Add Company
      </button>
    </div>
    ${empty ? `
      <div class="empty-state">
        <div>No Active Suppliers For ${escapeHtml(cat)}</div>
        <div class="text-small mt-8">Click <strong>Add Company</strong> above to add one.</div>
      </div>` : `
      <div class="rfq-supplier-grid" id="rfq-supplier-grid">
        ${suppliers.map(s => `
          <label class="rfq-supplier-row" data-id="${escapeHtml(s.id)}">
            <input type="checkbox" ${wizard.selectedSupplierIds.has(s.id) ? 'checked' : ''} />
            <div class="rfq-supplier-info">
              <div class="rfq-supplier-name">${escapeHtml(s.companyName)}</div>
              <div class="rfq-supplier-meta">${escapeHtml(s.contactName || '?')} · ${escapeHtml(s.email)}</div>
            </div>
          </label>
        `).join('')}
      </div>`}
    <div class="btn-row mt-16">
      <button class="btn-secondary" id="rfq-back-1">Back</button>
      <button class="btn-primary" id="rfq-next-3">Next</button>
    </div>
  `;
  document.getElementById('rfq-back-1').addEventListener('click', () => { wizard.step = 1; renderWizard(); });
  document.getElementById('rfq-next-3').addEventListener('click', () => {
    if (wizard.selectedSupplierIds.size === 0) { showToast('Select At Least One Supplier', 'error'); return; }
    wizard.step = 3;
    renderWizard();
  });
  document.getElementById('rfq-add-company').addEventListener('click', () => {
    // Open the company editor preset to this trade. After save, auto-tick the
    // new supplier and re-render Step 2 so it appears in the list.
    openCompanyEditor(null, cat, (newSupplier) => {
      if (newSupplier && newSupplier.id) wizard.selectedSupplierIds.add(newSupplier.id);
      renderWizard();
    });
  });
  if (!empty) {
    document.querySelectorAll('.rfq-supplier-row').forEach(row => {
      const cb = row.querySelector('input[type="checkbox"]');
      cb.addEventListener('change', () => {
        const id = row.dataset.id;
        if (cb.checked) wizard.selectedSupplierIds.add(id);
        else wizard.selectedSupplierIds.delete(id);
      });
    });
  }
}

// --------- Step 3: Requirements + dates + budget row ---------
function renderStep3Details(root) {
  const trade = wizard.trade;
  const rows = trade.availableRows || [];
  const rowOpts = rows.map(r => {
    const label = r.no ? `${r.no} — ${r.description || ''}` : (r.description || '');
    return `<option value="${escapeHtml(r.no || '')}"${wizard.budgetRowNo === r.no ? ' selected' : ''}>${escapeHtml(label)}</option>`;
  }).join('');
  root.innerHTML = `
    <div class="info-card mb-12">
      <div style="font-weight:600;">${escapeHtml(trade.category)} — RFQ Details</div>
      <div class="text-muted text-small mt-4">${wizard.selectedSupplierIds.size} supplier(s) selected.</div>
    </div>
    <div class="form-group">
      <label class="form-label">Job-Specific Requirements</label>
      <textarea id="rfq-requirements" rows="6" placeholder="1. ">${escapeHtml(wizard.requirements)}</textarea>
      <div class="form-hint">Numbered list — press Enter to start the next line. Inserted into the SOW Word doc (not the email body). Leave empty if there are none.</div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Days To Respond</label>
        <input id="rfq-days-respond" type="number" min="1" max="60" value="${wizard.daysToRespond}" />
        <div class="form-hint">Used in <code>{respondByDate}</code></div>
      </div>
      <div class="form-group">
        <label class="form-label">Days Until Follow-Up</label>
        <input id="rfq-days-followup" type="number" min="1" max="60" value="${wizard.daysToFollowup}" />
        <div class="form-hint">Auto-reminder if no reply</div>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Budget Row To Update With Quote</label>
      <select id="rfq-budget-row">
        <option value="">— Select —</option>
        ${rowOpts}
      </select>
      <div class="form-hint">When a quote arrives, this is the row in the budget Excel to update.</div>
    </div>
    <div class="btn-row mt-16">
      <button class="btn-secondary" id="rfq-back-2">Back</button>
      <button class="btn-primary" id="rfq-next-4">Next: Preview</button>
    </div>
  `;
  // Auto-numbering: pre-fill with "1. " if empty, and on Enter add next number
  const reqEl = document.getElementById('rfq-requirements');
  if (!reqEl.value) reqEl.value = '1. ';
  reqEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      const before = reqEl.value.slice(0, reqEl.selectionStart);
      const after = reqEl.value.slice(reqEl.selectionEnd);
      // Find the highest number used so far (anywhere in the textarea)
      const nums = (reqEl.value.match(/^\s*(\d+)\./gm) || []).map(s => parseInt(s.match(/\d+/)[0], 10));
      const next = nums.length ? Math.max(...nums) + 1 : 1;
      const insert = `\n${next}. `;
      reqEl.value = before + insert + after;
      const pos = (before + insert).length;
      reqEl.setSelectionRange(pos, pos);
    }
  });
  document.getElementById('rfq-back-2').addEventListener('click', () => { wizard.step = 2; renderWizard(); });
  document.getElementById('rfq-next-4').addEventListener('click', () => {
    // Capture verbatim — preserve newlines for SOW insertion
    const raw = document.getElementById('rfq-requirements').value;
    // Strip if it's only "1. " (the empty default placeholder)
    wizard.requirements = (raw.trim() === '1.' || raw.trim() === '1. ') ? '' : raw;
    wizard.daysToRespond = parseInt(document.getElementById('rfq-days-respond').value, 10) || CONFIG.defaultDaysToRespond;
    wizard.daysToFollowup = parseInt(document.getElementById('rfq-days-followup').value, 10) || CONFIG.defaultDaysToFollowup;
    wizard.budgetRowNo = document.getElementById('rfq-budget-row').value || null;
    // Requirements may be empty (becomes "None" in SOW)
    if (!wizard.budgetRowNo) { showToast('Pick A Budget Row', 'error'); return; }
    wizard.step = 4;
    renderWizard();
  });
}

// --------- Step 4: Preview ---------
async function renderStep4Preview(root) {
  root.innerHTML = `
    <div class="info-card mb-12">
      <div style="font-weight:600;">Preview Email</div>
      <div class="text-muted text-small mt-4">Showing the email as it will arrive for the first selected supplier. Each supplier gets the same body with their own greeting.</div>
    </div>
    <div id="rfq-preview-area"><div class="loading"><div class="spinner"></div><div>Building Preview...</div></div></div>
    <div class="btn-row mt-16">
      <button class="btn-secondary" id="rfq-back-3">Back</button>
      <button class="btn-primary" id="rfq-send-btn" disabled>Send</button>
    </div>
  `;
  document.getElementById('rfq-back-3').addEventListener('click', () => { wizard.step = 3; renderWizard(); });
  // Build snapshot (drawings list, share link, SOW, recipients)
  try {
    const snapshot = await buildSnapshot();
    wizard.snapshot = snapshot;
    renderPreview();
    document.getElementById('rfq-send-btn').disabled = false;
    document.getElementById('rfq-send-btn').addEventListener('click', () => doBatchSend());
  } catch (err) {
    console.error('Preview build error:', err);
    document.getElementById('rfq-preview-area').innerHTML = `
      <div class="info-card" style="border-color: var(--red);">
        <div style="color: var(--red); font-weight: 600;">Could Not Build Preview</div>
        <div class="text-small mt-8">${escapeHtml(err.message)}</div>
      </div>`;
  }
}

async function buildSnapshot() {
  const job = state.currentJob;
  const trade = wizard.trade;
  const siteId = await getAhSiteId();

  // Drawings folder path within AH Site Documents library
  const tradiesFolderName = `AAA Docs for Tradies ${job.jobName}`;
  const tradiesFolderPath = `${job.folderName}/${tradiesFolderName}`;

  // List files (folders excluded)
  const files = await listFiles(siteId, tradiesFolderPath);
  // Anonymous-read share link for the folder
  const tradiesShareLink = await createAnonymousReadLink(siteId, tradiesFolderPath);

  // Read project team CC list from rfq-tracker
  const quotePath = `${job.folderName}/Quote`;
  const tracker = (await readJson(siteId, quotePath, 'rfq-tracker.json')) || {};
  const projectTeamCC = tracker.projectTeamEmails || [];

  // SOW: try to read [Category].docx from SOW Templates folder.
  // If not present, prompt user before continuing.
  // The SOW template should contain {{REQUIREMENTS}} where the estimator's
  // job-specific requirements should be inserted. Replaced before sending —
  // a fresh modified copy is built per-RFQ in memory, never saved to disk.
  const sowFilename = `${trade.category}.docx`;
  const sowExists = await fileExists(siteId, `${CONFIG.commonDocsPath}/SOW Templates`, sowFilename);
  let sowAttachment = null;
  if (sowExists) {
    const buf = await readBinary(siteId, `${CONFIG.commonDocsPath}/SOW Templates`, sowFilename);
    const modifiedBuf = injectRequirementsIntoDocx(buf, wizard.requirements);
    sowAttachment = {
      name: sowFilename,
      contentBytes: arrayBufferToBase64(modifiedBuf),
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    };
  } else {
    // Warning popup, user decides whether to proceed without SOW
    const proceed = await confirmModal(
      'No SOW Found',
      `No SOW Word document found for <strong>${escapeHtml(trade.category)}</strong>.<br><br>Expected file: <code>${escapeHtml(CONFIG.commonDocsPath)}/SOW Templates/${escapeHtml(sowFilename)}</code><br><br>Send the RFQ <strong>without an SOW attachment</strong>?`,
      'Send Without SOW', 'Cancel'
    );
    if (!proceed) {
      const err = new Error('Cancelled — upload SOW and try again');
      throw err;
    }
  }

  // Compute respond-by date
  const respondByDate = computeRespondByDate(wizard.daysToRespond);

  // Build subject line (single-line, no HTML)
  const subjectTemplate = getSubjectTemplate();
  const streetAddress = parseStreetFromAddress(job.address);
  const subjectValues = {
    streetAddress, fullAddress: job.address,
    jobCode: job.jobCode, jobName: job.jobName,
    tradeName: trade.category
  };
  const subject = renderTemplate(subjectTemplate, subjectValues);

  // Build body template (per-trade override or default)
  const bodyTemplate = getTemplateForCategory(trade.category);

  // Build {filesList} HTML table
  const filesListHtml = buildFilesListHtml(files);

  // Build {tradiesLink} as a clickable link
  // target="_blank" so clicking the link in the preview opens a new tab
  // rather than navigating away from the wizard. The recipient's email client
  // decides behaviour for the actual sent email.
  const tradiesLinkHtml = `<a href="${escapeHtml(tradiesShareLink)}" target="_blank" rel="noopener">${escapeHtml(tradiesFolderName)}</a>`;

  // Resolve signature
  const sig = state.appConfig.signature || {};
  const signatureHtml = (sig.body || '').replace(/\n/g, '<br>');

  // Recipient details
  const allSuppliers = state.suppliersData.suppliers;
  const selected = Array.from(wizard.selectedSupplierIds)
    .map(id => allSuppliers.find(s => s.id === id))
    .filter(Boolean);

  return {
    subject, bodyTemplate,
    tradiesShareLink, tradiesLinkHtml,
    filesListHtml, files,
    sowAttachment, sowFilename: sowExists ? sowFilename : null,
    respondByDate,
    selected,
    streetAddress, projectTeamCC,
    signatureHtml,
    job: { ...job }, trade: { ...trade }
  };
}

function parseStreetFromAddress(fullAddress) {
  // Take everything up to but not including a street suffix (St, Rd, Ave, etc.)
  // Falls back to the first 3 words if no suffix detected.
  const m = fullAddress.match(/^(.+?)\s+(?:St|Street|Rd|Road|Ave|Avenue|Cres|Crescent|Pl|Place|Dr|Drive|Ct|Court|Ln|Lane|Hwy|Highway|Way|Pde|Parade|Tce|Terrace|Blvd|Boulevard)\b/i);
  if (m) return m[1].trim();
  return fullAddress.split(/\s+/).slice(0, 3).join(' ');
}

function computeRespondByDate(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  // Format: "Mon 12 May 2026"
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${dayNames[d.getDay()]} ${d.getDate()} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

// Modify a SOW Word doc (.docx ArrayBuffer) by replacing the {{REQUIREMENTS}}
// placeholder with the estimator's job-specific requirements. Returns a
// modified ArrayBuffer ready to attach. Original on SharePoint is untouched.
//
// How this works (plain language):
//   A .docx is a zip file containing word/document.xml among others. The
//   text the user sees lives in that XML inside <w:t> tags. Word sometimes
//   splits a placeholder like {{REQUIREMENTS}} across multiple <w:t> runs
//   (e.g. if it was typed with intermediate formatting changes). We work
//   around that by joining adjacent <w:t> contents inside the same paragraph
//   before searching for the placeholder.
//
// If the placeholder isn't found, the doc is returned unchanged. If
// requirements is empty/blank, "None" is inserted.
function injectRequirementsIntoDocx(arrayBuf, requirementsText) {
  if (typeof PizZip === 'undefined') {
    throw new Error('PizZip library not loaded — cannot modify SOW Word doc');
  }
  const zip = new PizZip(arrayBuf);
  const docXml = zip.file('word/document.xml');
  if (!docXml) {
    console.warn('Unexpected docx structure — word/document.xml missing. Sending unchanged.');
    return arrayBuf;
  }
  let xml = docXml.asText();

  // First pass: try a simple replace in case the placeholder is intact
  // within a single <w:t> run.
  const PLACEHOLDER = '{{REQUIREMENTS}}';
  const wantsListBlock = !!(requirementsText && requirementsText.trim());
  const replacement = wantsListBlock
    ? buildRequirementsXml(requirementsText)
    : 'None';

  let replaced = false;
  if (xml.includes(PLACEHOLDER)) {
    // Need to substitute INSIDE a <w:t> run if multi-line list, else plain text
    if (wantsListBlock) {
      // Replacing with a multi-paragraph block requires breaking out of the
      // containing <w:p> paragraph. Find the surrounding paragraph and split.
      xml = replaceInsideParagraphWithBlock(xml, PLACEHOLDER, replacement);
    } else {
      xml = xml.replace(PLACEHOLDER, escapeXml(replacement));
    }
    replaced = true;
  } else {
    // Second pass: placeholder might be split across <w:t> runs (Word can
    // do this if styling changed mid-token). Reconstruct visible text per
    // paragraph, locate the placeholder span, and rewrite it.
    const result = replaceAcrossRuns(xml, PLACEHOLDER, replacement, wantsListBlock);
    if (result.replaced) { xml = result.xml; replaced = true; }
  }

  if (!replaced) {
    // Placeholder not found — leave doc unchanged. Estimator will see the
    // unmodified template; this is recoverable behaviour.
    console.warn(`SOW template missing ${PLACEHOLDER} placeholder — sending unchanged`);
    return arrayBuf;
  }

  zip.file('word/document.xml', xml);
  return zip.generate({ type: 'arraybuffer' });
}

// Build OOXML for a numbered-list block. Each line of `text` becomes its own
// paragraph. We use plain numbering inline (not a Word numId list) because
// the estimator already typed numbers like "1. ", "2. " in the textarea —
// the simplest, most reliable rendering is to preserve them as-is.
function buildRequirementsXml(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length === 0) return '<w:p><w:r><w:t>None</w:t></w:r></w:p>';
  return lines.map(line =>
    `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`
  ).join('');
}

function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Replace a placeholder that lives inside a <w:p>...{{X}}...</w:p> with a
// block of <w:p> paragraphs. We split the original paragraph at the
// placeholder, drop the placeholder text, and insert the new paragraphs
// after closing the original paragraph.
function replaceInsideParagraphWithBlock(xml, placeholder, blockXml) {
  const idx = xml.indexOf(placeholder);
  if (idx < 0) return xml;
  // Find the enclosing <w:p>...</w:p>
  const pStart = xml.lastIndexOf('<w:p ', idx);
  const pStartAlt = xml.lastIndexOf('<w:p>', idx);
  const paragraphStart = Math.max(pStart, pStartAlt);
  const paragraphEndIdx = xml.indexOf('</w:p>', idx);
  if (paragraphStart < 0 || paragraphEndIdx < 0) {
    // Fallback: just substitute as text
    return xml.replace(placeholder, blockXml);
  }
  const paragraphEnd = paragraphEndIdx + '</w:p>'.length;
  // Drop the entire enclosing paragraph (which contained the placeholder)
  // and replace it with the block. This is simpler than trying to keep the
  // surrounding text, which is rarely useful for a placeholder paragraph.
  return xml.slice(0, paragraphStart) + blockXml + xml.slice(paragraphEnd);
}

// Best-effort replacement when the placeholder is split across multiple
// <w:t> runs. We walk paragraph-by-paragraph; for each, we extract the
// concatenation of its <w:t> contents; if that concatenation contains the
// placeholder, we rewrite the entire paragraph.
function replaceAcrossRuns(xml, placeholder, replacement, isBlock) {
  let replaced = false;
  // Split xml by paragraph boundaries (<w:p ...> ... </w:p>)
  const out = xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    if (replaced) return paragraph; // only do first match
    // Collect text from <w:t>...</w:t>
    const texts = [];
    paragraph.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, t) => { texts.push(t); return m; });
    const combined = texts.join('');
    if (!combined.includes(placeholder)) return paragraph;
    replaced = true;
    if (isBlock) {
      // Replace the entire paragraph with the block
      return replacement;
    } else {
      // Single-line replacement — emit a fresh simple paragraph
      return `<w:p><w:r><w:t xml:space="preserve">${escapeXml(combined.replace(placeholder, replacement))}</w:t></w:r></w:p>`;
    }
  });
  return { xml: out, replaced };
}


  if (!files || files.length === 0) {
    return '<p style="color:#777;font-style:italic;margin:0;">No files in drawings folder yet.</p>';
  }
  const rows = files.map(f =>
    `<tr><td style="padding:6px 12px 6px 0;border-bottom:1px solid #eee;">${escapeHtml(f.name)}</td><td style="padding:6px 0;border-bottom:1px solid #eee;color:#666;white-space:nowrap;">${formatDate(f.lastModifiedDateTime)}</td></tr>`
  ).join('');
  return `<table style="border-collapse:collapse;font-size:13px;margin:0;"><thead><tr><th style="text-align:left;padding:6px 12px 6px 0;border-bottom:2px solid #999;">File Name</th><th style="text-align:left;padding:6px 0;border-bottom:2px solid #999;">Date Modified</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function formatDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return '';
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${d.getDate()} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
}

function buildEmailHtml(snapshot, supplier) {
  // Plain-text placeholders are HTML-escaped; HTML-bearing placeholders
  // (tradiesLink, filesList, signature) are passed through as-is.
  const safeRequirements = escapeHtml(wizard.requirements).replace(/\n/g, '<br>');
  const safeValues = {
    firstName: escapeHtml(supplier.contactName || ''),
    companyName: escapeHtml(supplier.companyName || ''),
    jobCode: escapeHtml(snapshot.job.jobCode),
    jobName: escapeHtml(snapshot.job.jobName),
    streetAddress: escapeHtml(snapshot.streetAddress),
    fullAddress: escapeHtml(snapshot.job.address),
    tradeName: escapeHtml(snapshot.trade.category),
    requirements: safeRequirements,
    tradiesLink: snapshot.tradiesLinkHtml,
    filesList: snapshot.filesListHtml,
    respondByDate: escapeHtml(snapshot.respondByDate),
    signature: snapshot.signatureHtml
  };
  // Convert linebreaks BEFORE substitution so HTML inside placeholder values
  // (especially the {filesList} table) doesn't get torn apart by \n → <br>.
  const bodyAsHtml = templateToHtml(snapshot.bodyTemplate);
  return renderTemplate(bodyAsHtml, safeValues);
}

function renderPreview() {
  const snap = wizard.snapshot;
  const sample = snap.selected[0];
  const html = buildEmailHtml(snap, sample);
  const ccLine = snap.projectTeamCC.length
    ? `<div class="email-preview-row"><span class="email-preview-label">CC:</span> <span>${escapeHtml(snap.projectTeamCC.join(', '))}</span></div>`
    : '';
  document.getElementById('rfq-preview-area').innerHTML = `
    <div class="email-preview">
      <div class="email-preview-headers">
        <div class="email-preview-row"><span class="email-preview-label">From:</span> <span>${escapeHtml(state.currentAccount.name || '')} &lt;${escapeHtml(state.currentUserEmail)}&gt;</span></div>
        <div class="email-preview-row"><span class="email-preview-label">To:</span> <span>${escapeHtml(sample.contactName || '')} &lt;${escapeHtml(sample.email)}&gt; <em style="color:#999;">(and ${snap.selected.length - 1} other${snap.selected.length === 2 ? '' : 's'})</em></span></div>
        ${ccLine}
        <div class="email-preview-row"><span class="email-preview-label">Subject:</span> <strong>${escapeHtml(snap.subject)}</strong></div>
        ${snap.sowFilename
          ? `<div class="email-preview-row"><span class="email-preview-label">Attached:</span> <span>📎 ${escapeHtml(snap.sowFilename)}</span></div>`
          : `<div class="email-preview-row"><span class="email-preview-label">Attached:</span> <span style="color:var(--amber);">⚠ No SOW</span></div>`}
      </div>
      <div class="email-preview-body">${html}</div>
    </div>
    <div class="info-card mt-12">
      <div class="text-small text-muted">
        Sending to <strong>${snap.selected.length}</strong> supplier${snap.selected.length === 1 ? '' : 's'}: ${snap.selected.map(s => escapeHtml(s.companyName)).join(', ')}
      </div>
    </div>
  `;
}

// --------- Send batch ---------
async function doBatchSend() {
  const snap = wizard.snapshot;
  const sendBtn = document.getElementById('rfq-send-btn');
  sendBtn.disabled = true;
  sendBtn.innerHTML = '<div class="spinner-sm"></div> Sending...';

  const rfqId = `rfq-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  const sentEntries = [];
  const failed = [];
  const previewArea = document.getElementById('rfq-preview-area');
  previewArea.innerHTML = '<div class="info-card"><div style="font-weight:600;margin-bottom:8px;">Sending RFQs...</div><div id="rfq-send-progress" class="progress-list"></div></div>';
  const progressList = document.getElementById('rfq-send-progress');

  for (const supplier of snap.selected) {
    const itemId = 'send-' + supplier.id;
    progressList.insertAdjacentHTML('beforeend', `
      <div class="progress-item active" id="prog-${itemId}">
        <div class="progress-icon"><div class="spinner-sm"></div></div>
        <div>${escapeHtml(supplier.companyName)} &lt;${escapeHtml(supplier.email)}&gt;</div>
      </div>`);
    const el = document.getElementById('prog-' + itemId);
    try {
      const html = buildEmailHtml(snap, supplier);
      await sendMail({
        subject: snap.subject,
        htmlBody: html,
        toRecipients: [supplier.email],
        ccRecipients: snap.projectTeamCC,
        replyToRecipients: [state.currentUserEmail],
        attachments: snap.sowAttachment ? [snap.sowAttachment] : [],
        // Custom header — admin can target this in Defender / mail flow rules
        // to whitelist all RFQs from this app.
        customHeaders: { 'x-ah-estimating': 'rfq-v1' }
      });
      sentEntries.push({
        id: supplier.id,
        companyName: supplier.companyName,
        contactName: supplier.contactName || '',
        email: supplier.email,
        status: 'sent',
        sentAt: new Date().toISOString(),
        lastFollowupAt: null,
        followupCount: 0,
        replies: []
      });
      el.classList.remove('active'); el.classList.add('done');
      el.querySelector('.progress-icon').innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>`;
    } catch (err) {
      console.error(`Send to ${supplier.email} failed:`, err);
      failed.push({ supplier, error: err.message || String(err) });
      el.classList.remove('active'); el.classList.add('failed');
      el.querySelector('.progress-icon').innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`;
      el.insertAdjacentHTML('beforeend', `<div style="color:var(--red);font-size:11px;margin-left:8px;">${escapeHtml(err.message || '')}</div>`);
    }
  }

  // Persist tracker entry if anything was sent successfully
  if (sentEntries.length > 0) {
    try {
      await persistRfqEntry(rfqId, snap, sentEntries);
      await logAudit('RFQ_SENT', `${snap.job.jobCode} ${snap.job.jobName} / ${snap.trade.category}`, {
        rfqId, supplierCount: sentEntries.length, failedCount: failed.length
      });
    } catch (err) {
      console.error('Tracker persist failed:', err);
      showToast('Sent But Tracker Save Failed', 'error');
    }
  }

  wizard.step = 5;
  wizard.snapshot._sendResult = { sentEntries, failed, rfqId };
  renderWizard();
}

async function persistRfqEntry(rfqId, snap, sentEntries) {
  const siteId = await getAhSiteId();
  const quotePath = `${snap.job.folderName}/Quote`;
  const tracker = (await readJson(siteId, quotePath, 'rfq-tracker.json')) || {
    version: 1,
    jobCode: snap.job.jobCode,
    jobName: snap.job.jobName,
    address: snap.job.address,
    projectTeamEmails: snap.projectTeamCC,
    rfqs: [],
    createdAt: new Date().toISOString(),
    createdBy: state.currentUserEmail
  };
  if (!Array.isArray(tracker.rfqs)) tracker.rfqs = [];
  tracker.rfqs.push({
    id: rfqId,
    category: snap.trade.category,
    budgetRowNo: wizard.budgetRowNo,
    subject: snap.subject,
    bodyTemplate: snap.bodyTemplate,
    filledRequirements: wizard.requirements,
    respondByDate: snap.respondByDate,
    daysToRespond: wizard.daysToRespond,
    daysToFollowup: wizard.daysToFollowup,
    tradiesShareLink: snap.tradiesShareLink,
    fileSnapshot: snap.files.map(f => ({ name: f.name, modified: f.lastModifiedDateTime })),
    sowAttached: snap.sowFilename || null,
    sentBy: state.currentUserEmail,
    sentAt: new Date().toISOString(),
    suppliers: sentEntries,
    projectTeamCC: snap.projectTeamCC,
    pickedSupplierId: null,
    status: 'open' // 'open' | 'given_up' | 'picked'
  });
  await uploadJson(siteId, quotePath, 'rfq-tracker.json', tracker);
}

// --------- Step 5: Summary ---------
function renderStep5Summary(root) {
  const result = wizard.snapshot._sendResult;
  const sentCount = result.sentEntries.length;
  const failedCount = result.failed.length;
  let html = `
    <div class="info-card mb-12">
      <div style="font-weight: 600; font-size: 17px; ${sentCount > 0 ? 'color: var(--green);' : 'color: var(--red);'}">
        ${sentCount > 0 ? '✓ Sent ' + sentCount + ' RFQ' + (sentCount === 1 ? '' : 's') : '✗ Nothing Sent'}
      </div>
      ${failedCount > 0 ? `<div class="text-small mt-4" style="color: var(--amber);">${failedCount} Failure${failedCount === 1 ? '' : 's'} Below</div>` : ''}
    </div>
    ${sentCount > 0 ? `
      <div class="section-title">Sent Successfully</div>
      <div class="info-card">
        ${result.sentEntries.map(e => `
          <div class="text-small" style="padding: 4px 0;">
            ✓ ${escapeHtml(e.companyName)} &lt;${escapeHtml(e.email)}&gt;
          </div>`).join('')}
      </div>` : ''}
    ${failedCount > 0 ? `
      <div class="section-title">Failed</div>
      <div class="info-card" style="border-color: var(--red);">
        ${result.failed.map(f => `
          <div style="padding: 6px 0; border-bottom: 1px solid var(--line);">
            <div class="text-small">✗ ${escapeHtml(f.supplier.companyName)} &lt;${escapeHtml(f.supplier.email)}&gt;</div>
            <div style="color: var(--red); font-size: 12px; margin-top: 2px;">${escapeHtml(f.error)}</div>
          </div>`).join('')}
      </div>
      <div class="btn-row mt-12">
        <button class="btn-primary small" id="rfq-retry-failed">Retry Failed Only</button>
      </div>` : ''}
    <div class="btn-row mt-24">
      <button class="btn-primary" id="rfq-done">Done</button>
    </div>
  `;
  root.innerHTML = html;
  document.getElementById('rfq-done').addEventListener('click', () => {
    state.navStack.pop();
    showScreen('job-detail-screen');
    // Reload job detail so the new RFQ entry shows
    import('./jobs.js').then(m => m.loadJobDetail());
  });
  if (failedCount > 0) {
    document.getElementById('rfq-retry-failed').addEventListener('click', async () => {
      // Replace selected suppliers with the failed ones, jump back to send step
      const failedIds = result.failed.map(f => f.supplier.id);
      wizard.selectedSupplierIds = new Set(failedIds);
      wizard.snapshot._sendResult = null;
      wizard.step = 4;
      renderWizard();
    });
  }
}

// Inline-onclick exposure
window.openSendRfq = openSendRfq;
