// js/catalog.js
// Trades / Suppliers catalog: read from budget Excel, render with nested
// companies, save mappings, and add new items by writing back into the Excel
// template (preserving formulas where possible).

import { CONFIG } from './config.js';
import { state } from './state.js';
import { getAhSiteId, encodeUriPath, readXlsx } from './graph.js';
import { getToken } from './auth.js';
import { loadAppConfig, saveAppConfig, loadSuppliers, saveSuppliers, logAudit } from './audit.js';
import { showToast, showModal, closeModal, confirmModal, escapeHtml } from './ui.js';
import { openCompanyEditor } from './companies.js';

export async function loadCatalogTab() {
  const container = document.getElementById('catalog-list');
  container.innerHTML = '<div class="loading"><div class="spinner"></div><div>Loading...</div></div>';
  try {
    await loadAppConfig();
    await loadSuppliers();
    // First-run OR old-schema detection: re-import to get availableRows populated
    const hasOldSchema = state.appConfig.trades && state.appConfig.trades.some(t => !t.availableRows);
    const needsImport = !state.appConfig.trades || state.appConfig.trades.length === 0 || hasOldSchema;
    if (needsImport) {
      container.innerHTML = '<div class="loading"><div class="spinner"></div><div>Importing From Budget Template...</div></div>';
      await importCatalogFromExcel(false);
    }
    document.getElementById('catalog-search').oninput = renderCatalog;
    renderCatalog();
  } catch (err) {
    console.error('Catalog load error:', err);
    container.innerHTML = `<div class="empty-state"><div style="color: var(--red);">Failed To Load</div><div class="text-small mt-8">${escapeHtml(err.message)}</div><button class="btn-secondary mt-16" onclick="loadCatalogTab()">Try Again</button></div>`;
  }
}

async function importCatalogFromExcel(isReimport) {
  const siteId = await getAhSiteId();
  const buf = await readXlsx(siteId, CONFIG.commonDocsPath, CONFIG.budgetTemplateName);
  const wb = XLSX.read(buf, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  // Column A=Category(0), B=No(1), C=Description(2), D=Type(3)
  const categoriesMap = new Map();
  for (let i = 1; i < rows.length; i++) {
    const cat = rows[i][0]; const no = rows[i][1]; const desc = rows[i][2]; const type = rows[i][3];
    if (cat && typeof cat === 'string') {
      const c = cat.trim();
      if (!categoriesMap.has(c)) categoriesMap.set(c, []);
      if (no != null || (desc && typeof desc === 'string' && desc.trim())) {
        // Strip non-breaking spaces and other whitespace from No.
        const cleanNo = no != null ? String(no).replace(/[\s\xa0]+/g, '').trim() : null;
        categoriesMap.get(c).push({
          no: cleanNo || null,
          description: desc ? String(desc).trim() : null,
          type: type ? String(type).trim() : null,
          rowIdx: i
        });
      }
    }
  }
  const existingByCategory = new Map((state.appConfig.trades || []).map(t => [t.category, t]));
  const newTrades = [];
  let addedCount = 0;
  for (const [cat, items] of categoriesMap) {
    const existing = existingByCategory.get(cat);
    if (existing) {
      existing.availableRows = items;
      // Migrate legacy fields from older builds
      if (existing.availableDescriptions) delete existing.availableDescriptions;
      if (existing.budgetRowDescription && !existing.budgetRowNo) {
        const match = items.find(it => it.description === existing.budgetRowDescription);
        if (match) existing.budgetRowNo = match.no;
        delete existing.budgetRowDescription;
      }
      newTrades.push(existing);
    } else {
      newTrades.push({
        category: cat, budgetRowNo: null,
        daysToRespond: CONFIG.defaultDaysToRespond,
        daysToFollowup: CONFIG.defaultDaysToFollowup,
        sowTemplate: null, emailTemplate: null,
        availableRows: items
      });
      addedCount++;
    }
  }
  // Preserve any custom-added trades that aren't in Excel
  for (const existing of state.appConfig.trades || []) {
    if (!categoriesMap.has(existing.category) && existing.custom) newTrades.push(existing);
  }
  state.appConfig.trades = newTrades;
  await saveAppConfig();
  if (isReimport) {
    showToast(addedCount > 0 ? `Added ${addedCount} New Categor${addedCount === 1 ? 'y' : 'ies'}` : 'No New Categories Found', 'success');
    await logAudit('CATALOG_REIMPORTED', 'Excel template', { totalCount: newTrades.length, newCount: addedCount });
  }
}

export async function reimportCatalogFromExcel() {
  const proceed = await confirmModal(
    'Re-Import From Excel',
    'This rescans the budget template and adds any new categories found. Existing settings (mappings, defaults, custom-added items) are preserved.<br><br>Continue?',
    'Re-Import', 'Cancel'
  );
  if (!proceed) return;
  const container = document.getElementById('catalog-list');
  container.innerHTML = '<div class="loading"><div class="spinner"></div><div>Re-importing...</div></div>';
  try { await importCatalogFromExcel(true); renderCatalog(); }
  catch (err) { console.error(err); showToast('Re-Import Failed', 'error'); renderCatalog(); }
}

export function getCompaniesForTrade(category) {
  return (state.suppliersData.suppliers || []).filter(s => (s.trades || []).includes(category));
}

export function renderCatalog() {
  const container = document.getElementById('catalog-list');
  const search = (document.getElementById('catalog-search').value || '').toLowerCase().trim();
  if (!state.appConfig.trades || state.appConfig.trades.length === 0) {
    container.innerHTML = `<div class="empty-state"><div>No Items Yet</div><div class="text-small mt-8">Click "Re-Import From Excel" or "Add Item" to begin.</div></div>`;
    return;
  }
  let items = [...state.appConfig.trades];
  if (search) {
    items = items.filter(t => {
      if (t.category.toLowerCase().includes(search)) return true;
      const companies = getCompaniesForTrade(t.category);
      return companies.some(c =>
        (c.companyName || '').toLowerCase().includes(search) ||
        (c.contactName || '').toLowerCase().includes(search) ||
        (c.email || '').toLowerCase().includes(search)
      );
    });
  }
  items.sort((a, b) => a.category.localeCompare(b.category));

  const headerHtml = `
    <div class="catalog-header">
      <div class="col-category">Category</div>
      <div>Budget Row</div>
      <div class="col-days">Respond Days</div>
      <div class="col-days">Follow-Up Days</div>
      <div class="col-actions">Actions</div>
    </div>`;

  const itemsHtml = items.map(t => {
    const realIdx = state.appConfig.trades.indexOf(t);
    const companies = getCompaniesForTrade(t.category);
    const expandedClass = state.expandedCatalogItems.has(t.category) ? ' expanded' : '';
    const rows = t.availableRows || [];
    const descOptions = rows.map(r => {
      const label = r.no ? `${r.no} — ${r.description || ''}` : (r.description || '');
      return `<option value="${escapeHtml(r.no || '')}"${t.budgetRowNo === r.no ? ' selected' : ''}>${escapeHtml(label)}</option>`;
    }).join('');
    return `
      <div class="catalog-item${expandedClass}" data-idx="${realIdx}">
        <div class="catalog-row">
          <div class="catalog-name-cell">
            <div class="catalog-toggle" data-toggle="${escapeHtml(t.category)}">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="9 18 15 12 9 6"/>
              </svg>
            </div>
            <div class="catalog-name">${escapeHtml(t.category)}</div>
            <button class="icon-btn-tiny catalog-rename-inline" data-action="rename" title="Rename Category">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>
              </svg>
            </button>
            <div class="catalog-companies-count">${companies.length}</div>
          </div>
          <select class="catalog-budget" title="Budget row to write quote into">
            <option value="">— Select Budget Row —</option>
            ${descOptions}
          </select>
          <input type="number" class="catalog-days-input catalog-days-respond" min="1" max="60" value="${t.daysToRespond || CONFIG.defaultDaysToRespond}" />
          <input type="number" class="catalog-days-input catalog-days-followup" min="1" max="60" value="${t.daysToFollowup || CONFIG.defaultDaysToFollowup}" />
          <div class="catalog-actions">
            <button class="icon-btn-tiny" data-action="save" title="Save">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
                <polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
              </svg>
            </button>
            <button class="icon-btn-tiny danger" data-action="delete" title="Delete">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="3 6 5 6 21 6"/>
                <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
              </svg>
            </button>
          </div>
        </div>
        <div class="catalog-companies">
          ${companies.length === 0
            ? '<div class="text-muted text-small" style="padding: 12px 0;">No companies under this item yet.</div>'
            : companies.map(c => {
                const supIdx = state.suppliersData.suppliers.indexOf(c);
                return `
                  <div class="company-row${c.active === false ? ' inactive' : ''}">
                    <div class="company-info">
                      <div class="company-name">${escapeHtml(c.companyName || '?')}${c.active === false ? ' <span class="trade-tag" style="background: var(--bg-3); color: var(--text-3);">Inactive</span>' : ''}</div>
                      <div class="company-meta">${escapeHtml(c.contactName || '')}${c.contactName && c.email ? ' · ' : ''}${escapeHtml(c.email || '')}${c.phone ? ' · ' + escapeHtml(c.phone) : ''}</div>
                    </div>
                    <div class="company-actions">
                      <button class="icon-btn-tiny" data-company-idx="${supIdx}" data-company-action="edit" title="Edit">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                          <path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>
                        </svg>
                      </button>
                    </div>
                  </div>`;
              }).join('')}
          <button class="add-company-btn" data-add-for="${escapeHtml(t.category)}">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
            Add Company
          </button>
        </div>
      </div>`;
  }).join('');

  container.innerHTML = headerHtml + itemsHtml;
  attachCatalogHandlers();
}

function attachCatalogHandlers() {
  document.querySelectorAll('.catalog-toggle').forEach(el => {
    el.addEventListener('click', (e) => {
      e.stopPropagation();
      const cat = el.dataset.toggle;
      if (state.expandedCatalogItems.has(cat)) state.expandedCatalogItems.delete(cat);
      else state.expandedCatalogItems.add(cat);
      el.closest('.catalog-item').classList.toggle('expanded');
    });
  });
  document.querySelectorAll('.catalog-name-cell').forEach(el => {
    el.addEventListener('click', (e) => {
      if (e.target.closest('.catalog-companies-count')) return;
      const toggleBtn = el.querySelector('.catalog-toggle');
      if (toggleBtn) toggleBtn.click();
    });
  });
  document.querySelectorAll('.catalog-item').forEach(item => {
    const idx = parseInt(item.dataset.idx, 10);
    item.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const action = btn.dataset.action;
        if (action === 'save') saveCatalogItem(idx);
        else if (action === 'rename') renameCatalogItem(idx);
        else if (action === 'delete') deleteCatalogItem(idx);
      });
    });
    item.querySelectorAll('[data-company-action]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const supIdx = parseInt(btn.dataset.companyIdx, 10);
        openCompanyEditor(supIdx, null);
      });
    });
    item.querySelectorAll('[data-add-for]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openCompanyEditor(null, btn.dataset.addFor);
      });
    });
  });
}

async function saveCatalogItem(idx) {
  const item = document.querySelector(`.catalog-item[data-idx="${idx}"]`);
  if (!item) return;
  const t = state.appConfig.trades[idx];
  t.budgetRowNo = item.querySelector('.catalog-budget').value || null;
  t.daysToRespond = parseInt(item.querySelector('.catalog-days-respond').value, 10) || CONFIG.defaultDaysToRespond;
  t.daysToFollowup = parseInt(item.querySelector('.catalog-days-followup').value, 10) || CONFIG.defaultDaysToFollowup;
  try {
    await saveAppConfig();
    await logAudit('CATALOG_ITEM_UPDATED', t.category, { budgetRowNo: t.budgetRowNo, days: `${t.daysToRespond}/${t.daysToFollowup}` });
    showToast(`${t.category} Saved`, 'success');
  } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
}

async function renameCatalogItem(idx) {
  const t = state.appConfig.trades[idx];
  const oldName = t.category;
  showModal(`
    <div class="modal-title">Rename Item</div>
    <div class="modal-body">Renaming will also update all companies that reference this item.</div>
    <div class="form-group">
      <label class="form-label">Item Name</label>
      <input id="rename-input" type="text" value="${escapeHtml(oldName)}" />
    </div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" id="rename-confirm">Rename</button>
    </div>`);
  setTimeout(() => document.getElementById('rename-input').focus(), 50);
  document.getElementById('rename-confirm').addEventListener('click', async () => {
    const newName = document.getElementById('rename-input').value.trim();
    if (!newName) { showToast('Name Required', 'error'); return; }
    if (newName === oldName) { closeModal(); return; }
    if (state.appConfig.trades.some((x, i) => i !== idx && x.category === newName)) {
      showToast('Name Already Used', 'error'); return;
    }
    t.category = newName;
    let updatedCount = 0;
    for (const s of state.suppliersData.suppliers) {
      if (s.trades && s.trades.includes(oldName)) {
        s.trades = s.trades.map(x => x === oldName ? newName : x);
        updatedCount++;
      }
    }
    try {
      await saveAppConfig();
      if (updatedCount > 0) await saveSuppliers();
      await logAudit('CATALOG_ITEM_RENAMED', oldName, { newName, suppliersUpdated: updatedCount });
      closeModal(); renderCatalog();
      showToast('Renamed', 'success');
    } catch (err) { console.error(err); showToast('Rename Failed', 'error'); }
  });
}

async function deleteCatalogItem(idx) {
  const t = state.appConfig.trades[idx];
  const companies = getCompaniesForTrade(t.category);
  let body = `Delete <strong>${escapeHtml(t.category)}</strong>?<br><br>This removes the item and its budget row mapping.`;
  if (companies.length > 0) {
    body += `<br><br><strong>${companies.length}</strong> compan${companies.length === 1 ? 'y has' : 'ies have'} this item assigned. They will keep their other trade assignments but this one will be removed.`;
  }
  const proceed = await confirmModal('Delete Item?', body, 'Delete', 'Cancel');
  if (!proceed) return;
  state.appConfig.trades.splice(idx, 1);
  let companiesUpdated = 0;
  for (const s of state.suppliersData.suppliers) {
    if (s.trades && s.trades.includes(t.category)) {
      s.trades = s.trades.filter(x => x !== t.category);
      companiesUpdated++;
    }
  }
  state.expandedCatalogItems.delete(t.category);
  try {
    await saveAppConfig();
    if (companiesUpdated > 0) await saveSuppliers();
    await logAudit('CATALOG_ITEM_DELETED', t.category, { companiesUpdated });
    renderCatalog();
    showToast('Deleted', 'success');
  } catch (err) { console.error(err); showToast('Delete Failed', 'error'); }
}

export function addNewCatalogItem() {
  const allNos = [];
  for (const t of state.appConfig.trades || []) {
    for (const r of t.availableRows || []) {
      if (r.no) allNos.push({ no: r.no, category: t.category });
    }
  }
  const existingCats = (state.appConfig.trades || []).map(t => t.category);
  const datalistOptions = existingCats.map(c => `<option value="${escapeHtml(c)}">`).join('');

  showModal(`
    <div class="modal-title">Add New Item</div>
    <div class="modal-body">Adds a row to the budget Excel template and creates a catalog entry. If the Category exists, the No. will continue its sequence; otherwise a new Category group is added at the bottom (above any Total row).</div>
    <div class="form-group">
      <label class="form-label">Category</label>
      <input id="new-cat-input" type="text" list="cat-suggestions" placeholder="Type or pick existing" autocomplete="off" />
      <datalist id="cat-suggestions">${datalistOptions}</datalist>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">No.</label>
        <input id="new-no-input" type="text" placeholder="Auto-Predicted" disabled style="opacity: 0.7;" />
        <div class="form-hint">Auto-fills based on Category</div>
      </div>
      <div class="form-group">
        <label class="form-label">Type</label>
        <select id="new-type-input">
          <option value="Material">Material</option>
          <option value="Labour">Labour</option>
          <option value="MatLab">MatLab</option>
          <option value="Equipment">Equipment</option>
        </select>
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Description</label>
      <input id="new-desc-input" type="text" placeholder="e.g. PC Allowance for Pavers" />
    </div>
    <div class="form-error hidden" id="new-cat-err"></div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" id="new-cat-confirm">Add To Excel</button>
    </div>`);

  const catInput = document.getElementById('new-cat-input');
  const noInput = document.getElementById('new-no-input');
  setTimeout(() => catInput.focus(), 50);

  function predictNo() {
    const typedCat = catInput.value.trim();
    if (!typedCat) { noInput.value = ''; return; }
    const exact = existingCats.find(c => c.toLowerCase() === typedCat.toLowerCase());
    const prefix = !exact ? existingCats.find(c => c.toLowerCase().startsWith(typedCat.toLowerCase())) : null;
    const matchedCat = exact || prefix;
    if (matchedCat) {
      const trade = state.appConfig.trades.find(t => t.category === matchedCat);
      const nos = (trade.availableRows || []).map(r => r.no).filter(Boolean);
      if (nos.length === 0) { noInput.value = ''; return; }
      const parsed = nos.map(n => {
        const m = String(n).match(/^(\d+)\.(\d+)$/);
        return m ? { major: parseInt(m[1], 10), minor: parseInt(m[2], 10) } : null;
      }).filter(Boolean);
      if (parsed.length === 0) { noInput.value = ''; return; }
      const major = parsed[0].major;
      const maxMinor = Math.max(...parsed.filter(p => p.major === major).map(p => p.minor));
      noInput.value = `${major}.${maxMinor + 1}`;
    } else {
      let maxMajor = 0;
      for (const item of allNos) {
        const m = String(item.no).match(/^(\d+)\./);
        if (m) maxMajor = Math.max(maxMajor, parseInt(m[1], 10));
      }
      noInput.value = `${maxMajor + 1}.1`;
    }
  }
  catInput.addEventListener('input', predictNo);
  catInput.addEventListener('change', predictNo);

  document.getElementById('new-cat-confirm').addEventListener('click', async () => {
    const errEl = document.getElementById('new-cat-err');
    function showFieldErr(msg) { errEl.textContent = msg; errEl.classList.remove('hidden'); }
    errEl.classList.add('hidden');
    const cat = catInput.value.trim();
    const no = noInput.value.trim();
    const desc = document.getElementById('new-desc-input').value.trim();
    const type = document.getElementById('new-type-input').value;
    if (!cat) { showFieldErr('Category Required'); return; }
    if (!no || !/^\d+\.\d+$/.test(no)) { showFieldErr('No. Must Be Format X.Y (e.g. 37.6)'); return; }
    if (!desc) { showFieldErr('Description Required'); return; }
    for (const t of state.appConfig.trades) {
      for (const r of t.availableRows || []) {
        if (r.no === no) { showFieldErr(`No. ${no} Already Used Under "${t.category}"`); return; }
      }
    }
    const btn = document.getElementById('new-cat-confirm');
    btn.disabled = true;
    btn.innerHTML = '<div class="spinner-sm"></div> Writing To Excel...';
    try {
      await addRowToBudgetTemplate(cat, no, desc, type);
      await importCatalogFromExcel(false);
      await logAudit('CATALOG_ITEM_ADDED', cat, { no, description: desc, type });
      closeModal();
      state.expandedCatalogItems.add(cat);
      renderCatalog();
      showToast('Item Added', 'success');
    } catch (err) {
      console.error('Add item error:', err);
      btn.disabled = false;
      btn.textContent = 'Add To Excel';
      showFieldErr(`Failed: ${err.message}`);
    }
  });
}

// Insert a new row into the budget Excel template, preserving formulas/styling.
// Inserts under the existing Category group if found, otherwise above the Total row at the bottom.
async function addRowToBudgetTemplate(category, no, description, type) {
  const siteId = await getAhSiteId();
  const buf = await readXlsx(siteId, CONFIG.commonDocsPath, CONFIG.budgetTemplateName);
  const wb = XLSX.read(buf, { type: 'array', cellStyles: true, cellFormula: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref']);

  let insertAtRow = -1;
  const catCol = 0;
  let lastRowOfCategory = -1;
  let totalRowIdx = -1;
  for (let r = 1; r <= range.e.r; r++) {
    const cellAddr = XLSX.utils.encode_cell({ c: catCol, r });
    const cell = sheet[cellAddr];
    const val = cell ? cell.v : null;
    if (val && typeof val === 'string') {
      const trimmed = val.trim();
      if (trimmed.toLowerCase() === 'total') totalRowIdx = r;
      if (trimmed === category) lastRowOfCategory = r;
    }
  }
  if (lastRowOfCategory >= 0) insertAtRow = lastRowOfCategory + 1;
  else if (totalRowIdx >= 0) insertAtRow = totalRowIdx;
  else insertAtRow = range.e.r + 1;

  const newCells = {};
  for (let r = range.e.r; r >= insertAtRow; r--) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const fromAddr = XLSX.utils.encode_cell({ c, r });
      const cell = sheet[fromAddr];
      if (!cell) continue;
      const toAddr = XLSX.utils.encode_cell({ c, r: r + 1 });
      const newCell = { ...cell };
      if (newCell.f) newCell.f = shiftFormulaRow(newCell.f, 1);
      newCells[toAddr] = newCell;
      if (!newCells[fromAddr]) newCells[fromAddr] = null;
    }
  }
  for (const [addr, cell] of Object.entries(newCells)) {
    if (cell === null) delete sheet[addr];
    else sheet[addr] = cell;
  }
  sheet[XLSX.utils.encode_cell({ c: 0, r: insertAtRow })] = { t: 's', v: category };
  const noParsed = parseFloat(no);
  if (!isNaN(noParsed) && /^\d+\.\d+$/.test(no)) {
    sheet[XLSX.utils.encode_cell({ c: 1, r: insertAtRow })] = { t: 'n', v: noParsed };
  } else {
    sheet[XLSX.utils.encode_cell({ c: 1, r: insertAtRow })] = { t: 's', v: no };
  }
  sheet[XLSX.utils.encode_cell({ c: 2, r: insertAtRow })] = { t: 's', v: description };
  sheet[XLSX.utils.encode_cell({ c: 3, r: insertAtRow })] = { t: 's', v: type };

  range.e.r += 1;
  sheet['!ref'] = XLSX.utils.encode_range(range);

  if (sheet['!merges']) {
    sheet['!merges'] = sheet['!merges'].map(m => {
      const newM = { s: { ...m.s }, e: { ...m.e } };
      if (newM.s.r >= insertAtRow) newM.s.r += 1;
      if (newM.e.r >= insertAtRow) newM.e.r += 1;
      return newM;
    });
  }

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const blob = new Uint8Array(out);
  const token = await getToken();
  const path = `${CONFIG.commonDocsPath}/${CONFIG.budgetTemplateName}`;
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
    {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      body: blob
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excel upload failed: ${res.status} ${text}`);
  }
}

// Shift row references in a formula by `delta`. Best-effort: doesn't shift
// absolute-row references ($1) and doesn't handle cross-sheet refs.
function shiftFormulaRow(formula, delta) {
  return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (m, colAbs, col, rowAbs, row) => {
    if (rowAbs === '$') return m;
    const newRow = parseInt(row, 10) + delta;
    return `${colAbs}${col}${rowAbs}${newRow}`;
  });
}

// Inline-onclick exposure
window.loadCatalogTab = loadCatalogTab;
window.reimportCatalogFromExcel = reimportCatalogFromExcel;
window.addNewCatalogItem = addNewCatalogItem;
