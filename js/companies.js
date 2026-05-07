// js/companies.js
// Modal editor for adding / editing / deleting company contacts under
// trade/supplier categories.

import { state } from './state.js';
import { saveSuppliers, logAudit } from './audit.js';
import { showToast, showModal, closeModal, confirmModal, escapeHtml } from './ui.js';
import { renderCatalog } from './catalog.js';

export function openCompanyEditor(idx, presetTrade, onSaved) {
  const isNew = idx === null;
  const sup = isNew
    ? { id: 'sup-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8), companyName: '', contactName: '', email: '', phone: '', trades: presetTrade ? [presetTrade] : [], notes: '', active: true }
    : { ...state.suppliersData.suppliers[idx] };
  state.editingSupplier = { idx, supplier: sup, onSaved: onSaved || null };
  state.supplierMultiSelectTrades = [...(sup.trades || [])];
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" title="Close" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">${isNew ? 'Add Company' : 'Edit Company'}</div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Company Name</label>
        <input id="sup-company" type="text" value="${escapeHtml(sup.companyName)}" placeholder="Bob's Concreting" />
      </div>
      <div class="form-group">
        <label class="form-label">Contact First Name</label>
        <input id="sup-contact" type="text" value="${escapeHtml(sup.contactName)}" placeholder="Bob" />
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Email</label>
        <input id="sup-email" type="email" value="${escapeHtml(sup.email)}" placeholder="bob@bobsconcreting.com.au" />
      </div>
      <div class="form-group">
        <label class="form-label">Phone (Optional)</label>
        <input id="sup-phone" type="text" value="${escapeHtml(sup.phone)}" />
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Trades / Suppliers Covered</label>
      <div class="multi-chips" id="sup-trades-chips"></div>
      <div class="mt-4" id="sup-trades-picker" style="display:flex;flex-wrap:wrap;gap:5px;max-height:100px;overflow-y:auto;padding:6px;background:var(--bg-3);border:1px solid var(--line);border-radius:8px;"></div>
    </div>
    <div class="form-row" style="align-items:flex-start;">
      <div class="form-group" style="flex:2;">
        <label class="form-label">Notes (Optional)</label>
        <textarea id="sup-notes" rows="2">${escapeHtml(sup.notes || '')}</textarea>
      </div>
      <div class="form-group" style="flex:1;">
        <label class="form-label">Status</label>
        <label class="flex-row" style="gap:10px;cursor:pointer;padding:6px 0;" onclick="toggleSupActive()">
          <span class="toggle${sup.active !== false ? ' on' : ''}" id="sup-active-toggle"></span>
          <span class="text-small">Active</span>
        </label>
      </div>
    </div>
    <div class="modal-actions">
      ${isNew ? '' : '<button class="btn-danger small" style="margin-right:auto;" onclick="deleteCompany()">Delete</button>'}
      <button class="btn-secondary" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" onclick="saveCompany()">Save</button>
    </div>`);
  renderCompanyTradesPicker();
  renderCompanyTradeChips();
}

function renderCompanyTradesPicker() {
  const picker = document.getElementById('sup-trades-picker');
  const allCategories = (state.appConfig.trades || []).map(t => t.category).sort();
  picker.innerHTML = allCategories.map(c => {
    const selected = state.supplierMultiSelectTrades.includes(c);
    return `<button type="button" class="${selected ? 'multi-chip' : 'btn-secondary small'}" data-cat="${escapeHtml(c)}" style="${selected ? '' : 'padding:4px 10px;font-size:12px;border-radius:14px;'}">${selected ? '<svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>' : ''}${escapeHtml(c)}</button>`;
  }).join('');
  picker.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const c = btn.dataset.cat;
      const i = state.supplierMultiSelectTrades.indexOf(c);
      if (i >= 0) state.supplierMultiSelectTrades.splice(i, 1);
      else state.supplierMultiSelectTrades.push(c);
      renderCompanyTradesPicker(); renderCompanyTradeChips();
    });
  });
}

function renderCompanyTradeChips() {
  const chipsEl = document.getElementById('sup-trades-chips');
  if (!chipsEl) return;
  if (state.supplierMultiSelectTrades.length === 0) {
    chipsEl.innerHTML = '<span class="text-muted text-small">No items selected</span>';
    return;
  }
  chipsEl.innerHTML = state.supplierMultiSelectTrades.map((t) => `
    <span class="multi-chip">
      ${escapeHtml(t)}
      <span class="x" data-cat="${escapeHtml(t)}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></span>
    </span>`).join('');
  chipsEl.querySelectorAll('.x').forEach(x => x.addEventListener('click', () => {
    const c = x.dataset.cat;
    const i = state.supplierMultiSelectTrades.indexOf(c);
    if (i >= 0) state.supplierMultiSelectTrades.splice(i, 1);
    renderCompanyTradesPicker(); renderCompanyTradeChips();
  }));
}

function toggleSupActive() {
  const tog = document.getElementById('sup-active-toggle');
  tog.classList.toggle('on');
}

async function saveCompany() {
  const sup = state.editingSupplier.supplier;
  sup.companyName = document.getElementById('sup-company').value.trim();
  sup.contactName = document.getElementById('sup-contact').value.trim();
  sup.email = document.getElementById('sup-email').value.trim().toLowerCase();
  sup.phone = document.getElementById('sup-phone').value.trim();
  sup.notes = document.getElementById('sup-notes').value.trim();
  sup.trades = [...state.supplierMultiSelectTrades];
  sup.active = document.getElementById('sup-active-toggle').classList.contains('on');
  if (!sup.companyName) { showToast('Company Name Required', 'error'); return; }
  if (!sup.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(sup.email)) { showToast('Valid Email Required', 'error'); return; }
  if (sup.trades.length === 0) { showToast('At Least One Item Required', 'error'); return; }
  if (state.editingSupplier.idx === null) state.suppliersData.suppliers.push(sup);
  else state.suppliersData.suppliers[state.editingSupplier.idx] = sup;
  try {
    await saveSuppliers();
    await logAudit(state.editingSupplier.idx === null ? 'COMPANY_ADDED' : 'COMPANY_UPDATED', sup.companyName, { email: sup.email, trades: sup.trades });
    const cb = state.editingSupplier.onSaved;
    closeModal();
    if (cb) {
      // Caller (e.g. send-rfq wizard) handles its own re-render
      cb(sup);
    } else {
      // Default: refresh the Settings catalog view
      sup.trades.forEach(t => state.expandedCatalogItems.add(t));
      renderCatalog();
    }
    showToast('Company Saved', 'success');
  } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
}

async function deleteCompany() {
  const proceed = await confirmModal('Delete Company?', `Delete <strong>${escapeHtml(state.editingSupplier.supplier.companyName)}</strong>? This cannot be undone.<br><br>Tip: marking as Inactive instead keeps history.`, 'Delete', 'Cancel');
  if (!proceed) return;
  const removed = state.suppliersData.suppliers.splice(state.editingSupplier.idx, 1)[0];
  try {
    await saveSuppliers();
    await logAudit('COMPANY_DELETED', removed.companyName, { email: removed.email });
    const cb = state.editingSupplier.onSaved;
    closeModal();
    if (cb) cb(null); else renderCatalog();
    showToast('Company Deleted', 'success');
  } catch (err) { console.error(err); showToast('Delete Failed', 'error'); }
}

// Inline-onclick exposure (modal HTML uses these handlers)
window.toggleSupActive = toggleSupActive;
window.saveCompany = saveCompany;
window.deleteCompany = deleteCompany;
