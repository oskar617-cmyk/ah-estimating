// js/pending-review.js
// "Pending Review" entries are quote auto-writes to the budget Excel that
// haven't yet been confirmed/edited/rejected by the user. They live on the
// Job Detail screen at the top — scoped to whichever job is currently open.
//
// Source of truth: each job's rfq-tracker.json carries a `pendingReviews`
// array. We append on auto-write (in inbox.js) and consume when the user
// clicks Confirm/Edit/Reject.
//
// Shape of a pending entry:
//   {
//     id,                 // unique
//     createdAt,
//     supplierId, supplierCompany, supplierEmail,
//     rfqId, rfqCategory,
//     budgetRowNo,
//     amount,             // number we wrote
//     currency,           // 'AUD'
//     pdfFilename,        // saved PDF in Quote/ folder, or null
//     emailSubject,
//     status: 'pending' | 'confirmed' | 'rejected',
//   }

import { state } from './state.js';
import { CONFIG } from './config.js';
import {
  getAhSiteId, encodeUriPath, readJson, uploadJson, readXlsx
} from './graph.js';
import { getToken } from './auth.js';
import { logAudit, writeTracker } from './audit.js';
import { recordReaction } from './decision-log.js';
import { showToast, showModal, closeModal, confirmModal, escapeHtml } from './ui.js';

// Render section into the job detail screen. Caller is jobs.js.
// Returns HTML string to inject; caller wires up handlers via attachHandlers.
export function renderPendingReviewSection(tracker) {
  const pending = (tracker && tracker.pendingReviews || []).filter(p => p.status === 'pending');
  if (!pending.length) return '';
  return `
    <div class="section-title">Pending Review <span class="text-muted text-small" style="font-weight:normal;text-transform:none;letter-spacing:0;">(${pending.length} New)</span></div>
    <div class="pending-review-grid">
      <div class="pending-review-list" id="pending-review-list">
        ${pending.map(p => `
          <div class="pending-review-row" data-id="${escapeHtml(p.id)}">
            <div class="pending-review-main">
              <div class="pending-review-title">${escapeHtml(p.supplierCompany || '?')} <span class="text-muted text-small">· ${escapeHtml(p.rfqCategory || '')}</span></div>
              <div class="pending-review-amount">${p.amount != null ? '$' + Number(p.amount).toLocaleString('en-AU') : '<span style="color:var(--amber);">TBC</span>'}</div>
              <div class="pending-review-meta text-small text-muted">→ Budget row ${escapeHtml(p.budgetRowNo || '?')}${p.pdfFilename ? ' · ' + escapeHtml(p.pdfFilename) : ''}</div>
            </div>
            <div class="pending-review-actions">
              <button class="btn-primary small" data-action="confirm">Confirm</button>
              <button class="btn-secondary small" data-action="edit">Edit</button>
              <button class="btn-danger small" data-action="reject">Reject</button>
            </div>
          </div>
        `).join('')}
      </div>
      <div class="pending-review-preview" id="pending-review-preview">
        <div class="text-muted text-small" style="text-align:center;padding:40px 16px;">Tap an entry to preview the saved PDF.</div>
      </div>
    </div>
  `;
}

// Wire up handlers after rendering. Caller passes the tracker reference and
// a refresh callback that re-renders the job detail (so confirm/reject
// actions update the UI immediately).
export function attachPendingReviewHandlers(tracker, jobFolderName, onChange) {
  document.querySelectorAll('.pending-review-row').forEach(row => {
    row.addEventListener('click', (e) => {
      // Don't preview when clicking action buttons
      if (e.target.closest('button')) return;
      const id = row.dataset.id;
      const entry = (tracker.pendingReviews || []).find(p => p.id === id);
      if (entry) showPreview(entry, jobFolderName);
      document.querySelectorAll('.pending-review-row.active').forEach(r => r.classList.remove('active'));
      row.classList.add('active');
    });
    row.querySelectorAll('button[data-action]').forEach(btn => {
      btn.addEventListener('click', async (e) => {
        e.stopPropagation();
        const id = row.dataset.id;
        const action = btn.dataset.action;
        const entry = (tracker.pendingReviews || []).find(p => p.id === id);
        if (!entry) return;
        if (action === 'confirm') return doConfirm(entry, jobFolderName, tracker, onChange);
        if (action === 'edit') return doEdit(entry, jobFolderName, tracker, onChange);
        if (action === 'reject') return doReject(entry, jobFolderName, tracker, onChange);
      });
    });
  });
}

async function showPreview(entry, jobFolderName) {
  const previewEl = document.getElementById('pending-review-preview');
  if (!previewEl) return;
  if (!entry.pdfFilename) {
    previewEl.innerHTML = `
      <div style="padding:24px;text-align:center;">
        <div class="text-muted">No PDF attached.</div>
        <div class="text-small mt-8">Quote came in the email body — see the email in Outlook.</div>
      </div>`;
    return;
  }
  // Load PDF via SharePoint download URL (browser handles the rendering)
  previewEl.innerHTML = '<div class="loading"><div class="spinner"></div><div>Loading PDF...</div></div>';
  try {
    const siteId = await getAhSiteId();
    const path = `${jobFolderName}/Quote/${entry.pdfFilename}`;
    const token = await getToken();
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
      { headers: { 'Authorization': `Bearer ${token}` } }
    );
    if (!res.ok) throw new Error(`Load failed: ${res.status}`);
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    previewEl.innerHTML = `<embed src="${url}" type="application/pdf" style="width:100%;height:100%;border-radius:8px;" />`;
  } catch (err) {
    console.error(err);
    previewEl.innerHTML = `<div class="text-small" style="padding:24px;color:var(--red);">Could Not Load PDF: ${escapeHtml(err.message)}</div>`;
  }
}

async function doConfirm(entry, jobFolderName, tracker, onChange) {
  entry.status = 'confirmed';
  entry.confirmedAt = new Date().toISOString();
  entry.confirmedBy = state.currentUserEmail;
  try {
    await persistTracker(jobFolderName, tracker);
    await logAudit('QUOTE_CONFIRMED', `${entry.supplierCompany} / ${entry.rfqCategory}`, {
      amount: entry.amount, budgetRowNo: entry.budgetRowNo
    });
    // Record the reaction against the originating Gemini decisions so the
    // export later shows "Gemini said X, user confirmed".
    if (entry.classifyDecisionId) {
      recordReaction(entry.classifyDecisionId, 'confirmed').catch(e => console.warn('recordReaction(classify):', e));
    }
    if (entry.amountDecisionId) {
      recordReaction(entry.amountDecisionId, 'confirmed').catch(e => console.warn('recordReaction(amount):', e));
    }
    showToast('Confirmed', 'success');
    if (onChange) onChange();
  } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
}

async function doEdit(entry, jobFolderName, tracker, onChange) {
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">Edit Quote Amount</div>
    <div class="modal-body text-small text-muted">${escapeHtml(entry.supplierCompany)} · ${escapeHtml(entry.rfqCategory || '')}</div>
    <div class="form-group">
      <label class="form-label">Amount (AUD)</label>
      <input id="edit-amount" type="number" step="0.01" value="${entry.amount != null ? entry.amount : ''}" />
      <div class="form-hint">The budget Excel will be re-written with this amount.</div>
    </div>
    <div class="form-group">
      <label class="form-label">Why edited? (Optional, helps tune AI)</label>
      <input id="edit-why" type="text" placeholder="e.g. AI missed the GST line" />
    </div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" id="edit-amount-save">Save</button>
    </div>
  `);
  document.getElementById('edit-amount-save').addEventListener('click', async () => {
    const v = parseFloat(document.getElementById('edit-amount').value);
    const why = (document.getElementById('edit-why').value || '').trim();
    if (isNaN(v) || v < 0) { showToast('Invalid Amount', 'error'); return; }
    try {
      const originalAmount = entry.amount;
      await rewriteBudgetCell(jobFolderName, entry, v);
      entry.amount = v;
      entry.status = 'confirmed';
      entry.confirmedAt = new Date().toISOString();
      entry.confirmedBy = state.currentUserEmail;
      entry.editedFromOriginal = true;
      await persistTracker(jobFolderName, tracker);
      await logAudit('QUOTE_EDITED', `${entry.supplierCompany} / ${entry.rfqCategory}`, {
        amount: v, budgetRowNo: entry.budgetRowNo
      });
      // Record both reactions: the amount was edited, classification was implicitly confirmed.
      if (entry.amountDecisionId) {
        recordReaction(entry.amountDecisionId, 'edited',
          { from: originalAmount, to: v }, why || null
        ).catch(e => console.warn('recordReaction(amount/edited):', e));
      }
      if (entry.classifyDecisionId) {
        recordReaction(entry.classifyDecisionId, 'confirmed').catch(e => console.warn('recordReaction(classify):', e));
      }
      closeModal();
      showToast('Saved', 'success');
      if (onChange) onChange();
    } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
  });
}

async function doReject(entry, jobFolderName, tracker, onChange) {
  const proceed = await confirmModal(
    'Reject Quote?',
    `Reject the quote from <strong>${escapeHtml(entry.supplierCompany)}</strong>? This will remove the auto-written amount from budget row ${escapeHtml(entry.budgetRowNo || '?')}.`,
    'Reject', 'Cancel'
  );
  if (!proceed) return;
  try {
    await rewriteBudgetCell(jobFolderName, entry, null);  // clear cell
    entry.status = 'rejected';
    entry.rejectedAt = new Date().toISOString();
    entry.rejectedBy = state.currentUserEmail;
    await persistTracker(jobFolderName, tracker);
    await logAudit('QUOTE_REJECTED', `${entry.supplierCompany} / ${entry.rfqCategory}`, {
      budgetRowNo: entry.budgetRowNo
    });
    // Reject = both the classify and the amount were wrong.
    if (entry.classifyDecisionId) {
      recordReaction(entry.classifyDecisionId, 'rejected').catch(e => console.warn('recordReaction(classify):', e));
    }
    if (entry.amountDecisionId) {
      recordReaction(entry.amountDecisionId, 'rejected').catch(e => console.warn('recordReaction(amount):', e));
    }
    showToast('Rejected', 'success');
    if (onChange) onChange();
  } catch (err) { console.error(err); showToast('Save Failed', 'error'); }
}

async function persistTracker(jobFolderName, tracker) {
  await writeTracker(jobFolderName, tracker);
}

// Re-write the budget Excel cell. If amount is null, the slot is cleared.
// We find the supplier-name + amount columns by scanning the first row for
// "Name 1" / "Quote 1" etc. and pick the one we originally wrote into.
async function rewriteBudgetCell(jobFolderName, entry, newAmount) {
  const siteId = await getAhSiteId();
  // Locate the budget Excel — name pattern is "0 Budget Control [JobName].xlsx"
  // We don't know the JobName from the entry; derive from jobFolderName.
  // jobFolderName format: "[code] [name] Site Docs - [address]"
  const m = jobFolderName.match(CONFIG.jobFolderPattern);
  const jobName = m ? m[2].trim() : '';
  const filename = `0 Budget Control ${jobName}.xlsx`;
  const buf = await readXlsx(siteId, `${jobFolderName}/Quote`, filename);
  const wb = XLSX.read(buf, { type: 'array', cellStyles: true, cellFormula: true });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref']);

  // Find header row's "Name N" / "Quote N" columns
  const slots = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const headerCell = sheet[XLSX.utils.encode_cell({ c, r: 0 })];
    const v = headerCell ? String(headerCell.v || '').trim() : '';
    let m1 = v.match(/^Name\s*(\d+)$/i);
    if (m1) slots.push({ idx: parseInt(m1[1], 10), nameCol: c });
    let m2 = v.match(/^Quote\s*(\d+)$/i);
    if (m2) {
      const slot = slots.find(s => s.idx === parseInt(m2[1], 10));
      if (slot) slot.quoteCol = c;
    }
  }
  // Find the row matching this entry's budgetRowNo
  const targetNo = String(entry.budgetRowNo || '').replace(/[\s\xa0]+/g, '').trim();
  let targetRow = -1;
  for (let r = 1; r <= range.e.r; r++) {
    const noCell = sheet[XLSX.utils.encode_cell({ c: 1, r })];
    const noVal = noCell ? String(noCell.v || '').replace(/[\s\xa0]+/g, '').trim() : '';
    if (noVal === targetNo) { targetRow = r; break; }
  }
  if (targetRow < 0) throw new Error(`Budget row ${targetNo} not found`);

  // Find which slot was written for this supplier — match by name.
  // (We saved the slot index on the entry when first written.)
  const slot = slots.find(s => s.idx === entry.budgetSlotIndex);
  if (!slot) throw new Error(`Budget slot ${entry.budgetSlotIndex} not found in Excel`);

  if (newAmount == null) {
    delete sheet[XLSX.utils.encode_cell({ c: slot.nameCol, r: targetRow })];
    delete sheet[XLSX.utils.encode_cell({ c: slot.quoteCol, r: targetRow })];
  } else {
    sheet[XLSX.utils.encode_cell({ c: slot.nameCol, r: targetRow })] = { t: 's', v: entry.supplierCompany || '' };
    sheet[XLSX.utils.encode_cell({ c: slot.quoteCol, r: targetRow })] = { t: 'n', v: newAmount };
  }

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
  const blob = new Uint8Array(out);
  const token = await getToken();
  const path = `${jobFolderName}/Quote/${filename}`;
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
    {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body: blob
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excel write failed: ${res.status} ${text}`);
  }
}
