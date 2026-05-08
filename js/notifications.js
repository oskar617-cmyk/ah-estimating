// js/notifications.js
// Bell-icon notification panel.
//
// Notification shape (created by inbox.js):
//   {
//     id,              // matches messageId for dedupe
//     createdAt,       // ISO when first added
//     read: false,
//     classification,  // 'Quote' | 'Question' | 'Suspicious' | 'OOO' | 'Decline' | 'Unrelated' | 'Pending Match'
//     subject,
//     fromName, fromEmail,
//     receivedAt,      // ISO from message
//     bodyPreview,
//     bodyHtml,        // full body when fetched
//     // Match metadata (null if Tier 3 / Pending Match)
//     job: { folderName, code, name, address } | null,
//     rfqId, rfqCategory,
//     supplier: { id, companyName, contactName, email } | null,
//     budgetRowNo,
//     // Quote-specific:
//     extractedAmount: number | null,
//     extractedCurrency: 'AUD' | ...,
//     savedAttachments: [{ name, path }],   // SharePoint paths after save
//     // Misc:
//     tier,            // matching tier 1/2/3
//     evidence,        // optional: filename or other hint
//     warning          // optional warning for the user (string)
//   }

import { state } from './state.js';
import { showModal, closeModal, escapeHtml, showToast } from './ui.js';

// ----- Public API -----

// Add a notification (or merge if id already present). Updates UI badge.
export function addNotification(n) {
  // Dedupe by id — replace if exists (so we can upgrade Pending Match
  // notifications to a real classification when the user manually matches).
  const existing = state.notifications.findIndex(x => x.id === n.id);
  if (existing >= 0) state.notifications[existing] = { ...state.notifications[existing], ...n };
  else state.notifications.unshift(n);
  refreshBellBadge();
  // If panel is open, re-render
  if (state.notificationPanelOpen) renderPanel();
}

// Remove a notification (e.g. user dismissed it).
export function removeNotification(id) {
  state.notifications = state.notifications.filter(n => n.id !== id);
  refreshBellBadge();
  if (state.notificationPanelOpen) renderPanel();
}

// Mark as read.
export function markRead(id) {
  const n = state.notifications.find(x => x.id === id);
  if (n) { n.read = true; refreshBellBadge(); }
}

// ----- Bell badge -----

export function refreshBellBadge() {
  const bell = document.getElementById('bell-btn');
  if (!bell) return;
  // Unread count excludes Out-of-Office (low-priority — user said don't show).
  const unread = state.notifications.filter(n =>
    !n.read && n.classification !== 'Out-of-Office'
  ).length;
  let badge = bell.querySelector('.bell-badge');
  if (unread > 0) {
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'bell-badge';
      bell.appendChild(badge);
    }
    badge.textContent = unread > 99 ? '99+' : String(unread);
  } else if (badge) {
    badge.remove();
  }
}

// ----- Panel -----

export function openPanel() {
  state.notificationPanelOpen = true;
  document.getElementById('notif-panel').classList.add('open');
  document.getElementById('notif-overlay').classList.add('open');
  renderPanel();
}

export function closePanel() {
  state.notificationPanelOpen = false;
  document.getElementById('notif-panel').classList.remove('open');
  document.getElementById('notif-overlay').classList.remove('open');
}

export function togglePanel() {
  if (state.notificationPanelOpen) closePanel();
  else openPanel();
}

function renderPanel() {
  const list = document.getElementById('notif-list');
  if (!list) return;
  if (!state.notifications.length) {
    list.innerHTML = `
      <div class="empty-state" style="padding:40px 16px;">
        <div>Nothing New</div>
        <div class="text-small mt-8">Replies to your RFQs will appear here.</div>
      </div>`;
    return;
  }
  // Sort: unread first, then newest
  const sorted = [...state.notifications].sort((a, b) => {
    if (a.read !== b.read) return a.read ? 1 : -1;
    return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
  });
  list.innerHTML = sorted.map(n => renderItem(n)).join('');
  list.querySelectorAll('.notif-item').forEach(el => {
    el.addEventListener('click', () => {
      const id = el.dataset.id;
      const n = state.notifications.find(x => x.id === id);
      if (n) onClickNotification(n);
    });
  });
}

function renderItem(n) {
  const cls = `notif-classification notif-cls-${(n.classification || '').toLowerCase().replace(/\s+/g, '-')}`;
  const amount = n.extractedAmount != null
    ? `<div class="notif-amount">$${formatAmount(n.extractedAmount)}</div>`
    : '';
  const subline = n.supplier
    ? `${escapeHtml(n.supplier.companyName)} · ${escapeHtml(n.rfqCategory || '')}`
    : (n.classification === 'Unrelated' || n.tier === 3
        ? 'Could not auto-match — tap to assign'
        : 'Unmatched');
  const job = n.job ? `<div class="notif-job">${escapeHtml(n.job.code)} ${escapeHtml(n.job.name)}</div>` : '';
  return `
    <div class="notif-item ${n.read ? 'read' : ''}" data-id="${escapeHtml(n.id)}">
      <div class="notif-row1">
        <span class="${cls}">${escapeHtml(n.classification || '?')}</span>
        ${job}
        <span class="notif-time">${formatRelative(n.receivedAt || n.createdAt)}</span>
      </div>
      <div class="notif-subline">${subline}</div>
      ${amount}
      <div class="notif-subject">${escapeHtml(n.subject || '(no subject)')}</div>
    </div>`;
}

function formatAmount(n) {
  if (n == null) return '';
  return Number(n).toLocaleString('en-AU', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function formatRelative(iso) {
  if (!iso) return '';
  const t = new Date(iso).getTime();
  const diff = Date.now() - t;
  const min = Math.round(diff / 60000);
  if (min < 1) return 'just now';
  if (min < 60) return `${min}m ago`;
  const hr = Math.round(min / 60);
  if (hr < 24) return `${hr}h ago`;
  const d = Math.round(hr / 24);
  if (d < 7) return `${d}d ago`;
  return new Date(iso).toLocaleDateString();
}

// ----- Click routing per classification -----

function onClickNotification(n) {
  markRead(n.id);
  refreshBellBadge();
  // Re-render to grey out (read-state class change)
  if (state.notificationPanelOpen) renderPanel();
  switch (n.classification) {
    case 'Quote': return openQuoteModal(n);
    case 'Question': return openQuestionModal(n);
    case 'Suspicious': return openSuspiciousModal(n);
    case 'Decline': return openSimpleModal(n, 'Decline', 'Supplier has declined.');
    case 'Out-of-Office': return openSimpleModal(n, 'Out-Of-Office', 'Auto-reply — usually safe to ignore.');
    case 'Unrelated':
    case 'Pending Match':
      return openManualMatchModal(n);
    default:
      return openSimpleModal(n, n.classification || 'Email', '');
  }
}

function openQuoteModal(n) {
  const amount = n.extractedAmount != null ? `$${formatAmount(n.extractedAmount)}` : '<span style="color:var(--amber);">TBC</span>';
  const supplier = n.supplier ? `${escapeHtml(n.supplier.companyName)}` : '(unknown supplier)';
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">Quote Received</div>
    <div class="modal-body">
      <div><strong>${supplier}</strong> · ${escapeHtml(n.rfqCategory || '')}</div>
      <div class="text-small text-muted mt-4">${escapeHtml(n.job ? n.job.code + ' ' + n.job.name : '')}</div>
      <div style="font-size:24px;font-weight:600;margin:16px 0;">${amount}</div>
      <div class="text-small text-muted">Auto-written to budget row <strong>${escapeHtml(n.budgetRowNo || '?')}</strong>. Review the entry on the job page to confirm.</div>
      ${n.warning ? `<div class="form-error mt-12" style="display:block;">${escapeHtml(n.warning)}</div>` : ''}
    </div>
    <div class="email-preview mt-12" style="max-height:240px;">${n.bodyHtml || escapeHtml(n.bodyPreview || '')}</div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Close</button>
    </div>
  `);
}

function openQuestionModal(n) {
  const supplier = n.supplier ? escapeHtml(n.supplier.companyName) : '(unknown)';
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">Question From ${supplier}</div>
    <div class="modal-body text-small text-muted">${escapeHtml(n.rfqCategory || '')} · ${escapeHtml(n.job ? n.job.code + ' ' + n.job.name : '')}</div>
    <div class="email-preview mt-12" style="max-height:360px;">${n.bodyHtml || escapeHtml(n.bodyPreview || '')}</div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Close</button>
      ${n.fromEmail ? `<button class="btn-primary" onclick="window.open('mailto:${encodeURIComponent(n.fromEmail)}?subject=${encodeURIComponent('RE: ' + (n.subject || ''))}','_blank')">Reply</button>` : ''}
    </div>
  `);
}

function openSuspiciousModal(n) {
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">⚠ Suspicious Email</div>
    <div class="form-error" style="display:block;">AI flagged this reply as unusual. Inspect carefully before acting.</div>
    <div class="modal-body text-small text-muted mt-12">From: ${escapeHtml(n.fromName || '')} &lt;${escapeHtml(n.fromEmail || '')}&gt;</div>
    <div class="email-preview mt-12" style="max-height:300px;">${n.bodyHtml || escapeHtml(n.bodyPreview || '')}</div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Close</button>
    </div>
  `);
}

function openSimpleModal(n, title, body) {
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">${escapeHtml(title)}</div>
    ${body ? `<div class="modal-body text-small">${escapeHtml(body)}</div>` : ''}
    <div class="email-preview mt-12" style="max-height:360px;">${n.bodyHtml || escapeHtml(n.bodyPreview || '')}</div>
    <div class="modal-actions"><button class="btn-secondary" onclick="closeModal()">Close</button></div>
  `);
}

// Manual matching: surface a list of all open RFQs across all jobs and let
// the user pick one. After picking, the inbox processor re-runs the rest
// of the per-message logic (download attachments, save to job folder,
// update tracker, file mail).
async function openManualMatchModal(n) {
  // Lazy-import to avoid cycle with inbox.js
  const { processNotificationManualMatch, listAllOpenRfqs } = await import('./inbox.js');
  const choices = await listAllOpenRfqs();
  showModal(`
    <button class="modal-close-x" onclick="closeModal()" type="button">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" width="16" height="16"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
    </button>
    <div class="modal-title">Match To An RFQ</div>
    <div class="modal-body text-small text-muted">From <strong>${escapeHtml(n.fromName || n.fromEmail || '')}</strong>. Pick which RFQ this reply belongs to.</div>
    <div class="email-preview mt-12" style="max-height:200px;">${n.bodyHtml || escapeHtml(n.bodyPreview || '')}</div>
    <div class="form-group mt-16">
      <label class="form-label">RFQ</label>
      <select id="manual-match-rfq">
        <option value="">— Select RFQ —</option>
        ${choices.map(c => `<option value="${escapeHtml(c.key)}">${escapeHtml(c.label)}</option>`).join('')}
      </select>
      <div class="form-hint">RFQs are listed newest-first.</div>
    </div>
    <div class="modal-actions">
      <button class="btn-secondary" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" id="manual-match-confirm">Match</button>
    </div>
  `);
  document.getElementById('manual-match-confirm').addEventListener('click', async () => {
    const key = document.getElementById('manual-match-rfq').value;
    if (!key) { showToast('Pick An RFQ First', 'error'); return; }
    closeModal();
    try {
      await processNotificationManualMatch(n.id, key);
      showToast('Matched', 'success');
    } catch (err) { console.error(err); showToast('Match Failed', 'error'); }
  });
}

// ----- Boot -----

export function initNotificationPanel() {
  const bell = document.getElementById('bell-btn');
  if (bell) bell.addEventListener('click', togglePanel);
  const closeBtn = document.getElementById('notif-close');
  if (closeBtn) closeBtn.addEventListener('click', closePanel);
  const overlay = document.getElementById('notif-overlay');
  if (overlay) overlay.addEventListener('click', closePanel);
  refreshBellBadge();
}

window.togglePanel = togglePanel;
window.closeNotifPanel = closePanel;
