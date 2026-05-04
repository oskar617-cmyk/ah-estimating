// js/ui.js
// Reusable UI helpers: toast, modal, confirmation dialog, HTML escaping.

export function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = '';
  if (type) t.classList.add(type);
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 3000);
}

export function showModal(html) {
  document.getElementById('modal-content').innerHTML = html;
  document.getElementById('modal-backdrop').classList.add('active');
}

export function closeModal() {
  document.getElementById('modal-backdrop').classList.remove('active');
}

export function confirmModal(title, body, okLabel, cancelLabel) {
  return new Promise(resolve => {
    showModal(`
      <div class="modal-title">${escapeHtml(title)}</div>
      <div class="modal-body">${body}</div>
      <div class="modal-actions">
        <button class="btn-secondary" id="cm-cancel">${escapeHtml(cancelLabel || 'Cancel')}</button>
        <button class="btn-primary" id="cm-ok">${escapeHtml(okLabel || 'OK')}</button>
      </div>`);
    document.getElementById('cm-ok').addEventListener('click', () => { closeModal(); resolve(true); });
    document.getElementById('cm-cancel').addEventListener('click', () => { closeModal(); resolve(false); });
  });
}

export function escapeHtml(s) {
  if (s == null) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Wire up the modal backdrop click-to-dismiss (called once at boot)
export function initModalBackdrop() {
  document.getElementById('modal-backdrop').addEventListener('click', (e) => {
    if (e.target.id === 'modal-backdrop') closeModal();
  });
}
