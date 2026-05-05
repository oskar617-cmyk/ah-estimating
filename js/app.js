// js/app.js
// Entry point. Imports register inline-onclick globals and side-effect setup;
// then we wire up post-auth, initialise the modal, and start the auth flow.
//
// Import order matters here: each module sets `window.<fn>` for inline
// onclick attributes used in index.html. By importing them at boot we
// guarantee those handlers exist before the first user click.

import { initAuth, setOnAuthed } from './auth.js';
import { initModalBackdrop } from './ui.js';
import { loadJobs } from './jobs.js';
import './nav.js';
import './new-job.js';
import './settings.js';
import './catalog.js';
import './companies.js';
import './email-templates.js';

// Wire up the post-sign-in callback (auth doesn't import jobs directly to
// keep auth as a leaf module).
setOnAuthed(() => loadJobs());

// Modal backdrop click-to-dismiss
initModalBackdrop();

// Force-uppercase Job Name as user types (input lives in index.html)
document.addEventListener('DOMContentLoaded', () => {
  const njName = document.getElementById('nj-name');
  if (njName) {
    njName.addEventListener('input', (e) => {
      const start = e.target.selectionStart;
      e.target.value = e.target.value.toUpperCase();
      e.target.setSelectionRange(start, start);
    });
  }
});

// Service Worker for PWA installability
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('service-worker.js').catch(err => console.warn('SW registration failed:', err));
  });
}

// Boot
initAuth();
