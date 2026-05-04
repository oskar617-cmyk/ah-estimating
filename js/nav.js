// js/nav.js
// Screen routing and navigation back-stack.
// Other modules call navigate() and goBack(); jobs.js wires up the
// "refresh on return to jobs" behaviour via a callback to keep nav decoupled.

import { state } from './state.js';

let onReturnToJobs = null;

export function setOnReturnToJobs(fn) { onReturnToJobs = fn; }

export function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

export function navigate(screen, navState) {
  state.navStack.push({ screen, state: navState });
  showScreen(screen);
}

export function goBack() {
  if (state.navStack.length > 1) {
    state.navStack.pop();
    const top = state.navStack[state.navStack.length - 1];
    showScreen(top.screen);
    if (top.screen === 'jobs-screen' && onReturnToJobs) onReturnToJobs();
  }
}

// Make goBack accessible to inline onclick attributes in HTML.
window.goBack = goBack;
