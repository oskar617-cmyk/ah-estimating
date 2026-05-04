// js/auth.js
// MSAL authentication: init, login, logout, token acquisition.
// onSignedIn() is the post-auth dispatcher — it sets up admin status,
// the user chip, and triggers the initial jobs load.

import { CONFIG } from './config.js';
import { state } from './state.js';
import { showScreen } from './nav.js';
import { showToast } from './ui.js';

if (typeof msal === 'undefined') {
  document.body.innerHTML = '<div style="padding:24px;color:#ff6b6b;font-family:Inter,sans-serif;text-align:center;"><h2>Authentication Library Failed To Load</h2><p style="color:#9aa89e;margin-top:8px;">Check your internet connection and refresh.</p><button onclick="location.reload()" style="margin-top:16px;background:#7fff7f;color:#0f1410;border:none;padding:10px 20px;border-radius:8px;font-weight:600;cursor:pointer;">Refresh</button></div>';
  throw new Error('MSAL library failed to load');
}

export const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
    navigateToLoginRequestUrl: false
  },
  cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
});

// Modules that need to react after sign-in (e.g. jobs.js to load jobs)
// register a callback here. We avoid hard imports to keep auth.js leaf-level.
let onAuthedCallback = null;
export function setOnAuthed(fn) { onAuthedCallback = fn; }

export async function initAuth() {
  try {
    await msalInstance.initialize();
    const response = await msalInstance.handleRedirectPromise();
    if (response && response.account) msalInstance.setActiveAccount(response.account);
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      state.currentAccount = accounts[0];
      msalInstance.setActiveAccount(state.currentAccount);
      await onSignedIn();
    }
  } catch (err) {
    console.error('Auth init error:', err);
    showToast('Sign In Issue - Please Try Again', 'error');
  }
}

export async function login() {
  try { await msalInstance.loginRedirect({ scopes: CONFIG.scopes, prompt: 'select_account' }); }
  catch (err) { console.error(err); showToast('Sign In Failed', 'error'); }
}

export async function logout() {
  try { await msalInstance.logoutRedirect({ account: state.currentAccount }); } catch (err) { console.error(err); }
}

export async function getToken() {
  const request = { scopes: CONFIG.scopes, account: state.currentAccount };
  try { return (await msalInstance.acquireTokenSilent(request)).accessToken; }
  catch (err) {
    if (err instanceof msal.InteractionRequiredAuthError) {
      await msalInstance.acquireTokenRedirect(request);
      return null;
    }
    throw err;
  }
}

async function onSignedIn() {
  state.currentUserEmail = (state.currentAccount.username || '').toLowerCase();
  state.isAdmin = CONFIG.hardcodedAdmins.includes(state.currentUserEmail);

  if (!state.isAdmin) {
    document.getElementById('awaiting-email').textContent = state.currentAccount.username;
    showScreen('awaiting-screen');
    return;
  }

  document.getElementById('user-name').textContent = state.currentAccount.name || state.currentAccount.username;
  document.getElementById('user-avatar').textContent = (state.currentAccount.name || 'U').charAt(0).toUpperCase();
  document.getElementById('admin-banner').classList.remove('hidden');

  state.navStack.length = 0;
  state.navStack.push({ screen: 'jobs-screen' });
  showScreen('jobs-screen');

  if (onAuthedCallback) await onAuthedCallback();
}

export function toggleUserMenu() {
  if (confirm(`Signed in as ${state.currentAccount.username}\n\nSign out?`)) logout();
}

// Inline-onclick exposure
window.login = login;
window.logout = logout;
window.toggleUserMenu = toggleUserMenu;
