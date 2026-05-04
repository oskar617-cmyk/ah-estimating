// js/graph.js
// All Microsoft Graph API interactions: REST fetch wrapper, site lookups,
// folder creation, file copying, JSON read/write, XLSX read.

import { CONFIG } from './config.js';
import { state } from './state.js';
import { getToken } from './auth.js';

export async function graphFetch(path, options = {}) {
  const token = await getToken();
  if (!token) return null;
  const url = path.startsWith('http') ? path : `https://graph.microsoft.com/v1.0${path}`;
  const headers = { 'Authorization': `Bearer ${token}`, ...(options.headers || {}) };
  if (options.body && !headers['Content-Type'] && typeof options.body === 'string') {
    headers['Content-Type'] = 'application/json';
  }
  const res = await fetch(url, { ...options, headers });
  if (!res.ok) {
    const text = await res.text();
    let parsed;
    try { parsed = JSON.parse(text); } catch (e) {}
    const err = new Error(`Graph ${res.status}: ${(parsed && parsed.error && parsed.error.message) || text}`);
    err.status = res.status; err.body = parsed || text;
    throw err;
  }
  if (res.status === 204) return null;
  const ct = res.headers.get('content-type') || '';
  if (ct.includes('application/json')) return res.json();
  return res;
}

export async function getAhSiteId() {
  if (state.cachedAhSiteId) return state.cachedAhSiteId;
  const r = await graphFetch(`/sites/${CONFIG.ahSitePath}`);
  state.cachedAhSiteId = r.id;
  return state.cachedAhSiteId;
}

export async function getAhOfficeId() {
  if (state.cachedAhOfficeId) return state.cachedAhOfficeId;
  const r = await graphFetch(`/sites/${CONFIG.ahOfficePath}`);
  state.cachedAhOfficeId = r.id;
  return state.cachedAhOfficeId;
}

export function encodeUriPath(path) {
  return path.split('/').map(seg => encodeURIComponent(seg)).join('/');
}

// Create a folder if it doesn't exist; return the folder item either way.
export async function ensureFolder(siteId, parentPath, folderName) {
  const parentRef = parentPath
    ? `/sites/${siteId}/drive/root:/${encodeUriPath(parentPath)}:`
    : `/sites/${siteId}/drive/root`;
  try {
    const body = JSON.stringify({ name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
    const result = await graphFetch(`${parentRef}/children`, { method: 'POST', body });
    return { item: result, created: true };
  } catch (err) {
    if (err.status === 409) {
      const path = parentPath ? `${parentPath}/${folderName}` : folderName;
      const existing = await graphFetch(`/sites/${siteId}/drive/root:/${encodeUriPath(path)}`);
      return { item: existing, created: false };
    }
    throw err;
  }
}

export async function copyFile(siteId, sourcePath, targetFolderPath, newName) {
  const sourceItem = await graphFetch(`/sites/${siteId}/drive/root:/${encodeUriPath(sourcePath)}`);
  const targetFolder = await graphFetch(`/sites/${siteId}/drive/root:/${encodeUriPath(targetFolderPath)}`);
  const body = JSON.stringify({
    parentReference: { driveId: targetFolder.parentReference.driveId, id: targetFolder.id },
    name: newName
  });
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${sourceItem.id}/copy`,
    { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body }
  );
  if (res.status !== 202 && !res.ok) {
    const text = await res.text();
    throw new Error(`Copy failed: ${res.status} ${text}`);
  }
  return true;
}

export async function uploadJson(siteId, parentPath, filename, data) {
  const content = JSON.stringify(data, null, 2);
  const path = parentPath ? `${parentPath}/${filename}` : filename;
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
    { method: 'PUT', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, body: content }
  );
  if (!res.ok) { const text = await res.text(); throw new Error(`Upload failed: ${res.status} ${text}`); }
  return res.json();
}

export async function readJson(siteId, parentPath, filename) {
  const path = parentPath ? `${parentPath}/${filename}` : filename;
  try {
    const token = await getToken();
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
      { headers: { 'Authorization': `Bearer ${token}` } }
    );
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Read failed: ${res.status}`);
    return res.json();
  } catch (err) {
    if (err.message && err.message.includes('404')) return null;
    throw err;
  }
}

export async function readXlsx(siteId, parentPath, filename) {
  const path = parentPath ? `${parentPath}/${filename}` : filename;
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
    { headers: { 'Authorization': `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Read XLSX failed: ${res.status}`);
  return res.arrayBuffer();
}
