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

// List files (excluding folders) inside a folder path. Returns array of
// { name, lastModifiedDateTime, size, webUrl } sorted by name.
export async function listFiles(siteId, folderPath) {
  const result = await graphFetch(
    `/sites/${siteId}/drive/root:/${encodeUriPath(folderPath)}:/children?$top=500&$select=id,name,file,folder,webUrl,size,lastModifiedDateTime`
  );
  return (result.value || [])
    .filter(it => it.file)
    .map(it => ({
      name: it.name,
      lastModifiedDateTime: it.lastModifiedDateTime,
      size: it.size,
      webUrl: it.webUrl
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

// Read a file's binary content as ArrayBuffer (used for SOW attachments).
export async function readBinary(siteId, parentPath, filename) {
  const path = parentPath ? `${parentPath}/${filename}` : filename;
  const token = await getToken();
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeUriPath(path)}:/content`,
    { headers: { 'Authorization': `Bearer ${token}` } }
  );
  if (res.status === 404) {
    const err = new Error('File not found');
    err.status = 404;
    throw err;
  }
  if (!res.ok) throw new Error(`Read binary failed: ${res.status}`);
  return res.arrayBuffer();
}

// Check whether a file exists at a given path. Returns true / false.
export async function fileExists(siteId, parentPath, filename) {
  const path = parentPath ? `${parentPath}/${filename}` : filename;
  try {
    await graphFetch(`/sites/${siteId}/drive/root:/${encodeUriPath(path)}`);
    return true;
  } catch (err) {
    if (err.status === 404) return false;
    throw err;
  }
}

// Create or fetch an "anyone with link can view" share link for a folder.
// Returns the shareable URL.
export async function createAnonymousReadLink(siteId, folderPath) {
  const folderItem = await graphFetch(
    `/sites/${siteId}/drive/root:/${encodeUriPath(folderPath)}`
  );
  const body = JSON.stringify({ type: 'view', scope: 'anonymous' });
  const result = await graphFetch(
    `/sites/${siteId}/drive/items/${folderItem.id}/createLink`,
    { method: 'POST', body }
  );
  return result.link.webUrl;
}

// Convert ArrayBuffer to base64 string for Graph attachment payload.
function arrayBufferToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let binary = '';
  // Process in chunks to avoid call stack overflow on large files
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

// Send an email via Microsoft Graph as the signed-in user.
// `payload` shape:
//   {
//     subject, htmlBody, toRecipients[], ccRecipients[], replyToRecipients[],
//     attachments: [{ name, contentBytes (base64), contentType }],
//     customHeaders: { 'x-foo': 'bar' }   // optional
//   }
// Returns true on success. /me/sendMail returns 202 with no body so we can't
// get the message id back from the call itself.
export async function sendMail(payload) {
  const headers = [];
  if (payload.customHeaders) {
    for (const [name, value] of Object.entries(payload.customHeaders)) {
      // Graph requires custom headers to start with 'x-' (case-insensitive)
      if (/^x-/i.test(name)) headers.push({ name, value: String(value) });
    }
  }
  const message = {
    subject: payload.subject,
    body: { contentType: 'HTML', content: payload.htmlBody },
    toRecipients: (payload.toRecipients || []).map(e => ({ emailAddress: { address: e } })),
    ccRecipients: (payload.ccRecipients || []).map(e => ({ emailAddress: { address: e } })),
    replyTo: (payload.replyToRecipients || []).map(e => ({ emailAddress: { address: e } })),
    attachments: (payload.attachments || []).map(a => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentType: a.contentType || 'application/octet-stream',
      contentBytes: a.contentBytes
    }))
  };
  if (headers.length > 0) message.internetMessageHeaders = headers;
  await graphFetch('/me/sendMail', {
    method: 'POST',
    body: JSON.stringify({ message, saveToSentItems: true })
  });
  return true;
}

// Lightweight helper: list filenames (no metadata) under a folder path.
// Used by Settings to check which categories already have a SOW template.
export async function listFilenames(siteId, folderPath) {
  try {
    const result = await graphFetch(
      `/sites/${siteId}/drive/root:/${encodeUriPath(folderPath)}:/children?$top=500&$select=name,file`
    );
    return (result.value || [])
      .filter(it => it.file)
      .map(it => it.name);
  } catch (err) {
    // If the folder doesn't exist, treat as empty list rather than failing
    if (err.status === 404) return [];
    throw err;
  }
}

// Helper exported for callers that need to base64-encode binary attachments.
export { arrayBufferToBase64 };
