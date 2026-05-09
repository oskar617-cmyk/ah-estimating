// js/audit.js
// Audit logging plus app-level config/supplier persistence.
// These three concerns live together because they all hit the same
// Common Docs folder via JSON read/write.

import { CONFIG } from './config.js';
import { state } from './state.js';
import { getAhSiteId, readJson, uploadJson } from './graph.js';

export async function logAudit(action, target, details) {
  try {
    const siteId = await getAhSiteId();
    const now = new Date();
    const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const filename = `audit-log-${monthKey}.json`;
    const folderPath = CONFIG.commonDocsPath;
    const existing = (await readJson(siteId, folderPath, filename)) || { entries: [] };
    existing.entries.push({
      timestamp: now.toISOString(),
      user: state.currentUserEmail,
      action,
      target,
      details: details || null
    });
    await uploadJson(siteId, folderPath, filename, existing);
  } catch (err) {
    // Audit failure must never break user-facing actions
    console.warn('Audit log write failed:', err);
  }
}

export async function loadAppConfig() {
  if (state.appConfig) return state.appConfig;
  const siteId = await getAhSiteId();
  state.appConfig = (await readJson(siteId, CONFIG.commonDocsPath, CONFIG.configFileName)) || createDefaultConfig();
  return state.appConfig;
}

function createDefaultConfig() {
  return {
    version: 1,
    // each trade: {category, budgetRowNo, daysToRespond, daysToFollowup, sowTemplate, emailTemplate, availableRows: [{no, description, type}]}
    trades: [],
    signature: {
      title: 'Estimator',
      body: 'Auzzie Homes Pty Ltd\n[Address]\nPhone: [Phone]'
    },
    suggestionHistory: {},
    createdAt: new Date().toISOString()
  };
}

export async function saveAppConfig() {
  const siteId = await getAhSiteId();
  await uploadJson(siteId, CONFIG.commonDocsPath, CONFIG.configFileName, state.appConfig);
}

export async function loadSuppliers() {
  if (state.suppliersData) return state.suppliersData;
  const siteId = await getAhSiteId();
  state.suppliersData = (await readJson(siteId, CONFIG.commonDocsPath, CONFIG.suppliersFileName)) || { suppliers: [] };
  return state.suppliersData;
}

export async function saveSuppliers() {
  const siteId = await getAhSiteId();
  await uploadJson(siteId, CONFIG.commonDocsPath, CONFIG.suppliersFileName, state.suppliersData);
}

// ---------- Tracker JSON helpers ----------
// Tracker lives at [Job]/Quote/_app/rfq-tracker.json. To smoothly migrate
// existing jobs whose tracker is at the old location ([Job]/Quote/),
// readTracker tries the new path first, falls back to the legacy path,
// and on a successful legacy read writes a copy to the new path so the
// next read finds it there. The legacy file is left in place — site
// admins can delete it manually once they're satisfied. We never delete
// SharePoint files automatically.

export async function readTracker(jobFolderName) {
  const siteId = await getAhSiteId();
  const newDir = `${jobFolderName}/Quote/${CONFIG.trackerSubfolder}`;
  const legacyDir = `${jobFolderName}/Quote`;
  // Try new location first
  let tracker = await readJson(siteId, newDir, CONFIG.trackerFilename);
  if (tracker) return tracker;
  // Fall back to legacy location and migrate forward
  tracker = await readJson(siteId, legacyDir, CONFIG.trackerFilename);
  if (tracker) {
    try { await uploadJson(siteId, newDir, CONFIG.trackerFilename, tracker); }
    catch (e) { console.warn(`Tracker migration to ${newDir} failed:`, e); }
    return tracker;
  }
  return null;
}

export async function writeTracker(jobFolderName, tracker) {
  const siteId = await getAhSiteId();
  const dir = `${jobFolderName}/Quote/${CONFIG.trackerSubfolder}`;
  await uploadJson(siteId, dir, CONFIG.trackerFilename, tracker);
}
