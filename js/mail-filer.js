// js/mail-filer.js
// After a reply is matched + classified, move the email from Inbox into
// "Inbox / Jobs / <ah-office-folder-name>" so the inbox stays clean.
//
// The "ah-office-folder-name" is the same shape we create when a job is
// scaffolded: `[code] [name] - [address]` — so this module just needs the
// matched job's metadata.

import { ensureJobMailFolder, moveMessage } from './graph.js';

// Cache folder ids per session — avoids re-resolving the same id repeatedly
// during a poll cycle.
const folderIdCache = new Map();

// Build the AH Office folder name (no " Site Docs" suffix).
// `match.ref` from reply-matcher contains jobCode/jobName/address derived
// from the AH Site folder name; the AH Office naming convention is the
// same code + name + address but without " Site Docs - ".
// Our job folder name in AH Site is "{code} {name} Site Docs - {address}";
// the AH Office equivalent is "{code} {name} - {address}".
export function buildOfficeFolderName(jobCode, jobName, address) {
  return `${jobCode} ${jobName} - ${address}`;
}

// Move a message into Inbox/Jobs/<officeFolderName>. Returns the new
// message id (Graph returns the moved message). Failures are logged and
// rethrown so caller can decide whether to retry next poll.
export async function fileMessageToJob(messageId, officeFolderName) {
  let folderId = folderIdCache.get(officeFolderName);
  if (!folderId) {
    folderId = await ensureJobMailFolder(officeFolderName);
    folderIdCache.set(officeFolderName, folderId);
  }
  return moveMessage(messageId, folderId);
}

export function clearFolderIdCache() {
  folderIdCache.clear();
}
