// js/state.js
// Shared mutable app state. Modules import this object and read/write fields.
// Using a single exported object lets us mutate in place across modules
// (ES module exports of primitives are read-only; object properties are not).

export const state = {
  currentAccount: null,
  currentUserEmail: null,
  isAdmin: false,
  cachedAhSiteId: null,
  cachedAhOfficeId: null,
  navStack: [],
  currentJob: null,
  emailSearchTimer: null,
  projectTeamEmails: [],
  appConfig: null,        // estimating-config.json contents
  suppliersData: null,    // suppliers.json contents
  editingSupplier: null,
  supplierMultiSelectTrades: [],
  expandedCatalogItems: new Set(),
  activeTemplateCategory: null,   // Email Templates tab: '__default__' or a category name
  sowFilenames: null,             // Cached array of files in SOW Templates folder
  // Inbox / notifications (Phase 4c-iv)
  inboxPollerHandle: null,        // setInterval handle (so we can stop on logout)
  notifications: [],              // array of notification objects (newest first)
  pendingReview: [],              // entries auto-written to budget Excel awaiting confirmation
  notificationPanelOpen: false,
  inboxLastPolledAt: null,        // ISO timestamp of last successful poll
  processedMessageIds: new Set(), // dedupe: ids of messages already turned into notifications this session
  jobTrackerCache: new Map()      // jobFolderName -> rfq-tracker.json (cached during inbox processing)
};
