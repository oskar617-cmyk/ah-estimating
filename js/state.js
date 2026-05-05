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
  activeTemplateCategory: null   // Email Templates tab: '__default__' or a category name
};
