// js/config.js
// All CONFIG constants in one place. Edit values here, never elsewhere.

export const CONFIG = {
  clientId: '07eef32f-8834-424d-b4fd-ad04c91a3fcf',
  tenantId: 'ff968505-cca0-4cd1-9f6d-68ce6eaf06c7',
  redirectUri: 'https://oskar617-cmyk.github.io/ah-estimating/',
  ahSitePath: 'auzziehomes.sharepoint.com:/sites/AHSite',
  ahOfficePath: 'auzziehomes.sharepoint.com:/sites/AHOffice',
  commonDocsPath: 'AAA Quote Common Docs',
  budgetTemplateName: '0 Budget Control Template.xlsx',
  configFileName: 'estimating-config.json',
  suppliersFileName: 'suppliers.json',
  jobFolderPattern: /^(\d{4})\s+([A-Z][A-Z0-9]{1,4})\s+Site\s+Docs\s+-\s+(.+)$/,
  codeRegex: /^\d{4}$/,
  nameRegex: /^[A-Z][A-Z0-9]{1,4}$/,
  ahSiteSubfolders: [
    'AAA Docs for Tradies',
    'Builder Advise',
    'Dilapidation Report',
    'Inspections',
    'Permit',
    'Quote',
    'RFI',
    'Service',
    'Take off',
    'Variations'
  ],
  ahOfficeSubfolders: ['Contract', 'INV', 'INV to Client', 'Q to Client', 'Tender'],
  hardcodedAdmins: ['oskar@auhs.com.au', 'est@auhs.com.au'],
  senderEmail: 'est@auhs.com.au', // only this email can send RFQs
  defaultDaysToRespond: 5,
  defaultDaysToFollowup: 3,
  // Tracker JSON lives inside an _app sub-folder so site manager users
  // browsing the Quote folder aren't tempted to edit/delete it.
  trackerSubfolder: '_app',
  trackerFilename: 'rfq-tracker.json',
  // Classifier decision log (Phase 4c-iv-tune-2). Lives in the hidden _app
  // subfolder of Common Docs so site-manager users don't see it. One file
  // per month for predictable size; the cursor file remembers the last
  // export point.
  decisionLogPath: 'AAA Quote Common Docs/_app',
  decisionLogPrefix: 'classifier-decisions-',  // e.g. classifier-decisions-2026-05.json
  decisionCursorFilename: 'classifier-decisions-cursor.json',
  decisionExportMaxEntries: 500,
  // Sent Items lookup tuning (Phase 4c-iv).
  // Exchange indexing can lag a few seconds after sendMail. We retry with
  // backoff so most lookups succeed without manual fix.
  sentLookupInitialDelayMs: 3000,
  sentLookupRetryDelayMs: 2500,
  sentLookupMaxAttempts: 4,
  // Inbox poller (Phase 4c-iv)
  inboxPollIntervalMs: 30000,            // 30s while app is open
  inboxLookbackMinutesOnFirstPoll: 1440, // catch up the last 24h on first poll
  // Classification proxy. Holds the Gemini API key server-side so it never
  // ships to the browser. We use a Cloudflare Worker on the free tier.
  classifierUrl: 'https://ah-estimating-classifier.oskar617.workers.dev',
  scopes: [
    'User.Read', 'User.ReadBasic.All',
    'Files.ReadWrite.All', 'Sites.ReadWrite.All',
    'Mail.Send', 'Mail.Read', 'Mail.ReadWrite',
    'Contacts.Read'
  ]
};
