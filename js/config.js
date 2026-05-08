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
  // Inbox poller (Phase 4c-iv)
  inboxPollIntervalMs: 30000,            // 30s while app is open
  inboxLookbackMinutesOnFirstPoll: 1440, // catch up the last 24h on first poll
  // Azure Function endpoint for Gemini classification proxy. Holds the
  // Gemini API key server-side so it never ships to the browser. User must
  // create the Function and paste the URL here before classification works.
  classifierUrl: '',  // e.g. 'https://ah-estimating-fn.azurewebsites.net/api/classify'
  scopes: [
    'User.Read', 'User.ReadBasic.All',
    'Files.ReadWrite.All', 'Sites.ReadWrite.All',
    'Mail.Send', 'Mail.Read', 'Mail.ReadWrite',
    'Contacts.Read'
  ]
};
