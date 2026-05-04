# AH Estimating

Internal estimating tool for Auzzie Homes — manages RFQs (Request For Quotes) sent to subcontractors and material suppliers, tracks responses, classifies replies with AI, and updates the per-job costing spreadsheet.

**Live:** https://oskar617-cmyk.github.io/ah-estimating/

## Tech Stack

- **Front end:** PWA written as plain HTML + ES modules. No frameworks, no build step, no npm.
- **Auth:** Microsoft Entra ID (single-tenant, Auzzie Homes only) via MSAL.js
- **Backend services:** Microsoft Graph API for SharePoint files, Outlook mail and contacts
- **Background automation:** Power Automate flows (planned)
- **AI:** Google Gemini API via Azure Function proxy (planned)
- **Hosting:** GitHub Pages
- **Storage:** SharePoint (AH Site only) — no separate database

## File Structure

```
index.html              HTML structure, CSS, screen markup
manifest.json           PWA manifest
service-worker.js       Minimal service worker for installability
icon-192.png            App icon (192×192)
icon-512.png            App icon (512×512)
js/
├── app.js              Entry point; wires modules and boots
├── config.js           CONFIG constants
├── state.js            Shared mutable app state
├── ui.js               Toast, modal, confirm dialog, escapeHtml
├── nav.js              Screen routing and navigation stack
├── auth.js             MSAL setup, login/logout, token, post-auth dispatch
├── graph.js            Microsoft Graph helpers (fetch, folders, files, JSON, XLSX)
├── audit.js            Audit log + app config + supplier persistence
├── jobs.js             Jobs list, job detail, migration prompt
├── new-job.js          New Job creation flow
├── settings.js         Settings screen + Signature tab
├── catalog.js          Trades / Suppliers catalog with Excel write-back
└── companies.js        Company editor modal
```

## SharePoint Structure

```
AH Site / Documents /
├── AAA Quote Common Docs /
│   ├── Email Templates /         per-trade RFQ body templates
│   ├── SOW Templates /           Word docs, one per trade
│   ├── Trade Contacts /          (legacy, unused — supplier data lives in suppliers.json)
│   ├── 0 Budget Control Template.xlsx
│   ├── estimating-config.json    trades, mappings, signature
│   ├── suppliers.json            companies grouped by trade
│   └── audit-log-YYYY-MM.json    monthly audit log
│
└── [JobCode] [JobName] Site Docs - [Address] /
    ├── AAA Docs for Tradies [JobName] /
    ├── Quote /
    │   ├── [Trade] - [Company] v[N] - [Amount].pdf
    │   ├── 0 Budget Control [JobName].xlsx
    │   └── rfq-tracker.json
    ├── Builder Advise /
    ├── Dilapidation Report /
    ├── Inspections /
    ├── Permit /
    ├── RFI /
    ├── Service /
    ├── Take off /
    └── Variations /
```

The app also creates an empty matching folder structure under `AH Office / Documents / [JobCode] [JobName] - [Address] /` (Contract, INV, INV to Client, Q to Client, Tender) on job creation.

## Authorisation

- App is publicly hosted on GitHub Pages but only Microsoft 365 accounts in the Auzzie Homes tenant can sign in.
- Two hardcoded admins: `oskar@auhs.com.au` and `est@auhs.com.au`. Both can use all features *except* sending RFQs — only `est@auhs.com.au` will be able to send RFQs.

## Setup (For Forking This Repo)

1. **Microsoft Entra App Registration** — at https://entra.microsoft.com → Applications → App registrations → New registration:
   - Name your app
   - Single tenant
   - Single-page application (SPA) redirect URI matching your GitHub Pages URL
   - Add Microsoft Graph delegated permissions: `User.Read`, `User.ReadBasic.All`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Mail.Send`, `Mail.Read`, `Mail.ReadWrite`, `Contacts.Read`
   - Grant admin consent
   - Copy the Client ID and Tenant ID

2. **SharePoint preparation** — create `AAA Quote Common Docs/` with subfolders `Email Templates/`, `SOW Templates/`, `Trade Contacts/`, and place a budget Excel template inside.

3. **Edit `js/config.js`** — replace `clientId`, `tenantId`, `redirectUri`, and SharePoint site paths.

4. **Enable GitHub Pages** — repo Settings → Pages → Deploy from branch `main` / root.
