# AH Estimating

Internal estimating tool for Auzzie Homes — manages RFQs (Request For Quotes) sent to subcontractors and material suppliers, tracks responses, classifies replies with AI, and updates the per-job costing spreadsheet.

**Live:** https://oskar617-cmyk.github.io/ah-estimating/

## What It Does

- Creates new construction job folders in SharePoint (both AH Site and AH Office) from a single form
- Sends RFQ emails to multiple suppliers per trade with a per-trade template, attached SOW Word doc, drawings folder link, and auto-generated file list
- Schedules automatic follow-ups per trade (configurable days, gives up after 2 reminders)
- Watches the shared estimator inbox, classifies incoming replies as Quote / Question / Decline / Suspicious using AI
- Auto-saves received quote PDFs into the right job folder, renames them in standard format, archives previous versions
- Extracts the quote dollar amount and writes it into the per-job costing Excel sheet
- Suspicious-email detection flags phishing attempts and dangerous attachments without auto-opening them

## Tech Stack

- **Front end:** Single-file PWA — `index.html` containing all HTML, CSS, JavaScript. No frameworks, no build step.
- **Auth:** Microsoft Entra ID (single-tenant, Auzzie Homes only) via MSAL.js
- **Backend services:** Microsoft Graph API for SharePoint files, Outlook mail and contacts
- **Background automation:** Power Automate flows (inbox watcher + daily follow-up runner)
- **AI:** Google Gemini API (free tier, 1500 req/day) called via a small Azure Function proxy
- **Hosting:** GitHub Pages
- **Storage:** SharePoint (AH Site only) — no separate database

## SharePoint Structure (Read & Write)

```
AH Site / Documents /
├── AAA Quote Common Docs /
│   ├── Email Templates /         per-trade RFQ body templates
│   ├── SOW Templates /           Word docs, one per trade
│   ├── Trade Contacts /          supplier contacts grouped by trade
│   ├── 0 Budget Control Template.xlsx
│   ├── estimating-config.json    users, roles, trades, defaults
│   ├── suggestion-history.json   smart-suggestion learning data
│   └── audit-log-YYYY-MM.json    monthly audit log
│
└── [JobCode] Site Docs - [Address] /
    ├── AAA Docs for Tradies [XX] /   drawings (linked in RFQs)
    ├── Quote /
    │   ├── [Trade] - [Company] v[N] - [Amount].pdf
    │   ├── rfq-tracker.json          per-job RFQ state
    │   └── Archived /                superseded quote versions
    ├── 0 Budget Control [XX].xlsx
    └── (other standard job folders, scaffolded by app)
```

The app also creates an empty matching folder structure under `AH Office / Documents / [JobCode] - [Address] /` (Contract, INV, INV to Client, Q to Client, Tender) on job creation, then never touches AH Office again.

## Authorisation

- App is publicly hosted on GitHub Pages but only Microsoft 365 accounts in the Auzzie Homes tenant can sign in.
- Two hardcoded admins: `oskar@auhs.com.au` and `est@auhs.com.au`.
- All other users must be added by an admin in the in-app settings panel and assigned a role: `estimator` or `site_manager`.
- Only `est@auhs.com.au` can send RFQ emails (replies funnel back to that shared mailbox); other admins can do everything else.

## Setup (For Forking This Repo)

1. **Microsoft Entra App Registration** — at https://entra.microsoft.com → Applications → App registrations → New registration:
   - Name your app
   - Single tenant
   - Single-page application (SPA) redirect URI matching your GitHub Pages URL
   - Add Microsoft Graph delegated permissions: `User.Read`, `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Mail.Send`, `Mail.Read`, `Mail.ReadWrite`, `Contacts.Read`
   - Grant admin consent
   - Copy the Client ID and Tenant ID

2. **SharePoint preparation** — create `AAA Quote Common Docs/` with subfolders `Email Templates/`, `SOW Templates/`, `Trade Contacts/`, and place a budget Excel template inside.

3. **Edit `index.html`** — replace `clientId`, `tenantId`, `redirectUri`, and SharePoint site paths in the `CONFIG` object near the top of the script.

4. **Enable GitHub Pages** — repo Settings → Pages → Deploy from branch `main` / root.

5. **(Later phases)** — Azure Function for Gemini proxy, Power Automate flows. Not needed for the basic auth + jobs-list version.

## Commit Conventions

Clear, present-tense, descriptive commit messages. Examples:
- `Add RFQ send flow with per-trade templates`
- `Fix quote amount extraction for multi-page PDFs`
- `Move suspicious detection to Azure Function`

Avoid: `update`, `wip`, `fix stuff`.
