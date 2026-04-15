# 📧 MailMerge-Pro

**Personalized bulk email directly from Outlook — powered by Microsoft Graph API.**

MailMerge-Pro is a free, open-source Outlook Web Add-in that lets you send personalized emails to hundreds of recipients using data from Excel or CSV files. It runs entirely in the browser — no server, no backend, no data leaves your machine.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Outlook Add-in](https://img.shields.io/badge/Platform-Outlook%20Web%20Add--in-0078D4)](https://docs.microsoft.com/en-us/office/dev/add-ins/)
[![GitHub Pages](https://img.shields.io/badge/Hosted-GitHub%20Pages-222)](https://itsjamessmith.github.io/mailmerge-pro/)

---

## ✨ Features (44 total)

### Core Mail Merge
- 📋 **Excel/CSV Upload** — Drag-and-drop or file picker for `.xlsx`, `.xls`, `.csv`
- 🔗 **Column Mapping** — Map spreadsheet columns to To, CC, BCC, Subject
- ✏️ **Rich Text Editor** — Bold, italic, lists, links, colors, merge fields
- 🚀 **Bulk Send or Draft** — Send emails or save as drafts via Microsoft Graph API
- 📊 **Live Progress** — Real-time progress bar with per-recipient status
- 📎 **Attachments** — Global and per-recipient file attachments

### Advanced Features
- 🔬 **A/B Testing** — Split recipients into A/B groups with different subject/body
- ⏰ **Scheduled Sending** — Schedule emails for a future date/time
- 📝 **Email Templates** — Save, load, and manage reusable templates (built-in + custom)
- 👤 **Contact Import** — Import recipients from Microsoft 365 contacts
- 📑 **Saved Lists** — Save and merge recipient lists for reuse
- 📄 **HTML Import** — Import custom HTML email templates
- ✒️ **Auto-Signature** — Fetch Outlook signature or paste custom
- 🔄 **Fallback Defaults** — Default values when merge fields are empty
- 📊 **Group by Email** — Many-to-one: combine multiple rows per recipient
- 📬 **Read Receipts** — Request read receipt on sent emails
- ❗ **High Importance** — Flag emails as high importance
- 🚫 **Unsubscribe Link** — Add List-Unsubscribe header
- 📈 **Email Tracking** — Read tracking via Graph API

### UI & Experience
- 🌙 **Dark Mode** — Auto-detects OS/Outlook theme + manual toggle
- 🌍 **7 Languages** — English, Spanish, French, German, Portuguese, Japanese, Chinese
- 📊 **Admin Dashboard** — Campaign history, success rates, top recipients
- 📊 **Rate Limit Dashboard** — Track daily send volume with visual gauge
- 🔍 **Recipient Search** — Filter recipients by name/email
- 📜 **Campaign History** — View past campaign results and details

### Security & Resilience
- 🔒 **DOMPurify** — HTML sanitization prevents XSS attacks
- 🔐 **XSS Protection** — `sanitizeHtml()` strips script tags, iframes, event handlers, javascript: URLs from all HTML content
- 🔄 **Auto-Retry** — Exponential backoff for network/server errors
- ⏱️ **Rate Limiting** — Token bucket enforces 30 emails/min
- 💾 **Checkpointing** — Send progress saved; resume after interruption
- 🛡️ **NAA Authentication** — Nested App Authentication (Microsoft-recommended for 2025+) via MSAL.js v3.27.0 with OAuth 2.0 + PKCE
- 🔑 **sessionStorage tokens** — Authentication tokens stored in sessionStorage, auto-cleared when tab closes
- 🚫 **Outlook-only execution** — Add-in refuses to run outside Outlook

---

## 🚀 Quick Start

### 1. Register the Add-in

**Option A: Sideload for Development**
1. Open [Outlook on the web](https://outlook.office.com)
2. Go to **Settings → Integrated Apps → Add-ins → My add-ins**
3. Click **+ Add a custom add-in → Add from URL**
4. Enter: `https://itsjamessmith.github.io/mailmerge-pro/manifest.xml`

**Option B: Admin Deployment**
1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com) → **Integrated Apps**
2. Upload `manifest.xml` and deploy to users/groups

### 2. Configure Azure AD (Required)

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps)
2. Register a new app:
   - **Name:** MailMerge-Pro
   - **Supported account types:** Single tenant (your org)
   - **Redirect URI:** `https://itsjamessmith.github.io/mailmerge-pro/taskpane.html` (SPA)
3. Grant **delegated** permissions:
   - `Mail.Send`, `Mail.ReadWrite`, `User.Read`, `Contacts.Read`, `Mail.Send.Shared`
4. **Grant admin consent** for all permissions
5. Update `taskpane.js` with your Client ID and Tenant ID

### 3. Use It
1. Open an email compose window in Outlook
2. Click the MailMerge-Pro icon in the ribbon
3. Upload your spreadsheet → Map columns → Compose → Send!

---

## 🛠️ Development

### Prerequisites
- Node.js 18+ (for build tools)
- Git

### Setup
```bash
git clone https://github.com/itsjamessmith/mailmerge-pro.git
cd mailmerge-pro
npm install
```

### Build
```bash
# Development (watch mode)
npm run dev

# Production build (minified)
npm run build

# Syntax check
npm run lint
```

### Project Structure
```
mailmerge-pro/
├── taskpane.html        # Main add-in UI
├── taskpane.js          # Application logic (~2,400 lines)
├── taskpane.css         # Styles with dark mode + CSS variables
├── manifest.xml         # Office Add-in manifest
├── index.html           # Landing/support page
├── function-file.html   # Office function file
├── assets/              # Icons (16x16, 32x32, 80x80, SVG)
├── dist/                # Minified output (after build)
├── .github/workflows/   # CI/CD pipelines
├── CHANGELOG.md         # Version history
├── LICENSE              # MIT License
└── README.md            # This file
```

---

## 🔧 Configuration

Edit `taskpane.js` to set your Azure AD details:

```javascript
const msalConfig = {
    auth: {
        clientId: "YOUR_CLIENT_ID",
        authority: "https://login.microsoftonline.com/YOUR_TENANT_ID",
        redirectUri: "https://YOUR_DOMAIN/taskpane.html"
    }
};
```

### Required API Permissions (Delegated)

| Permission | Purpose |
|------------|---------|
| `Mail.Send` | Send emails on behalf of user |
| `Mail.ReadWrite` | Create and manage draft emails |
| `User.Read` | Get user profile/email address |
| `Contacts.Read` | Import contacts from address book |
| `Mail.Send.Shared` | Send from shared mailboxes |

---

## 🌍 Supported Languages

| Language | Code | Status |
|----------|------|--------|
| English | `en` | ✅ Complete |
| Spanish | `es` | ✅ Complete |
| French | `fr` | ✅ Complete |
| German | `de` | ✅ Complete |
| Portuguese | `pt` | ✅ Complete |
| Japanese | `ja` | ✅ Complete |
| Chinese | `zh` | ✅ Complete |

---

## 📊 Rate Limits

MailMerge-Pro enforces Microsoft Graph API limits:

| Limit | Value | Enforcement |
|-------|-------|-------------|
| Per minute | 30 emails | Token bucket (automatic) |
| Per day | 10,000 emails | Dashboard warning |
| Retry on 429 | Automatic | Respects `Retry-After` header |
| Network errors | 3 retries | Exponential backoff |

---

## 🔒 Security

- **No server / no backend** — All processing happens in the browser
- **NAA (Nested App Authentication)** — Microsoft-recommended auth for Office add-ins (2025+), uses `createNestablePublicClientApplication` for seamless SSO in Outlook's task pane
- **MSAL.js v3.27.0** — Upgraded from v2.35.0; loaded from jsDelivr CDN (`cdn.jsdelivr.net/npm/@azure/msal-browser@3.27.0`)
- **OAuth 2.0 + PKCE** — Industry-standard authentication via Microsoft's identity platform
- **sessionStorage for tokens** — MSAL cache uses `sessionStorage` instead of `localStorage`; tokens are automatically cleared when the browser tab closes, preventing token theft
- **XSS Protection** — `sanitizeHtml()` strips script tags, iframes, event handlers, and `javascript:` URLs; merge field values are HTML-escaped; link insertion validates URL schemes
- **No secrets in code** — Client ID is public; tokens managed by MSAL
- **DOMPurify** — All user-provided HTML is sanitized before rendering
- **Safe localStorage** — All JSON parsing wrapped in try/catch
- **Outlook-only execution** — App refuses to run outside Outlook, blocking standalone browser access
- **Clean sign-out** — All PII cleared from localStorage on sign-out
- **CDN integrity** — All CDN scripts have `crossorigin="anonymous"` attribute
- **CSP-ready** — No inline event handlers

---

## 📝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

*© 2026 MailMerge-Pro. All rights reserved.*

---

## 🙏 Acknowledgments

- [MSAL.js v3](https://github.com/AzureAD/microsoft-authentication-library-for-js) — Microsoft Authentication Library (NAA with `createNestablePublicClientApplication`)
- [SheetJS](https://sheetjs.com/) — Excel/CSV parsing
- [DOMPurify](https://github.com/cure53/DOMPurify) — HTML sanitization
- [Office.js](https://docs.microsoft.com/en-us/office/dev/add-ins/) — Office Add-in API
