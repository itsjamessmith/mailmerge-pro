# MailMerge-Pro — Security & Privacy Overview

> MailMerge-Pro processes ALL data locally in your browser. No data is ever sent to third-party servers.

---

## Architecture: How Data Flows

```
┌─────────────────────────────────────────────────────────┐
│                  USER'S BROWSER / OUTLOOK                │
│                                                         │
│  ┌──────────┐    ┌──────────────┐    ┌──────────────┐  │
│  │ Excel/CSV│───>│  SheetJS      │───>│ Merge Engine │  │
│  │  File    │    │ (JavaScript)  │    │ (JavaScript) │  │
│  └──────────┘    └──────────────┘    └──────┬───────┘  │
│                                             │          │
│  ┌──────────┐    ┌──────────────┐           │          │
│  │Attachment│───>│ FileReader   │───────────>│          │
│  │  Files   │    │ (base64)     │           │          │
│  └──────────┘    └──────────────┘           │          │
│                                             ▼          │
│                                    ┌──────────────┐    │
│                                    │ Microsoft    │    │
│                                    │ Graph API    │    │
│                                    │ (HTTPS)      │    │
│                                    └──────┬───────┘    │
│                                           │            │
└───────────────────────────────────────────┼────────────┘
                                            │
                                            ▼
                               ┌─────────────────────┐
                               │  Microsoft 365       │
                               │  Exchange Online     │
                               │  (Your Tenant)       │
                               └─────────────────────┘
```

**There is NO MailMerge-Pro server.** The add-in is static HTML/JS/CSS files served from GitHub Pages (or your chosen host). These files run entirely inside the user's browser session within Outlook.

---

## Data Processing — Where Each Piece Goes

| Data | Processed Where | Stored Where | Sent To |
|---|---|---|---|
| Excel/CSV recipient data | Browser (SheetJS) | Browser memory only | ❌ Nowhere |
| Email body text | Browser (JavaScript) | Browser memory only | Microsoft Exchange (as email) |
| Email subject | Browser (JavaScript) | Browser memory only | Microsoft Exchange (as email) |
| File attachments | Browser (FileReader API) | Browser memory (base64) | Microsoft Exchange (as email attachment) |
| User's OAuth token | Browser (MSAL.js) | Browser sessionStorage | Microsoft Azure AD only |
| Campaign history | Browser (JavaScript) | Browser localStorage | ❌ Nowhere |
| Merge field values | Browser (JavaScript) | Browser memory only | Microsoft Exchange (merged into email) |
| Email templates | Browser (JavaScript) | Browser localStorage | ❌ Nowhere |
| Contact groups | Browser (JavaScript) | Browser localStorage | ❌ Nowhere |
| Language preference | Browser (JavaScript) | Browser localStorage | ❌ Nowhere |
| Email signature (Graph) | Browser (MSAL.js + Graph API) | Browser localStorage (cached) | ❌ Nowhere (fetched from user's own Exchange settings) |
| Email signature (manual) | Browser (JavaScript) | Browser localStorage | Microsoft Exchange (appended to email body) |
| Dashboard / campaign stats | Browser (JavaScript) | Browser localStorage | ❌ Nowhere |
| A/B test configuration | Browser (JavaScript) | Browser memory only | ❌ Nowhere |
| Scheduled send settings | Browser (JavaScript) | Browser memory only | ❌ Nowhere (timer runs in-browser) |

### What Happens When You Upload a Spreadsheet
1. You select a file using the browser's file picker
2. The file is read **entirely in JavaScript** using the SheetJS library
3. The data exists only in browser memory (`appState.rows` variable)
4. When you close the task pane, the data is **gone** — nothing is persisted

### What Happens When You Send Emails
1. For each recipient, the add-in constructs a JSON message object in memory
2. It calls `POST https://graph.microsoft.com/v1.0/me/sendMail` with the user's OAuth token
3. Microsoft Graph routes this to Exchange Online — exactly the same as if the user clicked "Send" in Outlook
4. The email appears in the user's **Sent Items** folder
5. All Exchange compliance policies apply (DLP, retention, journaling, transport rules)

### What Happens With Attachments
1. You select files using the browser file picker
2. Each file is read as **base64** using the browser's FileReader API
3. The base64 data exists only in browser memory
4. When sending, the base64 is included in the Graph API call to create the email
5. **Attachment files are never uploaded to any external server**

### What Happens With localStorage Data (v3.0)

Several v3.0 features store data in the browser's localStorage:

1. **Templates** — When you save a custom email template, the subject, body, and options are serialized as JSON and stored in localStorage under `mailmergepro_templates`.
2. **Contact Groups** — When you save a recipient list as a contact group, the rows and column headers are stored in localStorage under `mailmergepro_groups`.
3. **Campaign History** — Each completed campaign's summary (date, subject, recipient count, success/fail counts) is stored in localStorage under `mailmergepro_campaigns`. This also powers the Local Admin Dashboard.
4. **Language Preference** — Your selected UI language is stored in localStorage under `mailmergepro_lang`.
5. **Signature** — Your fetched or manually pasted signature is cached in localStorage under `mailmergepro_signature`.

**Key privacy facts about localStorage:**
- localStorage is **sandboxed per origin** — only MailMerge-Pro running on the same domain can read this data.
- Data **stays on the device** — it is NOT synced to OneDrive, Microsoft 365, your organization's servers, or any cloud service.
- Data is **NOT accessible by administrators** — your IT admin cannot view your templates, groups, or campaign history.
- Data is **NOT encrypted** by default — anyone with physical access to the device and browser could inspect localStorage via browser developer tools.
- Clearing the browser's site data (Settings → Clear browsing data) removes ALL localStorage items.
- localStorage data does **NOT roam** across devices — switching to a different computer means starting fresh.

> **Note on authentication tokens:** MSAL authentication tokens (access tokens, refresh tokens) are stored in **`sessionStorage`**, NOT `localStorage`. This means tokens are automatically cleared when the browser tab is closed. This is a deliberate security measure — it prevents token theft from other same-origin pages and ensures credentials do not persist beyond the active session.

> **Sign-out cleanup:** When the user signs out, MailMerge-Pro clears all PII (personally identifiable information) from localStorage, including cached user name and email.

---

## XSS Protection & Input Sanitization

MailMerge-Pro implements comprehensive XSS (Cross-Site Scripting) protection:

| Protection | Details |
|---|---|
| **`sanitizeHtml()` function** | Strips `<script>` tags, `<iframe>` tags, inline event handlers (e.g., `onclick`, `onerror`), and `javascript:` URLs from all HTML content |
| **Template loading** | All loaded templates (built-in and custom) are sanitized before rendering |
| **Signature loading** | Signatures fetched from Graph API or loaded from localStorage are sanitized before display |
| **HTML template import** | Imported `.html` files are sanitized before insertion into the editor |
| **Merge field escaping** | Merge field values are HTML-escaped when building HTML emails to prevent injection |
| **Link URL validation** | Link insertion validates URL schemes — blocks `javascript:`, `data:`, and `vbscript:` protocols |

### CDN Script Integrity

All external CDN scripts are loaded with the `crossorigin="anonymous"` attribute to prevent credential leakage to CDN providers.

### Outlook-Only Execution

The add-in checks for the Office.js runtime environment at startup and refuses to run outside of Outlook. This prevents standalone browser access, reducing the attack surface.

### Large Attachment Upload

Attachments larger than 3 MB use a chunked upload session via the Graph API `createUploadSession` endpoint, rather than inline base64 encoding. This prevents memory issues and supports attachments up to 150 MB.

---

## Authentication Security

| Aspect | Details |
|---|---|
| **Protocol** | OAuth 2.0 with PKCE (Proof Key for Code Exchange) |
| **Library** | MSAL.js v3.27.0 (Microsoft Authentication Library) — loaded from jsDelivr CDN |
| **Auth method** | Nested App Authentication (NAA) via `createNestablePublicClientApplication` — the Microsoft-recommended approach for Office add-ins (2025+). NAA enables seamless SSO inside Outlook's task pane iframe, eliminating popup-blocked and redirect_in_iframe errors |
| **Token storage** | Browser `sessionStorage` (tokens are automatically cleared when the browser tab closes — prevents token theft from other same-origin pages) |
| **Token lifetime** | Access token: ~1 hour; Refresh token: ~24 hours |
| **Login flow** | NAA (seamless SSO brokered by Outlook) → Microsoft login page (only if needed) → token returned to add-in |
| **Credentials** | MailMerge-Pro NEVER sees or handles user passwords |
| **Scopes** | Only `Mail.Send`, `Mail.ReadWrite`, `User.Read` — minimum required |
| **CDN integrity** | All CDN scripts loaded with `crossorigin="anonymous"` attribute |
| **brk-multihub** | Azure AD app requires `brk-multihub://yourdomain.com` SPA redirect URI to enable NAA brokering |

### What MailMerge-Pro Can and Cannot Do

| ✅ CAN (with user consent) | ❌ CANNOT |
|---|---|
| Send emails as the signed-in user | Read other users' email |
| Create drafts in user's mailbox | Access mailboxes without permission |
| Read user's own email address | Store passwords or credentials |
| Read user's contacts (if permitted) | Send email without user signing in |
| Add attachments to outgoing emails | Access files on user's computer (beyond what user selects) |

---

## Compliance & Regulatory

### GDPR Compliance
- ✅ No personal data is collected, stored, or processed by MailMerge-Pro
- ✅ No cookies (except Microsoft's own MSAL cookies for auth)
- ✅ No analytics or tracking of any kind
- ✅ No data leaves the EU if your M365 tenant is in the EU
- ✅ Users can revoke access at any time (Azure AD → Enterprise applications → remove consent)

### HIPAA Compliance
- ✅ No PHI is processed by or transmitted to MailMerge-Pro infrastructure
- ✅ All data flows through Microsoft's HIPAA-compliant Graph API and Exchange Online
- ✅ If your M365 tenant has a BAA with Microsoft, MailMerge-Pro emails are covered

### SOC 2 / ISO 27001
- MailMerge-Pro itself is static files with no backend — no SOC 2 certification needed for the add-in
- All email processing occurs within Microsoft 365, which is SOC 2 and ISO 27001 certified
- The hosting (GitHub Pages) serves only static files over HTTPS — no data processing

### Data Residency
- Your data never leaves your Microsoft 365 tenant
- Email routing follows your Exchange Online configuration
- If your tenant is in a specific region (EU, US, etc.), emails stay in that region
- GitHub Pages serves the add-in code from global CDN — but code is not data

---

## Network Security

### URLs Accessed by MailMerge-Pro

| URL | Purpose | Data Sent |
|---|---|---|
| `login.microsoftonline.com` | User authentication | Login credentials (to Microsoft only) |
| `graph.microsoft.com` | Send emails, read contacts | Email content (to Microsoft only) |
| `cdn.jsdelivr.net` | Load MSAL.js v3.27.0 library | None (downloads JavaScript library) |
| `cdn.sheetjs.com` | Load SheetJS library | None (downloads JavaScript library) |
| Your hosting URL | Load add-in HTML/JS/CSS | None (downloads add-in code) |

### What a Network Admin Sees
- HTTPS requests to Microsoft domains (login, graph) — same as normal Outlook usage
- HTTPS request to your hosting URL to load the add-in files
- **No requests to any unknown or third-party domains**

### Firewall / Proxy Whitelist
```
login.microsoftonline.com
graph.microsoft.com
cdn.jsdelivr.net
cdn.sheetjs.com
YOUR-HOSTING-DOMAIN (e.g., username.github.io)
appsforoffice.microsoft.com
```

---

## Comparison With Competitors

| Security Aspect | MailMerge-Pro | SecureMailMerge | Mail Merge Toolkit |
|---|---|---|---|
| Data processing | Client-side only | Client-side only | Client-side (desktop) |
| Backend server | ❌ None | ⚠️ Has licensing server | ❌ None |
| Open source | ✅ Yes (auditable) | ❌ No | ❌ No |
| User data collected | ❌ None | ⚠️ License/usage data | ❌ None |
| Analytics/tracking | ❌ None | ⚠️ Unknown | ❌ None |
| Code auditable | ✅ Public GitHub repo | ❌ Minified/obfuscated | ❌ Compiled binary |
| Auth method | MSAL.js v3 (NAA — OAuth 2.0) | OAuth 2.0 | Desktop COM (no OAuth) |

**MailMerge-Pro's open-source nature means any security team can audit the code** — something competitors cannot offer.

---

## Admin Controls

### How to Revoke Access
1. **Remove the add-in:** M365 admin center → Integrated apps → Remove MailMerge-Pro
2. **Revoke app consent:** Azure AD → Enterprise applications → MailMerge-Pro → Delete
3. **Block the app:** Azure AD → Enterprise applications → Properties → "Enabled for users to sign-in" → No

### How to Audit Usage
- All emails sent via MailMerge-Pro appear in **Exchange message trace** (admin.exchange.microsoft.com)
- Emails are in each user's **Sent Items** folder
- Azure AD sign-in logs show when users authenticate to MailMerge-Pro

### How to Limit Scope
- Deploy only to specific groups (not entire organization)
- Use Exchange transport rules to limit bulk sending
- Use DLP policies to prevent sensitive data in bulk emails

---

## Frequently Asked Security Questions

**Q: Can MailMerge-Pro read my emails?**
A: No. It has `Mail.Send` and `Mail.ReadWrite` permissions. `Mail.ReadWrite` is used ONLY to create draft messages and add attachments — not to read existing emails.

**Q: Where is my spreadsheet data stored?**
A: Only in your browser's memory while the task pane is open. When you close it, the data is gone. Nothing is saved to disk, cloud, or any server.

**Q: Can the add-in developer see my data?**
A: No. The add-in is static files on GitHub Pages. There is no server to receive data, no analytics, no logging. The developer has zero visibility into your usage.

**Q: What if GitHub Pages goes down?**
A: The add-in won't load until it's back up, but no data is lost. You can self-host the files on your own server for maximum uptime.

**Q: Is the add-in safe for HR/payroll/legal emails?**
A: Yes. The data handling is equivalent to manually composing emails in Outlook. All Exchange compliance policies (encryption, DLP, retention) apply equally.

**Q: Where are my email templates stored?**
A: In your browser's localStorage on your device only. Templates are NOT sent to any server, NOT synced to the cloud, and NOT accessible by your IT admin. Clearing browser data deletes them.

**Q: Where are my contact groups stored?**
A: Same as templates — in your browser's localStorage on your device. They are local-only, not synced across devices, and not visible to administrators.

**Q: Is my campaign history visible to my admin?**
A: No. Campaign history and dashboard data are stored in your browser's localStorage. Your IT admin cannot see your campaign stats, top recipients, or send history through MailMerge-Pro. However, all emails sent via MailMerge-Pro appear in your Sent Items and Exchange message trace, which admins CAN access — this is the same as any email you send from Outlook.

**Q: Does the scheduled sending feature send data to a server?**
A: No. The schedule timer runs entirely in your browser's JavaScript runtime. No scheduling data is sent to any server. The trade-off is that Outlook and the task pane must remain open for the scheduled send to execute.

**Q: Does the signature auto-fetch feature access my mailbox?**
A: The auto-fetch uses the Microsoft Graph API to retrieve your configured email signature from your Exchange Online settings. This is the same API used to display your signature in Outlook on the Web. MailMerge-Pro only reads the signature — it does not modify it.

**Q: Is the multi-language feature sending my data to a translation service?**
A: No. All translations are pre-built and bundled into the add-in's static JavaScript files. No data is sent to Google Translate, DeepL, or any other translation service. The language selection only changes which pre-built UI strings are displayed.
