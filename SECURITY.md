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
| User's OAuth token | Browser (MSAL.js) | Browser localStorage | Microsoft Azure AD only |
| Campaign history | Browser (JavaScript) | Browser localStorage | ❌ Nowhere |
| Merge field values | Browser (JavaScript) | Browser memory only | Microsoft Exchange (merged into email) |

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

---

## Authentication Security

| Aspect | Details |
|---|---|
| **Protocol** | OAuth 2.0 with PKCE (Proof Key for Code Exchange) |
| **Library** | MSAL.js 2.0 (Microsoft Authentication Library) |
| **Token storage** | Browser localStorage (encrypted by MSAL) |
| **Token lifetime** | Access token: ~1 hour; Refresh token: ~24 hours |
| **Login flow** | Popup window → Microsoft login page → token returned to add-in |
| **Credentials** | MailMerge-Pro NEVER sees or handles user passwords |
| **Scopes** | Only `Mail.Send`, `Mail.ReadWrite`, `User.Read` — minimum required |

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
| `alcdn.msauth.net` | Load MSAL.js library | None (downloads JavaScript library) |
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
alcdn.msauth.net
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
| Auth method | MSAL.js (OAuth 2.0) | OAuth 2.0 | Desktop COM (no OAuth) |

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
