# MailMerge-Pro — Admin Deployment Guide

> Free mail merge add-in for Microsoft Outlook with CC, BCC, per-recipient attachments, and personalized bulk email.

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Step 1: Azure AD App Registration](#2-step-1-azure-ad-app-registration)
3. [Step 2: Host the Add-in Files](#3-step-2-host-the-add-in-files)
4. [Step 3: Update the Manifest](#4-step-3-update-the-manifest)
5. [Step 4: Deploy via M365 Admin Center](#5-step-4-deploy-via-m365-admin-center)
6. [Step 5: Verify & Test](#6-step-5-verify--test)
7. [User Guide](#7-user-guide)
8. [Troubleshooting](#8-troubleshooting)
9. [Security & Compliance](#9-security--compliance)
10. [Feature List](#10-feature-list)

---

## 1. Prerequisites

| Requirement | Details |
|---|---|
| **Microsoft 365 tenant** | Business Basic or higher (Exchange Online required) |
| **Admin access** | Global Admin or Exchange Admin role |
| **Azure AD access** | To register the app (same admin account works) |
| **Web hosting (HTTPS)** | GitHub Pages (free), Azure Static Web Apps (free), or any HTTPS server |

**Estimated setup time:** 10-15 minutes

---

## 2. Step 1: Azure AD App Registration

This is a **one-time setup** per tenant.

### 2.1 Create the App

1. Go to **https://entra.microsoft.com**
2. Navigate to: **Identity** → **Applications** → **App registrations**
3. Click **"+ New registration"**
4. Fill in:
   - **Name:** `MailMerge-Pro`
   - **Supported account types:** "Accounts in this organizational directory only" (single tenant)
   - **Redirect URI:**
     - Platform: **Single-page application (SPA)**
     - URI: `https://YOUR-DOMAIN/mailmerge-pro/taskpane.html`
       - For GitHub Pages: `https://USERNAME.github.io/mailmerge-pro/taskpane.html`
       - For Azure: `https://YOURAPP.azurestaticapps.net/taskpane.html`
5. Click **Register**
6. **Add brk-multihub Redirect URI (required for NAA):**
   - After registration, go to **Authentication** in the left sidebar
   - Under "Single-page application," click **"Add URI"**
   - Add: `brk-multihub://YOUR-DOMAIN` (e.g., `brk-multihub://itsjamessmith.github.io`)
   - Click **Save**
   - This tells Microsoft that Outlook can broker authentication for the add-in via Nested App Authentication (NAA). Without this URI, NAA will not work.

### 2.2 Copy the IDs

From the app's **Overview** page, copy:
- **Application (client) ID** — e.g., `360e4343-614f-4f70-a650-c020868516fc`
- **Directory (tenant) ID** — e.g., `e67c588e-f654-4727-b794-1ca5df7b6ee9`

### 2.3 Add API Permissions

1. Go to **API permissions** (left sidebar)
2. Click **"+ Add a permission"**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Add these permissions:

| Permission | Why |
|---|---|
| `Mail.Send` | Send emails on behalf of the user |
| `Mail.ReadWrite` | Create drafts, add attachments |
| `User.Read` | Get user's email for test email feature |
| `Contacts.Read` | (Optional) Import contacts from address book |

5. Click **"Grant admin consent for [Your Organization]"**
6. Verify all permissions show ✅ "Granted"

### 2.4 (Optional) Multi-Tenant Setup

If you want ANY M365 tenant to use this add-in (not just yours):

1. In App registration → **Authentication**
2. Change "Supported account types" to **"Accounts in any organizational directory"**
3. In the manifest and JS code, change the authority URL from:
   ```
   https://login.microsoftonline.com/YOUR-TENANT-ID
   ```
   to:
   ```
   https://login.microsoftonline.com/common
   ```

---

## 3. Step 2: Host the Add-in Files

The add-in consists of static files (HTML, JS, CSS, images) that must be hosted on HTTPS.

### Option A: GitHub Pages (Free — Recommended)

1. Create a GitHub repository (e.g., `mailmerge-pro`)
2. Copy these files to the repo:
   ```
   index.html
   taskpane.html
   taskpane.js
   taskpane.css
   function-file.html
   assets/icon-16.png
   assets/icon-32.png
   assets/icon-64.png
   assets/icon-80.png
   assets/icon-128.png
   ```
3. Push to `main` branch
4. Enable GitHub Pages: Repo → Settings → Pages → Source: `main` / `root`
5. Your URL: `https://USERNAME.github.io/mailmerge-pro/`

### Option B: Azure Static Web Apps (Free Tier)

1. Create a Static Web App in Azure Portal
2. Upload the same files
3. Your URL: `https://YOURAPP.azurestaticapps.net/`

### Option C: Company Web Server / IIS

1. Create an HTTPS-enabled virtual directory
2. Copy all files to the directory
3. Ensure MIME types are configured for `.js`, `.css`, `.svg`, `.png`

---

## 4. Step 3: Update the Manifest

Edit `manifest.xml` and replace these values with YOUR details:

### 4.1 Replace URLs

Find and replace ALL occurrences of the hosting URL:
```
FIND:    https://itsjamessmith.github.io/mailmerge-pro/
REPLACE: https://YOUR-HOSTING-URL/
```

This affects: `IconUrl`, `HighResolutionIconUrl`, `SupportUrl`, `SourceLocation`, `bt:Image`, `bt:Url`

### 4.2 Update the App ID (Optional)

If you want a unique add-in ID per tenant, generate a new GUID:
```powershell
[guid]::NewGuid().ToString()
```
Replace the `<Id>` value in manifest.xml.

### 4.3 Update the JavaScript

Edit `taskpane.js` and update the MSAL configuration (lines 8-19):
```javascript
const msalConfig = {
    auth: {
        clientId: "YOUR-CLIENT-ID-HERE",
        authority: "https://login.microsoftonline.com/YOUR-TENANT-ID-HERE",
        redirectUri: "https://YOUR-HOSTING-URL/taskpane.html"
    },
    cache: { cacheLocation: "sessionStorage" }
};
```

> **Note:** MailMerge-Pro uses **MSAL v3.27.0** (loaded from jsDelivr CDN) and **Nested App Authentication (NAA)** via `createNestablePublicClientApplication`. The MSAL cache uses `sessionStorage` (not `localStorage`) so tokens are automatically cleared when the browser tab closes. This is a security improvement — tokens cannot be stolen by other same-origin pages.

### 4.4 Validate the Manifest

```powershell
npm install -g office-addin-manifest
npx office-addin-manifest validate manifest.xml
```
Ensure it says **"The manifest is valid."**

---

## 5. Step 4: Deploy via M365 Admin Center

### 5.1 Centralized Deployment (Recommended)

1. Go to **https://admin.microsoft.com**
2. Navigate to: **Settings** → **Integrated apps**
3. Click **"Upload custom apps"**
4. Select **"I have the manifest file (.xml) on this device"**
5. Click **"Choose File"** → select your `manifest.xml`
6. Click **Next**
7. **Assign users:**
   - "Entire organization" — for all users
   - Or select specific users/groups
8. Click **Deploy**
9. Wait 5-10 minutes for propagation

### 5.2 Verify Deployment

1. Open **Outlook on the Web** (outlook.office.com)
2. Click on any email or compose new
3. Look for **"Mail Merge"** button in ribbon or under **Apps**
4. Click it → task pane opens → click **Sign In**

### 5.3 Alternative: User Self-Install (Sideload)

For individual user testing without admin deployment:

1. Go to **https://aka.ms/olksideload**
2. Sign in with your M365 account
3. **My add-ins** → **Custom Add-ins** → **"+ Add a custom add-in"** → **"Add from file"**
4. Upload `manifest.xml`

**Note:** Sideloading must be enabled in Exchange. If blocked:
```powershell
Connect-ExchangeOnline
New-ManagementRoleAssignment -Role "My Custom Apps" -User "user@domain.com"
```

---

## 6. Step 5: Verify & Test

### Test Checklist

| # | Test | Expected Result |
|---|---|---|
| 1 | Open add-in in Outlook Web | Task pane loads, "Sign In" visible |
| 2 | Click Sign In | Microsoft login popup → signed in |
| 3 | Upload sample Excel | Data preview shows rows and columns |
| 4 | Map columns | Auto-detects Email, CC, BCC, Subject |
| 5 | Compose email with {FirstName} | Merge fields insert correctly |
| 6 | Click "Test Email" | Sends first row to your own inbox |
| 7 | Send with "Save as Draft" checked | Emails appear in Drafts folder |
| 8 | Send with per-recipient attachment | Correct file attached per recipient |
| 9 | Send with global + per-recipient attachment | Both files attached |
| 10 | Send to internal recipient | Email received successfully |
| 11 | Open in Outlook Desktop (classic) | Add-in loads in task pane |
| 12 | Open in New Outlook (Monarch) | Add-in available under Apps |
| 13 | Save and load a custom email template | Template saves to localStorage and loads into composer |
| 14 | Schedule a send for 5 minutes in the future | Countdown appears; emails send at scheduled time |
| 15 | Enable email tracking (read receipts) | isReadReceiptRequested flag set on sent emails |
| 16 | Create an A/B test with 50/50 split | Both versions sent; results summary shows per-version stats |
| 17 | Save and load a contact group | Group saves to localStorage and loads into data table |
| 18 | Import an HTML template file | HTML renders in editor; merge fields auto-detected |
| 19 | Fetch signature from Graph API | Signature appears in panel and appends to emails |
| 20 | Verify rate limit dashboard | Daily counter and color-coded bar display correctly |
| 21 | Switch language to Spanish | All UI labels switch to Spanish; preference persists |
| 22 | View the local admin dashboard | Stats, charts, and recent campaigns display correctly |

### Sample Excel for Testing

| FirstName | LastName | Email | CC | BCC | Subject | Attachments |
|---|---|---|---|---|---|---|
| John | Doe | john@yourdomain.com | manager@yourdomain.com | | Welcome John! | john_report.pdf |
| Jane | Smith | jane@yourdomain.com | | audit@yourdomain.com | Welcome Jane! | jane_report.pdf;benefits.pdf |

---

## 7. User Guide

### For End Users

1. **Open Outlook** (Web, Windows, or Mac)
2. Click on any email → click **"Mail Merge"** in the ribbon (or find it under **Apps**)
3. **Step 1 — Data:** Upload your Excel/CSV file
4. **Step 2 — Map:** Verify column mapping (auto-detected)
5. **Step 3 — Compose:** Write your email, use `{ColumnName}` for personalization
6. **Step 4 — Send:** Review, test, then send

### Tips for Users

- **Always test first** — use "Test Email" or "Save as Draft" before bulk sending
- **Merge fields** — click the buttons above the editor to insert `{ColumnName}`
- **Attachments** — upload files in Step 3, filenames must match your spreadsheet
- **Delay** — set 2-3 seconds between emails to avoid Exchange throttling
- **From alias** — enter an alias email if you have Send As permissions

---

## 8. Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| "Auth library not loaded" | MSAL CDN blocked by proxy/firewall | Whitelist `cdn.jsdelivr.net` |
| "Pop-up blocked" | Browser blocking sign-in popup (legacy — NAA eliminates this in most cases) | Allow popups for `login.microsoftonline.com` |
| "Permission denied" (403) | Admin consent not granted | Re-grant admin consent in Azure AD |
| "Session expired" (401) | Token expired | Click Sign In again |
| "Rate limited" (429) | Sending too fast | Increase delay to 3-5 seconds |
| "550 5.7.501 Spam detected" | Exchange spam filter | Check Security center → Restricted entities |
| Add-in not visible | Not deployed to user | Deploy via M365 admin center to the user |
| Per-recipient attachment missing | Filename mismatch | Ensure spreadsheet filenames match uploaded file names (just filename, not full path) |
| Sign-in redirect error | Wrong redirect URI | Verify SPA redirect URI in Azure AD matches your hosting URL exactly |
| Scheduled send didn't execute | Outlook or task pane was closed | Outlook and the add-in task pane must remain open for scheduled sends to execute |

### Required URLs to Whitelist (Firewall/Proxy)

| URL | Purpose |
|---|---|
| `login.microsoftonline.com` | MSAL authentication |
| `graph.microsoft.com` | Microsoft Graph API |
| `cdn.jsdelivr.net` | MSAL.js v3 CDN (jsDelivr) |
| `YOUR-HOSTING-URL` | Add-in HTML/JS/CSS |
| `appsforoffice.microsoft.com` | Office.js library |
| `cdn.sheetjs.com` | SheetJS Excel parser |

---

## 9. Security & Compliance

### Data Privacy

| Aspect | Details |
|---|---|
| **Data processing** | 100% client-side — no data leaves the browser/Outlook |
| **No backend server** | Static files only — no server-side processing |
| **Authentication** | MSAL.js v3.27.0 with Nested App Authentication (NAA) — Microsoft-recommended OAuth 2.0 for Office add-ins |
| **Token storage** | `sessionStorage` — tokens cleared when browser tab closes (security improvement over localStorage) |
| **XSS protection** | `sanitizeHtml()` strips script tags, iframes, event handlers, javascript: URLs from all HTML content |
| **Email sending** | Through user's own Exchange Online mailbox via Graph API |
| **Spreadsheet data** | Read in-browser by SheetJS — never uploaded anywhere |
| **Attachments** | Read in-browser as base64 — sent directly via Graph API |
| **No tracking** | No analytics, telemetry, or usage tracking |
| **No third-party services** | Only Microsoft services (Azure AD, Graph API) |

### Compliance

- ✅ Data stays within your Microsoft 365 tenant
- ✅ Emails go through normal Exchange Online compliance (DLP, retention, journal)
- ✅ Admin can revoke access by removing the app registration or add-in deployment
- ✅ Audit logs: all emails appear in Sent Items and Exchange message trace
- ✅ Compatible with GDPR, HIPAA (when M365 is configured for compliance)

### Local Storage Usage (v3.0)

Several v3.0 features use browser `localStorage` to persist data on the user's device:

| Data | localStorage Key | Purpose | Synced Across Devices? |
|---|---|---|---|
| Email templates | `mailmergepro_templates` | Custom saved templates | ❌ No — device-only |
| Contact groups | `mailmergepro_groups` | Saved recipient lists | ❌ No — device-only |
| Campaign history | `mailmergepro_campaigns` | Past campaign records and dashboard data | ❌ No — device-only |
| Language preference | `mailmergepro_lang` | User's selected UI language | ❌ No — device-only |
| Signature | `mailmergepro_signature` | Cached email signature | ❌ No — device-only |

**Important notes:**
- All localStorage data resides **only on the user's device** and specific browser profile. It is NOT synced to the cloud, NOT accessible by admins, and NOT visible to other users.
- Clearing browser data (cache/cookies) will erase all localStorage items.
- Scheduled sends require Outlook and the add-in task pane to remain open until the send time.

### Permissions Used

| Permission | Type | Purpose | Sensitivity |
|---|---|---|---|
| `Mail.Send` | Delegated | Send emails on behalf of user | Medium |
| `Mail.ReadWrite` | Delegated | Create drafts, add attachments | Medium |
| `User.Read` | Delegated | Get user's email address | Low |
| `Contacts.Read` | Delegated (optional) | Import address book contacts | Low |

All permissions are **delegated** (act as the signed-in user), not application-level.

---

## 10. Feature List

### Core Features (44 total)

| # | Feature | Description |
|---|---|---|
| 1 | Excel/CSV upload | Parse spreadsheets client-side |
| 2 | Auto column detection | Finds To, CC, BCC, Subject, Attachments columns |
| 3 | Personalized email body | `{ColumnName}` merge fields |
| 4 | Personalized subject | Merge fields in subject line |
| 5 | Per-recipient CC | From spreadsheet column |
| 6 | Per-recipient BCC | From spreadsheet column |
| 7 | Global CC | Applied to all emails |
| 8 | Global BCC | Applied to all emails |
| 9 | HTML formatted emails | Rich text with formatting |
| 10 | Send from alias | Send As / Send on Behalf |
| 11 | Draft/Preview mode | Save to Drafts without sending |
| 12 | Test email | Send first row to yourself |
| 13 | Progress bar | Real-time sending progress |
| 14 | Error reporting | Per-recipient status with details |
| 15 | Send throttling | Configurable delay between emails |
| 16 | Per-recipient attachments | Different files per recipient |
| 17 | Global attachments | Same files for all recipients |
| 18 | Contacts import | Load from Outlook address book |
| 19 | Rich text editor | WYSIWYG with formatting toolbar |
| 20 | Email preview carousel | Cycle through all recipients |
| 21 | Fallback default values | Defaults for empty merge fields |
| 22 | Read receipts | Request read confirmation |
| 23 | High importance flag | Mark emails as important |
| 24 | Unsubscribe header | List-Unsubscribe compliance |
| 25 | Shared mailbox sending | Send from shared mailboxes |
| 26 | Many-to-one merge | Group rows by email |
| 27 | Campaign history | Save and view past campaigns |
| 28 | Export results CSV | Download send results |
| 29 | Onboarding | Welcome guide on first launch |
| 30 | Step validation | Green badges on completed steps |
| 31 | Recipient search | Filter data table |
| 32 | Dark mode | System and Office theme support |
| 33 | Responsive design | Works in narrow task panes |
| 34 | Keyboard shortcuts | Ctrl+Enter to send |
| 35 | Email templates library | 3 built-in + save/load/delete custom templates (localStorage) |
| 36 | Scheduled sending | Date/time picker, countdown timer, cancel button |
| 37 | Email tracking | Read receipts via Graph API isReadReceiptRequested flag |
| 38 | A/B testing | Tabbed editors (A/B), configurable split ratio, results per version |
| 39 | Contact groups/segments | Save/load/delete/merge recipient lists (localStorage) |
| 40 | HTML template import | File picker + drag-and-drop .html files with auto-detection |
| 41 | Signature auto-insert | Fetches from Graph API or manual paste, auto-append toggle |
| 42 | Rate limit dashboard | Daily send counter, color-coded bar, auto-suggested delay |
| 43 | Multi-language (i18n) | 6 languages: EN, ES, FR, DE, PT, JA with header selector |
| 44 | Local admin dashboard | Campaign stats, top recipients, monthly chart, success rate |

---

## Quick Start Summary

```
1. Register Azure AD app (entra.microsoft.com)
   → Get Client ID + Tenant ID

2. Host files on HTTPS (GitHub Pages / Azure / IIS)
   → Update URLs in manifest.xml + taskpane.js

3. Deploy via M365 admin center
   → Settings → Integrated apps → Upload manifest.xml

4. Users open Outlook → Apps → Mail Merge → Sign In → Send!
```

**Total setup time: ~15 minutes**

---

*© 2026 MailMerge-Pro. All rights reserved.*
