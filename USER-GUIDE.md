# MailMerge-Pro User Guide

> **Version:** 3.0 | **Platform:** Outlook 365 (Web & Desktop) | **License:** Free / Pro / Enterprise | **Features:** 44

---

## Quick Start in 60 Seconds

1. **Open Outlook** → Click **Get Add-ins** → Search **MailMerge-Pro** → **Add**
2. Open a **New Email** → Click the **MailMerge-Pro** icon in the ribbon
3. **Upload** an Excel file with columns like `Email`, `FirstName`, `Company`
4. **Compose** your email using merge fields: `Hello {FirstName}, welcome to {Company}!`
5. Click **Preview** to verify personalization → Click **Send All**
6. **New in v3.0:** Save email templates, schedule sends, A/B test subject lines, and view your dashboard — all from the same task pane.

That's it — personalized emails sent to every row in your spreadsheet.

---

## Table of Contents

- [1. Getting Started](#1-getting-started)
- [2. Data Source](#2-data-source)
- [3. Email Composition](#3-email-composition)
- [4. Recipients](#4-recipients)
- [5. Attachments](#5-attachments)
- [6. Sending Options](#6-sending-options)
- [7. Preview & Testing](#7-preview--testing)
- [8. Advanced Features](#8-advanced-features)
- [9. UI Features](#9-ui-features)
- [10. New in v3.0](#10-new-in-v30)
  - [10.1 Email Templates Library](#101-email-templates-library)
  - [10.2 Scheduled Sending](#102-scheduled-sending)
  - [10.3 Email Tracking](#103-email-tracking)
  - [10.4 A/B Testing](#104-ab-testing)
  - [10.5 Contact Groups & Segments](#105-contact-groups--segments)
  - [10.6 HTML Template Import](#106-html-template-import)
  - [10.7 Signature Auto-Insert](#107-signature-auto-insert)
  - [10.8 Rate Limit Dashboard](#108-rate-limit-dashboard)
  - [10.9 Multi-Language (i18n)](#109-multi-language-i18n)
  - [10.10 Local Admin Dashboard](#1010-local-admin-dashboard)

---

## 1. Getting Started

### 1.1 Sign-In

**What it does:** Authenticates you with your Microsoft 365 account so MailMerge-Pro can send emails on your behalf using the Outlook API.

**Authentication method:** MailMerge-Pro uses **Nested App Authentication (NAA)** — the Microsoft-recommended approach for Office add-ins (2025+). NAA enables seamless single sign-on (SSO) inside Outlook's task pane without popup windows or redirects. Under the hood, MSAL v3 calls `createNestablePublicClientApplication` so Outlook can broker authentication on behalf of the add-in. This eliminates common issues like "popup blocked" and "redirect_in_iframe" errors.

**Step-by-step:**

1. Open **Outlook** (web at outlook.office.com or the desktop app).
2. Compose a **New Email** (or open an existing draft).
3. Click the **MailMerge-Pro** icon in the ribbon toolbar. On Outlook Web, look under the **"…" (More actions)** menu if the icon isn't visible.
4. The add-in task pane opens on the right side of the compose window.
5. Click the **"Sign in with Microsoft"** button.
6. NAA authenticates you seamlessly through Outlook — in most cases, you're signed in automatically without a popup. If prompted, select your **Microsoft 365 account** or enter your credentials.
7. On first sign-in, you'll see a **permissions consent screen**. Review the permissions:
   - *Send mail as you* — required to send merge emails
   - *Read your profile* — to display your name and email
8. Click **Accept**.
9. The task pane updates to show your name and email address at the top.

**Tips & Gotchas:**
- NAA provides seamless SSO — most users are signed in automatically without a popup.
- Admin-consented tenants skip the individual consent step.
- If you see "Need admin approval," contact your IT administrator to grant tenant-wide consent.
- Sign-in persists across sessions — you won't need to re-authenticate each time.
- Authentication tokens are stored in **sessionStorage** (not localStorage), which means they are automatically cleared when the browser tab is closed for added security.

### 1.1.1 Security Features

MailMerge-Pro includes several built-in security protections:

- **XSS Protection** — A `sanitizeHtml()` function strips script tags, iframes, event handlers, and `javascript:` URLs from all HTML content (templates, signatures, HTML imports). Merge field values are HTML-escaped when building emails. Link insertion validates URL schemes (blocks `javascript:`, `data:`, `vbscript:`).
- **Session-scoped tokens** — MSAL tokens are stored in `sessionStorage`, not `localStorage`. Tokens are automatically cleared when the browser tab closes, preventing token theft from other same-origin pages.
- **Outlook-only execution** — The add-in refuses to run outside of Outlook, blocking standalone browser access.
- **Clean sign-out** — Signing out clears all personally identifiable information (PII) from localStorage.

---

### 1.2 First-Time Onboarding

**What it does:** Walks new users through the core workflow with an interactive guided tour.

**Step-by-step:**

1. After your first successful sign-in, an **onboarding overlay** appears automatically.
2. The overlay highlights each step of the mail merge process:
   - **Step 1 badge** — "Upload your data" (the Data Source panel)
   - **Step 2 badge** — "Compose your message" (the Email Editor)
   - **Step 3 badge** — "Preview & Send" (the Preview panel)
3. Click **Next** to advance through each tooltip, or click **Skip Tour** to dismiss.
4. A sample Excel file link is provided: click **"Download sample spreadsheet"** to get a pre-built demo file.

**Tips & Gotchas:**
- You can re-trigger the tour anytime from **Settings (⚙) → Show Onboarding Tour**.
- The sample spreadsheet contains 5 rows with `Email`, `FirstName`, `LastName`, `Company`, and `Amount` columns — perfect for testing.

---

## 2. Data Source

### 2.1 Excel / CSV Upload

**What it does:** Imports your recipient list from an Excel (.xlsx, .xls) or CSV (.csv) file. Each row becomes one email recipient. Each column becomes a merge field.

**Step-by-step:**

1. In the task pane, click the **"Upload Data"** button (Step 1 section) or drag-and-drop a file onto the upload area.
2. Select your file from the file picker dialog. Supported formats:
   - `.xlsx` (Excel 2007+)
   - `.xls` (Excel 97-2003)
   - `.csv` (Comma-Separated Values, UTF-8 recommended)
3. The file uploads and parses. A **data preview table** appears showing the first 5 rows.
4. Above the table, you see the detected columns with chip badges (e.g., `Email`, `FirstName`, `Company`).
5. The row count displays: **"48 recipients loaded"**.

**Example Data (Excel):**

| Email | FirstName | LastName | Company | InvoiceAmount |
|---|---|---|---|---|
| alice@contoso.com | Alice | Johnson | Contoso Ltd | $1,200.00 |
| bob@fabrikam.com | Bob | Smith | Fabrikam Inc | $3,450.00 |
| carol@northwind.com | Carol | Lee | Northwind Traders | $780.00 |

**Tips & Gotchas:**
- The **first row must be headers**. MailMerge-Pro uses these as merge field names.
- Avoid special characters in column headers (use `FirstName` not `First Name!`). Spaces are OK but avoid `{`, `}`, `#`, `/`.
- Maximum file size: **10 MB**. For larger lists, split into multiple files.
- CSV files must use UTF-8 encoding to properly handle accented characters (é, ñ, ü).
- If your Excel file has multiple sheets, MailMerge-Pro reads the **first sheet** only.
- Blank rows at the bottom of the spreadsheet are automatically skipped.

---

### 2.2 Column Mapping

**What it does:** Maps your spreadsheet columns to their functional role (e.g., which column contains the email address, which contains the CC list).

**Step-by-step:**

1. After uploading, the **Column Mapping** panel appears below the data preview.
2. You'll see dropdowns for key fields:
   - **Email Address** — Select the column containing recipient email addresses (required).
   - **Display Name** — Select the column for the recipient's display name (optional).
   - **CC** — Select a column containing per-recipient CC addresses (optional).
   - **BCC** — Select a column containing per-recipient BCC addresses (optional).
   - **Attachment Filename** — Select a column containing attachment filenames (optional).
3. Click each dropdown and choose the matching column from your spreadsheet.
4. Unmapped columns remain available as merge fields `{ColumnName}` in your email body.

**Example:**
- Your spreadsheet has columns: `EmailAddr`, `Name`, `Dept`, `CopyTo`
- Map: **Email Address** → `EmailAddr`, **Display Name** → `Name`, **CC** → `CopyTo`
- `Dept` is still available as `{Dept}` in your email.

**Tips & Gotchas:**
- If no Email column is mapped, the **Send** button is disabled.
- You can change mappings at any time — the preview updates instantly.
- If a CC or BCC column contains multiple addresses, separate them with semicolons: `john@co.com; jane@co.com`.

---

### 2.3 Auto-Detection

**What it does:** Automatically detects which column is the email address, display name, etc., based on column header names and data patterns.

**Step-by-step:**

1. Upload your file — auto-detection runs immediately.
2. Columns named `Email`, `E-mail`, `EmailAddress`, or `email_address` are auto-mapped to **Email Address**.
3. Columns named `Name`, `FullName`, `DisplayName`, `First Name` are auto-mapped to **Display Name**.
4. Columns named `CC`, `Copy`, `CopyTo` are auto-mapped to **CC**.
5. Columns named `BCC`, `BlindCopy` are auto-mapped to **BCC**.
6. Columns named `Attachment`, `File`, `FileName`, `AttachmentFile` are auto-mapped to **Attachment Filename**.
7. A green check (✓) appears next to auto-detected mappings. A yellow warning (⚠) appears if a mapping couldn't be auto-detected.
8. Review the auto-detected mappings and correct any mistakes using the dropdowns.

**Tips & Gotchas:**
- Auto-detection also scans the first 5 data rows. If a column contains mostly `@` symbols and valid email patterns, it's detected as the Email column even if the header doesn't say "Email."
- Auto-detection is a suggestion — always review before sending.

---

### 2.4 Contacts Import

**What it does:** Imports recipients directly from your Outlook contacts or address book, so you don't need a spreadsheet.

**Step-by-step:**

1. In the Data Source panel, click the **"Import from Contacts"** button (people icon).
2. A contact picker dialog opens showing your Outlook contacts.
3. Browse or search for contacts. You can select:
   - **Individual contacts** by clicking each one.
   - **Contact groups / distribution lists** by selecting the group name.
4. Click **"Import Selected"** to add them to your merge list.
5. Imported contacts appear in the data preview table with columns: `Email`, `FirstName`, `LastName`, `Company`, `JobTitle`.

**Tips & Gotchas:**
- Contacts without an email address are skipped with a warning.
- You can combine contacts import with a spreadsheet upload — the data merges together.
- Contact group members are expanded into individual rows.

---

### 2.5 Search & Filter

**What it does:** Lets you search and filter your loaded recipients so you can quickly find specific rows or send to a subset.

**Step-by-step:**

1. After loading data, a **search bar** appears above the data preview table.
2. Type a search term (e.g., `contoso`) — the table filters to matching rows in real-time.
3. The search checks **all columns**, not just the email column.
4. Use the **column filter dropdowns** (funnel icon on each column header) to filter by specific values:
   - Click the funnel icon on the `Company` column.
   - Check/uncheck values: ☑ Contoso Ltd ☐ Fabrikam Inc ☑ Northwind Traders.
   - Click **Apply**.
5. The recipient count updates to reflect the filter: **"24 of 48 recipients selected"**.
6. Only filtered/selected recipients receive the merge email.
7. Click **"Clear Filters"** to reset.

**Tips & Gotchas:**
- Filters are AND-based: filtering `Company = Contoso` and searching `alice` shows only Alice at Contoso.
- Use the **Select All / Deselect All** checkboxes to quickly toggle.
- You can manually uncheck individual rows using the checkbox in the first column.

---

## 3. Email Composition

### 3.1 Rich Text Editor

**What it does:** Provides a WYSIWYG email editor with formatting options — bold, italic, links, images, lists, and more.

**Step-by-step:**

1. Click in the **Email Body** editor (Step 2 section of the task pane).
2. Use the **formatting toolbar** above the editor:
   - **B** — Bold (Ctrl+B)
   - *I* — Italic (Ctrl+I)
   - **U** — Underline (Ctrl+U)
   - **Link** (🔗) — Insert hyperlink
   - **Image** (🖼) — Insert inline image
   - **Bulleted List** — Unordered list
   - **Numbered List** — Ordered list
   - **Font Size** — Dropdown (10pt, 12pt, 14pt, etc.)
   - **Font Color** — Color picker
   - **Highlight** — Background highlight color
3. Type your email content. The editor supports full HTML formatting.
4. To insert an image, click the image icon and paste a URL or upload a file.

**Example email body:**

```
Dear {FirstName},

Thank you for your recent purchase of **{ProductName}**.

Your order total was **{OrderTotal}** and it will ship to:
{ShippingAddress}

Best regards,
The Sales Team
```

**Tips & Gotchas:**
- The editor renders exactly as the recipient will see it in their inbox.
- Copy-paste from Word or Google Docs preserves most formatting.
- Avoid pasting from Notepad — it strips formatting. Use Shift+Ctrl+V for plain text paste.
- Images are embedded as URLs, not attachments. Ensure image URLs are publicly accessible.

---

### 3.2 Merge Fields `{ColumnName}`

**What it does:** Inserts personalized placeholders that get replaced with each recipient's data when sending.

**Step-by-step:**

1. In the email body editor, place your cursor where you want to insert a merge field.
2. Type `{` (opening brace) — an **autocomplete dropdown** appears showing available columns from your spreadsheet.
3. Select a column name (e.g., `FirstName`) from the dropdown, or type the full field name.
4. The merge field appears as a styled chip/badge in the editor: `{FirstName}`.
5. Alternatively, click the **"Insert Field"** button (fx icon) in the toolbar and select from the list.

**Example:**
- Spreadsheet column: `FirstName` with value `Alice`
- Email text: `Hello {FirstName}, your invoice is ready.`
- Sent email: `Hello Alice, your invoice is ready.`

**Tips & Gotchas:**
- Field names are **case-sensitive**: `{firstname}` won't match a column named `FirstName`.
- If a merge field doesn't match any column, it appears literally as `{UnknownField}` in the sent email — always preview first!
- You can use merge fields in the **subject line** too (see 3.3).
- Merge fields work inside hyperlinks: `https://portal.example.com/user/{UserID}`.

---

### 3.3 Personalized Subject Line

**What it does:** Allows merge fields in the email subject line so each recipient sees a customized subject.

**Step-by-step:**

1. Click in the **Subject** field at the top of the compose area.
2. Type your subject with merge fields, e.g.: `Invoice #{InvoiceNumber} for {Company}`
3. The autocomplete dropdown works in the subject field too — type `{` to trigger it.

**Example:**
- Subject template: `Your {Month} report is ready, {FirstName}`
- For Alice in January: `Your January report is ready, Alice`
- For Bob in February: `Your February report is ready, Bob`

**Tips & Gotchas:**
- Keep subject lines under 60 characters for best display on mobile.
- Test with the Preview feature to verify subject personalization before sending.
- Emoji in subjects are supported: `🎉 Welcome aboard, {FirstName}!`

---

### 3.4 Fallback Defaults

**What it does:** Defines default values for merge fields when a cell in the spreadsheet is blank. Prevents embarrassing emails like "Hello , your order is ready."

**Step-by-step:**

1. Click the **"Fallback Defaults"** button (⚙ icon next to the "Insert Field" button) or go to **Settings → Fallback Defaults**.
2. A panel shows all merge fields from your data with a text input next to each.
3. Enter a default value for any field:
   - `FirstName` → `Valued Customer`
   - `Company` → `your organization`
   - `InvoiceAmount` → `(amount pending)`
4. Click **Save Defaults**.
5. Now if Alice's `FirstName` cell is blank, the email reads: `Hello Valued Customer,` instead of `Hello ,`.

**Example:**

| Merge Field | Default Value | Used When |
|---|---|---|
| `{FirstName}` | `Valued Customer` | FirstName cell is empty |
| `{Company}` | `your organization` | Company cell is empty |
| `{Discount}` | `10%` | Discount cell is empty |

**Tips & Gotchas:**
- Fallback defaults apply to both the subject line and the email body.
- A cell containing only spaces is treated as blank — the fallback activates.
- Defaults are saved per session (cleared when you close the add-in). Save your defaults to a template if needed.

---

## 4. Recipients

### 4.1 Per-Recipient CC

**What it does:** Sends a CC (carbon copy) to different addresses for each recipient, based on a column in your spreadsheet.

**Step-by-step:**

1. Add a column to your spreadsheet (e.g., `CC` or `CopyTo`) with the CC addresses for each row.
2. Upload the spreadsheet in MailMerge-Pro.
3. In **Column Mapping**, set **CC** → your CC column (e.g., `CopyTo`).
4. Compose and send as usual. Each email includes the row's CC address.

**Example Data:**

| Email | FirstName | CopyTo |
|---|---|---|
| alice@contoso.com | Alice | alice.manager@contoso.com |
| bob@fabrikam.com | Bob | bob.manager@fabrikam.com; finance@fabrikam.com |
| carol@northwind.com | Carol | *(blank — no CC)* |

- Alice's email CCs her manager.
- Bob's email CCs his manager AND finance (separated by `;`).
- Carol's email has no CC.

**Tips & Gotchas:**
- Multiple CC addresses in one cell: separate with semicolons (`;`).
- Blank CC cells simply send without CC — no errors.
- CC recipients see the original "To" recipient and all other CC addresses.

---

### 4.2 Per-Recipient BCC

**What it does:** Same as per-recipient CC, but recipients on the BCC line are hidden from the To and CC recipients.

**Step-by-step:**

1. Add a `BCC` column to your spreadsheet.
2. In **Column Mapping**, set **BCC** → your BCC column.
3. Compose and send. Each email's BCC is set per-row.

**Example Data:**

| Email | FirstName | BCC |
|---|---|---|
| vendor@supplier.com | Vendor Corp | legal@mycompany.com |
| client@bigco.com | BigCo | compliance@mycompany.com |

**Tips & Gotchas:**
- BCC recipients cannot see each other or the other CC/BCC addresses.
- Useful for compliance, auditing, or CRM auto-capture.

---

### 4.3 Global CC / BCC

**What it does:** Adds the same CC or BCC address(es) to **every** email in the merge — without needing a spreadsheet column.

**Step-by-step:**

1. In the **Sending Options** panel (Step 3 section), find the **Global CC** and **Global BCC** fields.
2. Enter one or more email addresses, separated by semicolons:
   - **Global CC:** `teamlead@company.com`
   - **Global BCC:** `crm-capture@company.com; archive@company.com`
3. These addresses are added to every email in the merge, in addition to any per-recipient CC/BCC.

**Tips & Gotchas:**
- Global and per-recipient CC/BCC combine (they don't override each other).
- Use Global BCC to capture all sent emails into a CRM system (e.g., Salesforce BCC address).
- If a Global CC address is the same as the To address for a particular row, it's automatically deduplicated.

---

## 5. Attachments

### 5.1 Global Attachments

**What it does:** Attaches the same file(s) to every email in the merge.

**Step-by-step:**

1. In the **Attachments** panel, click **"Add Global Attachment"** (📎 button).
2. Select one or more files from the file picker.
3. Attached files appear as chips below the button: `Brochure.pdf ✕`, `PriceList.xlsx ✕`.
4. Click the **✕** on any chip to remove it.
5. All recipients receive these files attached to their email.

**Tips & Gotchas:**
- Maximum attachment size per email: **25 MB total** (Outlook/Exchange limit).
- Supported formats: PDF, DOCX, XLSX, PPTX, PNG, JPG, GIF, ZIP, TXT, CSV, and more.
- Executable files (.exe, .bat, .cmd, .ps1) are **blocked** by Exchange and will fail.
- If you exceed the 25 MB limit, the send fails for that recipient — others still send.

---

### 5.2 Per-Recipient Attachments

**What it does:** Attaches different files to each recipient's email based on a column in your spreadsheet that specifies the filename.

**Step-by-step:**

1. Prepare your files and name them systematically (e.g., `Invoice-Alice.pdf`, `Invoice-Bob.pdf`).
2. In your spreadsheet, add a column (e.g., `AttachmentFile`) with the filename for each row:

   | Email | FirstName | AttachmentFile |
   |---|---|---|
   | alice@contoso.com | Alice | Invoice-Alice.pdf |
   | bob@fabrikam.com | Bob | Invoice-Bob.pdf |
   | carol@northwind.com | Carol | Invoice-Carol.pdf |

3. Upload the spreadsheet in MailMerge-Pro.
4. In **Column Mapping**, set **Attachment Filename** → `AttachmentFile`.
5. Click the **"Upload Attachment Files"** button (folder icon) in the Attachments panel.
6. Select **all the attachment files** at once from the file picker (multi-select with Ctrl+Click).
7. MailMerge-Pro matches each row's `AttachmentFile` value to the uploaded files by filename.
8. Matched files show a green check (✓). Unmatched files show a red warning (⚠).
9. Send your merge — each recipient gets their specific file.

**Example filename matching:**

| Row's `AttachmentFile` value | Uploaded file | Match? |
|---|---|---|
| `Invoice-Alice.pdf` | `Invoice-Alice.pdf` | ✅ Match |
| `invoice-bob.pdf` | `Invoice-Bob.pdf` | ✅ Match (case-insensitive) |
| `Report Q3.docx` | `Report Q3.docx` | ✅ Match |
| `Missing.pdf` | *(not uploaded)* | ❌ No match — warning shown |

**Tips & Gotchas:**
- Filename matching is **case-insensitive**: `report.pdf` matches `Report.PDF`.
- Multiple attachments per recipient: put semicolon-separated filenames: `Invoice.pdf; Contract.pdf`.
- If a row's attachment file isn't found, you get a warning but can still send (that row sends without attachment).
- Per-recipient and global attachments combine — a recipient gets both.

---

### 5.3 Supported Formats

| Format | Extension | Notes |
|---|---|---|
| PDF | .pdf | Most common — universally viewable |
| Word | .docx, .doc | Office document |
| Excel | .xlsx, .xls | Spreadsheets |
| PowerPoint | .pptx, .ppt | Presentations |
| Images | .png, .jpg, .gif, .bmp | Inline or attachment |
| Archives | .zip, .7z | Compressed files |
| Text | .txt, .csv, .log | Plain text files |
| HTML | .html, .htm | Web pages |
| Calendar | .ics | Calendar invites |

**Blocked by Exchange:** `.exe`, `.bat`, `.cmd`, `.ps1`, `.vbs`, `.js`, `.msi`, `.scr`, `.com`

**Tip:** To send blocked file types, put them in a `.zip` archive first.

---

## 6. Sending Options

### 6.1 Send from Alias

**What it does:** Sends the email from a different email address (alias) configured on your Exchange account, instead of your primary address.

**Step-by-step:**

1. In the **Sending Options** panel, find the **"Send From"** dropdown.
2. Click the dropdown — it lists all aliases configured on your Exchange account (e.g., `you@company.com`, `sales@company.com`, `noreply@company.com`).
3. Select the alias you want to send from.
4. Recipients see the alias as the "From" address, and replies go to that alias.

**Tips & Gotchas:**
- Only aliases that your Exchange admin has configured for your mailbox appear in the list.
- This is NOT the same as sending from a shared mailbox (see 6.2).
- If the alias isn't in the list, ask your Exchange admin to add it to your mailbox.

---

### 6.2 Shared Mailbox

**What it does:** Sends the email from a shared mailbox (e.g., `info@company.com`, `support@company.com`) that you have "Send As" or "Send on Behalf" permissions for.

**Step-by-step:**

1. In the **Sending Options** panel, toggle **"Send from Shared Mailbox"** ON.
2. A text field appears. Enter the shared mailbox address: `support@company.com`.
3. MailMerge-Pro validates that you have permission to send from this mailbox.
4. A green check (✓) confirms permission. A red error (✗) means you lack permission.
5. Compose and send as usual. The "From" line shows the shared mailbox.

**Tips & Gotchas:**
- You need **"Send As"** permission for the email to appear as if it came directly from the shared mailbox.
- **"Send on Behalf"** permission shows: `Your Name on behalf of support@company.com`.
- Contact your Exchange admin to grant the appropriate permission.
- Sent emails appear in the **Sent Items of the shared mailbox** (with "Send As") or your personal Sent Items (with "Send on Behalf").

---

### 6.3 Read Receipts

**What it does:** Requests a read receipt from each recipient. When they open the email, their mail client prompts them to send a receipt notification back to you.

**Step-by-step:**

1. In the **Sending Options** panel, toggle **"Request Read Receipts"** ON.
2. That's it — each sent email includes the read receipt request header.
3. When a recipient opens the email, their mail client may prompt: "The sender requested a read receipt. Send one?"
4. If they click "Yes," you receive an email confirming they read your message.

**Tips & Gotchas:**
- Recipients can **decline** to send a read receipt — this is not a guaranteed tracking mechanism.
- Some organizations auto-suppress read receipt prompts.
- Read receipts appear as new emails in your inbox with the subject: `Read: [Original Subject]`.
- For reliable open tracking, consider dedicated email tracking tools. Read receipts are a lightweight option.

---

### 6.4 High Importance

**What it does:** Marks every email in the merge as "High Importance" — recipients see a red exclamation mark (❗) next to the email in their inbox.

**Step-by-step:**

1. In the **Sending Options** panel, toggle **"High Importance"** ON.
2. A red exclamation mark icon appears on the toggle to confirm.
3. All merge emails are sent with the importance flag set to High.

**Tips & Gotchas:**
- Use sparingly — frequent "High Importance" emails train recipients to ignore them.
- Good use cases: urgent invoices, time-sensitive deadlines, system outage notifications.
- There is no "Low Importance" option in MailMerge-Pro; emails default to Normal importance when the toggle is off.

---

### 6.5 Unsubscribe Header

**What it does:** Adds a `List-Unsubscribe` header to every email, which causes mail clients to show a one-click "Unsubscribe" link. This improves deliverability and compliance.

**Step-by-step:**

1. In the **Sending Options** panel, toggle **"Include Unsubscribe Header"** ON.
2. Enter the unsubscribe URL or email address:
   - **URL:** `https://yoursite.com/unsubscribe?email={Email}` (merge fields work here!)
   - **Email:** `mailto:unsubscribe@yoursite.com?subject=Unsubscribe-{Email}`
3. Recipients see an "Unsubscribe" link at the top of the email in Gmail, Outlook, and other clients.

**Tips & Gotchas:**
- The `List-Unsubscribe` header is a **best practice** for bulk email — it reduces spam complaints.
- It does NOT automatically remove people from your list. You need to handle the unsubscribe URL/email on your end.
- Gmail and Outlook prominently display the Unsubscribe link — recipients appreciate the easy opt-out.
- Required for compliance with CAN-SPAM (US) and GDPR (EU) regulations if sending marketing content.

---

### 6.6 Send Delay (Throttling)

**What it does:** Adds a delay between each email send to avoid Exchange throttling limits and to appear more like natural sending behavior.

**Step-by-step:**

1. In the **Sending Options** panel, find the **"Send Delay"** slider or input.
2. Set the delay in seconds between sends:
   - **0 seconds** — sends as fast as possible (may trigger throttling for large lists).
   - **2-5 seconds** — recommended for lists of 50-200 recipients.
   - **10-30 seconds** — recommended for lists of 200+ recipients.
3. The estimated total send time displays: `"48 emails × 3s delay = ~2 min 24 sec"`.
4. Click **Send All** — emails send one by one with the configured delay.

**Tips & Gotchas:**
- Exchange Online limits: ~30 emails/minute or ~10,000 emails/day. Exceeding these triggers temporary blocks.
- A 3-second delay keeps you safely under the 30/minute limit.
- During sending, a **progress bar** shows how many emails have been sent and the estimated time remaining.
- You can **cancel** mid-send by clicking the **"Stop Sending"** button — already-sent emails are not recalled.

---

## 7. Preview & Testing

### 7.1 Email Preview Carousel

**What it does:** Shows you exactly how each recipient's email will look after merge fields are replaced, before you send anything.

**Step-by-step:**

1. After composing your email with merge fields, click **"Preview"** (👁 icon) in the toolbar or press **Ctrl+P** in the task pane.
2. The preview panel opens, showing the first recipient's fully-merged email.
3. Use the **← Previous** and **Next →** buttons (or left/right arrow keys) to cycle through recipients.
4. The current recipient indicator shows: `"Viewing 3 of 48: carol@northwind.com"`.
5. Check each email for:
   - Correct merge field replacement
   - Proper formatting
   - Attachment indicators
   - Subject line personalization
6. Click **"Exit Preview"** to return to editing.

**Example preview:**

> **To:** alice@contoso.com
> **Subject:** Invoice #1042 for Contoso Ltd
> **CC:** alice.manager@contoso.com
>
> Dear Alice,
>
> Please find attached your invoice for **$1,200.00**.
>
> Best regards,
> The Finance Team
>
> 📎 Invoice-Alice.pdf

**Tips & Gotchas:**
- Preview shows real data from your spreadsheet — it's not a mock.
- If a merge field shows as `{UnknownField}`, it means the field name doesn't match any column. Fix the typo in your email.
- Preview also shows CC, BCC, and attachments for each recipient.
- Use preview to catch blank fields (where fallback defaults should be set).

---

### 7.2 Test Email

**What it does:** Sends a single test email to yourself (or any address you specify) using the first row's data, so you can verify the final result in your actual inbox.

**Step-by-step:**

1. Click the **"Send Test"** button (🧪 icon) in the toolbar.
2. A dialog appears:
   - **Send to:** (pre-filled with your email address — editable)
   - **Use data from row:** (dropdown, defaults to Row 1)
3. Click **"Send Test Email"**.
4. Check your inbox — the test email arrives with merge fields replaced using the selected row's data.
5. Verify formatting, attachments, subject line, and sender address.
6. Send additional tests with different rows to verify edge cases.

**Tips & Gotchas:**
- Test emails count toward your Exchange sending limit, but it's just one email.
- The test email includes `[TEST]` prefix in the subject line so you can easily identify it.
- Always send at least one test before a full merge.
- Check the test email on mobile too — formatting can differ between desktop and mobile clients.

---

### 7.3 Draft Mode

**What it does:** Instead of sending emails immediately, creates them as **drafts** in your Outlook Drafts folder. You can review each one individually before hitting Send.

**Step-by-step:**

1. In the **Sending Options** panel, toggle **"Save as Drafts"** ON.
2. Click **"Send All"** — the button changes to **"Create Drafts"**.
3. MailMerge-Pro creates each email as a draft. Progress bar shows: `"Creating draft 12 of 48..."`.
4. When complete, open your **Drafts** folder in Outlook.
5. Review each draft. You can manually edit individual emails before sending.
6. Send drafts individually or select multiple and send.

**Tips & Gotchas:**
- Draft mode is the **safest option** for first-time users — review everything before committing.
- Drafts include all merge data, CC/BCC, attachments, and formatting.
- Creating 50 drafts takes about 30 seconds.
- Drafts are regular Outlook drafts — you can forward, edit, or delete them freely.

---

## 8. Advanced Features

### 8.1 Many-to-One Merge with `{#rows}...{/rows}`

**What it does:** Groups multiple spreadsheet rows by email address and generates a **single email per unique recipient** containing a repeated section for each row. Perfect for sending a single invoice email listing multiple line items.

**Step-by-step:**

1. Prepare your spreadsheet with multiple rows per recipient:

   | Email | FirstName | Product | Quantity | Price |
   |---|---|---|---|---|
   | alice@contoso.com | Alice | Widget A | 5 | $50.00 |
   | alice@contoso.com | Alice | Widget B | 2 | $30.00 |
   | alice@contoso.com | Alice | Widget C | 10 | $15.00 |
   | bob@fabrikam.com | Bob | Gadget X | 1 | $200.00 |

2. Upload the spreadsheet. MailMerge-Pro detects multiple rows per email and shows: `"4 rows → 2 unique recipients"`.
3. In the email body, wrap the repeating section with `{#rows}` and `{/rows}`:

   ```
   Dear {FirstName},

   Here are your items:

   {#rows}
   • {Product} — Qty: {Quantity} — Price: {Price}
   {/rows}

   Best regards,
   Sales Team
   ```

4. Preview to verify. Alice's email shows:

   > Dear Alice,
   >
   > Here are your items:
   >
   > • Widget A — Qty: 5 — Price: $50.00
   > • Widget B — Qty: 2 — Price: $30.00
   > • Widget C — Qty: 10 — Price: $15.00
   >
   > Best regards,
   > Sales Team

5. Bob's email shows only his one item.

**Tips & Gotchas:**
- Only merge fields **inside** `{#rows}...{/rows}` repeat. Fields outside use the first row's data.
- The `{#rows}` tag supports HTML tables too:
  ```html
  <table>
    <tr><th>Product</th><th>Qty</th><th>Price</th></tr>
    {#rows}
    <tr><td>{Product}</td><td>{Quantity}</td><td>{Price}</td></tr>
    {/rows}
  </table>
  ```
- Grouping is by the **Email Address** column — all rows with the same email become one email.
- You can use this for: invoice line items, event attendee lists, order summaries, expense reports.

---

### 8.2 Campaign History

**What it does:** Keeps a log of all your past mail merge campaigns with details like date sent, recipient count, success/failure counts, and the email content used.

**Step-by-step:**

1. Click the **"History"** tab (📋 icon) in the task pane navigation.
2. A list of past campaigns appears, newest first:

   | Date | Subject | Recipients | Sent | Failed |
   |---|---|---|---|---|
   | 2024-01-15 2:30 PM | Invoice #{InvoiceNumber} | 48 | 47 | 1 |
   | 2024-01-10 9:00 AM | Welcome to {Company} | 120 | 120 | 0 |
   | 2024-01-05 4:15 PM | Monthly Report - {Month} | 35 | 35 | 0 |

3. Click any row to expand details:
   - **Email body** (with merge fields, as composed)
   - **Sending options** used (CC, BCC, aliases, etc.)
   - **Failed recipients** with error messages
   - **Timestamp** for each send
4. Click **"Re-use Template"** to pre-fill the composer with that campaign's email and settings.

**Tips & Gotchas:**
- History is stored locally in your browser's storage. Clearing browser data deletes history.
- Failed sends show error reasons: `Mailbox not found`, `Throttled`, `Attachment too large`, etc.
- Use "Re-use Template" to quickly run the same campaign with new data.

---

### 8.3 Export Results CSV

**What it does:** Exports a detailed log of the completed merge as a CSV file showing each recipient's send status (success/failure), timestamp, and error message if applicable.

**Step-by-step:**

1. After a merge completes, a **"Export Results"** button appears in the completion summary.
2. Click **"Export Results"** — a CSV file downloads named `MailMerge-Results-2024-01-15.csv`.
3. Open the CSV in Excel. Columns include:

   | Email | Name | Status | SentAt | Error |
   |---|---|---|---|---|
   | alice@contoso.com | Alice | Success | 2024-01-15T14:30:05Z | |
   | bob@fabrikam.com | Bob | Success | 2024-01-15T14:30:08Z | |
   | invalid@nowhere.xyz | Carol | Failed | 2024-01-15T14:30:11Z | Mailbox not found |

4. Use this file for auditing, troubleshooting, or importing into your CRM.

**Tips & Gotchas:**
- The export includes ALL rows, including successful ones.
- You can also access past results from Campaign History (8.2) → click a campaign → **"Export"**.
- CSV is UTF-8 encoded with BOM for proper Excel compatibility.

---

## 9. UI Features

### 9.1 Dark Mode

**What it does:** Switches the add-in interface to a dark color scheme that matches Outlook's dark mode.

**Step-by-step:**

1. Click the **Settings** gear icon (⚙) in the top-right of the task pane.
2. Find the **"Theme"** toggle.
3. Options:
   - **Auto** (default) — follows Outlook's theme setting.
   - **Light** — always light background.
   - **Dark** — always dark background.
4. The change applies immediately — no restart needed.

**Tips & Gotchas:**
- Auto mode detects Outlook's theme using the Office.js `Office.context.officeTheme` API.
- Dark mode also affects the email preview panel — but NOT the actual sent email.
- The email body editor stays light-background for accurate WYSIWYG preview.

---

### 9.2 Keyboard Shortcuts

**What it does:** Provides keyboard shortcuts for common actions to speed up your workflow.

| Shortcut | Action |
|---|---|
| **Ctrl+Enter** | Send All (or Create Drafts if Draft Mode is on) |
| **Ctrl+P** | Open Preview |
| **Ctrl+T** | Send Test Email |
| **Ctrl+Shift+F** | Insert Merge Field picker |
| **←** / **→** | Previous / Next recipient in Preview |
| **Esc** | Close Preview or Dialog |
| **Ctrl+B** | Bold (in editor) |
| **Ctrl+I** | Italic (in editor) |
| **Ctrl+U** | Underline (in editor) |
| **Ctrl+K** | Insert Link (in editor) |
| **Ctrl+Z** | Undo (in editor) |
| **Ctrl+Y** | Redo (in editor) |

**Tips & Gotchas:**
- **Ctrl+Enter** is the fastest way to launch a merge — no need to click buttons.
- Keyboard shortcuts only work when the MailMerge-Pro task pane is focused.
- On Mac, substitute **Cmd** for **Ctrl**.

---

### 9.3 Step Badges

**What it does:** Visual step indicators (numbered badges) in the task pane navigation that show your progress through the mail merge workflow and highlight which steps are complete.

**Step-by-step:**

1. The task pane shows 3 main steps with circular number badges:
   - **①** Data Source — Gray when empty, Blue when data is loaded, Green (✓) when mapped.
   - **②** Compose — Gray when empty, Blue when content is entered, Green (✓) when merge fields are valid.
   - **③** Send — Gray when prerequisites aren't met, Blue when ready, Green (✓) after successful send.
2. Click any step badge to jump directly to that section.
3. Incomplete steps show a tooltip explaining what's needed: `"Upload a data file to continue"`.

**Tips & Gotchas:**
- All three steps must be Green (✓) or Blue before the Send button enables.
- Step badges animate (bounce) when their status changes — drawing attention to what's next.
- On narrow screens, steps collapse to icons only.

---

### 9.4 Responsive Layout

**What it does:** Adapts the task pane UI to different widths — from the narrow Outlook sidebar to wider pop-out windows.

**Step-by-step:**

1. The task pane works in the default **narrow sidebar** (~350px wide).
2. To get more room, click the **"Pop Out"** button (↗ icon) at the top of the task pane — opens the add-in in a larger window.
3. The layout automatically adjusts:
   - **Narrow (< 400px):** Single-column layout, stacked sections, compact toolbar.
   - **Medium (400-700px):** Side-by-side data preview and field list.
   - **Wide (> 700px):** Full three-column layout with data, editor, and preview side-by-side.
4. The data preview table becomes horizontally scrollable in narrow mode.

**Tips & Gotchas:**
- Pop-out mode gives you the best experience for composing — use it for complex merges.
- The add-in remembers your last window size.
- On mobile Outlook (iOS/Android), the add-in opens full-screen in a bottom sheet.

---

## 10. New in v3.0

> **v3.0 brings 10 new features** — expanding MailMerge-Pro from 34 to **44 total features**. Templates, scheduling, A/B testing, contact groups, multi-language support, and more.

---

### 10.1 Email Templates Library

**What it does:** Provides 3 built-in email templates and lets you save, load, and delete your own custom templates. Templates are stored in your browser's localStorage — no server or cloud sync involved.

**Step-by-step:**

1. In the email composer, click the **"Templates"** button (📄 icon) in the toolbar.
2. The Templates panel opens with two tabs: **Built-in** and **My Templates**.
3. **Built-in templates** include:
   - **Professional Invoice** — Formal invoice notification with `{FirstName}`, `{Company}`, `{InvoiceAmount}` merge fields.
   - **Event Invitation** — Friendly event invite with `{FirstName}`, `{EventName}`, `{EventDate}`, `{Location}`.
   - **Follow-Up Reminder** — Polite follow-up with `{FirstName}`, `{Company}`, `{Topic}`, `{LastContactDate}`.
4. Click a template to preview it. Click **"Use This Template"** to load it into the composer.
5. **Save a custom template:** Compose your email, then click **"Save as Template"** in the Templates panel. Enter a name (e.g., "Q1 Marketing Blast") and click **Save**.
6. **Load a custom template:** Open the **My Templates** tab, click any saved template to preview, then click **"Use This Template"**.
7. **Delete a custom template:** Hover over a saved template and click the **🗑 trash icon**.

**Example custom template:**

| Template Name | Subject | Body (excerpt) |
|---|---|---|
| Q1 Marketing Blast | {Company} — Special Q1 Offer | Dear {FirstName}, We're excited to offer... |
| Onboarding Welcome | Welcome to {Company}, {FirstName}! | Hi {FirstName}, Congratulations on joining... |

**Tips & Gotchas:**
- Built-in templates cannot be deleted or edited — but you can load one, modify it, and save as a new custom template.
- Templates are stored in **localStorage** — they stay on this device and browser only. They are NOT synced across devices or browsers.
- Clearing your browser data deletes saved templates. Export important templates by copying the email content before clearing data.
- Templates include the subject line, body, and sending options (CC, BCC, importance, etc.).

---

### 10.2 Scheduled Sending

**What it does:** Lets you schedule your mail merge to send at a specific future date and time. Includes a countdown timer and a cancel button. Requires Outlook to remain open until the scheduled time.

**Step-by-step:**

1. Compose your email and upload your data as usual.
2. Instead of clicking **"Send All"**, click the **clock icon (🕐)** next to the Send button — or click the **"Schedule Send"** button.
3. A **date/time picker** appears:
   - **Date:** Select a future date from the calendar.
   - **Time:** Select the send time (in your local timezone). Use the hour/minute dropdowns or type directly.
4. Click **"Schedule"**.
5. A **countdown timer** appears in the task pane: `"Scheduled: Sending in 2h 34m 12s"`.
6. The task pane shows a **"Cancel Scheduled Send"** button (red). Click it anytime before the scheduled time to cancel.
7. When the countdown reaches zero, emails send automatically with the configured delay between each send.

**Example schedule:**
- Current time: Monday 9:00 AM
- Scheduled time: Tuesday 8:00 AM
- Countdown shows: `"23h 00m 00s"`
- At Tuesday 8:00 AM, all 48 emails begin sending.

**Tips & Gotchas:**
- **⚠️ Outlook must remain open** for the scheduled send to execute. If you close Outlook or the add-in task pane, the scheduled send is cancelled.
- The schedule uses your **local timezone** — verify the displayed time if you're scheduling for recipients in different timezones.
- You can schedule up to 7 days in advance.
- If your computer goes to sleep, the send may be delayed until it wakes up.
- Scheduled sends still respect your send delay (throttle) setting.
- You can continue using Outlook normally while a send is scheduled — just keep the task pane open.

---

### 10.3 Email Tracking

**What it does:** Requests read receipts via the Microsoft Graph API `isReadReceiptRequested` flag, so you can track whether recipients have opened your email.

**Step-by-step:**

1. In the **Sending Options** panel, toggle **"Email Tracking (Read Receipts)"** ON.
2. This sets the `isReadReceiptRequested` flag on each email via the Graph API.
3. When a recipient opens the email, their mail client may send a read receipt back to you.
4. Read receipts arrive as emails in your inbox: `"Read: [Original Subject]"`.
5. In **Campaign History** (Section 8.2), tracked campaigns show a **"Tracking"** badge. Click to see which recipients sent read receipts.

**Tips & Gotchas:**
- This uses the same Graph API flag as the existing Read Receipts toggle (Section 6.3). The v3.0 enhancement adds tracking visibility in Campaign History.
- Recipients can **decline** to send a read receipt — tracking is not guaranteed.
- Some organizations suppress read receipt prompts entirely.
- For compliance: read receipts are a standard email feature, not pixel tracking.

---

### 10.4 A/B Testing

**What it does:** Lets you create two versions of your email (Version A and Version B) and split your recipients between them to test which performs better. Includes tabbed editors and configurable split ratios.

**Step-by-step:**

1. In the email composer, click the **"A/B Test"** toggle (🔬 icon) in the toolbar.
2. The composer switches to a **tabbed view** with two tabs: **Version A** and **Version B**.
3. **Version A tab:** Write your first email version (e.g., formal tone, specific subject line).
4. **Version B tab:** Write your second email version (e.g., casual tone, different subject line).
5. Below the editors, set the **Split Ratio**:
   - **50/50** — Half your recipients get A, half get B (default).
   - **70/30** — 70% get A, 30% get B (use when you have a preferred version).
   - **80/20** — 80% get A, 20% get B (small test of version B).
6. Click **Preview** to see which recipients get which version. The preview carousel shows **"[A]"** or **"[B]"** next to each recipient.
7. Click **"Send All"** — recipients are randomly assigned to A or B based on the split ratio.
8. After sending, the **completion summary** shows results per version:

   | Version | Recipients | Sent | Failed |
   |---|---|---|---|
   | A (Formal) | 24 | 24 | 0 |
   | B (Casual) | 24 | 23 | 1 |

9. In **Campaign History**, A/B test campaigns show results for each version separately.

**Example A/B Test:**
- **Version A Subject:** `Your January Invoice from {Company}`
- **Version B Subject:** `{FirstName}, your invoice is ready! 📄`
- **Split:** 50/50
- After sending, compare open rates (via read receipts) to see which subject performs better.

**Tips & Gotchas:**
- Both versions can have different subjects, body content, and formatting — but they share the same recipient list, CC/BCC, and attachments.
- The split is randomized — you can't manually assign specific recipients to A or B.
- A/B testing works with all other features (attachments, scheduling, read receipts, etc.).
- For meaningful results, use at least 50 recipients per version.
- You can only test 2 versions (A and B) per campaign — not 3 or more.

---

### 10.5 Contact Groups & Segments

**What it does:** Lets you save, load, delete, and merge recipient lists as named contact groups. Groups are stored in your browser's localStorage for quick reuse across campaigns.

**Step-by-step:**

1. Upload your data (Excel/CSV) as usual.
2. Click the **"Contact Groups"** button (👥 icon) in the Data Source panel.
3. The Contact Groups panel opens with options:
   - **Save Current List:** Click **"Save as Group"**, enter a name (e.g., "Marketing Team" or "Q1 Prospects"), and click **Save**. All currently loaded recipients are saved.
   - **Load a Group:** Select a saved group from the list and click **"Load"**. The recipients replace the current data in the task pane.
   - **Merge Groups:** Select two or more groups and click **"Merge"**. Recipients are combined, and duplicates (by email address) are removed automatically.
   - **Delete a Group:** Hover over a group and click the **🗑 trash icon**.
4. The group list shows: group name, recipient count, and date saved:

   | Group Name | Recipients | Saved |
   |---|---|---|
   | Marketing Team | 35 | Jan 15, 2024 |
   | Q1 Prospects | 120 | Jan 10, 2024 |
   | VIP Clients | 15 | Jan 5, 2024 |

5. Click **"Load"** on any group to instantly populate the data table without re-uploading a file.

**Tips & Gotchas:**
- Contact groups are stored in **localStorage** — they persist across sessions on the same device and browser, but are NOT synced across devices.
- Merging groups performs automatic de-duplication based on the email address column.
- Maximum group size: ~5,000 recipients (limited by localStorage capacity, typically 5-10 MB).
- Groups save ALL columns from your spreadsheet, not just the email address.
- Clearing browser data deletes all saved groups. Export critical groups to Excel/CSV first.

---

### 10.6 HTML Template Import

**What it does:** Lets you import pre-designed HTML email templates from `.html` files. Supports file picker and drag-and-drop. Automatically detects `{MergeField}` placeholders in the HTML.

**Step-by-step:**

1. In the email composer, click the **"Import HTML"** button (📥 icon) in the toolbar.
2. Choose one of two methods:
   - **File picker:** Click **"Choose File"** and select a `.html` file from your computer.
   - **Drag-and-drop:** Drag a `.html` file directly onto the composer area. A blue drop zone appears: `"Drop HTML file here"`.
3. The HTML is loaded into the rich text editor with full formatting preserved.
4. MailMerge-Pro **auto-detects merge field placeholders** in the HTML:
   - Scans for `{FieldName}` patterns in the HTML.
   - Displays detected fields as chips above the editor: `Found: {FirstName}, {Company}, {Amount}`.
   - Unmatched fields (not in your spreadsheet) show a yellow warning.
5. Edit the imported HTML as needed using the WYSIWYG editor.

**Example:**
- You design a beautiful email in an HTML editor (e.g., MJML, Unlayer, Stripo).
- The HTML contains: `<p>Hello {FirstName}, welcome to {Company}!</p>`.
- Import the file → MailMerge-Pro detects `{FirstName}` and `{Company}` as merge fields.
- Upload your spreadsheet → the merge fields match your columns → send!

**Tips & Gotchas:**
- Only `.html` and `.htm` files are accepted. Other formats (`.txt`, `.docx`) are rejected with an error.
- Complex HTML with external CSS may not render identically in all email clients. Test with Preview and send a Test Email.
- Inline CSS is recommended for email HTML (most email clients strip `<style>` tags).
- Imported HTML replaces the current editor content. Save your work as a template first if needed.
- The drag-and-drop zone is only visible when you're dragging a file over the composer area.

---

### 10.7 Signature Auto-Insert

**What it does:** Automatically appends your email signature to every merge email. Fetches your signature from the Microsoft Graph API, or lets you paste one manually. Includes an auto-append toggle.

**Step-by-step:**

1. Click the **"Signature"** button (✒️ icon) in the email composer toolbar, or go to **Settings (⚙) → Signature**.
2. The Signature panel shows two options:
   - **Auto-fetch from Outlook:** Click **"Fetch My Signature"**. MailMerge-Pro calls the Graph API to retrieve your default Outlook signature. It appears in the preview area.
   - **Manual paste:** Click **"Paste Signature"** and paste your HTML or plain-text signature into the text area.
3. Toggle **"Auto-append signature"** ON (default: ON).
4. When enabled, your signature is automatically appended to the bottom of every merge email.
5. In Preview mode, you can see the signature at the bottom of each email.

**Example signature:**

> **Jane Smith**
> Marketing Manager | Contoso Ltd
> jane@contoso.com | +1 (555) 123-4567
> *Sent with MailMerge-Pro*

**Tips & Gotchas:**
- The Graph API fetches the signature configured in **Outlook on the Web** → Settings → Compose and reply → Email signature.
- If the Graph API doesn't return a signature (e.g., signature is configured only in desktop Outlook), use the manual paste option.
- The auto-append toggle lets you disable signature insertion for specific campaigns (e.g., when the template already includes a signature).
- Signature formatting (images, links, fonts) is preserved. However, signature images must be hosted online (not embedded base64) for reliable display.
- The signature is appended AFTER the email body and BEFORE any unsubscribe footer.

---

### 10.8 Rate Limit Dashboard

**What it does:** Displays a real-time dashboard showing your daily email send count against Exchange Online limits. Includes a color-coded progress bar and auto-suggested delay to stay under limits.

**Step-by-step:**

1. The Rate Limit Dashboard appears in the **Sending Options** panel, just above the Send button.
2. It shows:
   - **Daily Send Counter:** `"142 / 10,000 emails sent today"` (updates in real-time as you send).
   - **Color-coded bar:**
     - 🟢 **Green** (0-70%): Safe zone — plenty of capacity remaining.
     - 🟡 **Yellow** (70-90%): Caution — approaching the daily limit.
     - 🔴 **Red** (90-100%): Danger — near or at the limit. Sending may be throttled.
   - **Auto-suggested delay:** Based on your current usage and remaining quota, the dashboard suggests an optimal send delay:
     - Green zone: `"Suggested delay: 1-2 seconds"`
     - Yellow zone: `"Suggested delay: 5-10 seconds"`
     - Red zone: `"Suggested delay: 15-30 seconds — consider waiting until tomorrow"`
3. Click **"Apply Suggested Delay"** to automatically set the recommended delay in the Send Delay slider.

**Tips & Gotchas:**
- Exchange Online limits are approximately **10,000 emails/day** and **30 emails/minute**. These limits vary by tenant configuration.
- The counter tracks sends from the current MailMerge-Pro session. It does NOT include emails sent from other apps or manually from Outlook.
- The counter resets at midnight (UTC).
- If you hit the daily limit, Exchange returns a 429 error. MailMerge-Pro pauses and shows: `"Daily limit reached. Resume tomorrow or reduce recipient count."`.
- The auto-suggested delay is a recommendation — you can still set a custom delay.

---

### 10.9 Multi-Language (i18n)

**What it does:** Translates the entire MailMerge-Pro interface into 6 languages. A language selector in the header lets you switch instantly.

**Supported languages:**

| Code | Language | Flag |
|---|---|---|
| `en` | English | 🇺🇸 |
| `es` | Spanish (Español) | 🇪🇸 |
| `fr` | French (Français) | 🇫🇷 |
| `de` | German (Deutsch) | 🇩🇪 |
| `pt` | Portuguese (Português) | 🇧🇷 |
| `ja` | Japanese (日本語) | 🇯🇵 |

**Step-by-step:**

1. In the **header bar** of the task pane, click the **language selector** (🌐 globe icon) next to your profile name.
2. A dropdown appears with the 6 available languages.
3. Select a language — the entire UI translates immediately (no reload required).
4. Your language preference is saved in **localStorage** and persists across sessions.

**Example:**
- Select **Español** → All buttons, labels, tooltips, and messages switch to Spanish.
- "Send All" → "Enviar Todo"
- "Upload Data" → "Subir Datos"
- "Preview" → "Vista Previa"

**Tips & Gotchas:**
- The language setting affects only the **add-in UI** — NOT the email content you compose. Your emails are sent in whatever language you write them in.
- Language preference is stored in **localStorage** — it's per-device, per-browser. Switching devices requires re-selecting the language.
- Error messages and tooltips are also translated.
- If a translation is missing for a specific UI element, it falls back to English.

---

### 10.10 Local Admin Dashboard

**What it does:** Provides a personal analytics dashboard showing your mail merge activity: total campaigns, total emails sent, success rate, top recipients, monthly activity chart, and recent campaigns. All data is stored locally in your browser.

**Step-by-step:**

1. Click the **"Dashboard"** tab (📊 icon) in the task pane navigation.
2. The dashboard displays:
   - **Summary Cards (top row):**
     - 📬 **Total Campaigns:** `12`
     - 📧 **Total Emails Sent:** `1,847`
     - ✅ **Success Rate:** `98.7%`
   - **Top Recipients (list):** The 10 email addresses you've sent to most frequently:

     | Recipient | Times Emailed |
     |---|---|
     | alice@contoso.com | 8 |
     | bob@fabrikam.com | 7 |
     | carol@northwind.com | 6 |

   - **Monthly Activity Chart (bar chart):** A 6-month bar chart showing emails sent per month:
     - Oct: 120 | Nov: 340 | Dec: 280 | Jan: 450 | Feb: 380 | Mar: 277
   - **Recent Campaigns (table):** Last 10 campaigns with date, subject, recipient count, and success/fail status.
3. Click any campaign row in "Recent Campaigns" to jump to its full details in Campaign History (Section 8.2).

**Tips & Gotchas:**
- Dashboard data is calculated from your **Campaign History** stored in localStorage.
- Data is **local only** — your admin, IT department, or anyone else cannot see your dashboard. It is NOT shared or synced.
- Clearing browser data resets the dashboard to zero.
- The monthly activity chart auto-scales to show the last 6 months with data.
- The "Top Recipients" list can help identify over-contacted recipients (useful for avoiding spam fatigue).
- Dashboard data does NOT include emails sent manually from Outlook — only those sent via MailMerge-Pro.

---

## Appendix: Troubleshooting

| Problem | Solution |
|---|---|
| "Sign-in popup blocked" | Allow popups for `outlook.office.com` in browser settings |
| "Need admin approval" | Ask IT admin to grant tenant-wide consent |
| Merge field shows literally | Check case-sensitivity; ensure column exists in data |
| Email stuck in Outbox | Check Exchange throttling; reduce send delay |
| Attachment too large | Total attachments must be < 25 MB per email |
| "Mailbox not found" error | Verify recipient email address is correct |
| Add-in not appearing | Ensure add-in is installed; try closing and reopening compose window |
| Slow performance with large files | Keep spreadsheets under 5,000 rows for optimal performance |

---

*© 2026 MailMerge-Pro. All rights reserved.*
