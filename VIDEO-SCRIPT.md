# MailMerge-Pro: Video Demo Scripts & Storyboard

> **Purpose:** Complete scripts for YouTube tutorial videos | **Format:** Scene-by-scene narration + screen directions

---

## Table of Contents

- [Video 1: Main Demo (10 min)](#video-1-mailmerge-pro-free-mail-merge-for-outlook-365)
- [Video 2: Advanced Features (5 min)](#video-2-advanced-mail-merge-attachments-ccbcc-aliases)
- [Video 3: IT Admin Setup (5 min)](#video-3-mailmerge-pro-setup-guide-for-it-admins)
- [Recording Tips & YouTube SEO](#recording-tips--youtube-seo)

---

## Video 1: "MailMerge-Pro: Free Mail Merge for Outlook 365"

**Duration:** 10 minutes | **Audience:** End users, professionals | **Goal:** Install-to-send complete walkthrough

### Metadata

- **Title:** MailMerge-Pro: Free Mail Merge for Outlook 365 — Full Tutorial
- **Description:** Learn how to send personalized bulk emails directly from Outlook using MailMerge-Pro, a free mail merge add-in. Upload an Excel spreadsheet, compose a personalized email with merge fields, preview each recipient's email, and send — all without leaving Outlook. No external service, no data sharing, no Word required. Works on Outlook Web, Desktop, and Mobile.
- **Tags:** mail merge outlook, outlook mail merge, free mail merge, bulk email outlook, personalized email outlook 365, mail merge add-in, outlook add-in, mail merge excel outlook, mass email outlook
- **Thumbnail:** Split-screen showing Excel spreadsheet on left, personalized email on right, with text "FREE Mail Merge" and the Outlook logo.

---

### Scene 1: Hook / Introduction

**Timestamp:** 0:00 – 0:45

**Screen:** Animated title card with MailMerge-Pro logo, then cut to Outlook inbox.

**Narration:**

> Need to send personalized emails to a hundred people — but don't want to write each one by hand? Maybe you're sending invoices, event invitations, or customer updates, and you need each email to say the right name, the right amount, the right details.
>
> In this video, I'll show you how to do a complete mail merge directly inside Outlook 365 — for free — using an add-in called MailMerge-Pro.
>
> No Microsoft Word. No external email service. Your data never leaves your mailbox.
>
> Let's dive in.

---

### Scene 2: Installing the Add-in

**Timestamp:** 0:45 – 2:00

**Screen:** Outlook Web (outlook.office.com), inbox view.

**Narration:**

> First, let's install MailMerge-Pro. I'm using Outlook on the web, but this works on the desktop app too.

**Screen Action:** Click "New mail" to open compose window.

> Open a new email. In the toolbar, click the three dots — "More actions" — and then click "Get Add-ins."

**Screen Action:** Click "…" → "Get Add-ins." The add-ins dialog opens.

> In the search box, type "MailMerge-Pro" and hit Enter.

**Screen Action:** Type "MailMerge-Pro" in search. Results appear.

> Here it is. Click "Add" — and confirm.

**Screen Action:** Click "Add" on the MailMerge-Pro listing. Confirmation dialog appears. Click "Continue."

> Done. You'll see the MailMerge-Pro icon appear in your toolbar. Click it to open the task pane.

**Screen Action:** Close the dialog. Click the MailMerge-Pro icon in the ribbon. The task pane opens on the right side.

> The task pane opens on the right. You'll see three steps: Upload Data, Compose, and Send. Let's start with step one — uploading our data.

---

### Scene 3: Preparing and Uploading the Spreadsheet

**Timestamp:** 2:00 – 3:30

**Screen:** Switch to Excel showing the sample spreadsheet.

**Narration:**

> Here's my spreadsheet in Excel. I have columns for Email, FirstName, LastName, Company, and InvoiceAmount. Each row is one recipient.

**Screen Action:** Show the spreadsheet with 5 sample rows:

| Email | FirstName | LastName | Company | InvoiceAmount |
|---|---|---|---|---|
| alice@contoso.com | Alice | Johnson | Contoso Ltd | $1,200.00 |
| bob@fabrikam.com | Bob | Smith | Fabrikam Inc | $3,450.00 |
| carol@northwind.com | Carol | Lee | Northwind Traders | $780.00 |
| david@adventure.com | David | Kim | Adventure Works | $2,100.00 |
| eve@litware.com | Eve | Chen | Litware Inc | $950.00 |

> The first row must be your column headers — these become your merge fields. Let me save this and switch back to Outlook.

**Screen Action:** Save the file. Switch back to Outlook with the add-in task pane open.

> In the task pane, I'll click "Upload Data" — or I can drag and drop my file right here.

**Screen Action:** Click "Upload Data." File picker opens. Select the Excel file. Upload completes.

> The file uploads, and I can see a preview of my data. It detected 5 recipients and found my columns: Email, FirstName, LastName, Company, InvoiceAmount.

**Screen Action:** Show the data preview table in the task pane. Column chips display above the table.

> Notice the column mapping section — MailMerge-Pro auto-detected that my "Email" column contains the email addresses. The green checkmark means it's mapped correctly. I don't need to change anything.

**Screen Action:** Hover over the column mapping section showing Email → auto-detected.

---

### Scene 4: Composing the Email with Merge Fields

**Timestamp:** 3:30 – 5:30

**Screen:** Outlook compose window with MailMerge-Pro task pane.

**Narration:**

> Now for step two — composing the email. I'll type my subject first.

**Screen Action:** Click the Subject field.

> For the subject, I'll type: "Invoice for" — and now I'll add a merge field. I type an opening brace — the curly bracket — and look! An autocomplete dropdown shows all my columns.

**Screen Action:** Type `Invoice for {` — autocomplete dropdown appears showing: Email, FirstName, LastName, Company, InvoiceAmount.

> I'll select "Company."

**Screen Action:** Select "Company" from the dropdown. Subject now reads: `Invoice for {Company}`

> The merge field appears as a highlighted badge. When the email sends, this will be replaced with each company's name. So Alice's email will say "Invoice for Contoso Ltd" and Bob's will say "Invoice for Fabrikam Inc."

> Now the email body. Let me type a professional message.

**Screen Action:** Click in the email body editor. Type the following, using the autocomplete for merge fields:

```
Dear {FirstName},

Please find below the details of your latest invoice.

Company: {Company}
Amount Due: {InvoiceAmount}

Payment is due within 30 days. If you have any questions, please don't hesitate to reach out.

Best regards,
The Finance Team
```

> I'm typing naturally and inserting merge fields with the curly brace shortcut. Let me also bold the "Amount Due" value for emphasis.

**Screen Action:** Select `{InvoiceAmount}` and click Bold.

> I can use all the standard formatting — bold, italic, bullet points, links, images. This is a full rich-text editor.

**Screen Action:** Briefly show the formatting toolbar, hovering over a few buttons.

> Let me also set a fallback default. If any FirstName is missing, I don't want the email to say "Dear comma." Click "Fallback Defaults" here.

**Screen Action:** Click the "Fallback Defaults" button (⚙ icon).

> I'll set FirstName's default to "Valued Customer." Now if a row has a blank name, it'll read "Dear Valued Customer" instead of an awkward blank.

**Screen Action:** Enter "Valued Customer" next to FirstName. Click Save.

---

### Scene 5: Previewing Emails

**Timestamp:** 5:30 – 7:00

**Screen:** MailMerge-Pro task pane, preview mode.

**Narration:**

> Before sending, let's preview. This is my favorite feature. I'll click the Preview button — or use the shortcut Control+P.

**Screen Action:** Click "Preview" (👁 icon). The preview panel opens showing Alice's email.

> Here's Alice's email. Look — the subject says "Invoice for Contoso Ltd," the body says "Dear Alice," and the amount is "$1,200.00." Everything is personalized.

**Screen Action:** Highlight the personalized parts.

> I can use these arrows to cycle through recipients. Next — here's Bob's email. "Invoice for Fabrikam Inc, Dear Bob, $3,450.00." Perfect.

**Screen Action:** Click "Next →" to show Bob's email. Click again for Carol's.

> Let me check Carol's — "Invoice for Northwind Traders, Dear Carol, $780.00." All good.

> I can also scroll through all five to make sure everything looks right. The preview shows the subject, body, and even the CC, BCC, and attachments if I had any.

**Screen Action:** Cycle through remaining recipients quickly.

> Everything looks perfect. Let me close the preview and send a test email first.

**Screen Action:** Click "Exit Preview."

---

### Scene 6: Sending a Test Email

**Timestamp:** 7:00 – 8:00

**Screen:** MailMerge-Pro task pane, test email dialog.

**Narration:**

> Before sending to all five people, I want to verify the email in my actual inbox. Click "Send Test" — the test tube icon.

**Screen Action:** Click "Send Test" (🧪 icon). Dialog opens.

> It pre-fills my own email address. I can choose which row's data to use — I'll keep it on Row 1 (Alice's data). Click "Send Test Email."

**Screen Action:** Keep defaults. Click "Send Test Email." A success toast appears: "Test email sent!"

> Let me check my inbox.

**Screen Action:** Navigate to Inbox. Show the test email arriving with subject "[TEST] Invoice for Contoso Ltd."

> Here it is — "[TEST] Invoice for Contoso Ltd." Let me open it. "Dear Alice, Company: Contoso Ltd, Amount Due: $1,200.00." The formatting is perfect, bold and everything. This is exactly what Alice would see.

**Screen Action:** Open the test email. Show the formatted content.

> Satisfied. Let me go back to the compose window and send the full merge.

---

### Scene 7: Sending the Merge

**Timestamp:** 8:00 – 9:15

**Screen:** Outlook compose window with MailMerge-Pro task pane.

**Narration:**

> Back in the add-in, I'll click "Send All." I can also use the keyboard shortcut Control+Enter.

**Screen Action:** Click "Send All" button. A confirmation dialog appears.

> A confirmation dialog asks: "Send to 5 recipients?" This is your last chance to double-check. I'll click "Confirm & Send."

**Screen Action:** Click "Confirm & Send." The progress bar appears.

> And it's sending! The progress bar shows each email being sent. "Sending 1 of 5... 2 of 5..." There's a slight delay between each send to avoid throttling.

**Screen Action:** Show the progress bar animating: 1/5, 2/5, 3/5, 4/5, 5/5.

> Done! All 5 emails sent successfully. The summary shows: 5 sent, 0 failed. I can click "Export Results" to download a CSV log of everything that was sent.

**Screen Action:** Show the completion summary. Click "Export Results" — a CSV downloads.

> And that's it. Five personalized emails sent from my Outlook account, each one unique, in under a minute.

---

### Scene 8: Wrap-Up

**Timestamp:** 9:15 – 10:00

**Screen:** Animated summary slide with key features listed.

**Narration:**

> Let's recap what we did:
>
> One — Installed MailMerge-Pro from the Outlook add-in store.
> Two — Uploaded an Excel spreadsheet with our recipient data.
> Three — Composed a personalized email using merge fields like FirstName and Company.
> Four — Previewed each recipient's email to make sure everything looked right.
> Five — Sent a test to ourselves, then sent the full merge.
>
> MailMerge-Pro is free for up to 50 emails per day. If you need more — per-recipient attachments, CC/BCC control, send-from-alias, and other advanced features — check out the Pro plan. Link in the description.
>
> In the next video, I'll show you those advanced features. If this was helpful, give it a thumbs up, subscribe, and drop a comment if you have questions. Thanks for watching!

**Screen Action:** End card with subscribe button, next video link, and MailMerge-Pro logo.

---

## Video 2: "Advanced Mail Merge: Attachments, CC/BCC, Aliases"

**Duration:** 5 minutes | **Audience:** Power users | **Goal:** Demonstrate Pro features

### Metadata

- **Title:** Advanced Outlook Mail Merge: Attachments, CC/BCC & Aliases — MailMerge-Pro
- **Description:** Go beyond basic mail merge in Outlook 365. Learn how to attach different files to each recipient, add per-recipient CC/BCC, send from an alias or shared mailbox, and use many-to-one merge for grouped data. Uses MailMerge-Pro, a free add-in for Outlook.
- **Tags:** mail merge attachments outlook, per-recipient attachments, bulk email cc bcc, outlook shared mailbox mail merge, send from alias outlook, advanced mail merge
- **Thumbnail:** Email icon with paperclip and multiple documents fanning out, text "Advanced Mail Merge."

---

### Scene 1: Introduction

**Timestamp:** 0:00 – 0:25

**Screen:** Animated title card, then Outlook with MailMerge-Pro open.

**Narration:**

> In the last video, we did a basic mail merge. Now let's get advanced. I'll show you how to attach different files to each recipient, add CC and BCC per row, send from a shared mailbox, and use an alias. These are Pro features in MailMerge-Pro. Let's go.

---

### Scene 2: Per-Recipient Attachments

**Timestamp:** 0:25 – 2:00

**Screen:** Excel spreadsheet with an AttachmentFile column.

**Narration:**

> First, per-recipient attachments. I've updated my spreadsheet to include an "AttachmentFile" column. Each row has the filename of the PDF I want to attach for that person.

**Screen Action:** Show updated spreadsheet:

| Email | FirstName | Company | AttachmentFile |
|---|---|---|---|
| alice@contoso.com | Alice | Contoso Ltd | Invoice-1042-Contoso.pdf |
| bob@fabrikam.com | Bob | Fabrikam Inc | Invoice-1043-Fabrikam.pdf |
| carol@northwind.com | Carol | Northwind Traders | Invoice-1044-Northwind.pdf |

> Back in MailMerge-Pro, I upload this spreadsheet. It auto-detects the Email column and — look — it also detects the AttachmentFile column as the attachment mapping.

**Screen Action:** Upload the file. Show auto-detected mappings with green checks.

> Now I click "Upload Attachment Files" — this folder icon — and select all my PDFs at once.

**Screen Action:** Click the folder icon. Multi-select 3 PDF files. Upload completes.

> MailMerge-Pro matches each row's filename to the uploaded files. Green checks mean matched. Alice gets Invoice-1042-Contoso.pdf, Bob gets Invoice-1043-Fabrikam, Carol gets Invoice-1044-Northwind.

**Screen Action:** Show the match list with green checks next to each file.

> I can also add a global attachment — something everyone gets. I'll add our Terms of Service PDF.

**Screen Action:** Click "Add Global Attachment." Select "Terms-of-Service.pdf." It appears as a chip.

> Now everyone gets their personal invoice PLUS the Terms of Service. Let me preview Alice's email to confirm.

**Screen Action:** Click Preview. Show Alice's email with: 📎 Invoice-1042-Contoso.pdf, 📎 Terms-of-Service.pdf.

> Two attachments — her personal invoice and the shared Terms file. 

---

### Scene 3: Per-Recipient CC/BCC

**Timestamp:** 2:00 – 3:00

**Screen:** Excel spreadsheet with CC and BCC columns.

**Narration:**

> Next — CC and BCC. I've added two more columns to my spreadsheet: "CC" and "BCC."

**Screen Action:** Show updated spreadsheet:

| Email | FirstName | CC | BCC |
|---|---|---|---|
| alice@contoso.com | Alice | alice.manager@contoso.com | crm@mycompany.com |
| bob@fabrikam.com | Bob | bob.manager@fabrikam.com; finance@fabrikam.com | crm@mycompany.com |
| carol@northwind.com | Carol | *(empty)* | crm@mycompany.com |

> Bob's row has two CC addresses separated by a semicolon — his manager and finance team. Carol has no CC — that's fine, it just sends without one.

> The BCC column copies our CRM system on every email. But since it's the same for everyone, I could also use Global BCC instead. Let me show both.

**Screen Action:** Upload the spreadsheet. In Column Mapping, map CC → CC column, BCC → BCC column.

> In the Column Mapping, I set CC to my CC column and BCC to my BCC column. But I also want to add a Global BCC — for our legal compliance team.

**Screen Action:** Scroll to Sending Options. Enter `legal@mycompany.com` in Global BCC.

> I've added legal@mycompany.com as a Global BCC. Now every email BCCs both our CRM system from the spreadsheet column AND our legal team from the Global BCC. They combine — they don't override.

**Screen Action:** Preview Bob's email. Show: To: bob@fabrikam.com, CC: bob.manager@fabrikam.com; finance@fabrikam.com, BCC: crm@mycompany.com; legal@mycompany.com.

---

### Scene 4: Send from Alias & Shared Mailbox

**Timestamp:** 3:00 – 4:15

**Screen:** MailMerge-Pro Sending Options panel.

**Narration:**

> Now let's send from a different address. In Sending Options, I see the "Send From" dropdown. It shows my aliases — I have my main address and a "sales" alias.

**Screen Action:** Click the "Send From" dropdown. Show: you@company.com, sales@company.com.

> I'll select "sales@company.com." Now all emails will come from the sales address. Recipients see "From: sales@company.com" and replies go there too.

**Screen Action:** Select sales@company.com.

> For a shared mailbox, I toggle "Send from Shared Mailbox" on and enter the shared mailbox address — "invoicing@company.com."

**Screen Action:** Toggle "Send from Shared Mailbox" ON. Enter `invoicing@company.com`. Green check appears.

> Green check means I have permission. If you get a red X, ask your Exchange admin to grant you "Send As" permission on that shared mailbox.

> Now the emails come from invoicing@company.com — not my personal address. Great for team mailboxes like support@, billing@, or info@.

---

### Scene 5: Quick Demo of Other Pro Features

**Timestamp:** 4:15 – 4:45

**Screen:** MailMerge-Pro Sending Options panel.

**Narration:**

> A few more Pro features, quickly:
>
> Read receipts — toggle this on and recipients get prompted to send a read confirmation.

**Screen Action:** Toggle "Request Read Receipts" ON.

> High importance — marks all emails with a red exclamation mark. Use sparingly.

**Screen Action:** Toggle "High Importance" ON briefly, then OFF.

> Unsubscribe header — adds a one-click unsubscribe link in Gmail and Outlook. Paste your unsubscribe URL here. Great for compliance.

**Screen Action:** Toggle "Include Unsubscribe Header" ON. Enter a sample URL.

> Send delay — I can control the seconds between each send. Three seconds is a safe default to avoid Exchange throttling.

**Screen Action:** Set the delay slider to 3 seconds.

---

### Scene 6: Wrap-Up

**Timestamp:** 4:45 – 5:00

**Screen:** Summary card.

**Narration:**

> That's the Pro toolkit — per-recipient attachments, CC/BCC, aliases, shared mailboxes, read receipts, and more. All inside Outlook, no external tools needed.
>
> Get MailMerge-Pro free from the Outlook add-in store. Upgrade to Pro for these advanced features. Links below. Thanks for watching!

**Screen Action:** End card with links and subscribe button.

---

## Video 3: "MailMerge-Pro Setup Guide for IT Admins"

**Duration:** 5 minutes | **Audience:** IT administrators | **Goal:** Deployment walkthrough

### Metadata

- **Title:** MailMerge-Pro IT Admin Setup: Azure AD, Manifest & Intune Deployment
- **Description:** IT admin guide for deploying the MailMerge-Pro Outlook add-in across your organization. Covers Azure AD app registration, manifest deployment via the M365 Admin Center, Intune configuration, group targeting, and troubleshooting. Step-by-step walkthrough for enterprise deployment.
- **Tags:** outlook add-in deployment, intune office add-in, deploy outlook add-in, m365 admin center add-in, azure ad app registration, office 365 add-in management, enterprise outlook add-in
- **Thumbnail:** Shield icon with gear/cog, Microsoft Intune logo, text "IT Admin Guide."

---

### Scene 1: Introduction

**Timestamp:** 0:00 – 0:25

**Screen:** Title card, then M365 admin center.

**Narration:**

> If you're an IT admin and your team needs MailMerge-Pro, this video shows you how to deploy it organization-wide. I'll cover the Azure AD app registration, deploying the manifest through the M365 admin center, and optional Intune integration. Let's get started.

---

### Scene 2: Azure AD App Registration

**Timestamp:** 0:25 – 1:30

**Screen:** Azure Portal (portal.azure.com) → Azure Active Directory.

**Narration:**

> First, if you're hosting MailMerge-Pro yourself — not from AppSource — you need an Azure AD app registration. Go to the Azure portal and navigate to Azure Active Directory — or "Microsoft Entra ID" as it's now called.

**Screen Action:** Navigate to portal.azure.com → Microsoft Entra ID → App registrations.

> Click "New registration." Name it "MailMerge-Pro." For supported account types, select "Accounts in this organizational directory only" — single tenant.

**Screen Action:** Click "New registration." Enter name. Select single tenant. Click Register.

> For the Redirect URI, select "Single-page application" and enter the add-in's URL — that's where it's hosted, like your GitHub Pages URL.

**Screen Action:** Set redirect URI type to SPA, enter `https://your-org.github.io/MailMerge-Pro/taskpane.html`.

> Click Register. Copy the Application (client) ID and Tenant ID — you'll need these in the manifest.

**Screen Action:** Show the overview page with Application ID and Tenant ID highlighted.

> Now go to "API permissions." Add the Microsoft Graph permissions the add-in needs: Mail.Send, Mail.ReadWrite, and User.Read. Then click "Grant admin consent" so users don't get individual consent prompts.

**Screen Action:** Click "API permissions" → "Add a permission" → Microsoft Graph → Delegated → search and add Mail.Send, Mail.ReadWrite, User.Read. Click "Grant admin consent for [tenant]." Confirm.

> Green checkmarks mean consent is granted for everyone.

**Screen Action:** Show green checkmarks next to all permissions.

---

### Scene 3: Deploying via M365 Admin Center

**Timestamp:** 1:30 – 3:15

**Screen:** M365 Admin Center (admin.microsoft.com).

**Narration:**

> Now let's deploy the add-in. Go to admin.microsoft.com and navigate to Settings, then Integrated Apps.

**Screen Action:** Navigate to admin.microsoft.com → Settings → Integrated Apps.

> Click "Upload custom apps." Select "Office Add-in" as the app type. Choose "Provide link to manifest file" and paste your manifest URL.

**Screen Action:** Click "Upload custom apps." Select "Office Add-in." Select "Provide link to manifest file." Paste: `https://your-org.github.io/MailMerge-Pro/manifest.xml`. Click Validate.

> Validation passes — green checkmark. It shows the add-in name, version, and description. Click Next.

**Screen Action:** Show validation success. Click Next.

> For user assignment, I recommend starting with a pilot group. Select "Specific users/groups" and search for your pilot group — I'll pick "IT-Team" with 10 members.

**Screen Action:** Select "Specific users/groups." Search "IT-Team." Select it. Shows "10 members."

> Set the deployment to "Fixed" — meaning users can't remove it, it's always there. For a softer rollout, choose "Available" so users can self-install from the admin-managed section.

**Screen Action:** Select "Fixed." Click Next.

> Accept the permissions — these match what we configured in Azure AD. Click Next. Review the summary and click "Finish deployment."

**Screen Action:** Click "Accept permissions." Click Next. Review summary. Click "Finish deployment." Success message appears.

> Done. It can take up to 24 hours for the add-in to propagate to all users, but it's usually faster — typically 1 to 4 hours for Outlook on the web.

---

### Scene 4: Group Targeting & Rollout Strategy

**Timestamp:** 3:15 – 3:50

**Screen:** M365 Admin Center, MailMerge-Pro detail page.

**Narration:**

> For a phased rollout, I recommend three stages. Week one — deploy to your IT team for testing. Week two — expand to an early adopter group of 50 to 100 business users. Week three onward — deploy to all users.

**Screen Action:** Show the Integrated Apps detail page. Click "Edit users." Change from "IT-Team" to "All Company Users."

> To change the target group, click on MailMerge-Pro in Integrated Apps, go to user assignment, and update the group. Changes propagate within hours.

---

### Scene 5: Updating the Add-in

**Timestamp:** 3:50 – 4:20

**Screen:** GitHub repository page.

**Narration:**

> Here's the beautiful thing about web add-ins — updates are automatic. The add-in code lives on your web server, like GitHub Pages. When you push a code update, users get the new version the next time they open the add-in. No redeployment needed in the admin center.

**Screen Action:** Show a GitHub push updating the code. Show the GitHub Pages deployment.

> You only need to touch the admin center if the manifest itself changes — for example, if you add new permissions or change the add-in's name. In that case, go to Integrated Apps, click the add-in, and click "Update" to revalidate the manifest.

**Screen Action:** Show the "Update" button in the admin center.

---

### Scene 6: Troubleshooting

**Timestamp:** 4:20 – 4:50

**Screen:** Outlook showing common issues.

**Narration:**

> Quick troubleshooting tips. If the add-in doesn't appear after 24 hours, verify the user is in the assigned group and check the deployment status in the admin center.
>
> If users see "This add-in has been disabled," re-enable it in the admin center.
>
> If the add-in loads but shows an error, check that the manifest URL is publicly accessible — try opening it in your browser.
>
> For desktop Outlook issues, clear the Office cache by deleting the Wef folder in the user's local app data. The path is in the description below.

**Screen Action:** Show bullet points on screen:
- Check group membership
- Verify deployment status in admin center
- Test manifest URL in browser
- Clear cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`

---

### Scene 7: Wrap-Up

**Timestamp:** 4:50 – 5:00

**Screen:** Summary card.

**Narration:**

> That's the IT admin setup for MailMerge-Pro. Azure AD registration, admin center deployment, group targeting, and automatic updates. Check the detailed Intune deployment guide linked in the description for Intune-specific configuration profiles and PowerShell scripts. Thanks for watching!

**Screen Action:** End card with links to INTUNE-DEPLOYMENT.md guide and subscribe button.

---

## Recording Tips & YouTube SEO

### Recommended Recording Tools

| Tool | Purpose | Cost | Notes |
|---|---|---|---|
| **OBS Studio** | Screen recording + webcam | Free | Open source; most popular; supports scenes, transitions |
| **ShareX** | Quick screen capture + GIF | Free | Great for short clips; not ideal for long recordings |
| **Camtasia** | Recording + built-in editor | $300 | Easiest editing; good for non-video-editors |
| **DaVinci Resolve** | Video editing (post-recording) | Free | Professional-grade editor; steep learning curve |
| **Audacity** | Audio editing / noise removal | Free | Clean up narration audio; remove background noise |
| **Descript** | AI-powered editing + captions | $24/month | Edit video by editing text transcript; auto-captions |

### Resolution & Quality Settings

| Setting | Recommended Value |
|---|---|
| **Resolution** | 1920 × 1080 (1080p) |
| **Frame rate** | 30 fps (screen recording doesn't need 60) |
| **Bitrate** | 8,000-12,000 Kbps for 1080p30 |
| **Audio** | 48 kHz, 320 kbps AAC or similar |
| **Format** | MP4 (H.264 video, AAC audio) |
| **Aspect ratio** | 16:9 (standard YouTube) |
| **Scaling** | Set Windows/Mac display to 100% (no DPI scaling) for crisp text |

### Audio Recording Tips

1. **Use a dedicated microphone** — even a $30 USB mic (e.g., Fifine K669) sounds dramatically better than a laptop mic.
2. **Record in a quiet room** — close windows, turn off fans, mute phone notifications.
3. **Maintain consistent distance** — stay 6-8 inches from the mic.
4. **Use a pop filter** — prevents plosive "p" and "b" sounds.
5. **Normalize audio in post** — use Audacity's "Normalize" effect to even out volume.
6. **Remove background noise** — Audacity → Effect → Noise Reduction → Sample noise → Apply.
7. **Speak slowly and clearly** — viewers can speed up; they can't slow down unclear speech.

### Screen Recording Tips

1. **Clean your desktop** — hide personal bookmarks, close unrelated tabs, use a clean browser profile.
2. **Use a clean Outlook account** — create a demo M365 account with sample data, not your real inbox.
3. **Increase font size** — in Outlook Web, use browser zoom (125-150%) so text is readable on mobile.
4. **Highlight the cursor** — use a cursor highlighter tool (built into OBS, or use "Mouse Pointer Highlight" utility) so viewers can follow clicks.
5. **Zoom in on key areas** — use OBS's "Window Capture + Crop" or post-production zoom to highlight the task pane.
6. **Pause before clicking** — give viewers a beat to see what you're about to click. Narrate first, then click.
7. **Use keyboard shortcuts** — shows efficiency and teaches viewers the shortcuts.

### Adding Captions (Subtitles)

| Method | Tool | Quality | Effort |
|---|---|---|---|
| YouTube auto-captions | YouTube Studio | 85-90% accurate | Minimal (auto-generated, but review and fix errors) |
| AI transcription | Descript, Otter.ai, Whisper | 95%+ accurate | Low (paste transcript, fix minor errors) |
| Manual SRT file | Subtitle Edit (free) | 100% accurate | High (type everything manually) |
| Professional service | Rev.com, GoTranscript | 99%+ accurate | $1-2/minute of video |

**Recommendation:** Use YouTube's auto-captions as a starting point, then review and fix errors in YouTube Studio. This takes 15-20 minutes per 10-minute video and dramatically improves accessibility and SEO.

### YouTube SEO & Optimization

#### Titles (follow this formula)

```
[Topic]: [Specific Benefit] — [Product Name]
```

Examples:
- ✅ "MailMerge-Pro: Free Mail Merge for Outlook 365 — Full Tutorial"
- ✅ "Advanced Outlook Mail Merge: Attachments, CC/BCC & Aliases"
- ❌ "MailMerge-Pro Tutorial" (too vague — no keywords)
- ❌ "How I Sent 500 Emails" (clickbait — no search value)

#### Descriptions (first 200 characters matter most)

```
Learn how to send personalized bulk emails from Outlook using MailMerge-Pro, 
a free mail merge add-in. Upload Excel, compose with merge fields, preview, 
and send — all inside Outlook 365.

⏱ TIMESTAMPS:
0:00 Introduction
0:45 Installing MailMerge-Pro
2:00 Uploading your spreadsheet
3:30 Composing with merge fields
5:30 Previewing emails
7:00 Sending a test email
8:00 Sending the full merge
9:15 Recap and next steps

🔗 LINKS:
Install MailMerge-Pro: [AppSource URL]
Advanced features video: [Video 2 URL]
IT Admin setup: [Video 3 URL]
Full documentation: [GitHub URL]

📌 RELATED VIDEOS:
• Advanced Mail Merge (Attachments, CC/BCC): [URL]
• IT Admin Deployment Guide: [URL]

#mailmerge #outlook #outlook365 #mailmergeoutlook #bulkemail
```

#### Tags

```
mail merge outlook, outlook mail merge, free mail merge, mail merge outlook 365,
bulk email outlook, personalized email outlook, outlook add-in, mail merge excel,
mass email outlook, mail merge tutorial, outlook mail merge add-in, 
MailMerge-Pro, send personalized emails outlook, mail merge attachments outlook
```

#### Thumbnails

| Element | Specification |
|---|---|
| **Resolution** | 1280 × 720 pixels (minimum) |
| **Format** | JPG or PNG, under 2 MB |
| **Text** | 3-5 words max (e.g., "FREE Mail Merge") |
| **Font** | Bold, sans-serif, white with dark outline for legibility |
| **Background** | Contrasting color (blue, orange) or screenshot with overlay |
| **Faces** | Human faces increase click-through rate (if using webcam) |
| **Branding** | Small MailMerge-Pro logo in corner |
| **Tools** | Canva (free), Photoshop, Figma |

#### Thumbnail Examples

**Video 1:** Split-screen: Excel spreadsheet (left) → Personalized email (right), with arrow between them. Bold text: "FREE Mail Merge." Outlook logo in corner.

**Video 2:** Email icon with paperclip and 3 different PDF documents fanning out. Bold text: "Advanced Mail Merge." Pro badge.

**Video 3:** Server rack / shield icon with Microsoft Intune and Azure AD logos. Bold text: "IT Admin Guide." Enterprise/professional look.

#### Publishing Schedule

| Day | Action |
|---|---|
| Monday | Upload Video 1 (main demo) — publish immediately |
| +7 days | Upload Video 2 (advanced) — reference Video 1 in cards |
| +14 days | Upload Video 3 (IT admin) — reference Videos 1 & 2 |
| +21 days | Upload Video 4 (specific use case: invoices) — cross-link all |

Publish on **Tuesday or Wednesday between 9-11 AM** in your target timezone (US Eastern or GMT) — this is when B2B / professional audiences are most active on YouTube.

---

*© 2024 MailMerge-Pro. Video production guide.*
