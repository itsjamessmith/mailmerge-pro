# MailMerge-Pro: Intune Deployment Guide

> **Audience:** IT Administrators | **Platform:** Microsoft 365 / Intune / Azure AD | **Features:** 44
> **Last Updated:** 2025-01 (v3.0)

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Method 1: Integrated Apps in M365 Admin Center (Recommended)](#method-1-integrated-apps-in-m365-admin-center-recommended)
- [Method 2: Intune Configuration Profiles](#method-2-intune-configuration-profiles)
- [Method 3: PowerShell Scripts via Intune](#method-3-powershell-scripts-via-intune)
- [Group Targeting](#group-targeting)
- [Updating the Add-in](#updating-the-add-in)
- [Removing / Uninstalling](#removing--uninstalling)
- [Monitoring Deployment Status](#monitoring-deployment-status)
- [Troubleshooting](#troubleshooting)

---

## Prerequisites

Before deploying, ensure you have:

- [ ] **Microsoft 365 Admin** or **Exchange Admin** role
- [ ] **Intune Admin** role (for Method 2 and 3)
- [ ] The add-in **manifest URL** hosted on a publicly accessible HTTPS endpoint:
  ```
  https://your-org.github.io/MailMerge-Pro/manifest.xml
  ```
  Or the **manifest XML file** downloaded locally.
- [ ] Azure AD **app registration** completed (if the add-in requires Graph API permissions)
- [ ] Azure AD app registration includes the **`brk-multihub://your-domain.com`** SPA redirect URI (required for Nested App Authentication)
- [ ] Target user groups created in Azure AD (for scoped deployment)

> **Authentication Note:** MailMerge-Pro uses **Nested App Authentication (NAA)** with **MSAL v3.27.0**. NAA is the Microsoft-recommended authentication approach for Office add-ins (2025+). It provides seamless SSO inside Outlook's task pane without popup windows. Ensure the Azure AD app registration has the `brk-multihub://` redirect URI configured — without it, NAA will not work and users may experience authentication failures. Authentication tokens are stored in `sessionStorage` (not `localStorage`), so tokens are automatically cleared when the browser tab closes.

---

## Method 1: Integrated Apps in M365 Admin Center (Recommended)

This is the **recommended approach**. The M365 Admin Center's "Integrated Apps" feature handles deployment and automatically syncs with Intune for management. It provides a single deployment interface that covers Outlook on the web, Windows, Mac, iOS, and Android.

### Step 1: Access the M365 Admin Center

1. Navigate to **https://admin.microsoft.com**.
2. Sign in with your **Global Admin** or **Exchange Admin** account.
3. In the left sidebar, click **Settings** → **Integrated Apps**.

> **What you see:** The Integrated Apps page shows a list of currently deployed apps/add-ins with their status, type, and assigned users. A blue **"Get apps"** button is at the top, plus an **"Upload custom apps"** link.

### Step 2: Upload the Add-in

1. Click **"Upload custom apps"** at the top of the Integrated Apps page.
2. In the "Upload custom apps" flyout panel, select **"Office Add-in"** as the app type.
3. Choose the deployment method:
   - **Option A — Provide URL to manifest file (recommended):**
     - Select **"Provide link to manifest file"**
     - Enter the manifest URL:
       ```
       https://your-org.github.io/MailMerge-Pro/manifest.xml
       ```
     - Click **Validate**. A green checkmark confirms the manifest is valid.
   - **Option B — Upload manifest file:**
     - Select **"Upload manifest file"**
     - Click **Browse** and select the downloaded `manifest.xml` file.
     - Click **Validate**.

> **What you see:** The validation step shows the add-in name ("MailMerge-Pro"), version, description, and required permissions. A green "Manifest is valid" badge appears if everything checks out. Red errors indicate issues with the manifest XML.

4. Click **Next**.

### Step 3: Assign Users

1. On the "Assign Users" page, choose who gets the add-in:
   - **Entire organization** — All users in the tenant.
   - **Specific users/groups** — Select Azure AD security groups or individual users.
   - **Just me** — For testing (deploys only to your account).

> **What you see:** A user/group picker with search functionality. Selected groups appear as chips below the search box. The page shows a count: "This app will be deployed to X users."

2. Set the deployment type:
   - **Fixed (default)** — Users cannot remove the add-in; it's always available.
   - **Available** — Users can optionally install it from the "Admin Managed" section in "Get Add-ins."
3. Click **Next**.

### Step 4: Accept Permissions

1. Review the permissions the add-in requests:
   - **ReadWriteMailbox** — Compose and send emails
   - **ReadItem** — Read the current email being composed
2. If the add-in requires Microsoft Graph permissions (e.g., for contacts import), you'll see an **"Accept permissions"** button. Click it to grant **admin consent** on behalf of the organization.

> **What you see:** A Microsoft consent dialog listing the permissions with descriptions. The page title says "Permissions requested by MailMerge-Pro." An "Accept" button is at the bottom.

3. Click **Next**.

### Step 5: Review and Deploy

1. Review the summary:
   - **App name:** MailMerge-Pro
   - **Host apps:** Outlook
   - **Assigned users:** [your selection]
   - **Deployment type:** Fixed / Available
   - **Permissions:** Accepted ✓
2. Click **"Finish deployment"**.

> **What you see:** A "Deployment complete" confirmation page with a green checkmark. A note says: "It might take up to 24 hours for the app to appear for assigned users." A link to view the app in the Integrated Apps list is provided.

3. Click **Done**.

### Step 6: Verify Deployment

1. Navigate to **Integrated Apps** in the admin center.
2. Find **MailMerge-Pro** in the list. Status should show **"Deployed"**.
3. Click the add-in to see details:
   - Assigned users/groups
   - Deployment method
   - Last updated date
   - Status per platform (Web, Windows, Mac, Mobile)

### Propagation Timeline

| Platform | Typical Time | Maximum |
|---|---|---|
| Outlook on the Web | 1-4 hours | 24 hours |
| Outlook Desktop (Windows) | 4-12 hours | 24 hours |
| Outlook Desktop (Mac) | 4-12 hours | 24 hours |
| Outlook Mobile (iOS/Android) | 12-24 hours | 48 hours |

---

## Method 2: Intune Configuration Profiles

Use this method when you need more granular control over deployment or when managing Office Add-ins alongside other Intune device configurations.

### Step 1: Prepare the Add-in Policy XML

Create an XML file that defines the add-in deployment policy. This policy instructs Office to install the add-in from the manifest URL.

```xml
<DefaultApps>
  <App>
    <Id>MailMerge-Pro-Id</Id>
    <StoreId>WA200000000</StoreId>
    <SourceUrl>https://your-org.github.io/MailMerge-Pro/manifest.xml</SourceUrl>
    <DefaultState>enabled</DefaultState>
    <RequestAdmin>true</RequestAdmin>
  </App>
</DefaultApps>
```

### Step 2: Create a Configuration Profile in Intune

1. Navigate to **https://intune.microsoft.com** (Microsoft Intune admin center).
2. Go to **Devices** → **Configuration profiles** → **Create profile**.

> **What you see:** The "Create a profile" page with platform and profile type dropdowns. The page shows a list of existing profiles with their assignment status.

3. Select:
   - **Platform:** Windows 10 and later
   - **Profile type:** Templates → **Administrative Templates**
4. Click **Create**.

### Step 3: Configure the Profile

1. **Basics tab:**
   - **Name:** `MailMerge-Pro Office Add-in Deployment`
   - **Description:** `Deploys the MailMerge-Pro mail merge add-in to Outlook for assigned users.`
   - Click **Next**.

2. **Configuration settings tab:**
   - Browse to: **User Configuration** → **Microsoft Office 2016** → **Security Settings** → **Trust Center** → **Trusted Add-in Catalogs**
   - Enable **"Trusted Web Add-in Catalogs"**
   - Add the catalog URL: `https://your-org.github.io/MailMerge-Pro/`
   - Set **"Block web add-ins not from trusted catalogs"** to **Not configured** (or Disabled).

   Alternatively, use **OMA-URI** settings:
   - Click **Add** → **OMA-URI**
   - **Name:** `MailMerge-Pro Add-in`
   - **OMA-URI:** `./User/Vendor/MSFT/Policy/Config/Office16v2~Policy~L_MicrosoftOfficeOutlook~L_Security/L_TrustedAddInCatalogs`
   - **Data type:** String
   - **Value:** `https://your-org.github.io/MailMerge-Pro/`

> **What you see:** The OMA-URI configuration panel with fields for Name, Description, OMA-URI path, Data type dropdown, and Value. A "Save" button at the bottom.

3. Click **Next**.

### Step 4: Assign the Profile

1. **Assignments tab:**
   - Under **Included groups**, click **Add groups**.
   - Search for and select your target Azure AD groups:
     - `MailMerge-Pro-Pilot` (for initial rollout)
     - `All Company Users` (for full rollout)
   - Optionally add **Excluded groups** to exempt specific users.

> **What you see:** The assignments page with "Included groups" and "Excluded groups" sections. Each section has an "Add groups" button that opens a group picker flyout with search. Selected groups appear as a list with remove buttons.

2. Click **Next**.

### Step 5: Review and Create

1. **Applicability Rules tab:** (Optional)
   - Add rules to target specific OS versions or editions if needed.
   - Click **Next**.
2. **Review + create tab:**
   - Verify all settings.
   - Click **Create**.

> **What you see:** A summary page showing profile name, platform, settings configured, and assigned groups. A "Create" button at the bottom. After creation, a green banner confirms: "Profile created successfully."

### Step 6: Deploy an Accompanying Script (Optional)

For the configuration profile to trigger add-in installation, you may need to deploy a companion PowerShell script (see Method 3) that registers the manifest with Office.

---

## Method 3: PowerShell Scripts via Intune

Use this method for maximum control, scripted deployments, or when the admin template approach doesn't work for your environment.

### Step 1: Create the Deployment Script

Create a PowerShell script named `Deploy-MailMergePro.ps1`:

```powershell
<#
.SYNOPSIS
    Deploys MailMerge-Pro Outlook Web Add-in for the current user.
.DESCRIPTION
    Registers the MailMerge-Pro manifest URL with the Office web add-in
    catalog via registry keys so that Outlook loads the add-in automatically.
.NOTES
    Deploy via Intune > Devices > Scripts
#>

param(
    [string]$ManifestUrl = "https://your-org.github.io/MailMerge-Pro/manifest.xml",
    [string]$AddInId = "MailMerge-Pro-Id"
)

$ErrorActionPreference = "Stop"

try {
    # Log start
    $logPath = "$env:ProgramData\MailMerge-Pro\deployment.log"
    New-Item -Path (Split-Path $logPath) -ItemType Directory -Force | Out-Null
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "$timestamp - Starting MailMerge-Pro deployment"

    # Set registry key for the trusted catalog
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\$AddInId"
    
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $regPath -Name "Id" -Value $AddInId
    Set-ItemProperty -Path $regPath -Name "Url" -Value $ManifestUrl
    Set-ItemProperty -Path $regPath -Name "Flags" -Value 1 -Type DWord

    # Register the manifest URL in the web extensions catalog
    $wefPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
    
    if (-not (Test-Path $wefPath)) {
        New-Item -Path $wefPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $wefPath -Name $AddInId -Value $ManifestUrl

    Add-Content -Path $logPath -Value "$timestamp - MailMerge-Pro deployment completed successfully"
    Write-Output "MailMerge-Pro add-in deployed successfully."
    exit 0
}
catch {
    $errorMsg = $_.Exception.Message
    Add-Content -Path $logPath -Value "$timestamp - ERROR: $errorMsg"
    Write-Error "Failed to deploy MailMerge-Pro: $errorMsg"
    exit 1
}
```

### Step 2: Upload Script to Intune

1. Navigate to **https://intune.microsoft.com**.
2. Go to **Devices** → **Scripts and remediations** → **Platform scripts** tab.
3. Click **Add** → **Windows 10 and later**.

> **What you see:** The "Add PowerShell script" wizard with tabs for Basics, Script settings, Assignments, and Review + create. The Basics tab has Name and Description fields.

4. **Basics tab:**
   - **Name:** `Deploy MailMerge-Pro Add-in`
   - **Description:** `Registers the MailMerge-Pro manifest URL for Outlook Web Add-in installation.`
   - Click **Next**.

### Step 3: Configure Script Settings

1. **Script settings tab:**
   - **Script location:** Click **Browse** and upload `Deploy-MailMergePro.ps1`.
   - **Run this script using the logged-on credentials:** **Yes** (critical — registry keys are per-user HKCU).
   - **Enforce script signature check:** **No** (unless you sign the script).
   - **Run script in 64-bit PowerShell host:** **Yes**.

> **What you see:** The script settings page with a file upload area, toggle switches for credential context and signature check, and a radio button for 32-bit vs 64-bit PowerShell. The uploaded script filename appears below the upload area.

2. Click **Next**.

### Step 4: Assign to Groups

1. **Assignments tab:**
   - Under **Included groups**, click **Add groups**.
   - Select target groups (e.g., `MailMerge-Pro-Users`).
   - Optionally set **Excluded groups**.
2. Click **Next**.

### Step 5: Review and Create

1. **Review + create tab:**
   - Verify script name, settings, and group assignments.
   - Click **Add**.

> **What you see:** Summary showing script name, file name, execution context (User), and assigned groups. A blue "Add" button at the bottom. After creation, the script appears in the list with "Pending" status.

### Script Execution Timeline

- Scripts run once per device per user (unless configured to retry).
- Initial execution: within **1-8 hours** of assignment (at next Intune sync).
- To force immediate execution: user can go to **Settings → Accounts → Access work or school → [their account] → Info → Sync**.

---

## Group Targeting

### Recommended Group Strategy

| Phase | Group Name | Members | Purpose |
|---|---|---|---|
| **Pilot** | `MailMerge-Pro-Pilot` | 5-10 IT staff and power users | Initial testing and validation |
| **Early Adopters** | `MailMerge-Pro-EarlyAccess` | 50-100 business users | Broader validation before org-wide |
| **Production** | `All Users` or `MailMerge-Pro-Users` | Entire organization | Full rollout |
| **Exclusion** | `MailMerge-Pro-Exclude` | Service accounts, shared mailboxes | Users who should NOT get the add-in |

### Creating Groups in Azure AD

1. Go to **https://entra.microsoft.com** (Azure AD admin center).
2. Navigate to **Groups** → **All groups** → **New group**.
3. Configure:
   - **Group type:** Security
   - **Group name:** `MailMerge-Pro-Pilot`
   - **Membership type:** Assigned (manual) or Dynamic (rule-based)
4. For **Dynamic membership**, use a rule like:
   ```
   (user.department -eq "Marketing") or (user.department -eq "Sales")
   ```
5. Click **Create**.

### Phased Rollout Plan

1. **Week 1:** Deploy to `MailMerge-Pro-Pilot`. Collect feedback.
2. **Week 2-3:** Deploy to `MailMerge-Pro-EarlyAccess`. Monitor for issues.
3. **Week 4+:** Deploy to `All Users`. Announce via company communication.

---

## Updating the Add-in

### For Code Changes (No Redeployment Needed)

Since the add-in is a **web application hosted on GitHub Pages** (or your web server), code updates require **zero redeployment through Intune or the Admin Center**.

1. Push your code changes to the GitHub repository's `main` branch (or your deployment branch).
2. GitHub Pages automatically builds and deploys the updated files.
3. The next time a user opens the add-in in Outlook, they get the latest version.
4. **That's it.** No Intune changes. No admin center changes. No user action required.

> **Why this works:** The Intune/Admin Center deployment only references the `manifest.xml` URL. The manifest points to the web app URL where the actual add-in code lives. When the code is updated at that URL, users automatically get the new version.

### For Manifest Changes

If you change the **manifest.xml** itself (e.g., adding new permissions, changing the add-in name, updating the version number):

1. Update the `manifest.xml` in your repository and push to your hosting.
2. **Method 1 (Admin Center):**
   - Go to **Integrated Apps** → Click **MailMerge-Pro** → Click **"Update"**.
   - Revalidate the manifest URL or re-upload the file.
   - Review any new permissions → **Accept** → **Update**.
3. **Method 2/3 (Intune):**
   - If the manifest URL hasn't changed, no action needed — Outlook fetches the latest manifest.
   - If the URL changed, update the configuration profile or script with the new URL.

### Version Numbering

Update the version in `manifest.xml` following semantic versioning:
```xml
<Version>3.0.0</Version>
```
- **Major** (2.0.0): Breaking changes or major feature overhaul
- **Minor** (1.2.0): New features, backward-compatible
- **Patch** (1.1.1): Bug fixes

---

## Removing / Uninstalling

### Method 1: Admin Center Removal

1. Go to **admin.microsoft.com** → **Settings** → **Integrated Apps**.
2. Click **MailMerge-Pro** in the list.
3. Click **"Remove"** at the top of the detail pane.
4. Confirm: **"Yes, remove this app"**.

> **What you see:** A confirmation dialog warning "This will remove the app for all assigned users. Users will no longer be able to access the app." with "Remove" and "Cancel" buttons.

5. The add-in disappears from users' Outlook within 24 hours.

### Method 2: Remove Intune Configuration Profile

1. Go to **intune.microsoft.com** → **Devices** → **Configuration profiles**.
2. Find the MailMerge-Pro profile.
3. Click **"Delete"** → Confirm.
4. The registry keys are removed at the next Intune sync.

### Method 3: Remove via PowerShell (Intune Script)

Deploy a removal script:

```powershell
<#
.SYNOPSIS
    Removes MailMerge-Pro Outlook Web Add-in for the current user.
#>

$AddInId = "MailMerge-Pro-Id"

try {
    # Remove trusted catalog entry
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\$AddInId"
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force
    }

    # Remove developer entry
    $wefPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
    if (Test-Path $wefPath) {
        Remove-ItemProperty -Path $wefPath -Name $AddInId -ErrorAction SilentlyContinue
    }

    Write-Output "MailMerge-Pro add-in removed successfully."
    exit 0
}
catch {
    Write-Error "Failed to remove MailMerge-Pro: $($_.Exception.Message)"
    exit 1
}
```

### Per-User Self-Removal (if Deployment Type = "Available")

If the add-in was deployed as "Available" (not "Fixed"):
1. In Outlook, click **"Get Add-ins"** in the ribbon.
2. Go to the **"Admin Managed"** tab.
3. Find MailMerge-Pro → Click **"Remove"**.

> **Note:** If deployed as "Fixed," users **cannot** remove it — only admins can.

---

## Monitoring Deployment Status

### In M365 Admin Center

1. Go to **admin.microsoft.com** → **Settings** → **Integrated Apps**.
2. Click **MailMerge-Pro**.
3. The detail pane shows:
   - **Status:** Deployed / Deploying / Failed
   - **Assigned users/groups:** List of groups with member counts
   - **Platform availability:** Check marks for Web ✓, Windows ✓, Mac ✓, Mobile ✓

### In Intune (for Method 2 & 3)

1. Go to **intune.microsoft.com** → **Devices** → **Configuration profiles** (Method 2) or **Scripts** (Method 3).
2. Click the MailMerge-Pro profile/script.
3. View the **"Device status"** or **"User status"** tab:

> **What you see:** A table with columns: User, Device, Status (Succeeded/Pending/Error/Conflict), Last check-in time. A pie chart summary shows the breakdown. Drill-down links for each status.

   | Status | Meaning |
   |---|---|
   | **Succeeded** | Script ran / profile applied successfully |
   | **Pending** | Not yet synced; waiting for device check-in |
   | **Error** | Script failed; click to see error details |
   | **Conflict** | Another profile conflicts with this one |
   | **Not applicable** | Device/user doesn't meet applicability rules |

### Using PowerShell for Bulk Status Check

```powershell
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"

# Get all device configuration profiles
$profiles = Get-MgDeviceManagementDeviceConfiguration | 
    Where-Object { $_.DisplayName -like "*MailMerge*" }

# Check deployment status
foreach ($profile in $profiles) {
    $statuses = Get-MgDeviceManagementDeviceConfigurationDeviceStatus -DeviceConfigurationId $profile.Id
    Write-Output "Profile: $($profile.DisplayName)"
    Write-Output "  Success: $(($statuses | Where-Object Status -eq 'succeeded').Count)"
    Write-Output "  Pending: $(($statuses | Where-Object Status -eq 'pending').Count)"
    Write-Output "  Error: $(($statuses | Where-Object Status -eq 'error').Count)"
}
```

### End-User Verification

Ask users to verify the add-in is available:

1. Open **Outlook** (web or desktop).
2. Open **New Email** compose window.
3. Look for the **MailMerge-Pro** icon in the ribbon toolbar.
4. If not visible on Outlook Web, check **"…" (More actions)** menu.
5. If not visible at all, click **"Get Add-ins"** → **"Admin Managed"** tab — it should be listed there.

---

## Troubleshooting

### Important: localStorage Data (v3.0)

MailMerge-Pro v3.0 stores several types of data in the browser's `localStorage` on each user's device. **This data is per-device and per-browser — it does NOT roam with the user's profile and is NOT synced via Intune, OneDrive, or Enterprise State Roaming.**

| localStorage Data | Description | Roaming? | Backed Up by Intune? |
|---|---|---|---|
| Email templates | User's saved custom email templates | ❌ No | ❌ No |
| Contact groups | User's saved recipient lists | ❌ No | ❌ No |
| Campaign history | Past campaign records and dashboard stats | ❌ No | ❌ No |
| Language preference | User's selected UI language (EN/ES/FR/DE/PT/JA) | ❌ No | ❌ No |
| Signature cache | User's cached email signature | ❌ No | ❌ No |

**Implications for IT admins:**
- If a user switches to a new device, their templates, contact groups, and campaign history do **not** transfer automatically. They will need to recreate them.
- Clearing the browser cache/data on a managed device (e.g., via Intune remediation script) will erase these items.
- **Scheduled sends** require Outlook and the add-in task pane to remain open. If a device restart policy or sleep policy interrupts Outlook, scheduled sends will not execute.
- localStorage data is **not visible** to administrators through any admin console. It is private to the user on that device.

| Issue | Cause | Solution |
|---|---|---|
| Add-in doesn't appear after 24h | Deployment not propagated | Check admin center status; verify user is in assigned group |
| "This add-in has been disabled" | Admin or user disabled it | Re-enable in admin center or Outlook add-in manager |
| "We can't load the add-in" | Manifest URL unreachable | Verify manifest URL is accessible (try in browser); check SSL cert |
| Script shows "Error" in Intune | Script execution failed | Check error details in Intune; verify script runs locally |
| Add-in works on Web but not Desktop | Desktop caching | Clear Office cache: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` |
| "Trusted catalog" error | Catalog URL not trusted | Add URL to trusted catalogs via Group Policy or Intune |
| Consent prompt appears for each user | Admin consent not granted | Grant admin consent in Azure AD → Enterprise Apps → MailMerge-Pro |
| Mobile not showing add-in | Mobile support lag | Wait 48 hours; verify manifest has `MobileFormFactor` entry |
| Templates/groups missing after device swap | localStorage is device-specific | Templates, contact groups, and campaign history are stored in browser localStorage — they do not roam across devices |
| Scheduled send did not execute | Outlook or task pane was closed | Scheduled sends require Outlook and the task pane to remain open. Device restart or sleep policies may interrupt. |
| Language reverts to English on new device | Language preference in localStorage | User must re-select their language on each new device/browser |

### Clearing the Office Web Add-in Cache (Windows)

If users experience issues with a stale version:

```powershell
# Close Outlook first
Stop-Process -Name "OUTLOOK" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# Clear the WEF cache
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force

# Restart Outlook
Start-Process "outlook.exe"
```

### Useful Diagnostic URLs

| Tool | URL |
|---|---|
| M365 Admin Center | https://admin.microsoft.com |
| Intune Admin Center | https://intune.microsoft.com |
| Azure AD Admin Center | https://entra.microsoft.com |
| Service Health Dashboard | https://admin.microsoft.com/Adminportal/Home#/servicehealth |
| Manifest Validator | https://dev.office.com/manifest-validator |
| Office Add-in Debug Mode | Append `?et=debug` to the add-in URL |

---

*© 2026 MailMerge-Pro. For internal IT administration use.*
