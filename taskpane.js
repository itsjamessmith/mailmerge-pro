/* MailMerge-Pro — Outlook Web Add-in JavaScript
 * MSAL.js 2.0 + Microsoft Graph API
 * NO alert/confirm/prompt — all feedback via DOM
 */
"use strict";

// ========== MSAL Configuration ==========
const msalConfig = {
    auth: {
        clientId: "360e4343-614f-4f70-a650-c020868516fc",
        authority: "https://login.microsoftonline.com/e67c588e-f654-4727-b794-1ca5df7b6ee9",
        redirectUri: "https://itsjamessmith.github.io/mailmerge-pro/taskpane.html"
    },
    cache: { cacheLocation: "localStorage" }
};
const loginRequest = { scopes: ["Mail.Send", "Mail.ReadWrite", "User.Read", "Contacts.Read", "Mail.Send.Shared"] };
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
let msalInstance = null;

// ========== Application State ==========
const appState = {
    currentStep: 1,
    headers: [],
    rows: [],
    mapping: { to: "", cc: "", bcc: "", subject: "", attachments: "" },
    globalAttachments: new Map(),
    perRecipientFiles: new Map(),
    defaults: {},
    results: [],
    sending: false,
    userEmail: "",
    previewIndex: 0,
    contactsData: []
};

// ========== Office.js Initialization ==========
Office.onReady(async function (info) {
    console.log("Office.onReady:", info.host, info.platform);
    detectDarkMode(info);
    await initMsal();
    initUI();
    showOnboarding();
});
if (typeof Office === "undefined") {
    document.addEventListener("DOMContentLoaded", async () => {
        detectDarkMode(null);
        await initMsal();
        initUI();
        showOnboarding();
    });
}

// ========== Dark Mode Detection ==========
function detectDarkMode(officeInfo) {
    const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    if (prefersDark) document.body.classList.add("dark-mode");
    try {
        if (officeInfo && typeof Office !== "undefined" && Office.context && Office.context.officeTheme) {
            const bg = Office.context.officeTheme.bodyBackgroundColor;
            if (bg && isColorDark(bg)) document.body.classList.add("dark-mode");
        }
    } catch (e) { console.log("Theme detection skipped:", e.message); }
}
function isColorDark(hex) {
    hex = hex.replace("#", "");
    if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
    const r = parseInt(hex.substr(0,2),16), g = parseInt(hex.substr(2,2),16), b = parseInt(hex.substr(4,2),16);
    return (r*299 + g*587 + b*114) / 1000 < 128;
}

// ========== Onboarding ==========
function showOnboarding() {
    if (localStorage.getItem("mailmerge-pro-onboarded")) return;
    const el = document.getElementById("onboardingOverlay");
    if (el) el.style.display = "flex";
}
function dismissOnboarding() {
    localStorage.setItem("mailmerge-pro-onboarded", "true");
    const el = document.getElementById("onboardingOverlay");
    if (el) el.style.display = "none";
}

// ========== MSAL Initialization ==========
async function initMsal() {
    console.log("initMsal: creating PublicClientApplication");
    const authEl = document.getElementById("authStatus");
    const btnIn = document.getElementById("btnSignIn");
    const btnOut = document.getElementById("btnSignOut");
    if (typeof msal === "undefined" || !msal.PublicClientApplication) {
        console.error("MSAL library not loaded");
        authEl.textContent = "Auth library not loaded. Refresh page.";
        authEl.classList.add("error");
        btnIn.style.display = "inline-block";
        return;
    }
    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        console.log("MSAL initialized");
        try {
            const resp = await msalInstance.handleRedirectPromise();
            if (resp) console.log("handleRedirectPromise: got token for", resp.account.username);
        } catch (e) { console.warn("handleRedirectPromise error:", e); }
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            console.log("Already signed in:", accounts[0].username);
            updateAuthUI(accounts[0]);
        } else {
            authEl.textContent = "Not signed in";
            btnIn.style.display = "inline-block";
            btnOut.style.display = "none";
        }
    } catch (err) {
        console.error("MSAL init error:", err);
        authEl.textContent = "Auth init failed";
        authEl.classList.add("error");
        btnIn.style.display = "inline-block";
    }
}

function updateAuthUI(account) {
    const authEl = document.getElementById("authStatus");
    const btnIn = document.getElementById("btnSignIn");
    const btnOut = document.getElementById("btnSignOut");
    if (account) {
        authEl.textContent = account.username;
        authEl.classList.add("signed-in");
        authEl.classList.remove("error");
        btnIn.style.display = "none";
        btnOut.style.display = "inline-block";
        appState.userEmail = account.username;
    } else {
        authEl.textContent = "Not signed in";
        authEl.classList.remove("signed-in");
        btnIn.style.display = "inline-block";
        btnOut.style.display = "none";
        appState.userEmail = "";
    }
}

async function signIn() {
    console.log("signIn: starting interactive login");
    const authEl = document.getElementById("authStatus");
    if (!msalInstance) {
        try {
            msalInstance = new msal.PublicClientApplication(msalConfig);
            await msalInstance.initialize();
            await msalInstance.handleRedirectPromise().catch(() => {});
        } catch (e) {
            authEl.textContent = "Auth failed: " + (e.message || String(e));
            authEl.classList.add("error");
            return;
        }
    }
    try {
        const result = await msalInstance.acquireTokenPopup(loginRequest);
        console.log("signIn success:", result.account.username);
        updateAuthUI(result.account);
        return result.accessToken;
    } catch (err) {
        console.error("signIn error:", err);
        if (err.message && err.message.includes("popup_window_error")) {
            authEl.textContent = "Pop-up blocked. Allow pop-ups and try again.";
        } else if (err.message && err.message.includes("user_cancelled")) {
            authEl.textContent = "Sign-in cancelled.";
        } else {
            authEl.textContent = "Sign-in failed: " + (err.message || String(err));
        }
        authEl.classList.add("error");
        throw err;
    }
}

function signOut() {
    console.log("signOut");
    const accounts = msalInstance ? msalInstance.getAllAccounts() : [];
    if (accounts.length > 0) msalInstance.clearCache();
    updateAuthUI(null);
}

async function getGraphToken() {
    if (!msalInstance) {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();
        await msalInstance.handleRedirectPromise().catch(() => {});
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const r = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            console.log("getGraphToken: silent OK");
            updateAuthUI(accounts[0]);
            return r.accessToken;
        } catch (e) { console.warn("Silent token failed:", e.message); }
    }
    return await signIn();
}

// ========== UI Initialization ==========
function initUI() {
    console.log("initUI");
    // Auth
    document.getElementById("btnSignIn").addEventListener("click", async () => { try { await signIn(); } catch(_){} });
    document.getElementById("btnSignOut").addEventListener("click", signOut);
    // Onboarding
    document.getElementById("btnDismissOnboarding").addEventListener("click", dismissOnboarding);
    // Step 1
    document.getElementById("fileInput").addEventListener("change", handleFileUpload);
    document.getElementById("btnImportContacts").addEventListener("click", importContacts);
    document.getElementById("recipientSearch").addEventListener("input", filterRecipients);
    document.getElementById("btnContactsCancel").addEventListener("click", () => { document.getElementById("contactsPanel").style.display = "none"; });
    document.getElementById("btnContactsSelectAll").addEventListener("click", toggleSelectAllContacts);
    document.getElementById("btnContactsUse").addEventListener("click", useSelectedContacts);
    // Navigation
    document.getElementById("btnStep1Next").addEventListener("click", () => goToStep(2));
    document.getElementById("btnStep2Back").addEventListener("click", () => goToStep(1));
    document.getElementById("btnStep2Next").addEventListener("click", () => goToStep(3));
    document.getElementById("btnStep3Back").addEventListener("click", () => goToStep(2));
    document.getElementById("btnStep3Next").addEventListener("click", () => goToStep(4));
    document.getElementById("btnStep4Back").addEventListener("click", () => goToStep(3));
    // Send
    document.getElementById("btnSend").addEventListener("click", () => executeMerge(false));
    document.getElementById("btnTestEmail").addEventListener("click", () => executeMerge(true));
    // Preview carousel
    document.getElementById("btnPrevPreview").addEventListener("click", () => changePreview(-1));
    document.getElementById("btnNextPreview").addEventListener("click", () => changePreview(1));
    // Attachments
    document.getElementById("globalAttachmentInput").addEventListener("change", handleGlobalAttachmentUpload);
    document.getElementById("perRecipientAttachmentInput").addEventListener("change", handlePerRecipientAttachmentUpload);
    // Rich editor
    initEditorToolbar();
    initEditorPlaceholder();
    document.getElementById("fontColorPicker").addEventListener("input", (e) => {
        document.execCommand("foreColor", false, e.target.value);
    });
    // Link dialog
    document.getElementById("btnLinkCancel").addEventListener("click", () => { document.getElementById("linkDialog").style.display = "none"; });
    document.getElementById("btnLinkInsert").addEventListener("click", insertLink);
    document.getElementById("linkUrlInput").addEventListener("keydown", (e) => { if (e.key === "Enter") insertLink(); });
    // Unsubscribe toggle
    document.getElementById("chkUnsubscribe").addEventListener("change", (e) => {
        document.getElementById("unsubscribeOptions").style.display = e.target.checked ? "block" : "none";
    });
    // Group by email toggle refreshes preview
    document.getElementById("chkGroupByEmail").addEventListener("change", () => {
        if (appState.currentStep === 4) { appState.previewIndex = 0; buildReview(); buildPreview(); }
    });
    // Campaign
    loadCampaignHistory();
    document.getElementById("pastCampaigns").addEventListener("change", showCampaignDetails);
    document.getElementById("btnCloseCampaignDetail").addEventListener("click", () => {
        document.getElementById("campaignDetailModal").style.display = "none";
    });
    // Keyboard shortcut: Ctrl+Enter to send
    document.addEventListener("keydown", (e) => {
        if (e.ctrlKey && e.key === "Enter" && appState.currentStep === 4 && !appState.sending) {
            e.preventDefault();
            executeMerge(false);
        }
    });
}

// ========== Rich Text Editor ==========
function initEditorToolbar() {
    document.getElementById("editorToolbar").querySelectorAll("button[data-cmd]").forEach(btn => {
        btn.addEventListener("mousedown", (e) => {
            e.preventDefault(); // prevent losing editor focus
            const cmd = btn.dataset.cmd;
            if (cmd === "createLink") { showLinkDialog(); return; }
            document.execCommand(cmd, false, null);
        });
    });
}
function initEditorPlaceholder() {
    const editor = document.getElementById("emailBody");
    const update = () => {
        const c = editor.innerHTML;
        if (!c || c === "<br>" || c === "<div><br></div>") editor.classList.add("is-empty");
        else editor.classList.remove("is-empty");
    };
    editor.addEventListener("input", update);
    editor.addEventListener("focus", update);
    editor.addEventListener("blur", update);
    editor.classList.add("is-empty");
}
function showLinkDialog() {
    document.getElementById("linkUrlInput").value = "";
    document.getElementById("linkDialog").style.display = "flex";
    document.getElementById("linkUrlInput").focus();
}
function insertLink() {
    const url = document.getElementById("linkUrlInput").value.trim();
    document.getElementById("linkDialog").style.display = "none";
    if (url) {
        document.getElementById("emailBody").focus();
        document.execCommand("createLink", false, url);
    }
}
function getEditorContent() {
    const editor = document.getElementById("emailBody");
    const html = editor.innerHTML;
    if (!html || html === "<br>" || html === "<div><br></div>") return "";
    return html;
}
function getEditorPlainText() {
    return document.getElementById("emailBody").innerText || "";
}
function insertTextInEditor(text) {
    const editor = document.getElementById("emailBody");
    editor.focus();
    const sel = window.getSelection();
    if (sel.rangeCount) {
        const range = sel.getRangeAt(0);
        range.deleteContents();
        range.insertNode(document.createTextNode(text));
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
    } else {
        editor.appendChild(document.createTextNode(text));
    }
    editor.classList.remove("is-empty");
}

// ========== Campaign History ==========
function loadCampaignHistory() {
    const stored = localStorage.getItem("mailmerge-pro-campaigns");
    const campaigns = stored ? JSON.parse(stored) : [];
    const sel = document.getElementById("pastCampaigns");
    sel.innerHTML = '<option value="">\u{1F4DC} History</option>';
    campaigns.forEach((c, i) => {
        const opt = document.createElement("option");
        opt.value = i;
        opt.textContent = c.name + " (" + c.date + ") " + c.sent + "/" + c.total;
        sel.appendChild(opt);
    });
}
function saveCampaign(name, total, sent, errors) {
    const stored = localStorage.getItem("mailmerge-pro-campaigns");
    const campaigns = stored ? JSON.parse(stored) : [];
    campaigns.unshift({ name: name || "Untitled", date: new Date().toLocaleDateString(), total: total, sent: sent, errors: errors });
    if (campaigns.length > 20) campaigns.length = 20;
    localStorage.setItem("mailmerge-pro-campaigns", JSON.stringify(campaigns));
    loadCampaignHistory();
}
function showCampaignDetails() {
    const sel = document.getElementById("pastCampaigns");
    const idx = sel.value;
    if (idx === "") return;
    const stored = localStorage.getItem("mailmerge-pro-campaigns");
    const campaigns = stored ? JSON.parse(stored) : [];
    const c = campaigns[parseInt(idx)];
    if (!c) return;
    document.getElementById("campaignDetailContent").innerHTML =
        "<p><strong>Name:</strong> " + escapeHtml(c.name) + "</p>" +
        "<p><strong>Date:</strong> " + escapeHtml(c.date) + "</p>" +
        "<p><strong>Total:</strong> " + c.total + "</p>" +
        "<p><strong>Sent:</strong> " + c.sent + "</p>" +
        "<p><strong>Errors:</strong> " + c.errors + "</p>";
    document.getElementById("campaignDetailModal").style.display = "flex";
    sel.value = "";
}

// ========== Step Navigation ==========
function goToStep(step) {
    if (appState.sending) return;
    if (appState.currentStep === 2) saveMapping();
    if (appState.currentStep === 3) saveDefaults();

    if (step > appState.currentStep) {
        if (appState.currentStep === 1 && appState.rows.length === 0) {
            showStatus("\u26A0\uFE0F Upload a data file or import contacts first.", "warning");
            return;
        }
        if (appState.currentStep === 2 && !document.getElementById("mapTo").value) {
            showStatus("\u26A0\uFE0F Select the To (Email) column.", "warning");
            return;
        }
        if (appState.currentStep === 3) {
            const subj = document.getElementById("emailSubject").value;
            const body = getEditorContent();
            if (!subj && !appState.mapping.subject) {
                showStatus("\u26A0\uFE0F Enter a subject line or map a Subject column.", "warning");
                return;
            }
            if (!body && !getEditorPlainText()) {
                showStatus("\u26A0\uFE0F Enter an email body.", "warning");
                return;
            }
        }
    }

    if (step === 2) populateColumnDropdowns();
    if (step === 3) { populateMergeFieldButtons(); updatePerRecipientAttachmentVisibility(); populateDefaults(); }
    if (step === 4) { appState.previewIndex = 0; buildReview(); buildPreview(); }

    document.querySelectorAll(".step-content").forEach(el => el.classList.remove("active"));
    document.getElementById("step" + step).classList.add("active");

    document.querySelectorAll(".step-item").forEach(el => {
        const s = parseInt(el.dataset.step);
        el.classList.remove("active", "done");
        if (s === step) el.classList.add("active");
        else if (s < step) el.classList.add("done");
    });

    appState.currentStep = step;
    hideResults();
}
function markStepComplete(s) {
    const item = document.querySelector('.step-item[data-step="' + s + '"]');
    if (item) item.classList.add("done");
}
function updateStepBadge(step, text) {
    const badge = document.getElementById("badge" + step);
    if (!badge) return;
    if (text) { badge.textContent = text; badge.classList.add("visible"); }
    else { badge.textContent = ""; badge.classList.remove("visible"); }
}

// ========== Step 1: File Upload ==========
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById("fileName").textContent = file.name;
    console.log("handleFileUpload:", file.name, file.size);
    const reader = new FileReader();
    reader.onload = function (evt) {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            if (jsonData.length === 0) { showStatus("\u26A0\uFE0F No data rows found.", "warning"); return; }
            appState.headers = Object.keys(jsonData[0]);
            appState.rows = jsonData;
            renderDataPreview();
            document.getElementById("btnStep1Next").disabled = false;
            updateStepBadge(1, String(appState.rows.length));
            console.log("Parsed", appState.rows.length, "rows,", appState.headers.length, "columns");
        } catch (err) {
            console.error("File parse error:", err);
            showStatus("\u274C Error reading file: " + err.message, "error");
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderDataPreview() {
    document.getElementById("dataStats").textContent =
        "\u2705 " + appState.rows.length + " recipients, " + appState.headers.length + " columns: " + appState.headers.join(", ");
    let t = "<table><tr>";
    appState.headers.forEach(h => { t += "<th>" + escapeHtml(h) + "</th>"; });
    t += "</tr>";
    appState.rows.forEach((row, i) => {
        t += '<tr data-idx="' + i + '">';
        appState.headers.forEach(h => { t += "<td>" + escapeHtml(String(row[h] || "")) + "</td>"; });
        t += "</tr>";
    });
    t += "</table>";
    document.getElementById("previewTable").innerHTML = t;
    document.getElementById("dataPreview").style.display = "block";
    document.getElementById("recipientSearch").value = "";
}

function filterRecipients() {
    const q = document.getElementById("recipientSearch").value.toLowerCase();
    const rows = document.querySelectorAll("#previewTable table tr[data-idx]");
    rows.forEach(tr => {
        if (!q) { tr.classList.remove("hidden-row"); return; }
        const text = tr.textContent.toLowerCase();
        tr.classList.toggle("hidden-row", !text.includes(q));
    });
}

// ========== Contacts Import ==========
async function importContacts() {
    const panel = document.getElementById("contactsPanel");
    const loading = document.getElementById("contactsLoading");
    const list = document.getElementById("contactsList");
    const actions = document.getElementById("contactsActions");
    panel.style.display = "block";
    loading.style.display = "block";
    list.innerHTML = "";
    actions.style.display = "none";
    try {
        const token = await getGraphToken();
        const resp = await graphFetch(GRAPH_BASE + "/me/contacts?$top=500&$select=displayName,emailAddresses", token, "GET");
        const contacts = resp.value || [];
        loading.style.display = "none";
        if (contacts.length === 0) { list.innerHTML = '<p class="hint">No contacts found.</p>'; return; }
        appState.contactsData = contacts;
        contacts.forEach((c, i) => {
            const email = c.emailAddresses && c.emailAddresses[0] ? c.emailAddresses[0].address : "";
            if (!email) return;
            const div = document.createElement("div");
            div.className = "contact-item";
            div.innerHTML = '<label><input type="checkbox" data-idx="' + i + '"/> ' + escapeHtml(c.displayName || "") + ' &lt;' + escapeHtml(email) + '&gt;</label>';
            list.appendChild(div);
        });
        actions.style.display = "flex";
    } catch (err) {
        loading.style.display = "none";
        list.innerHTML = '<p class="error-text">\u274C ' + escapeHtml(err.message) + "</p>";
    }
}
function toggleSelectAllContacts() {
    const boxes = document.querySelectorAll("#contactsList input[type=checkbox]");
    const allChecked = Array.from(boxes).every(b => b.checked);
    boxes.forEach(b => { b.checked = !allChecked; });
}
function useSelectedContacts() {
    const checked = document.querySelectorAll("#contactsList input:checked");
    if (checked.length === 0) return;
    const rows = [];
    checked.forEach(cb => {
        const idx = parseInt(cb.dataset.idx);
        const c = appState.contactsData[idx];
        if (!c) return;
        const email = c.emailAddresses[0].address;
        rows.push({ Name: c.displayName || "", Email: email });
    });
    appState.headers = ["Name", "Email"];
    appState.rows = rows;
    document.getElementById("contactsPanel").style.display = "none";
    renderDataPreview();
    document.getElementById("btnStep1Next").disabled = false;
    updateStepBadge(1, String(rows.length));
}

// ========== Step 2: Column Mapping ==========
function populateColumnDropdowns() {
    const selects = ["mapTo", "mapCC", "mapBCC", "mapSubject", "mapAttachments"];
    const autoMatch = {
        mapTo: ["email", "e-mail", "to", "emailaddress", "mail"],
        mapCC: ["cc", "carbon copy"],
        mapBCC: ["bcc", "blind"],
        mapSubject: ["subject", "subj"],
        mapAttachments: ["attachment", "attachments", "files", "file"]
    };
    selects.forEach(id => {
        const sel = document.getElementById(id);
        const prev = sel.value;
        sel.innerHTML = '<option value="">(none)</option>';
        appState.headers.forEach(h => {
            const opt = document.createElement("option");
            opt.value = h; opt.textContent = h;
            sel.appendChild(opt);
        });
        if (prev && appState.headers.includes(prev)) { sel.value = prev; return; }
        const matches = autoMatch[id] || [];
        for (const header of appState.headers) {
            if (matches.some(m => header.toLowerCase().includes(m))) { sel.value = header; break; }
        }
    });
}
function saveMapping() {
    appState.mapping.to = document.getElementById("mapTo").value;
    appState.mapping.cc = document.getElementById("mapCC").value;
    appState.mapping.bcc = document.getElementById("mapBCC").value;
    appState.mapping.subject = document.getElementById("mapSubject").value;
    appState.mapping.attachments = document.getElementById("mapAttachments").value;
}

// ========== Step 3: Compose ==========
function populateMergeFieldButtons() {
    const container = document.getElementById("mergeFieldBtns");
    container.innerHTML = "";
    appState.headers.forEach(h => {
        const btn = document.createElement("button");
        btn.textContent = "{" + h + "}";
        btn.type = "button";
        btn.addEventListener("click", () => insertTextInEditor("{" + h + "}"));
        container.appendChild(btn);
    });
}
function populateDefaults() {
    const container = document.getElementById("defaultsContainer");
    container.innerHTML = "";
    appState.headers.forEach(h => {
        const div = document.createElement("div");
        div.className = "default-row";
        const lbl = document.createElement("label");
        lbl.textContent = h + ":";
        const inp = document.createElement("input");
        inp.type = "text";
        inp.className = "input";
        inp.dataset.col = h;
        inp.value = appState.defaults[h] || "";
        inp.placeholder = "Default for {" + h + "}";
        div.appendChild(lbl);
        div.appendChild(inp);
        container.appendChild(div);
    });
}
function saveDefaults() {
    document.querySelectorAll("#defaultsContainer input").forEach(inp => {
        appState.defaults[inp.dataset.col] = inp.value;
    });
}
function updatePerRecipientAttachmentVisibility() {
    saveMapping();
    const section = document.getElementById("perRecipientAttachmentSection");
    if (appState.mapping.attachments) { section.style.display = "block"; checkMissingAttachments(); }
    else { section.style.display = "none"; }
}

// ========== Attachment Handling ==========
function handleGlobalAttachmentUpload(e) {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    console.log("Global attachments:", files.map(f => f.name));
    files.forEach(file => {
        readFileAsBase64(file).then(base64 => {
            appState.globalAttachments.set(file.name, { name: file.name, contentBytes: base64, contentType: file.type || "application/octet-stream" });
            renderGlobalAttachmentList();
        }).catch(err => console.error("Read error:", file.name, err));
    });
    e.target.value = "";
}
function renderGlobalAttachmentList() {
    const container = document.getElementById("globalAttachmentList");
    container.innerHTML = "";
    for (const [name, att] of appState.globalAttachments) {
        const sizeKB = Math.round(att.contentBytes.length * 3 / 4 / 1024);
        const div = document.createElement("div");
        div.className = "attachment-item";
        div.innerHTML = '<span class="att-name" title="' + escapeHtml(name) + '">' + escapeHtml(name) + '</span>' +
            '<span class="att-size">' + sizeKB + ' KB</span>' +
            '<button class="att-remove" title="Remove">&times;</button>';
        div.querySelector(".att-remove").addEventListener("click", () => { appState.globalAttachments.delete(name); renderGlobalAttachmentList(); });
        container.appendChild(div);
    }
}
function handlePerRecipientAttachmentUpload(e) {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    console.log("Per-recipient files:", files.map(f => f.name));
    files.forEach(file => {
        readFileAsBase64(file).then(base64 => {
            appState.perRecipientFiles.set(file.name.toLowerCase(), { name: file.name, contentBytes: base64, contentType: file.type || "application/octet-stream" });
            renderPerRecipientAttachmentList();
            checkMissingAttachments();
        }).catch(err => console.error("Read error:", file.name, err));
    });
    e.target.value = "";
}
function renderPerRecipientAttachmentList() {
    const container = document.getElementById("perRecipientAttachmentList");
    container.innerHTML = "";
    for (const [key, att] of appState.perRecipientFiles) {
        const sizeKB = Math.round(att.contentBytes.length * 3 / 4 / 1024);
        const div = document.createElement("div");
        div.className = "attachment-item";
        div.innerHTML = '<span class="att-name" title="' + escapeHtml(att.name) + '">' + escapeHtml(att.name) + '</span>' +
            '<span class="att-size">' + sizeKB + ' KB</span>' +
            '<button class="att-remove" title="Remove">&times;</button>';
        div.querySelector(".att-remove").addEventListener("click", () => {
            appState.perRecipientFiles.delete(key); renderPerRecipientAttachmentList(); checkMissingAttachments();
        });
        container.appendChild(div);
    }
}
function checkMissingAttachments() {
    if (!appState.mapping.attachments) return;
    const allNeeded = new Set();
    appState.rows.forEach(row => {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) val.split(";").forEach(f => { const n = f.trim(); if (n) allNeeded.add(n); });
    });
    const missing = [];
    for (const name of allNeeded) { if (!appState.perRecipientFiles.has(name.toLowerCase())) missing.push(name); }
    const w = document.getElementById("missingAttachmentWarning");
    if (missing.length > 0) {
        w.style.display = "block"; w.style.borderColor = "#f9a825"; w.style.background = "var(--warning-light)";
        w.innerHTML = "\u26A0\uFE0F Missing " + missing.length + " file(s): <strong>" + missing.map(escapeHtml).join(", ") + "</strong>";
    } else if (allNeeded.size > 0) {
        w.style.display = "block"; w.style.borderColor = "var(--success)"; w.style.background = "var(--success-light)";
        w.innerHTML = "\u2705 All " + allNeeded.size + " referenced file(s) uploaded.";
    } else { w.style.display = "none"; }
}
// ========== Step 4: Review & Preview ==========
function buildReview() {
    saveMapping();
    const subject = document.getElementById("emailSubject").value || "(no subject)";
    const globalCC = document.getElementById("globalCC").value;
    const globalBCC = document.getElementById("globalBCC").value;
    const fromAlias = document.getElementById("fromAlias").value.trim();
    const sharedMb = document.getElementById("sharedMailbox").value.trim();
    const isGrouped = document.getElementById("chkGroupByEmail").checked;
    const recipientCount = isGrouped ? getGroupedRecipients().size : appState.rows.length;

    let html = "<p><strong>Recipients:</strong> " + recipientCount;
    if (isGrouped) html += " (grouped from " + appState.rows.length + " rows)";
    html += "</p>";
    html += "<p><strong>To column:</strong> " + escapeHtml(appState.mapping.to) + "</p>";
    if (fromAlias) html += "<p><strong>From alias:</strong> " + escapeHtml(fromAlias) + "</p>";
    if (sharedMb) html += "<p><strong>Shared mailbox:</strong> " + escapeHtml(sharedMb) + "</p>";
    if (appState.mapping.cc) html += "<p><strong>CC column:</strong> " + escapeHtml(appState.mapping.cc) + "</p>";
    if (appState.mapping.bcc) html += "<p><strong>BCC column:</strong> " + escapeHtml(appState.mapping.bcc) + "</p>";
    if (globalCC) html += "<p><strong>Global CC:</strong> " + escapeHtml(globalCC) + "</p>";
    if (globalBCC) html += "<p><strong>Global BCC:</strong> " + escapeHtml(globalBCC) + "</p>";
    html += "<p><strong>Subject:</strong> " + escapeHtml(subject) + "</p>";
    if (appState.globalAttachments.size > 0) html += "<p><strong>Global attachments:</strong> " + appState.globalAttachments.size + " file(s)</p>";
    if (appState.mapping.attachments) html += "<p><strong>Per-recipient attachment column:</strong> " + escapeHtml(appState.mapping.attachments) + "</p>";
    document.getElementById("reviewSummary").innerHTML = html;
}

function changePreview(delta) {
    const isGrouped = document.getElementById("chkGroupByEmail").checked;
    const total = isGrouped ? getGroupedRecipients().size : appState.rows.length;
    appState.previewIndex = Math.max(0, Math.min(total - 1, appState.previewIndex + delta));
    buildPreview();
}

function buildPreview() {
    const isGrouped = document.getElementById("chkGroupByEmail").checked;
    const total = isGrouped ? getGroupedRecipients().size : appState.rows.length;
    if (total === 0) { document.getElementById("previewEmail").innerHTML = "<p>No recipients.</p>"; return; }
    const idx = Math.min(appState.previewIndex, total - 1);
    document.getElementById("previewIndex").textContent = (idx + 1) + " / " + total;
    document.getElementById("btnPrevPreview").disabled = idx <= 0;
    document.getElementById("btnNextPreview").disabled = idx >= total - 1;

    const subject = document.getElementById("emailSubject").value || "(no subject)";
    const bodyHtml = getEditorContent();
    const fromAlias = document.getElementById("fromAlias").value.trim();
    const sendAsHtml = document.getElementById("chkSendAsHtml").checked;
    let row, mergedSubj, mergedBody;

    if (isGrouped) {
        const groups = getGroupedRecipients();
        const keys = Array.from(groups.keys());
        const groupRows = groups.get(keys[idx]);
        row = groupRows[0];
        mergedSubj = appState.mapping.subject
            ? mergeFieldsWithGroup(String(row[appState.mapping.subject] || subject), groupRows)
            : mergeFieldsWithGroup(subject, groupRows);
        mergedBody = mergeFieldsWithGroup(bodyHtml, groupRows);
    } else {
        row = appState.rows[idx];
        mergedSubj = appState.mapping.subject
            ? mergeFields(String(row[appState.mapping.subject] || subject), row)
            : mergeFields(subject, row);
        mergedBody = mergeFields(bodyHtml, row);
    }

    let p = '<p class="label">Preview (recipient ' + (idx+1) + '):</p>';
    p += "<p><strong>To:</strong> " + escapeHtml(String(row[appState.mapping.to] || "")) + "</p>";
    if (fromAlias) p += "<p><strong>From:</strong> " + escapeHtml(fromAlias) + "</p>";
    p += "<p><strong>Subject:</strong> " + escapeHtml(mergedSubj) + "</p><hr/>";
    if (sendAsHtml) { p += "<div>" + mergedBody + "</div>"; }
    else { p += "<div>" + escapeHtml(mergedBody).replace(/\n/g, "<br/>") + "</div>"; }
    const attNames = getAttachmentNamesForRow(row);
    if (attNames.length > 0) p += '<hr/><p><strong>\u{1F4CE} Attachments:</strong> ' + attNames.map(escapeHtml).join(", ") + "</p>";
    document.getElementById("previewEmail").innerHTML = p;
}

function getAttachmentNamesForRow(row) {
    const names = [];
    for (const [name] of appState.globalAttachments) names.push(name);
    if (appState.mapping.attachments) {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) val.split(";").forEach(f => { const n = f.trim(); if (n) names.push(n); });
    }
    return names;
}

// ========== Merge Engine ==========
function mergeFields(template, row) {
    if (!template) return "";
    let result = template;
    for (const key of appState.headers) {
        const regex = new RegExp("\\{" + escapeRegExp(key) + "\\}", "g");
        let value = String(row[key] === undefined || row[key] === null ? "" : row[key]);
        if (!value && appState.defaults[key]) value = appState.defaults[key];
        result = result.replace(regex, value);
    }
    return result;
}

function getGroupedRecipients() {
    const groups = new Map();
    for (const row of appState.rows) {
        const email = String(row[appState.mapping.to] || "").trim().toLowerCase();
        if (!email) continue;
        if (!groups.has(email)) groups.set(email, []);
        groups.get(email).push(row);
    }
    return groups;
}

function mergeFieldsWithGroup(template, rows) {
    if (!template) return "";
    let result = template;
    // Handle {#rows}...{/rows} repeating sections
    result = result.replace(/\{#rows\}([\s\S]*?)\{\/rows\}/g, function(match, inner) {
        return rows.map(function(row) { return mergeFields(inner, row); }).join("");
    });
    // Replace remaining fields with first row values
    result = mergeFields(result, rows[0]);
    return result;
}

// ========== Graph API Helpers ==========
async function graphFetch(url, token, method, body) {
    const options = {
        method: method || "GET",
        headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" }
    };
    if (body) options.body = JSON.stringify(body);
    console.log("Graph " + method + " " + url);
    const response = await fetch(url, options);
    if (!response.ok) {
        let errMsg = "HTTP " + response.status + " " + response.statusText;
        try { const eb = await response.json(); if (eb.error && eb.error.message) errMsg += ": " + eb.error.message; } catch(_){}
        if (response.status === 401) throw new Error("SESSION_EXPIRED:" + errMsg);
        if (response.status === 403) throw new Error("Permission denied. Ask admin to grant required permissions.");
        if (response.status === 429) throw new Error("THROTTLED:" + errMsg);
        throw new Error(errMsg);
    }
    const ct = response.headers.get("content-type");
    if (ct && ct.includes("application/json")) return response.json();
    return null;
}

// ========== Email Building ==========
function buildRecipientList(addressStr) {
    if (!addressStr) return [];
    return addressStr.split(/[;,]/).map(e => e.trim()).filter(Boolean).map(email => ({ emailAddress: { address: email } }));
}

function buildGraphMessage(to, cc, bcc, subject, body, asHtml, fromAlias, opts) {
    const message = {
        subject: subject,
        body: { contentType: asHtml ? "HTML" : "Text", content: body },
        toRecipients: buildRecipientList(to)
    };
    const ccR = buildRecipientList(cc);
    if (ccR.length > 0) message.ccRecipients = ccR;
    const bccR = buildRecipientList(bcc);
    if (bccR.length > 0) message.bccRecipients = bccR;
    if (fromAlias) message.from = { emailAddress: { address: fromAlias } };
    if (opts.readReceipt) message.isReadReceiptRequested = true;
    if (opts.highImportance) message.importance = "high";
    if (opts.unsubscribeEmail) {
        message.internetMessageHeaders = [
            { name: "List-Unsubscribe", value: "<mailto:" + opts.unsubscribeEmail + ">" }
        ];
        if (asHtml) {
            message.body.content += '<br/><hr style="border:none;border-top:1px solid #ddd;margin:16px 0 8px"/>' +
                '<p style="font-size:11px;color:#888;">To unsubscribe, email <a href="mailto:' +
                escapeHtml(opts.unsubscribeEmail) + '">' + escapeHtml(opts.unsubscribeEmail) + '</a></p>';
        } else {
            message.body.content += "\n\n---\nTo unsubscribe, email: " + opts.unsubscribeEmail;
        }
    }
    return message;
}

function buildGraphAttachment(att) {
    return { "@odata.type": "#microsoft.graph.fileAttachment", name: att.name, contentType: att.contentType, contentBytes: att.contentBytes };
}

function collectAttachmentsForRow(row) {
    const attachments = [];
    for (const [, att] of appState.globalAttachments) attachments.push(att);
    if (appState.mapping.attachments) {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) val.split(";").forEach(f => {
            const name = f.trim();
            if (!name) return;
            const att = appState.perRecipientFiles.get(name.toLowerCase());
            if (att) attachments.push(att);
            else console.warn("Per-recipient attachment not found:", name);
        });
    }
    return attachments;
}

// ========== Send One Email ==========
async function sendOneEmail(token, to, cc, bcc, subject, body, asHtml, fromAlias, attachments, draftOnly, opts) {
    const sharedMb = opts.sharedMailbox || "";
    const baseUrl = sharedMb ? GRAPH_BASE + "/users/" + encodeURIComponent(sharedMb) : GRAPH_BASE + "/me";
    const message = buildGraphMessage(to, cc, bcc, subject, body, asHtml, fromAlias, opts);

    if (attachments.length === 0 && !draftOnly) {
        await graphFetch(baseUrl + "/sendMail", token, "POST", { message: message, saveToSentItems: true });
        return;
    }
    // Create draft -> add attachments -> send
    const draft = await graphFetch(baseUrl + "/messages", token, "POST", message);
    const msgId = draft.id;
    console.log("Created draft:", msgId);
    for (const att of attachments) {
        await graphFetch(baseUrl + "/messages/" + encodeURIComponent(msgId) + "/attachments", token, "POST", buildGraphAttachment(att));
        console.log("Added attachment:", att.name);
    }
    if (!draftOnly) {
        await graphFetch(baseUrl + "/messages/" + encodeURIComponent(msgId) + "/send", token, "POST", null);
        console.log("Sent:", msgId);
    } else {
        console.log("Draft saved:", msgId);
    }
}

// ========== Execute Mail Merge ==========
async function executeMerge(testMode) {
    if (appState.sending) return;
    try {
        saveMapping();
        saveDefaults();
        const subject = document.getElementById("emailSubject").value;
        const bodyContent = getEditorContent();
        const plainBody = getEditorPlainText();
        const globalCC = document.getElementById("globalCC").value.trim();
        const globalBCC = document.getElementById("globalBCC").value.trim();
        const fromAlias = document.getElementById("fromAlias").value.trim();
        const sendAsHtml = document.getElementById("chkSendAsHtml").checked;
        const draftOnly = document.getElementById("chkDraftOnly").checked;
        const readReceipt = document.getElementById("chkReadReceipt").checked;
        const highImportance = document.getElementById("chkHighImportance").checked;
        const groupByEmail = document.getElementById("chkGroupByEmail").checked;
        const unsubscribe = document.getElementById("chkUnsubscribe").checked;
        const unsubscribeEmail = unsubscribe ? document.getElementById("unsubscribeEmail").value.trim() : "";
        const sharedMailbox = document.getElementById("sharedMailbox").value.trim();
        let delay = parseInt(document.getElementById("sendDelay").value) || 1;

        const opts = { readReceipt: readReceipt, highImportance: highImportance, unsubscribeEmail: unsubscribeEmail, sharedMailbox: sharedMailbox };

        const body = sendAsHtml ? bodyContent : plainBody;
        if (!subject && !appState.mapping.subject) { showStatus("\u26A0\uFE0F Enter a subject line.", "warning"); return; }
        if (!body) { showStatus("\u26A0\uFE0F Enter an email body.", "warning"); return; }

        // Build send list
        let sendItems = [];
        if (testMode) {
            sendItems = [{ rows: [appState.rows[0]], to: null }]; // to set later
        } else if (groupByEmail) {
            const groups = getGroupedRecipients();
            for (const [email, rows] of groups) sendItems.push({ rows: rows, to: email });
        } else {
            appState.rows.forEach(row => {
                sendItems.push({ rows: [row], to: String(row[appState.mapping.to] || "").trim() });
            });
        }
        const total = sendItems.length;
        if (total === 0) { showStatus("\u26A0\uFE0F No recipients.", "warning"); return; }

        // Lock UI
        appState.sending = true;
        setButtonsDisabled(true);
        const sendBtn = document.getElementById("btnSend");
        sendBtn.classList.add("sending");
        document.getElementById("progressContainer").style.display = "block";
        hideResults();

        const modeLabel = testMode ? "Test" : (draftOnly ? "Drafting" : "Sending");
        updateProgress(0, total, modeLabel + ": authenticating...");

        let token;
        try {
            token = await getGraphToken();
            console.log("Graph token acquired");
        } catch (tokenErr) {
            showStatus("\u274C Authentication failed: " + escapeHtml(tokenErr.message || String(tokenErr)), "error");
            appState.sending = false; setButtonsDisabled(false); sendBtn.classList.remove("sending");
            document.getElementById("progressContainer").style.display = "none";
            return;
        }

        if (testMode) {
            let testTo = appState.userEmail;
            if (!testTo) {
                try {
                    const profile = await graphFetch(GRAPH_BASE + "/me", token, "GET");
                    testTo = profile.mail || profile.userPrincipalName;
                    appState.userEmail = testTo;
                } catch (e) {
                    showStatus("\u274C Could not determine your email: " + escapeHtml(e.message), "error");
                    appState.sending = false; setButtonsDisabled(false); sendBtn.classList.remove("sending");
                    document.getElementById("progressContainer").style.display = "none";
                    return;
                }
            }
            updateProgress(0, 1, modeLabel + ": sending to " + testTo + "...");
            const row = appState.rows[0];
            const mSubj = appState.mapping.subject ? mergeFields(String(row[appState.mapping.subject] || subject), row) : mergeFields(subject, row);
            const mBody = sendAsHtml ? mergeFields(bodyContent, row) : mergeFields(plainBody, row);
            let ccL = ""; if (appState.mapping.cc && row[appState.mapping.cc]) ccL = String(row[appState.mapping.cc]);
            if (globalCC) ccL = ccL ? ccL + ";" + globalCC : globalCC;
            let bccL = ""; if (appState.mapping.bcc && row[appState.mapping.bcc]) bccL = String(row[appState.mapping.bcc]);
            if (globalBCC) bccL = bccL ? bccL + ";" + globalBCC : globalBCC;
            const atts = collectAttachmentsForRow(row);
            try {
                await sendOneEmail(token, testTo, "", "", "[TEST] " + mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                updateProgress(1, 1, "Done!");
                showStatus("\u2705 Test email " + (draftOnly ? "drafted" : "sent") + " to " + escapeHtml(testTo), "info");
            } catch (err) {
                showStatus("\u274C Test failed: " + escapeHtml(err.message), "error");
            }
            appState.sending = false; setButtonsDisabled(false); sendBtn.classList.remove("sending");
            document.getElementById("progressContainer").style.display = "none";
            return;
        }

        // ===== Bulk send =====
        appState.results = [];
        let sent = 0, errors = 0;
        for (let i = 0; i < total; i++) {
            const item = sendItems[i];
            const toAddr = item.to;
            if (!toAddr) {
                errors++;
                appState.results.push({ row: i + 2, to: "(empty)", status: "Error", error: "No email address" });
                continue;
            }
            const row = item.rows[0];
            const isGroup = item.rows.length > 1;
            const mSubj = appState.mapping.subject
                ? (isGroup ? mergeFieldsWithGroup(String(row[appState.mapping.subject] || subject), item.rows) : mergeFields(String(row[appState.mapping.subject] || subject), row))
                : (isGroup ? mergeFieldsWithGroup(subject, item.rows) : mergeFields(subject, row));
            const mBody = isGroup
                ? (sendAsHtml ? mergeFieldsWithGroup(bodyContent, item.rows) : mergeFieldsWithGroup(plainBody, item.rows))
                : (sendAsHtml ? mergeFields(bodyContent, row) : mergeFields(plainBody, row));
            let ccL = ""; if (appState.mapping.cc && row[appState.mapping.cc]) ccL = String(row[appState.mapping.cc]);
            if (globalCC) ccL = ccL ? ccL + ";" + globalCC : globalCC;
            let bccL = ""; if (appState.mapping.bcc && row[appState.mapping.bcc]) bccL = String(row[appState.mapping.bcc]);
            if (globalBCC) bccL = bccL ? bccL + ";" + globalBCC : globalBCC;
            const atts = collectAttachmentsForRow(row);
            updateProgress(i, total, modeLabel + " " + (i+1) + " of " + total + " \u2014 " + escapeHtml(toAddr));
            try {
                await sendOneEmail(token, toAddr, ccL, bccL, mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                sent++;
                appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent" });
            } catch (err) {
                const errMsg = err.message || String(err);
                if (errMsg.startsWith("THROTTLED:")) {
                    const old = delay; delay = Math.max(delay * 2, 5);
                    console.warn("Rate limited at " + (i+2) + ". Delay: " + old + "s -> " + delay + "s");
                    updateProgress(i, total, "Rate limited. Delay -> " + delay + "s. Retrying...");
                    document.getElementById("sendDelay").value = delay;
                    await sleep(delay * 1000);
                    try {
                        await sendOneEmail(token, toAddr, ccL, bccL, mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                        sent++;
                        appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent" });
                        continue;
                    } catch (retryErr) {
                        errors++;
                        appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: retryErr.message || String(retryErr) });
                    }
                } else if (errMsg.startsWith("SESSION_EXPIRED:")) {
                    console.warn("Token expired, re-acquiring...");
                    try {
                        token = await getGraphToken();
                        await sendOneEmail(token, toAddr, ccL, bccL, mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                        sent++;
                        appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent" });
                        continue;
                    } catch (retryErr) {
                        errors++;
                        appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: retryErr.message || String(retryErr) });
                    }
                } else {
                    errors++;
                    appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: errMsg });
                    console.error("Error sending to " + toAddr + ":", err);
                }
            }
            if (i < total - 1 && delay > 0) await sleep(delay * 1000);
        }

        updateProgress(total, total, "Complete!");
        document.getElementById("progressContainer").style.display = "none";
        sendBtn.classList.remove("sending");
        sendBtn.classList.add("complete");
        setTimeout(() => sendBtn.classList.remove("complete"), 3000);
        showResultsTable(total, sent, errors, draftOnly);
        // Save campaign
        const campName = document.getElementById("campaignName").value.trim();
        saveCampaign(campName, total, sent, errors);

    } catch (outerErr) {
        console.error("executeMerge error:", outerErr);
        showStatus("\u274C Mail merge error: " + escapeHtml(outerErr.message || String(outerErr)), "error");
    } finally {
        appState.sending = false;
        setButtonsDisabled(false);
        document.getElementById("btnSend").classList.remove("sending");
    }
}

// ========== UI Helpers ==========
function setButtonsDisabled(disabled) {
    document.getElementById("btnSend").disabled = disabled;
    document.getElementById("btnTestEmail").disabled = disabled;
    document.getElementById("btnStep4Back").disabled = disabled;
}

function updateProgress(current, total, text) {
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;
    document.getElementById("progressFill").style.width = pct + "%";
    document.getElementById("progressText").textContent = text;
}

function showStatus(message, type) {
    const el = document.getElementById("resultsContainer");
    el.style.display = "block";
    el.className = "results " + (type || "info");
    el.innerHTML = "<p>" + message + "</p>";
}

function hideResults() {
    document.getElementById("resultsContainer").style.display = "none";
}

function showResultsTable(total, sent, errors, draftOnly) {
    const el = document.getElementById("resultsContainer");
    el.style.display = "block";
    el.className = errors > 0 ? "results error" : "results success";
    let html = "<h3>" + (errors === 0 ? "\u2705" : "\u26A0\uFE0F") + " Mail Merge Complete</h3>";
    html += "<p><strong>Total:</strong> " + total + "</p>";
    if (sent > 0) html += "<p><strong>" + (draftOnly ? "Drafts" : "Sent") + ":</strong> " + sent + "</p>";
    if (errors > 0) html += "<p><strong>Errors:</strong> " + errors + "</p>";
    html += '<table class="results-table"><tr><th>Row</th><th>Email</th><th>Status</th></tr>';
    for (const r of appState.results) {
        const cls = r.status === "Error" ? "status-err" : "status-ok";
        const statusTxt = r.status === "Error" ? escapeHtml(r.status + ": " + (r.error || "")) : escapeHtml(r.status);
        html += "<tr><td>" + r.row + "</td><td>" + escapeHtml(r.to) + "</td><td class=\"" + cls + "\">" + statusTxt + "</td></tr>";
    }
    html += "</table>";
    html += '<button class="btn btn-secondary btn-small btn-export mt-8" onclick="exportResultsCsv()">\u{1F4E5} Download Results CSV</button>';
    el.innerHTML = html;
}

// ========== Export CSV ==========
function exportResultsCsv() {
    let csv = "Row,Email,Status,Error\n";
    for (const r of appState.results) {
        csv += r.row + ',"' + (r.to || "").replace(/"/g, '""') + '","' + (r.status || "") + '","' + (r.error || "").replace(/"/g, '""') + '"\n';
    }
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "mailmerge-results.csv";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}

// ========== Utility Functions ==========
function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = String(str);
    return div.innerHTML;
}
function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function() { resolve(reader.result.split(",")[1]); };
        reader.onerror = function() { reject(new Error("Failed to read: " + file.name)); };
        reader.readAsDataURL(file);
    });
}