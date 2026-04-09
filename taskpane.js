/* MailMerge-Pro — Outlook Web Add-in JavaScript
 * Uses Outlook REST API v2.0 via getCallbackTokenAsync for bulk sending.
 * NO alert/confirm/prompt — all feedback via DOM.
 */

"use strict";

// ========== Application State ==========

const appState = {
    currentStep: 1,
    headers: [],
    rows: [],
    mapping: { to: "", cc: "", bcc: "", subject: "", attachments: "" },
    globalAttachments: new Map(),   // filename -> { name, contentBytes (base64), contentType }
    perRecipientFiles: new Map(),   // filename (lowercase) -> { name, contentBytes, contentType }
    results: [],
    sending: false
};

// ========== Office.js Initialization ==========

Office.onReady(function (info) {
    console.log("Office.onReady:", info.host, info.platform);
    initUI();
});

if (typeof Office === "undefined") {
    document.addEventListener("DOMContentLoaded", initUI);
}

// ========== UI Initialization ==========

function initUI() {
    console.log("initUI: binding event listeners");

    // Step 1: File input
    document.getElementById("fileInput").addEventListener("change", handleFileUpload);

    // Navigation
    document.getElementById("btnStep1Next").addEventListener("click", () => goToStep(2));
    document.getElementById("btnStep2Back").addEventListener("click", () => goToStep(1));
    document.getElementById("btnStep2Next").addEventListener("click", () => goToStep(3));
    document.getElementById("btnStep3Back").addEventListener("click", () => goToStep(2));
    document.getElementById("btnStep3Next").addEventListener("click", () => goToStep(4));
    document.getElementById("btnStep4Back").addEventListener("click", () => goToStep(3));

    // Send / Test
    document.getElementById("btnSend").addEventListener("click", () => executeMerge(false));
    document.getElementById("btnTestEmail").addEventListener("click", () => executeMerge(true));

    // Global attachment upload
    document.getElementById("globalAttachmentInput").addEventListener("change", handleGlobalAttachmentUpload);

    // Per-recipient attachment upload
    document.getElementById("perRecipientAttachmentInput").addEventListener("change", handlePerRecipientAttachmentUpload);
}

// ========== Step Navigation ==========

function goToStep(step) {
    if (appState.sending) return;

    // Save current step data before leaving
    if (appState.currentStep === 2) saveMapping();
    if (appState.currentStep === 3) saveCompose();

    // Validate before advancing
    if (step > appState.currentStep) {
        if (appState.currentStep === 1 && appState.rows.length === 0) {
            showStatus("⚠️ Please upload a data file first.", "error");
            return;
        }
        if (appState.currentStep === 2 && !document.getElementById("mapTo").value) {
            showStatus("⚠️ Please select the To (Email) column.", "error");
            return;
        }
    }

    if (step === 2) populateColumnDropdowns();
    if (step === 3) {
        populateMergeFieldButtons();
        updatePerRecipientAttachmentVisibility();
    }
    if (step === 4) buildReview();

    // Toggle visibility
    document.querySelectorAll(".step-content").forEach(el => el.classList.remove("active"));
    document.getElementById("step" + step).classList.add("active");

    // Update step indicators
    document.querySelectorAll(".steps-indicator .step").forEach(el => {
        const s = parseInt(el.dataset.step);
        el.classList.remove("active", "done");
        if (s === step) el.classList.add("active");
        else if (s < step) el.classList.add("done");
    });

    appState.currentStep = step;
    hideResults();
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
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            if (jsonData.length === 0) {
                showStatus("⚠️ No data rows found in the file.", "error");
                return;
            }

            appState.headers = Object.keys(jsonData[0]);
            appState.rows = jsonData;

            document.getElementById("dataStats").textContent =
                `✅ ${appState.rows.length} recipients, ${appState.headers.length} columns: ${appState.headers.join(", ")}`;

            let tableHtml = "<table><tr>";
            appState.headers.forEach(h => { tableHtml += `<th>${escapeHtml(h)}</th>`; });
            tableHtml += "</tr>";
            const previewRows = appState.rows.slice(0, 5);
            previewRows.forEach(row => {
                tableHtml += "<tr>";
                appState.headers.forEach(h => { tableHtml += `<td>${escapeHtml(String(row[h] || ""))}</td>`; });
                tableHtml += "</tr>";
            });
            if (appState.rows.length > 5) {
                tableHtml += `<tr><td colspan="${appState.headers.length}" style="text-align:center;color:#888;">... and ${appState.rows.length - 5} more rows</td></tr>`;
            }
            tableHtml += "</table>";
            document.getElementById("previewTable").innerHTML = tableHtml;
            document.getElementById("dataPreview").style.display = "block";
            document.getElementById("btnStep1Next").disabled = false;

            console.log("Parsed", appState.rows.length, "rows,", appState.headers.length, "columns");
        } catch (err) {
            console.error("File parse error:", err);
            showStatus("❌ Error reading file: " + err.message, "error");
        }
    };
    reader.readAsArrayBuffer(file);
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
        sel.innerHTML = '<option value="">(none)</option>';
        appState.headers.forEach(h => {
            const opt = document.createElement("option");
            opt.value = h;
            opt.textContent = h;
            sel.appendChild(opt);
        });

        const matches = autoMatch[id] || [];
        for (const header of appState.headers) {
            if (matches.some(m => header.toLowerCase().includes(m))) {
                sel.value = header;
                break;
            }
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
        btn.textContent = `{${h}}`;
        btn.type = "button";
        btn.addEventListener("click", () => {
            const textarea = document.getElementById("emailBody");
            const pos = textarea.selectionStart;
            const text = textarea.value;
            textarea.value = text.slice(0, pos) + `{${h}}` + text.slice(pos);
            textarea.focus();
            textarea.selectionStart = textarea.selectionEnd = pos + h.length + 2;
        });
        container.appendChild(btn);
    });
}

function updatePerRecipientAttachmentVisibility() {
    saveMapping();
    const section = document.getElementById("perRecipientAttachmentSection");
    if (appState.mapping.attachments) {
        section.style.display = "block";
        checkMissingAttachments();
    } else {
        section.style.display = "none";
    }
}

function saveCompose() {
    // Values are read directly from DOM at send time
}

// ========== Attachment Handling ==========

function handleGlobalAttachmentUpload(e) {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    console.log("Global attachments selected:", files.map(f => f.name));

    files.forEach(file => {
        readFileAsBase64(file).then(base64 => {
            appState.globalAttachments.set(file.name, {
                name: file.name,
                contentBytes: base64,
                contentType: file.type || "application/octet-stream"
            });
            renderGlobalAttachmentList();
        }).catch(err => {
            console.error("Error reading attachment:", file.name, err);
        });
    });

    // Reset input so same file can be re-selected
    e.target.value = "";
}

function renderGlobalAttachmentList() {
    const container = document.getElementById("globalAttachmentList");
    container.innerHTML = "";
    for (const [name, att] of appState.globalAttachments) {
        const sizeKB = Math.round(att.contentBytes.length * 3 / 4 / 1024);
        const div = document.createElement("div");
        div.className = "attachment-item";
        div.innerHTML = `<span class="att-name" title="${escapeHtml(name)}">${escapeHtml(name)}</span>
            <span class="att-size">${sizeKB} KB</span>
            <button class="att-remove" data-name="${escapeHtml(name)}" title="Remove">&times;</button>`;
        div.querySelector(".att-remove").addEventListener("click", () => {
            appState.globalAttachments.delete(name);
            renderGlobalAttachmentList();
        });
        container.appendChild(div);
    }
}

function handlePerRecipientAttachmentUpload(e) {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    console.log("Per-recipient files selected:", files.map(f => f.name));

    files.forEach(file => {
        readFileAsBase64(file).then(base64 => {
            appState.perRecipientFiles.set(file.name.toLowerCase(), {
                name: file.name,
                contentBytes: base64,
                contentType: file.type || "application/octet-stream"
            });
            renderPerRecipientAttachmentList();
            checkMissingAttachments();
        }).catch(err => {
            console.error("Error reading per-recipient file:", file.name, err);
        });
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
        div.innerHTML = `<span class="att-name" title="${escapeHtml(att.name)}">${escapeHtml(att.name)}</span>
            <span class="att-size">${sizeKB} KB</span>
            <button class="att-remove" title="Remove">&times;</button>`;
        div.querySelector(".att-remove").addEventListener("click", () => {
            appState.perRecipientFiles.delete(key);
            renderPerRecipientAttachmentList();
            checkMissingAttachments();
        });
        container.appendChild(div);
    }
}

function checkMissingAttachments() {
    if (!appState.mapping.attachments) return;

    const allNeeded = new Set();
    appState.rows.forEach(row => {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) {
            val.split(";").forEach(f => {
                const name = f.trim();
                if (name) allNeeded.add(name);
            });
        }
    });

    const missing = [];
    for (const name of allNeeded) {
        if (!appState.perRecipientFiles.has(name.toLowerCase())) {
            missing.push(name);
        }
    }

    const warningEl = document.getElementById("missingAttachmentWarning");
    if (missing.length > 0) {
        warningEl.style.display = "block";
        warningEl.innerHTML = `⚠️ Missing ${missing.length} file(s): <strong>${missing.map(escapeHtml).join(", ")}</strong>`;
    } else if (allNeeded.size > 0) {
        warningEl.style.display = "block";
        warningEl.innerHTML = `✅ All ${allNeeded.size} referenced file(s) uploaded.`;
        warningEl.style.borderColor = "#107c10";
        warningEl.style.background = "#dff6dd";
    } else {
        warningEl.style.display = "none";
    }
}

// ========== Step 4: Review ==========

function buildReview() {
    saveMapping();

    const subject = document.getElementById("emailSubject").value || "(no subject)";
    const globalCC = document.getElementById("globalCC").value;
    const globalBCC = document.getElementById("globalBCC").value;
    const fromAlias = document.getElementById("fromAlias").value.trim();

    let html = `<p><strong>Recipients:</strong> ${appState.rows.length}</p>`;
    html += `<p><strong>To column:</strong> ${escapeHtml(appState.mapping.to)}</p>`;
    if (fromAlias) html += `<p><strong>From alias:</strong> ${escapeHtml(fromAlias)}</p>`;
    if (appState.mapping.cc) html += `<p><strong>CC column:</strong> ${escapeHtml(appState.mapping.cc)}</p>`;
    if (appState.mapping.bcc) html += `<p><strong>BCC column:</strong> ${escapeHtml(appState.mapping.bcc)}</p>`;
    if (globalCC) html += `<p><strong>Global CC:</strong> ${escapeHtml(globalCC)}</p>`;
    if (globalBCC) html += `<p><strong>Global BCC:</strong> ${escapeHtml(globalBCC)}</p>`;
    html += `<p><strong>Subject:</strong> ${escapeHtml(subject)}</p>`;
    if (appState.globalAttachments.size > 0) {
        html += `<p><strong>Global attachments:</strong> ${appState.globalAttachments.size} file(s)</p>`;
    }
    if (appState.mapping.attachments) {
        html += `<p><strong>Per-recipient attachment column:</strong> ${escapeHtml(appState.mapping.attachments)}</p>`;
    }
    document.getElementById("reviewSummary").innerHTML = html;

    // Preview first email
    if (appState.rows.length > 0) {
        const row = appState.rows[0];
        const mergedSubj = appState.mapping.subject
            ? mergeFields(String(row[appState.mapping.subject] || subject), row)
            : mergeFields(subject, row);
        const mergedBody = mergeFields(document.getElementById("emailBody").value, row);

        let preview = `<p class="label">Preview (recipient 1):</p>`;
        preview += `<p><strong>To:</strong> ${escapeHtml(String(row[appState.mapping.to] || ""))}</p>`;
        if (fromAlias) preview += `<p><strong>From:</strong> ${escapeHtml(fromAlias)}</p>`;
        preview += `<p><strong>Subject:</strong> ${escapeHtml(mergedSubj)}</p>`;
        preview += `<hr/><div>${escapeHtml(mergedBody).replace(/\n/g, "<br/>")}</div>`;

        // Show attachments for first row
        const attachNames = getAttachmentNamesForRow(row);
        if (attachNames.length > 0) {
            preview += `<hr/><p><strong>Attachments:</strong> ${attachNames.map(escapeHtml).join(", ")}</p>`;
        }

        document.getElementById("previewEmail").innerHTML = preview;
    }
}

function getAttachmentNamesForRow(row) {
    const names = [];
    // Global attachments
    for (const [name] of appState.globalAttachments) {
        names.push(name);
    }
    // Per-recipient attachments
    if (appState.mapping.attachments) {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) {
            val.split(";").forEach(f => {
                const n = f.trim();
                if (n) names.push(n);
            });
        }
    }
    return names;
}

// ========== Merge Engine ==========

function mergeFields(template, row) {
    let result = template;
    for (const key of appState.headers) {
        const regex = new RegExp(`\\{${escapeRegExp(key)}\\}`, "g");
        result = result.replace(regex, String(row[key] || ""));
    }
    return result;
}

// ========== REST API Helpers ==========

function getRestToken() {
    return new Promise((resolve, reject) => {
        if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.getCallbackTokenAsync) {
            reject(new Error("Office.context.mailbox.getCallbackTokenAsync is not available. Ensure Mailbox requirement set 1.5+ is supported."));
            return;
        }
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log("REST token obtained successfully");
                resolve(result.value);
            } else {
                console.error("Failed to get REST token:", result.error);
                reject(new Error("Failed to get REST token: " + (result.error.message || result.error.code)));
            }
        });
    });
}

function getRestUrl() {
    return Office.context.mailbox.restUrl;
}

async function restFetch(url, token, method, body) {
    const options = {
        method: method || "GET",
        headers: {
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json"
        }
    };
    if (body) {
        options.body = JSON.stringify(body);
    }

    console.log(`REST ${method} ${url}`);
    const response = await fetch(url, options);

    if (!response.ok) {
        let errMsg = `HTTP ${response.status} ${response.statusText}`;
        try {
            const errBody = await response.json();
            if (errBody.error && errBody.error.message) {
                errMsg += ": " + errBody.error.message;
            }
        } catch (_) { /* ignore parse error */ }
        throw new Error(errMsg);
    }

    // Some endpoints (like /send) return 202 with no body
    const contentType = response.headers.get("content-type");
    if (contentType && contentType.includes("application/json")) {
        return response.json();
    }
    return null;
}

// ========== Email Sending via REST API ==========

function buildRecipientList(addressStr) {
    if (!addressStr) return [];
    return addressStr.split(/[;,]/)
        .map(e => e.trim())
        .filter(Boolean)
        .map(email => ({ EmailAddress: { Address: email } }));
}

function buildMessagePayload(to, cc, bcc, subject, body, asHtml, fromAlias) {
    const message = {
        Subject: subject,
        Body: {
            ContentType: asHtml ? "HTML" : "Text",
            Content: asHtml ? body.replace(/\n/g, "<br/>") : body
        },
        ToRecipients: buildRecipientList(to)
    };

    const ccRecipients = buildRecipientList(cc);
    if (ccRecipients.length > 0) message.CcRecipients = ccRecipients;

    const bccRecipients = buildRecipientList(bcc);
    if (bccRecipients.length > 0) message.BccRecipients = bccRecipients;

    if (fromAlias) {
        message.From = { EmailAddress: { Address: fromAlias } };
    }

    return message;
}

function buildAttachmentPayload(att) {
    return {
        "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
        Name: att.name,
        ContentType: att.contentType,
        ContentBytes: att.contentBytes
    };
}

function collectAttachmentsForRow(row) {
    const attachments = [];

    // Global attachments
    for (const [, att] of appState.globalAttachments) {
        attachments.push(att);
    }

    // Per-recipient attachments
    if (appState.mapping.attachments) {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) {
            val.split(";").forEach(f => {
                const name = f.trim();
                if (!name) return;
                const att = appState.perRecipientFiles.get(name.toLowerCase());
                if (att) {
                    attachments.push(att);
                } else {
                    console.warn("Per-recipient attachment not found:", name);
                }
            });
        }
    }

    return attachments;
}

async function sendOneEmail(restUrl, token, to, cc, bcc, subject, body, asHtml, fromAlias, attachments, draftOnly) {
    const message = buildMessagePayload(to, cc, bcc, subject, body, asHtml, fromAlias);

    if (attachments.length === 0 && !draftOnly) {
        // Simple send — no attachments, no draft needed
        await restFetch(restUrl + "/v2.0/me/sendmail", token, "POST", {
            Message: message,
            SaveToSentItems: true
        });
        return;
    }

    // Create draft first (needed for attachments or draft-only mode)
    const draft = await restFetch(restUrl + "/v2.0/me/messages", token, "POST", message);
    const messageId = draft.Id;
    console.log("Created draft:", messageId);

    // Add attachments
    for (const att of attachments) {
        await restFetch(
            restUrl + "/v2.0/me/messages/" + encodeURIComponent(messageId) + "/attachments",
            token,
            "POST",
            buildAttachmentPayload(att)
        );
        console.log("Added attachment:", att.name);
    }

    if (!draftOnly) {
        // Send the draft
        await restFetch(
            restUrl + "/v2.0/me/messages/" + encodeURIComponent(messageId) + "/send",
            token,
            "POST",
            null
        );
        console.log("Sent message:", messageId);
    } else {
        console.log("Draft saved:", messageId);
    }
}

// ========== Execute Mail Merge ==========

async function executeMerge(testMode) {
    if (appState.sending) return;

    try {
        saveMapping();
        const subject = document.getElementById("emailSubject").value;
        const body = document.getElementById("emailBody").value;
        const globalCC = document.getElementById("globalCC").value.trim();
        const globalBCC = document.getElementById("globalBCC").value.trim();
        const fromAlias = document.getElementById("fromAlias").value.trim();
        const sendAsHtml = document.getElementById("chkSendAsHtml").checked;
        const draftOnly = document.getElementById("chkDraftOnly").checked;
        const delay = parseInt(document.getElementById("sendDelay").value) || 1;

        if (!subject && !appState.mapping.subject) {
            showStatus("⚠️ Please enter a subject line or map a Subject column.", "error");
            return;
        }
        if (!body) {
            showStatus("⚠️ Please enter an email body.", "error");
            return;
        }

        const rowsToSend = testMode ? [appState.rows[0]] : appState.rows;
        const total = rowsToSend.length;
        if (total === 0) {
            showStatus("⚠️ No recipients loaded.", "error");
            return;
        }

        // Lock UI
        appState.sending = true;
        setButtonsDisabled(true);
        document.getElementById("progressContainer").style.display = "block";
        hideResults();

        const modeLabel = testMode ? "Test" : (draftOnly ? "Drafting" : "Sending");
        updateProgress(0, total, `${modeLabel}: getting API token...`);

        // Get REST API token once
        let token, restUrl;
        try {
            token = await getRestToken();
            restUrl = getRestUrl();
            console.log("REST URL:", restUrl);
        } catch (tokenErr) {
            showStatus("❌ Could not get API token: " + tokenErr.message, "error");
            appState.sending = false;
            setButtonsDisabled(false);
            document.getElementById("progressContainer").style.display = "none";
            return;
        }

        if (testMode) {
            // For test mode, send the first row to the user's own mailbox
            const testTo = Office.context.mailbox.userProfile.emailAddress;
            updateProgress(0, 1, `${modeLabel}: sending to ${testTo}...`);

            const row = rowsToSend[0];
            const mergedSubject = appState.mapping.subject
                ? mergeFields(String(row[appState.mapping.subject] || subject), row)
                : mergeFields(subject, row);
            const mergedBody = mergeFields(body, row);

            // Build CC/BCC
            let ccList = "";
            if (appState.mapping.cc && row[appState.mapping.cc]) ccList = String(row[appState.mapping.cc]);
            if (globalCC) ccList = ccList ? ccList + ";" + globalCC : globalCC;

            let bccList = "";
            if (appState.mapping.bcc && row[appState.mapping.bcc]) bccList = String(row[appState.mapping.bcc]);
            if (globalBCC) bccList = bccList ? bccList + ";" + globalBCC : globalBCC;

            const attachments = collectAttachmentsForRow(row);
            const testSubject = "[TEST] " + mergedSubject;

            try {
                await sendOneEmail(restUrl, token, testTo, "", "", testSubject, mergedBody, sendAsHtml, fromAlias, attachments, draftOnly);
                updateProgress(1, 1, "Done!");
                showStatus(
                    `✅ Test email ${draftOnly ? "saved as draft" : "sent"} to ${escapeHtml(testTo)} — check your ${draftOnly ? "Drafts" : "inbox"}!`,
                    "info"
                );
            } catch (err) {
                showStatus("❌ Test email failed: " + escapeHtml(err.message), "error");
            }

            appState.sending = false;
            setButtonsDisabled(false);
            document.getElementById("progressContainer").style.display = "none";
            return;
        }

        // ===== Bulk send =====
        appState.results = [];
        let sent = 0, errors = 0;

        for (let i = 0; i < total; i++) {
            const row = rowsToSend[i];
            const toAddr = String(row[appState.mapping.to] || "").trim();

            if (!toAddr) {
                errors++;
                appState.results.push({ row: i + 2, to: "(empty)", status: "Error", error: "No email address" });
                continue;
            }

            const mergedSubject = appState.mapping.subject
                ? mergeFields(String(row[appState.mapping.subject] || subject), row)
                : mergeFields(subject, row);
            const mergedBody = mergeFields(body, row);

            // Build CC
            let ccList = "";
            if (appState.mapping.cc && row[appState.mapping.cc]) ccList = String(row[appState.mapping.cc]);
            if (globalCC) ccList = ccList ? ccList + ";" + globalCC : globalCC;

            // Build BCC
            let bccList = "";
            if (appState.mapping.bcc && row[appState.mapping.bcc]) bccList = String(row[appState.mapping.bcc]);
            if (globalBCC) bccList = bccList ? bccList + ";" + globalBCC : globalBCC;

            const attachments = collectAttachmentsForRow(row);

            updateProgress(i, total, `${modeLabel} ${i + 1} of ${total} — ${escapeHtml(toAddr)}`);

            try {
                await sendOneEmail(restUrl, token, toAddr, ccList, bccList, mergedSubject, mergedBody, sendAsHtml, fromAlias, attachments, draftOnly);
                sent++;
                appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent" });
            } catch (err) {
                errors++;
                appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: err.message || String(err) });
                console.error(`Error sending to ${toAddr}:`, err);
            }

            // Throttle between emails
            if (i < total - 1 && delay > 0) {
                await sleep(delay * 1000);
            }
        }

        // Show final results
        updateProgress(total, total, "Complete!");
        document.getElementById("progressContainer").style.display = "none";
        showResultsTable(total, sent, errors, draftOnly);

    } catch (outerErr) {
        console.error("executeMerge outer error:", outerErr);
        showStatus("❌ Mail merge error: " + escapeHtml(outerErr.message || String(outerErr)), "error");
    } finally {
        appState.sending = false;
        setButtonsDisabled(false);
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
    el.innerHTML = `<p>${message}</p>`;
}

function hideResults() {
    document.getElementById("resultsContainer").style.display = "none";
}

function showResultsTable(total, sent, errors, draftOnly) {
    const el = document.getElementById("resultsContainer");
    el.style.display = "block";
    el.className = errors > 0 ? "results error" : "results success";

    let html = `<h3>${errors === 0 ? "✅" : "⚠️"} Mail Merge Complete</h3>`;
    html += `<p><strong>Total:</strong> ${total}</p>`;
    if (sent > 0) html += `<p><strong>${draftOnly ? "Drafts created" : "Sent"}:</strong> ${sent}</p>`;
    if (errors > 0) html += `<p><strong>Errors:</strong> ${errors}</p>`;

    // Detailed results table
    html += `<table class="results-table"><tr><th>Row</th><th>To</th><th>Status</th></tr>`;
    for (const r of appState.results) {
        const statusClass = r.status === "Error" ? "status-err" : "status-ok";
        const statusText = r.status === "Error"
            ? escapeHtml(r.status + ": " + (r.error || ""))
            : escapeHtml(r.status);
        html += `<tr><td>${r.row}</td><td>${escapeHtml(r.to)}</td><td class="${statusClass}">${statusText}</td></tr>`;
    }
    html += `</table>`;

    el.innerHTML = html;
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
        reader.onload = function () {
            // result is "data:<mime>;base64,<data>" — extract just the base64 part
            const base64 = reader.result.split(",")[1];
            resolve(base64);
        };
        reader.onerror = function () {
            reject(new Error("Failed to read file: " + file.name));
        };
        reader.readAsDataURL(file);
    });
}
