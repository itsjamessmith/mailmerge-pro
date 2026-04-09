/* MailMerge-Pro — Outlook Web Add-in JavaScript */
/* Uses Office.js + SheetJS (XLSX) for client-side Excel reading + EWS for sending */

"use strict";

let appState = {
    currentStep: 1,
    headers: [],
    rows: [],
    mapping: { to: "", cc: "", bcc: "", subject: "" },
    results: []
};

// Initialize Office.js
Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        initUI();
    }
});

// ========== UI Initialization ==========

function initUI() {
    // File input
    document.getElementById("fileInput").addEventListener("change", handleFileUpload);

    // Navigation buttons
    document.getElementById("btnStep1Next").addEventListener("click", () => goToStep(2));
    document.getElementById("btnStep2Back").addEventListener("click", () => goToStep(1));
    document.getElementById("btnStep2Next").addEventListener("click", () => goToStep(3));
    document.getElementById("btnStep3Back").addEventListener("click", () => goToStep(2));
    document.getElementById("btnStep3Next").addEventListener("click", () => goToStep(4));
    document.getElementById("btnStep4Back").addEventListener("click", () => goToStep(3));
    document.getElementById("btnSend").addEventListener("click", executeMerge);
}

// ========== Step Navigation ==========

function goToStep(step) {
    // Save current step data
    if (appState.currentStep === 2) saveMapping();
    if (appState.currentStep === 3) saveCompose();

    // Validate before advancing
    if (step > appState.currentStep) {
        if (appState.currentStep === 1 && appState.rows.length === 0) {
            alert("Please upload a data file first."); return;
        }
        if (appState.currentStep === 2 && !document.getElementById("mapTo").value) {
            alert("Please select the To (Email) column."); return;
        }
    }

    // If going to step 2, populate column dropdowns
    if (step === 2) populateColumnDropdowns();
    // If going to step 3, populate merge field buttons
    if (step === 3) populateMergeFieldButtons();
    // If going to step 4, build review summary
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
}

// ========== Step 1: File Upload ==========

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById("fileName").textContent = file.name;

    const reader = new FileReader();
    reader.onload = function (evt) {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            if (jsonData.length === 0) {
                alert("No data rows found in the file.");
                return;
            }

            appState.headers = Object.keys(jsonData[0]);
            appState.rows = jsonData;

            // Show preview
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

        } catch (err) {
            alert("Error reading file: " + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ========== Step 2: Column Mapping ==========

function populateColumnDropdowns() {
    const selects = ["mapTo", "mapCC", "mapBCC", "mapSubject"];
    const autoMatch = {
        mapTo: ["email", "e-mail", "to", "emailaddress", "mail"],
        mapCC: ["cc", "carbon copy"],
        mapBCC: ["bcc", "blind"],
        mapSubject: ["subject", "subj"]
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

        // Auto-select matching columns
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

function saveCompose() {
    // Values are read directly from the DOM during send
}

// ========== Step 4: Review & Send ==========

function buildReview() {
    saveMapping();

    const subject = document.getElementById("emailSubject").value || "(no subject)";
    const globalCC = document.getElementById("globalCC").value;
    const globalBCC = document.getElementById("globalBCC").value;
    const bodyPreview = document.getElementById("emailBody").value.substring(0, 200);

    let html = `<p><strong>Recipients:</strong> ${appState.rows.length}</p>`;
    html += `<p><strong>To column:</strong> ${appState.mapping.to}</p>`;
    if (appState.mapping.cc) html += `<p><strong>CC column:</strong> ${appState.mapping.cc}</p>`;
    if (appState.mapping.bcc) html += `<p><strong>BCC column:</strong> ${appState.mapping.bcc}</p>`;
    if (globalCC) html += `<p><strong>Global CC:</strong> ${escapeHtml(globalCC)}</p>`;
    if (globalBCC) html += `<p><strong>Global BCC:</strong> ${escapeHtml(globalBCC)}</p>`;
    html += `<p><strong>Subject:</strong> ${escapeHtml(subject)}</p>`;
    document.getElementById("reviewSummary").innerHTML = html;

    // Preview first email
    if (appState.rows.length > 0) {
        const row = appState.rows[0];
        const mergedSubj = mergeFields(subject, row);
        const mergedBody = mergeFields(document.getElementById("emailBody").value, row);

        let preview = `<p class="label">Preview (recipient 1):</p>`;
        preview += `<p><strong>To:</strong> ${escapeHtml(row[appState.mapping.to] || "")}</p>`;
        preview += `<p><strong>Subject:</strong> ${escapeHtml(mergedSubj)}</p>`;
        preview += `<hr/><p>${escapeHtml(mergedBody).replace(/\n/g, "<br/>")}</p>`;
        document.getElementById("previewEmail").innerHTML = preview;
    }
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

// ========== Email Sending via EWS ==========

async function executeMerge() {
    const subject = document.getElementById("emailSubject").value;
    const body = document.getElementById("emailBody").value;
    const globalCC = document.getElementById("globalCC").value.trim();
    const globalBCC = document.getElementById("globalBCC").value.trim();
    const sendAsHtml = document.getElementById("chkSendAsHtml").checked;
    const draftOnly = document.getElementById("chkDraftOnly").checked;
    const delay = parseInt(document.getElementById("sendDelay").value) || 1;

    if (!subject && !appState.mapping.subject) {
        alert("Please enter a subject line."); return;
    }
    if (!body) {
        alert("Please enter an email body."); return;
    }

    const total = appState.rows.length;
    if (!confirm(`Ready to ${draftOnly ? "save to Drafts" : "send"} ${total} emails. Continue?`)) return;

    // Show progress
    document.getElementById("progressContainer").style.display = "block";
    document.getElementById("btnSend").disabled = true;
    document.getElementById("btnStep4Back").disabled = true;

    appState.results = [];
    let sent = 0, errors = 0;

    for (let i = 0; i < total; i++) {
        const row = appState.rows[i];
        const toAddr = row[appState.mapping.to];
        if (!toAddr) { errors++; continue; }

        const mergedSubject = appState.mapping.subject
            ? mergeFields(String(row[appState.mapping.subject] || subject), row)
            : mergeFields(subject, row);
        const mergedBody = mergeFields(body, row);

        // Build CC list
        let ccList = "";
        if (appState.mapping.cc && row[appState.mapping.cc]) ccList = row[appState.mapping.cc];
        if (globalCC) ccList = ccList ? ccList + ";" + globalCC : globalCC;

        // Build BCC list
        let bccList = "";
        if (appState.mapping.bcc && row[appState.mapping.bcc]) bccList = row[appState.mapping.bcc];
        if (globalBCC) bccList = bccList ? bccList + ";" + globalBCC : globalBCC;

        // Update progress
        const pct = Math.round(((i + 1) / total) * 100);
        document.getElementById("progressFill").style.width = pct + "%";
        document.getElementById("progressText").textContent =
            `${draftOnly ? "Drafting" : "Sending"} ${i + 1} of ${total} — ${toAddr}`;

        try {
            if (draftOnly) {
                await createDraft(toAddr, ccList, bccList, mergedSubject, mergedBody, sendAsHtml);
            } else {
                await sendViaEWS(toAddr, ccList, bccList, mergedSubject, mergedBody, sendAsHtml);
            }
            sent++;
            appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent" });
        } catch (err) {
            errors++;
            appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: err.message || String(err) });
        }

        // Throttle
        if (i < total - 1 && delay > 0) {
            await sleep(delay * 1000);
        }
    }

    // Show results
    document.getElementById("progressContainer").style.display = "none";
    const resultsDiv = document.getElementById("resultsContainer");
    resultsDiv.style.display = "block";
    resultsDiv.className = errors > 0 ? "results error" : "results success";
    resultsDiv.innerHTML = `
        <h3>✅ Mail Merge Complete</h3>
        <p><strong>Total:</strong> ${total}</p>
        ${sent > 0 ? `<p><strong>${draftOnly ? "Drafts" : "Sent"}:</strong> ${sent}</p>` : ""}
        ${errors > 0 ? `<p><strong>Errors:</strong> ${errors}</p>` : ""}
        ${errors > 0 ? `<p style="font-size:11px;color:#d32f2f;">Failed: ${appState.results.filter(r => r.status === "Error").map(r => r.to + " (" + r.error + ")").join(", ")}</p>` : ""}
    `;

    document.getElementById("btnSend").disabled = false;
    document.getElementById("btnStep4Back").disabled = false;
}

// Create email using Office.js displayNewMessageFormAsync (for drafts/preview)
function createDraft(to, cc, bcc, subject, body, asHtml) {
    return new Promise((resolve, reject) => {
        try {
            const options = {
                toRecipients: to.split(/[;,]/).map(e => e.trim()).filter(Boolean),
                ccRecipients: cc ? cc.split(/[;,]/).map(e => e.trim()).filter(Boolean) : [],
                bccRecipients: bcc ? bcc.split(/[;,]/).map(e => e.trim()).filter(Boolean) : [],
                subject: subject,
                htmlBody: asHtml ? body.replace(/\n/g, "<br/>") : undefined,
            };

            // displayNewMessageFormAsync opens a compose window (user can review & send)
            Office.context.mailbox.displayNewMessageFormAsync(options, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(new Error(result.error.message || "Failed to create draft"));
                }
            });
        } catch (err) {
            reject(err);
        }
    });
}

// Send email using EWS (Exchange Web Services) via makeEwsRequestAsync
function sendViaEWS(to, cc, bcc, subject, body, asHtml) {
    return new Promise((resolve, reject) => {
        const bodyType = asHtml ? "HTML" : "Text";
        const htmlBody = asHtml ? body.replace(/\n/g, "<br/>") : body;

        // Build EWS CreateItem SOAP request
        let ewsRequest = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:Message>
          <t:Subject>${escapeXml(subject)}</t:Subject>
          <t:Body BodyType="${bodyType}">${escapeXml(htmlBody)}</t:Body>
          <t:ToRecipients>
            ${to.split(/[;,]/).filter(Boolean).map(e =>
              `<t:Mailbox><t:EmailAddress>${escapeXml(e.trim())}</t:EmailAddress></t:Mailbox>`
            ).join("\n            ")}
          </t:ToRecipients>`;

        if (cc) {
            ewsRequest += `
          <t:CcRecipients>
            ${cc.split(/[;,]/).filter(Boolean).map(e =>
              `<t:Mailbox><t:EmailAddress>${escapeXml(e.trim())}</t:EmailAddress></t:Mailbox>`
            ).join("\n            ")}
          </t:CcRecipients>`;
        }

        if (bcc) {
            ewsRequest += `
          <t:BccRecipients>
            ${bcc.split(/[;,]/).filter(Boolean).map(e =>
              `<t:Mailbox><t:EmailAddress>${escapeXml(e.trim())}</t:EmailAddress></t:Mailbox>`
            ).join("\n            ")}
          </t:BccRecipients>`;
        }

        ewsRequest += `
        </t:Message>
      </m:Items>
    </m:CreateItem>
  </soap:Body>
</soap:Envelope>`;

        Office.context.mailbox.makeEwsRequestAsync(ewsRequest, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Check EWS response for errors
                const response = result.value;
                if (response.includes("NoError") || response.includes("Success")) {
                    resolve();
                } else if (response.includes("Error")) {
                    const match = response.match(/ResponseCode>([^<]+)</);
                    reject(new Error(match ? match[1] : "EWS error"));
                } else {
                    resolve(); // Assume success if no explicit error
                }
            } else {
                reject(new Error(result.error.message || "EWS request failed"));
            }
        });
    });
}

// ========== Utilities ==========

function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
}

function escapeXml(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}
