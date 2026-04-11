/* MailMerge-Pro — Outlook Web Add-in JavaScript
 * MSAL.js 2.0 + Microsoft Graph API
 * NO alert/confirm/prompt — all feedback via DOM
 */
"use strict";

// ========== Internationalization (i18n) ==========
let currentLang = localStorage.getItem("mailmergepro_language") || "en";
const translations = {
    en: {
        appTitle:"📧 MailMerge-Pro", stepData:"Data", stepMap:"Map", stepCompose:"Compose", stepSend:"Send",
        uploadTitle:"📋 Upload Recipient Data", uploadDesc:"Select an Excel (.xlsx) or CSV file with your recipients.",
        chooseFile:"📁 Choose File", noFileSelected:"No file selected", orDivider:"— or —",
        importContacts:"👤 Import from Contacts", savedLists:"📑 Saved Lists", saveCurrentList:"💾 Save Current List",
        loadSavedList:"Load saved list...", deleteSavedList:"🗑️", mergeLists:"🔗 Merge Lists",
        savedListsEmpty:"No saved lists yet.", sizeWarning:"⚠️ Data exceeds 3MB. localStorage may be full.",
        mapTitle:"🔗 Map Columns", mapDesc:"Match your spreadsheet columns to email fields.",
        toLabel:"To (Email)", ccLabel:"CC", bccLabel:"BCC", subjectLabel:"Subject", attachmentsLabel:"Attachments",
        composeTitle:"✏️ Compose Email", fromLabel:"From / Send As", sharedMailboxLabel:"Shared Mailbox",
        subjectLineLabel:"Subject Line", globalCCLabel:"Global CC", globalBCCLabel:"Global BCC",
        emailBodyLabel:"Email Body", sendAsHtml:"Send as HTML", templatesTitle:"📝 Templates",
        saveAsTemplate:"💾 Save as Template", loadTemplate:"Load template...", deleteTemplate:"🗑️",
        builtInLabel:"Built-in", customLabel:"Custom", abTestTitle:"🔬 A/B Testing",
        abTestEnable:"🔬 Enable A/B Test", versionA:"Version A", versionB:"Version B", splitRatio:"Split Ratio",
        importHtml:"📄 Import HTML", insertSignature:"✒️ Sig", signatureTitle:"✍️ Signature",
        autoAppendSig:"Auto-append signature", fetchSignature:"Fetch from Outlook",
        pasteManually:"Or paste your signature below:", signatureSaved:"Signature saved!",
        fallbackDefaults:"🔄 Fallback Default Values",
        fallbackHint:"Used when a merge field resolves to empty for a recipient.",
        globalAttachments:"📎 Global Attachments", perRecipientAttachments:"📎 Per-Recipient Attachments",
        reviewTitle:"🚀 Review & Send", optionsTitle:"⚙️ Options", delayLabel:"Delay between emails (seconds)",
        draftOnly:"📝 Save as Drafts only", readReceipt:"📬 Request read receipt",
        highImportance:"❗ High importance", groupByEmail:"📊 Group by email (many-to-one)",
        unsubscribe:"🚫 Add unsubscribe link", trackingTitle:"📈 Email Tracking",
        addTracking:"Add read tracking",
        trackingNote:"Uses Graph API read receipt. For advanced tracking (open rates, click rates), consider upgrading to a hosted version.",
        scheduleTitle:"⏰ Scheduled Sending", scheduleSend:"📅 Schedule Send", scheduleBtn:"Schedule",
        cancelSchedule:"Cancel Schedule",
        scheduleWarning:"⚠️ Outlook must remain open for scheduled sends to work.",
        schedulePast:"Scheduled time is in the past — sending immediately.",
        scheduleSet:"Sending in", scheduleInterrupted:"Schedule was interrupted (page refreshed).",
        rateLimitTitle:"📊 Rate Limit Dashboard", sentToday:"Emails sent today",
        dailyLimit:"Recommended daily limit: 10,000", perMinute:"Recommended per-minute: 30",
        suggestedDelay:"Suggested delay", sendAll:"🚀 Send All Emails", test:"🧪 Test",
        back:"← Back", next:"Next →", prev:"◀ Prev", nextPreview:"Next ▶",
        ctrlEnterHint:"Ctrl+Enter to send", dashboard:"📊 Dashboard",
        dashboardTitle:"📊 Admin Dashboard", totalCampaigns:"Total Campaigns",
        totalEmails:"Total Emails Sent", successRate:"Success Rate", avgCampaignSize:"Avg Campaign Size",
        topRecipients:"Top Recipients", monthlyActivity:"Monthly Activity",
        recentCampaigns:"Recent Campaigns", closeDashboard:"Close",
        signIn:"Sign In", signOut:"Sign Out", getStarted:"Get Started",
        welcome:"👋 Welcome to MailMerge-Pro!",
        welcomeDesc:"Upload your spreadsheet to get started with personalized bulk email.",
        recipients:"Recipients", columns:"columns", required:"*", optional:"(optional)",
        none:"(none)", history:"📜 History", campaignDetails:"📜 Campaign Details",
        close:"Close", insertLink:"Insert Link", cancel:"Cancel", insert:"Insert",
        searchRecipients:"🔍 Search recipients...", selectContacts:"👤 Select Contacts",
        selectAll:"Select All", useSelected:"Use Selected", addFiles:"+ Add Files",
        uploadFiles:"+ Upload Files", templateNamePrompt:"Template Name",
        templateNamePlaceholder:"Enter template name...", save:"Save",
        listNamePrompt:"List Name", listNamePlaceholder:"Enter list name...",
        mergeSelectPrompt:"Select a list to merge with current data:", merge:"Merge",
        noRecipients:"No recipients.", enterSubject:"⚠️ Enter a subject line.",
        enterBody:"⚠️ Enter an email body.", uploadFirst:"⚠️ Upload a data file or import contacts first.",
        selectToColumn:"⚠️ Select the To (Email) column.", authFailed:"❌ Authentication failed",
        sendComplete:"Mail Merge Complete", error:"Error", sent:"Sent", draft:"Draft",
        dragDropHtml:"Drop .html file here or click Import HTML",
        htmlImported:"HTML template imported!", scheduledFor:"Scheduled for", sending:"Sending",
        notSignedIn:"Not signed in"
    },
    es: {
        appTitle:"📧 MailMerge-Pro", stepData:"Datos", stepMap:"Mapear", stepCompose:"Redactar", stepSend:"Enviar",
        uploadTitle:"📋 Subir Datos", uploadDesc:"Seleccione un archivo Excel (.xlsx) o CSV.",
        chooseFile:"📁 Elegir Archivo", noFileSelected:"Ningún archivo", orDivider:"— o —",
        importContacts:"👤 Importar Contactos", savedLists:"📑 Listas Guardadas", saveCurrentList:"💾 Guardar Lista",
        mergeLists:"🔗 Fusionar", savedListsEmpty:"Sin listas guardadas.",
        sizeWarning:"⚠️ Datos superan 3MB.", mapTitle:"🔗 Mapear Columnas",
        mapDesc:"Asocie las columnas a campos de correo.", toLabel:"Para (Email)",
        composeTitle:"✏️ Redactar", subjectLineLabel:"Asunto", emailBodyLabel:"Cuerpo",
        sendAsHtml:"Enviar como HTML", templatesTitle:"📝 Plantillas", saveAsTemplate:"💾 Guardar Plantilla",
        abTestEnable:"🔬 Prueba A/B", versionA:"Versión A", versionB:"Versión B", splitRatio:"Proporción",
        importHtml:"📄 Importar HTML", insertSignature:"✒️ Firma", signatureTitle:"✍️ Firma",
        autoAppendSig:"Agregar firma auto", fetchSignature:"Obtener de Outlook",
        pasteManually:"O pegue su firma:", signatureSaved:"¡Firma guardada!",
        reviewTitle:"🚀 Revisar y Enviar", optionsTitle:"⚙️ Opciones",
        delayLabel:"Retraso entre correos (seg)", draftOnly:"📝 Solo borradores",
        readReceipt:"📬 Acuse de lectura", highImportance:"❗ Alta importancia",
        addTracking:"Seguimiento de lectura",
        trackingNote:"Usa acuse de lectura de Graph API. Para seguimiento avanzado, considere actualizar.",
        scheduleTitle:"⏰ Envío Programado", scheduleBtn:"Programar", cancelSchedule:"Cancelar",
        scheduleWarning:"⚠️ Outlook debe permanecer abierto.", scheduleSet:"Enviando en",
        scheduleInterrupted:"Programación interrumpida.", rateLimitTitle:"📊 Límites de Envío",
        sentToday:"Enviados hoy", dailyLimit:"Límite diario: 10,000", perMinute:"Por minuto: 30",
        sendAll:"🚀 Enviar Todos", test:"🧪 Prueba", back:"← Atrás", next:"Siguiente →",
        dashboard:"📊 Panel", dashboardTitle:"📊 Panel de Admin", totalCampaigns:"Campañas",
        totalEmails:"Correos Enviados", successRate:"Tasa de Éxito", avgCampaignSize:"Tamaño Promedio",
        topRecipients:"Top Destinatarios", monthlyActivity:"Actividad Mensual",
        recentCampaigns:"Campañas Recientes", closeDashboard:"Cerrar", signIn:"Iniciar Sesión",
        signOut:"Cerrar Sesión", getStarted:"Comenzar", cancel:"Cancelar", save:"Guardar",
        merge:"Fusionar", close:"Cerrar", sent:"Enviado", draft:"Borrador", error:"Error",
        sendComplete:"Combinación Completa", dragDropHtml:"Arrastre .html aquí",
        htmlImported:"¡Plantilla importada!", sending:"Enviando", notSignedIn:"No conectado",
        history:"📜 Historial", searchRecipients:"🔍 Buscar destinatarios...",
        noFileSelected:"Ningún archivo seleccionado", orDivider:"— o —",
        authFailed:"❌ Error de autenticación", noRecipients:"Sin destinatarios.",
        enterSubject:"⚠️ Introduzca un asunto.", enterBody:"⚠️ Introduzca un cuerpo.",
        uploadFirst:"⚠️ Suba un archivo primero.", selectToColumn:"⚠️ Seleccione la columna Para (Email)."
    },
    fr: {
        appTitle:"📧 MailMerge-Pro", stepData:"Données", stepMap:"Mapper", stepCompose:"Composer", stepSend:"Envoyer",
        uploadTitle:"📋 Télécharger les Données", uploadDesc:"Sélectionnez un fichier Excel (.xlsx) ou CSV.",
        chooseFile:"📁 Choisir", noFileSelected:"Aucun fichier", orDivider:"— ou —",
        importContacts:"👤 Importer Contacts", savedLists:"📑 Listes Sauvées", saveCurrentList:"💾 Sauver Liste",
        mergeLists:"🔗 Fusionner", savedListsEmpty:"Aucune liste sauvée.",
        mapTitle:"🔗 Mapper Colonnes", mapDesc:"Associez les colonnes aux champs email.",
        composeTitle:"✏️ Composer", subjectLineLabel:"Objet", emailBodyLabel:"Corps",
        sendAsHtml:"Envoyer en HTML", templatesTitle:"📝 Modèles", saveAsTemplate:"💾 Sauver Modèle",
        abTestEnable:"🔬 Test A/B", versionA:"Version A", versionB:"Version B", splitRatio:"Ratio",
        importHtml:"📄 Importer HTML", insertSignature:"✒️ Sig.", signatureTitle:"✍️ Signature",
        autoAppendSig:"Ajouter signature auto", fetchSignature:"Récupérer d'Outlook",
        pasteManually:"Ou collez votre signature:", signatureSaved:"Signature sauvée!",
        reviewTitle:"🚀 Vérifier et Envoyer", optionsTitle:"⚙️ Options",
        delayLabel:"Délai entre emails (sec)", draftOnly:"📝 Brouillons seulement",
        readReceipt:"📬 Accusé de réception", highImportance:"❗ Haute importance",
        addTracking:"Suivi de lecture", scheduleTitle:"⏰ Envoi Programmé",
        scheduleBtn:"Programmer", cancelSchedule:"Annuler",
        scheduleWarning:"⚠️ Outlook doit rester ouvert.", scheduleSet:"Envoi dans",
        rateLimitTitle:"📊 Limites d'envoi", sentToday:"Envoyés aujourd'hui",
        sendAll:"🚀 Envoyer Tout", test:"🧪 Test", back:"← Retour", next:"Suivant →",
        dashboard:"📊 Tableau", dashboardTitle:"📊 Tableau de Bord", totalCampaigns:"Campagnes",
        totalEmails:"Emails Envoyés", successRate:"Taux de Réussite", avgCampaignSize:"Taille Moyenne",
        topRecipients:"Top Destinataires", recentCampaigns:"Campagnes Récentes",
        closeDashboard:"Fermer", signIn:"Connexion", signOut:"Déconnexion",
        getStarted:"Commencer", cancel:"Annuler", save:"Sauver", merge:"Fusionner",
        close:"Fermer", sent:"Envoyé", draft:"Brouillon", error:"Erreur",
        sendComplete:"Fusion Terminée", sending:"Envoi", notSignedIn:"Non connecté"
    },
    de: {
        appTitle:"📧 MailMerge-Pro", stepData:"Daten", stepMap:"Zuordnen", stepCompose:"Verfassen", stepSend:"Senden",
        uploadTitle:"📋 Daten Hochladen", uploadDesc:"Wählen Sie eine Excel- (.xlsx) oder CSV-Datei.",
        chooseFile:"📁 Datei Wählen", noFileSelected:"Keine Datei", orDivider:"— oder —",
        importContacts:"👤 Kontakte Importieren", savedLists:"📑 Gespeicherte Listen",
        saveCurrentList:"💾 Liste Speichern", mergeLists:"🔗 Zusammenführen",
        savedListsEmpty:"Keine Listen gespeichert.", mapTitle:"🔗 Spalten Zuordnen",
        composeTitle:"✏️ Verfassen", subjectLineLabel:"Betreff", emailBodyLabel:"Inhalt",
        sendAsHtml:"Als HTML senden", templatesTitle:"📝 Vorlagen", saveAsTemplate:"💾 Vorlage Speichern",
        abTestEnable:"🔬 A/B-Test", versionA:"Version A", versionB:"Version B", splitRatio:"Aufteilung",
        importHtml:"📄 HTML Importieren", insertSignature:"✒️ Sig.", signatureTitle:"✍️ Signatur",
        autoAppendSig:"Signatur automatisch anhängen", fetchSignature:"Von Outlook abrufen",
        pasteManually:"Oder fügen Sie Ihre Signatur ein:", signatureSaved:"Signatur gespeichert!",
        reviewTitle:"🚀 Prüfen & Senden", optionsTitle:"⚙️ Optionen",
        delayLabel:"Verzögerung zwischen E-Mails (Sek.)", draftOnly:"📝 Nur Entwürfe",
        readReceipt:"📬 Lesebestätigung", highImportance:"❗ Hohe Priorität",
        addTracking:"Lesebestätigung aktivieren", scheduleTitle:"⏰ Zeitgesteuert Senden",
        scheduleBtn:"Planen", cancelSchedule:"Abbrechen",
        scheduleWarning:"⚠️ Outlook muss geöffnet bleiben.", scheduleSet:"Senden in",
        rateLimitTitle:"📊 Sendebeschränkungen", sentToday:"Heute gesendet",
        sendAll:"🚀 Alle Senden", test:"🧪 Test", back:"← Zurück", next:"Weiter →",
        dashboard:"📊 Dashboard", dashboardTitle:"📊 Admin-Dashboard", totalCampaigns:"Kampagnen",
        totalEmails:"Gesendete E-Mails", successRate:"Erfolgsrate", avgCampaignSize:"Durchschnittsgröße",
        topRecipients:"Top-Empfänger", recentCampaigns:"Letzte Kampagnen",
        closeDashboard:"Schließen", signIn:"Anmelden", signOut:"Abmelden",
        getStarted:"Loslegen", cancel:"Abbrechen", save:"Speichern", merge:"Zusammenführen",
        close:"Schließen", sent:"Gesendet", draft:"Entwurf", error:"Fehler",
        sendComplete:"Zusammenführung Abgeschlossen", sending:"Senden", notSignedIn:"Nicht angemeldet"
    },
    pt: {
        appTitle:"📧 MailMerge-Pro", stepData:"Dados", stepMap:"Mapear", stepCompose:"Compor", stepSend:"Enviar",
        uploadTitle:"📋 Carregar Dados", uploadDesc:"Selecione um arquivo Excel (.xlsx) ou CSV.",
        chooseFile:"📁 Escolher", noFileSelected:"Nenhum arquivo", orDivider:"— ou —",
        importContacts:"👤 Importar Contatos", savedLists:"📑 Listas Salvas",
        saveCurrentList:"💾 Salvar Lista", mergeLists:"🔗 Mesclar",
        savedListsEmpty:"Nenhuma lista salva.", mapTitle:"🔗 Mapear Colunas",
        composeTitle:"✏️ Compor", subjectLineLabel:"Assunto", emailBodyLabel:"Corpo",
        sendAsHtml:"Enviar como HTML", templatesTitle:"📝 Modelos", saveAsTemplate:"💾 Salvar Modelo",
        abTestEnable:"🔬 Teste A/B", versionA:"Versão A", versionB:"Versão B", splitRatio:"Proporção",
        importHtml:"📄 Importar HTML", insertSignature:"✒️ Assin.", signatureTitle:"✍️ Assinatura",
        autoAppendSig:"Adicionar assinatura auto", fetchSignature:"Buscar do Outlook",
        pasteManually:"Ou cole sua assinatura:", signatureSaved:"Assinatura salva!",
        reviewTitle:"🚀 Revisar e Enviar", optionsTitle:"⚙️ Opções",
        delayLabel:"Atraso entre emails (seg)", draftOnly:"📝 Apenas rascunhos",
        readReceipt:"📬 Confirmação de leitura", highImportance:"❗ Alta importância",
        addTracking:"Rastreamento de leitura", scheduleTitle:"⏰ Envio Agendado",
        scheduleBtn:"Agendar", cancelSchedule:"Cancelar",
        scheduleWarning:"⚠️ Outlook deve permanecer aberto.", scheduleSet:"Enviando em",
        rateLimitTitle:"📊 Limites de Envio", sentToday:"Enviados hoje",
        sendAll:"🚀 Enviar Todos", test:"🧪 Teste", back:"← Voltar", next:"Próximo →",
        dashboard:"📊 Painel", dashboardTitle:"📊 Painel Admin", totalCampaigns:"Campanhas",
        totalEmails:"Emails Enviados", successRate:"Taxa de Sucesso", avgCampaignSize:"Tamanho Médio",
        topRecipients:"Top Destinatários", recentCampaigns:"Campanhas Recentes",
        closeDashboard:"Fechar", signIn:"Entrar", signOut:"Sair",
        getStarted:"Começar", cancel:"Cancelar", save:"Salvar", merge:"Mesclar",
        close:"Fechar", sent:"Enviado", draft:"Rascunho", error:"Erro",
        sendComplete:"Mesclagem Concluída", sending:"Enviando", notSignedIn:"Não conectado"
    },
    ja: {
        appTitle:"📧 MailMerge-Pro", stepData:"データ", stepMap:"マップ", stepCompose:"作成", stepSend:"送信",
        uploadTitle:"📋 データアップロード", uploadDesc:"Excel (.xlsx) または CSV ファイルを選択してください。",
        chooseFile:"📁 ファイル選択", noFileSelected:"未選択", orDivider:"— または —",
        importContacts:"👤 連絡先インポート", savedLists:"📑 保存リスト",
        saveCurrentList:"💾 リスト保存", mergeLists:"🔗 結合",
        savedListsEmpty:"保存されたリストはありません。", mapTitle:"🔗 列マッピング",
        composeTitle:"✏️ メール作成", subjectLineLabel:"件名", emailBodyLabel:"本文",
        sendAsHtml:"HTMLで送信", templatesTitle:"📝 テンプレート", saveAsTemplate:"💾 テンプレート保存",
        abTestEnable:"🔬 A/Bテスト", versionA:"バージョンA", versionB:"バージョンB", splitRatio:"分割比率",
        importHtml:"📄 HTMLインポート", insertSignature:"✒️ 署名", signatureTitle:"✍️ 署名",
        autoAppendSig:"署名を自動追加", fetchSignature:"Outlookから取得",
        pasteManually:"または署名を貼り付け:", signatureSaved:"署名を保存しました！",
        reviewTitle:"🚀 確認＆送信", optionsTitle:"⚙️ オプション",
        delayLabel:"メール間の遅延（秒）", draftOnly:"📝 下書きのみ",
        readReceipt:"📬 開封確認", highImportance:"❗ 高重要度",
        addTracking:"開封トラッキング", scheduleTitle:"⏰ 予約送信",
        scheduleBtn:"予約", cancelSchedule:"キャンセル",
        scheduleWarning:"⚠️ 予約送信にはOutlookを開いたままにしてください。", scheduleSet:"送信まで",
        rateLimitTitle:"📊 送信制限", sentToday:"本日の送信数",
        sendAll:"🚀 全て送信", test:"🧪 テスト", back:"← 戻る", next:"次へ →",
        dashboard:"📊 ダッシュボード", dashboardTitle:"📊 管理ダッシュボード",
        totalCampaigns:"キャンペーン数", totalEmails:"送信メール数",
        successRate:"成功率", avgCampaignSize:"平均サイズ",
        topRecipients:"トップ受信者", recentCampaigns:"最近のキャンペーン",
        closeDashboard:"閉じる", signIn:"サインイン", signOut:"サインアウト",
        getStarted:"始める", cancel:"キャンセル", save:"保存", merge:"結合",
        close:"閉じる", sent:"送信済", draft:"下書き", error:"エラー",
        sendComplete:"メール結合完了", sending:"送信中", notSignedIn:"未サインイン"
    }
};

function t(key) {
    return (translations[currentLang] && translations[currentLang][key]) || translations.en[key] || key;
}

function setLanguage(lang) {
    currentLang = lang;
    localStorage.setItem("mailmergepro_language", lang);
    applyTranslations();
}

function applyTranslations() {
    document.querySelectorAll("[data-i18n]").forEach(function(el) {
        var key = el.getAttribute("data-i18n");
        var val = t(key);
        if (val) el.textContent = val;
    });
}

// ========== MSAL Configuration ==========
const msalConfig = {
    auth: {
        clientId: "360e4343-614f-4f70-a650-c020868516fc",
        authority: "https://login.microsoftonline.com/common",
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
    fileName: "",
    previewIndex: 0,
    contactsData: [],
    abTestEnabled: false,
    abVersion: "A",
    scheduledTimer: null,
    scheduledTime: null,
    scheduleCountdownInterval: null,
    signatureHtml: localStorage.getItem("mailmergepro_signature") || "",
    autoSignature: localStorage.getItem("mailmergepro_autosignature") === "true"
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
    // Restore saved preference first
    const savedPref = localStorage.getItem("mailmergepro_darkmode");
    if (savedPref === "true") { document.body.classList.add("dark-mode"); return; }
    if (savedPref === "false") { return; }
    // Auto-detect from OS or Office theme
    const prefersDark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
    if (prefersDark) document.body.classList.add("dark-mode");
    try {
        if (officeInfo && typeof Office !== "undefined" && Office.context && Office.context.officeTheme) {
            const bg = Office.context.officeTheme.bodyBackgroundColor;
            if (bg && isColorDark(bg)) document.body.classList.add("dark-mode");
        }
    } catch (e) { console.log("Theme detection skipped:", e.message); }
}
function toggleDarkMode() {
    document.body.classList.toggle("dark-mode");
    const isDark = document.body.classList.contains("dark-mode");
    localStorage.setItem("mailmergepro_darkmode", String(isDark));
    const btn = document.getElementById("darkModeToggle");
    if (btn) btn.textContent = isDark ? "☀️" : "🌙";
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
            authEl.textContent = t("notSignedIn");
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
        authEl.textContent = t("notSignedIn");
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
            authEl.textContent = t("authFailed") + ": " + (e.message || String(e));
            authEl.classList.add("error");
            return;
        }
    }
    try {
        // Clear stuck interaction state before attempting new login
        const activeAccount = msalInstance.getActiveAccount();
        if (activeAccount) {
            try {
                const result = await msalInstance.acquireTokenSilent({ ...loginRequest, account: activeAccount });
                updateAuthUI(activeAccount);
                return result.accessToken;
            } catch (_) { /* silent failed, proceed to popup */ }
        }
        const result = await msalInstance.acquireTokenPopup(loginRequest);
        console.log("signIn success:", result.account.username);
        updateAuthUI(result.account);
        return result.accessToken;
    } catch (err) {
        console.error("signIn error:", err);
        const msg = err.message || String(err);
        if (msg.includes("interaction_in_progress")) {
            // Clear stuck MSAL interaction state
            try {
                const keys = Object.keys(sessionStorage);
                keys.forEach(k => { if (k.indexOf("interaction") !== -1) sessionStorage.removeItem(k); });
            } catch (_) {}
            authEl.textContent = t("signIn") + " — " + t("error") + ". Please try again.";
        } else if (msg.includes("popup_window_error")) {
            authEl.textContent = "Pop-up blocked. Allow pop-ups and try again.";
        } else if (msg.includes("user_cancelled")) {
            authEl.textContent = "Sign-in cancelled.";
        } else if (msg.includes("AADSTS50020") || msg.includes("AADSTS700016")) {
            authEl.textContent = "Account not authorized for this app. Check your tenant.";
        } else {
            authEl.textContent = t("authFailed") + ": " + msg.substring(0, 120);
        }
        authEl.classList.add("error");
        throw err;
    }
}

async function signOut() {
    console.log("signOut: starting");
    try {
        if (msalInstance) {
            const accounts = msalInstance.getAllAccounts();
            console.log("signOut: found", accounts.length, "accounts");
            
            // Method 1: Try MSAL logout (works in modern browsers / new Outlook)
            if (accounts.length > 0) {
                try {
                    await msalInstance.logoutPopup({ account: accounts[0] });
                    console.log("signOut: logoutPopup succeeded");
                } catch (popupErr) {
                    console.warn("signOut: logoutPopup failed (expected in classic Outlook):", popupErr.message);
                    // Method 2: Manual cache clear (fallback for classic Outlook)
                    try {
                        accounts.forEach(acct => {
                            try { msalInstance.getTokenCache().removeAccount(acct); } catch(_) {}
                        });
                    } catch(_) {}
                }
            }
            
            // Method 3: Brute force clear all MSAL keys from storage
            try {
                const keysToRemove = [];
                for (let i = 0; i < localStorage.length; i++) {
                    const key = localStorage.key(i);
                    if (key && (key.indexOf("msal") !== -1 || key.indexOf("login.microsoftonline") !== -1 ||
                        key.indexOf("authority") !== -1 || key.indexOf("token") !== -1)) {
                        keysToRemove.push(key);
                    }
                }
                keysToRemove.forEach(k => { try { localStorage.removeItem(k); } catch(_) {} });
                console.log("signOut: cleared", keysToRemove.length, "localStorage keys");
            } catch (storageErr) {
                console.warn("signOut: localStorage clear failed:", storageErr.message);
            }
            
            // Method 4: Recreate MSAL instance (nuclear option)
            try {
                msalInstance = new msal.PublicClientApplication(msalConfig);
                await msalInstance.initialize();
                console.log("signOut: MSAL instance recreated");
            } catch(_) {}
        }
    } catch (err) {
        console.warn("signOut error:", err);
    }
    
    appState.userEmail = "";
    updateAuthUI(null);
    console.log("signOut: complete");
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
    document.getElementById("btnSignOut").addEventListener("click", async () => { await signOut(); });
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

    // ===== New Feature Event Listeners =====

    // Feature 9: Language selector
    const langSel = document.getElementById("languageSelect");
    if (langSel) {
        langSel.value = currentLang;
        langSel.addEventListener("change", (e) => setLanguage(e.target.value));
    }

    // Feature 10: Dashboard
    document.getElementById("btnDashboard").addEventListener("click", openDashboard);
    document.getElementById("btnCloseDashboard").addEventListener("click", closeDashboard);

    // Feature 1: Templates
    document.getElementById("templatesHeader").addEventListener("click", () => toggleCollapsible("templatesHeader", "templatesBody"));
    document.getElementById("btnSaveTemplate").addEventListener("click", showSaveTemplateDialog);
    document.getElementById("btnTemplateNameCancel").addEventListener("click", () => { document.getElementById("templateNameDialog").style.display = "none"; });
    document.getElementById("btnTemplateNameSave").addEventListener("click", saveCurrentTemplate);
    document.getElementById("templateNameInput").addEventListener("keydown", (e) => { if (e.key === "Enter") saveCurrentTemplate(); });
    loadTemplatesUI();

    // Feature 5: Contact Groups / Saved Lists
    document.getElementById("savedListsHeader").addEventListener("click", () => toggleCollapsible("savedListsHeader", "savedListsBody"));
    document.getElementById("btnSaveCurrentList").addEventListener("click", showSaveListDialog);
    document.getElementById("btnMergeLists").addEventListener("click", showMergeListDialog);
    document.getElementById("btnListNameCancel").addEventListener("click", () => { document.getElementById("listNameDialog").style.display = "none"; });
    document.getElementById("btnListNameSave").addEventListener("click", saveCurrentList);
    document.getElementById("listNameInput").addEventListener("keydown", (e) => { if (e.key === "Enter") saveCurrentList(); });
    document.getElementById("btnMergeListCancel").addEventListener("click", () => { document.getElementById("mergeListDialog").style.display = "none"; });
    document.getElementById("btnMergeListConfirm").addEventListener("click", mergeSelectedList);
    loadSavedListsUI();

    // Feature 4: A/B Testing
    document.getElementById("chkABTest").addEventListener("change", toggleABTest);
    document.querySelectorAll(".ab-tab").forEach(tab => {
        tab.addEventListener("click", () => switchABTab(tab.dataset.ab));
    });

    // Feature 6: HTML Import
    document.getElementById("htmlImportHeader").addEventListener("click", () => toggleCollapsible("htmlImportHeader", "htmlImportBody"));
    document.getElementById("htmlFileInput").addEventListener("change", handleHtmlImport);
    initHtmlDragDrop();

    // Feature 7: Signature
    document.getElementById("signatureHeader").addEventListener("click", () => toggleCollapsible("signatureHeader", "signatureBody"));
    document.getElementById("btnInsertSignature").addEventListener("click", showSignatureDialog);
    document.getElementById("btnManageSignature").addEventListener("click", showSignatureDialog);
    document.getElementById("btnFetchSignature").addEventListener("click", fetchOutlookSignature);
    document.getElementById("btnSignatureCancel").addEventListener("click", () => { document.getElementById("signatureDialog").style.display = "none"; });
    document.getElementById("btnSignatureSave").addEventListener("click", saveSignature);
    document.getElementById("chkAutoSignatureInline").checked = appState.autoSignature;
    document.getElementById("chkAutoSignatureInline").addEventListener("change", (e) => {
        appState.autoSignature = e.target.checked;
        localStorage.setItem("mailmergepro_autosignature", e.target.checked ? "true" : "false");
    });
    loadSignaturePreview();

    // Feature 3: Email Tracking
    document.getElementById("chkTracking").addEventListener("change", (e) => {
        document.getElementById("trackingNote").style.display = e.target.checked ? "block" : "none";
        if (e.target.checked) document.getElementById("chkReadReceipt").checked = true;
    });

    // Feature 2: Scheduled Sending
    document.getElementById("scheduleHeader").addEventListener("click", () => toggleCollapsible("scheduleHeader", "scheduleBody"));
    document.getElementById("btnScheduleSend").addEventListener("click", scheduleSend);
    document.getElementById("btnCancelSchedule").addEventListener("click", cancelScheduledSend);
    checkInterruptedSchedule();

    // Feature 8: Rate Limit
    document.getElementById("rateLimitHeader").addEventListener("click", () => toggleCollapsible("rateLimitHeader", "rateLimitBody"));
    updateRateLimitDisplay();

    // Apply i18n
    applyTranslations();
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
function saveCampaign(name, total, sent, errors, extra) {
    const stored = localStorage.getItem("mailmerge-pro-campaigns");
    const campaigns = stored ? JSON.parse(stored) : [];
    const record = {
        name: name || "Untitled",
        date: new Date().toLocaleDateString(),
        dateISO: new Date().toISOString(),
        total: total,
        sent: sent,
        errors: errors,
        recipients: (extra && extra.recipients) ? extra.recipients : [],
        abTest: (extra && extra.abTest) ? extra.abTest : null
    };
    campaigns.unshift(record);
    if (campaigns.length > 50) campaigns.length = 50;
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
    appState.fileName = file.name;
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
            document.getElementById("btnSaveCurrentList").disabled = false;
            document.getElementById("btnMergeLists").disabled = false;
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
    document.getElementById("btnSaveCurrentList").disabled = false;
    document.getElementById("btnMergeLists").disabled = false;
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
    const displayNames = new Map(); // lowered filename -> original display name
    appState.rows.forEach(row => {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) val.split(";").forEach(f => {
            const n = f.trim();
            if (n) {
                const fileName = n.replace(/^.*[\\\/]/, "").toLowerCase();
                allNeeded.add(fileName);
                displayNames.set(fileName, n);
            }
        });
    });
    const missing = [];
    for (const name of allNeeded) { if (!appState.perRecipientFiles.has(name)) missing.push(displayNames.get(name) || name); }
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
    // A/B Test info
    if (appState.abTestEnabled) {
        const ratio = parseInt(document.getElementById("abSplitRatio").value) || 50;
        const groupA = Math.ceil(recipientCount * ratio / 100);
        const groupB = recipientCount - groupA;
        html += '<p style="color:var(--primary);font-weight:600;">🔬 A/B Test: Group A: ' + groupA + ' recipients, Group B: ' + groupB + ' recipients (' + ratio + '/' + (100-ratio) + ')</p>';
    }
    document.getElementById("reviewSummary").innerHTML = html;
    // Update rate limit display when entering Step 4
    updateRateLimitDisplay();
    suggestDelay(recipientCount);
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
    // Add global attachments
    for (const [, att] of appState.globalAttachments) attachments.push(att);
    // Add per-recipient attachments from spreadsheet column
    if (appState.mapping.attachments) {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) val.split(";").forEach(f => {
            const rawName = f.trim();
            if (!rawName) return;
            // Extract just the filename from full path (handle both / and \)
            const fileName = rawName.replace(/^.*[\\\/]/, "").toLowerCase();
            // Try exact match first, then filename-only match
            let att = appState.perRecipientFiles.get(rawName.toLowerCase());
            if (!att) att = appState.perRecipientFiles.get(fileName);
            if (att) {
                attachments.push(att);
                console.log("Per-recipient attachment matched:", rawName, "->", att.name);
            } else {
                console.warn("Per-recipient attachment not found:", rawName, "| Tried key:", fileName, "| Available:", Array.from(appState.perRecipientFiles.keys()).join(", "));
            }
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

        // Auto-append signature if enabled
        let signatureBlock = "";
        if (appState.autoSignature && appState.signatureHtml) {
            signatureBlock = sendAsHtml
                ? '<br/><div class="email-signature">' + appState.signatureHtml + '</div>'
                : "\n\n" + appState.signatureHtml;
        }

        // A/B test config
        const abEnabled = document.getElementById("chkABTest").checked;
        const abRatio = abEnabled ? (parseInt(document.getElementById("abSplitRatio").value) || 50) : 100;
        const subjectB = abEnabled ? document.getElementById("emailSubjectB").value : "";
        const bodyBEditor = document.getElementById("emailBodyB");
        const bodyBHtml = abEnabled ? bodyBEditor.innerHTML : "";
        const bodyBPlain = abEnabled ? bodyBEditor.innerText : "";

        const body = sendAsHtml ? bodyContent + signatureBlock : plainBody + signatureBlock;
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
            const mBody = sendAsHtml ? mergeFields(bodyContent + signatureBlock, row) : mergeFields(plainBody + signatureBlock, row);
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
        let abSentA = 0, abSentB = 0, abErrA = 0, abErrB = 0;
        const recipientEmails = [];
        for (let i = 0; i < total; i++) {
            const item = sendItems[i];
            const toAddr = item.to;
            if (!toAddr) {
                errors++;
                appState.results.push({ row: i + 2, to: "(empty)", status: "Error", error: "No email address" });
                continue;
            }
            recipientEmails.push(toAddr);
            // A/B split: deterministic by index
            const isGroupB = abEnabled && (i >= Math.ceil(total * abRatio / 100));
            const useSubjectB = isGroupB && subjectB;
            const useBodyB = isGroupB && (bodyBHtml || bodyBPlain);

            const activeSubject = useSubjectB ? subjectB : subject;
            const activeBodyHtml = useBodyB ? bodyBHtml + signatureBlock : bodyContent + signatureBlock;
            const activeBodyPlain = useBodyB ? bodyBPlain + signatureBlock : plainBody + signatureBlock;
            const activeBody = sendAsHtml ? activeBodyHtml : activeBodyPlain;

            const row = item.rows[0];
            const isGroup = item.rows.length > 1;
            const mSubj = appState.mapping.subject
                ? (isGroup ? mergeFieldsWithGroup(String(row[appState.mapping.subject] || activeSubject), item.rows) : mergeFields(String(row[appState.mapping.subject] || activeSubject), row))
                : (isGroup ? mergeFieldsWithGroup(activeSubject, item.rows) : mergeFields(activeSubject, row));
            const mBody = isGroup
                ? (sendAsHtml ? mergeFieldsWithGroup(activeBodyHtml, item.rows) : mergeFieldsWithGroup(activeBodyPlain, item.rows))
                : (sendAsHtml ? mergeFields(activeBodyHtml, row) : mergeFields(activeBodyPlain, row));
            let ccL = ""; if (appState.mapping.cc && row[appState.mapping.cc]) ccL = String(row[appState.mapping.cc]);
            if (globalCC) ccL = ccL ? ccL + ";" + globalCC : globalCC;
            let bccL = ""; if (appState.mapping.bcc && row[appState.mapping.bcc]) bccL = String(row[appState.mapping.bcc]);
            if (globalBCC) bccL = bccL ? bccL + ";" + globalBCC : globalBCC;
            const atts = collectAttachmentsForRow(row);
            updateProgress(i, total, modeLabel + " " + (i+1) + " of " + total + " \u2014 " + escapeHtml(toAddr));
            try {
                await sendOneEmail(token, toAddr, ccL, bccL, mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                sent++;
                if (isGroupB) abSentB++; else abSentA++;
                appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent", abGroup: isGroupB ? "B" : "A" });
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
                        if (isGroupB) abSentB++; else abSentA++;
                        appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent", abGroup: isGroupB ? "B" : "A" });
                        continue;
                    } catch (retryErr) {
                        errors++;
                        if (isGroupB) abErrB++; else abErrA++;
                        appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: retryErr.message || String(retryErr), abGroup: isGroupB ? "B" : "A" });
                    }
                } else if (errMsg.startsWith("SESSION_EXPIRED:")) {
                    console.warn("Token expired, re-acquiring...");
                    try {
                        token = await getGraphToken();
                        await sendOneEmail(token, toAddr, ccL, bccL, mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                        sent++;
                        if (isGroupB) abSentB++; else abSentA++;
                        appState.results.push({ row: i + 2, to: toAddr, status: draftOnly ? "Draft" : "Sent", abGroup: isGroupB ? "B" : "A" });
                        continue;
                    } catch (retryErr) {
                        errors++;
                        if (isGroupB) abErrB++; else abErrA++;
                        appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: retryErr.message || String(retryErr), abGroup: isGroupB ? "B" : "A" });
                    }
                } else {
                    errors++;
                    if (isGroupB) abErrB++; else abErrA++;
                    appState.results.push({ row: i + 2, to: toAddr, status: "Error", error: errMsg, abGroup: isGroupB ? "B" : "A" });
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
        // Show A/B results if enabled
        if (abEnabled) {
            showABResults(abSentA, abErrA, abSentB, abErrB);
        }
        // Update rate limit counter
        updateDailySentCount(sent);
        updateRateLimitDisplay();
        // Save campaign with extra data
        const campName = document.getElementById("campaignName").value.trim();
        const abData = abEnabled ? { sentA: abSentA, errA: abErrA, sentB: abSentB, errB: abErrB, ratio: abRatio } : null;
        saveCampaign(campName, total, sent, errors, { recipients: recipientEmails, abTest: abData });

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

// ========== Collapsible Sections ==========
function toggleCollapsible(headerId, bodyId) {
    var header = document.getElementById(headerId);
    var body = document.getElementById(bodyId);
    if (!header || !body) return;
    var isOpen = body.classList.contains("open");
    if (isOpen) {
        body.classList.remove("open");
        header.classList.remove("open");
    } else {
        body.classList.add("open");
        header.classList.add("open");
    }
}

// ========== Feature 1: Email Templates Library ==========
function getBuiltInTemplates() {
    return [
        {
            name: "Welcome Email",
            subject: "Welcome {FirstName}!",
            body: '<h2 style="color:#0078d4;">Welcome aboard, {FirstName}!</h2>' +
                '<p>We\'re thrilled to have you join us. Here\'s what you can expect:</p>' +
                '<ul><li>Personalized onboarding experience</li><li>Access to our full suite of tools</li><li>Dedicated support team ready to help</li></ul>' +
                '<p>If you have any questions, don\'t hesitate to reach out.</p>' +
                '<p>Best regards,<br/>The Team</p>',
            createdAt: "built-in",
            builtIn: true
        },
        {
            name: "Product Update",
            subject: "New updates from our team",
            body: '<h2 style="color:#0078d4;">What\'s New This Month</h2>' +
                '<p>Hi {FirstName},</p>' +
                '<p>We\'ve been hard at work improving our product. Here are the highlights:</p>' +
                '<ul><li><strong>Feature 1:</strong> Improved performance across the board</li>' +
                '<li><strong>Feature 2:</strong> New dashboard with real-time analytics</li>' +
                '<li><strong>Feature 3:</strong> Enhanced security and compliance</li></ul>' +
                '<p>Update now to take advantage of these improvements!</p>' +
                '<p>— The Product Team</p>',
            createdAt: "built-in",
            builtIn: true
        },
        {
            name: "Invoice/Statement",
            subject: "Your statement for {Month}",
            body: '<h2 style="color:#0078d4;">Monthly Statement</h2>' +
                '<p>Dear {FirstName},</p>' +
                '<p>Please find below your statement for the current billing period.</p>' +
                '<table style="border-collapse:collapse;width:100%;margin:16px 0;">' +
                '<tr style="background:#f5f5f5;"><td style="padding:8px;border:1px solid #ddd;"><strong>Account</strong></td><td style="padding:8px;border:1px solid #ddd;">{Email}</td></tr>' +
                '<tr><td style="padding:8px;border:1px solid #ddd;"><strong>Period</strong></td><td style="padding:8px;border:1px solid #ddd;">{Month}</td></tr>' +
                '</table>' +
                '<p>If you have questions about this statement, please contact our billing team.</p>' +
                '<p>Thank you for your business!</p>',
            createdAt: "built-in",
            builtIn: true
        }
    ];
}

function getSavedTemplates() {
    var stored = localStorage.getItem("mailmergepro_templates");
    return stored ? JSON.parse(stored) : [];
}

function saveTemplatesStorage(templates) {
    localStorage.setItem("mailmergepro_templates", JSON.stringify(templates));
}

function loadTemplatesUI() {
    var container = document.getElementById("templateCards");
    if (!container) return;
    container.innerHTML = "";
    var builtIn = getBuiltInTemplates();
    var custom = getSavedTemplates();
    var all = builtIn.concat(custom);
    if (all.length === 0) {
        container.innerHTML = '<p class="hint">No templates available.</p>';
        return;
    }
    all.forEach(function(tmpl, idx) {
        var card = document.createElement("div");
        card.className = "template-card";
        var info = document.createElement("div");
        info.className = "template-card-info";
        info.innerHTML = '<div class="template-card-name">' + escapeHtml(tmpl.name) + '</div>' +
            '<div class="template-card-meta">' + escapeHtml(tmpl.subject || "") + '</div>';
        card.appendChild(info);
        var badge = document.createElement("span");
        badge.className = "template-card-badge";
        badge.textContent = tmpl.builtIn ? t("builtInLabel") : t("customLabel");
        card.appendChild(badge);
        if (!tmpl.builtIn) {
            var del = document.createElement("button");
            del.className = "btn-icon";
            del.title = "Delete";
            del.innerHTML = "&times;";
            del.addEventListener("click", function(e) {
                e.stopPropagation();
                deleteTemplate(idx - builtIn.length);
            });
            card.appendChild(del);
        }
        card.addEventListener("click", function() { loadTemplate(tmpl); });
        container.appendChild(card);
    });
}

function loadTemplate(tmpl) {
    document.getElementById("emailSubject").value = tmpl.subject || "";
    var editor = document.getElementById("emailBody");
    editor.innerHTML = tmpl.body || "";
    editor.classList.remove("is-empty");
}

function showSaveTemplateDialog() {
    document.getElementById("templateNameInput").value = "";
    document.getElementById("templateNameDialog").style.display = "flex";
    document.getElementById("templateNameInput").focus();
}

function saveCurrentTemplate() {
    var name = document.getElementById("templateNameInput").value.trim();
    if (!name) return;
    var templates = getSavedTemplates();
    templates.push({
        name: name,
        subject: document.getElementById("emailSubject").value,
        body: getEditorContent(),
        createdAt: new Date().toISOString(),
        builtIn: false
    });
    saveTemplatesStorage(templates);
    document.getElementById("templateNameDialog").style.display = "none";
    loadTemplatesUI();
}

function deleteTemplate(customIdx) {
    var templates = getSavedTemplates();
    if (customIdx >= 0 && customIdx < templates.length) {
        templates.splice(customIdx, 1);
        saveTemplatesStorage(templates);
        loadTemplatesUI();
    }
}

// ========== Feature 2: Scheduled Sending ==========
function scheduleSend() {
    var dtInput = document.getElementById("scheduleDateTime");
    if (!dtInput.value) return;
    var scheduledDate = new Date(dtInput.value);
    var now = new Date();
    var diff = scheduledDate.getTime() - now.getTime();

    if (diff <= 0) {
        showStatus(t("schedulePast"), "warning");
        executeMerge(false);
        return;
    }

    appState.scheduledTime = scheduledDate;
    // Store in localStorage for interrupted detection
    localStorage.setItem("mailmergepro_scheduled", JSON.stringify({
        time: scheduledDate.toISOString(),
        campaign: document.getElementById("campaignName").value || "Untitled"
    }));

    appState.scheduledTimer = setTimeout(function() {
        localStorage.removeItem("mailmergepro_scheduled");
        clearScheduleUI();
        executeMerge(false);
    }, diff);

    document.getElementById("btnScheduleSend").style.display = "none";
    document.getElementById("btnCancelSchedule").style.display = "inline-block";
    document.getElementById("scheduleCountdown").style.display = "block";
    document.getElementById("scheduleInterruptedMsg").style.display = "none";

    // Countdown timer
    updateScheduleCountdown();
    appState.scheduleCountdownInterval = setInterval(updateScheduleCountdown, 1000);
}

function updateScheduleCountdown() {
    if (!appState.scheduledTime) return;
    var now = new Date();
    var diff = appState.scheduledTime.getTime() - now.getTime();
    if (diff <= 0) {
        document.getElementById("scheduleCountdown").textContent = t("sending") + "...";
        return;
    }
    var hours = Math.floor(diff / 3600000);
    var minutes = Math.floor((diff % 3600000) / 60000);
    var seconds = Math.floor((diff % 60000) / 1000);
    var text = t("scheduleSet") + " ";
    if (hours > 0) text += hours + "h ";
    if (minutes > 0) text += minutes + "m ";
    text += seconds + "s";
    document.getElementById("scheduleCountdown").textContent = text;
}

function cancelScheduledSend() {
    if (appState.scheduledTimer) {
        clearTimeout(appState.scheduledTimer);
        appState.scheduledTimer = null;
    }
    if (appState.scheduleCountdownInterval) {
        clearInterval(appState.scheduleCountdownInterval);
        appState.scheduleCountdownInterval = null;
    }
    appState.scheduledTime = null;
    localStorage.removeItem("mailmergepro_scheduled");
    clearScheduleUI();
}

function clearScheduleUI() {
    document.getElementById("btnScheduleSend").style.display = "inline-block";
    document.getElementById("btnCancelSchedule").style.display = "none";
    document.getElementById("scheduleCountdown").style.display = "none";
}

function checkInterruptedSchedule() {
    var stored = localStorage.getItem("mailmergepro_scheduled");
    if (!stored) return;
    try {
        var data = JSON.parse(stored);
        var scheduledTime = new Date(data.time);
        if (scheduledTime.getTime() < Date.now()) {
            document.getElementById("scheduleInterruptedMsg").style.display = "block";
            localStorage.removeItem("mailmergepro_scheduled");
        } else {
            localStorage.removeItem("mailmergepro_scheduled");
        }
    } catch(e) {
        localStorage.removeItem("mailmergepro_scheduled");
    }
}

// ========== Feature 3: Email Tracking ==========
// Tracking is primarily handled via the existing readReceipt checkbox.
// The chkTracking checkbox auto-enables readReceipt (see initUI listener).

// ========== Feature 4: A/B Testing ==========
function toggleABTest() {
    var enabled = document.getElementById("chkABTest").checked;
    appState.abTestEnabled = enabled;
    document.getElementById("abTestPanel").style.display = enabled ? "block" : "none";
    document.getElementById("emailBodyBContainer").style.display = enabled ? "block" : "none";
    document.getElementById("abEditorLabel").style.display = enabled ? "inline" : "none";
    if (enabled) switchABTab("A");
}

function switchABTab(version) {
    appState.abVersion = version;
    document.querySelectorAll(".ab-tab").forEach(function(tab) {
        tab.classList.toggle("active", tab.dataset.ab === version);
    });
    document.getElementById("abVersionB").style.display = version === "B" ? "block" : "none";
}

function showABResults(sentA, errA, sentB, errB) {
    var el = document.getElementById("resultsContainer");
    var abHtml = '<div style="margin-top:8px;padding:8px;background:var(--primary-light);border-radius:var(--radius-sm);">' +
        '<h3 style="font-size:12px;color:var(--primary);">🔬 A/B Test Results</h3>' +
        '<p style="font-size:11px;"><strong>Version A:</strong> ' + sentA + ' sent, ' + errA + ' errors</p>' +
        '<p style="font-size:11px;"><strong>Version B:</strong> ' + sentB + ' sent, ' + errB + ' errors</p>' +
        '</div>';
    el.innerHTML += abHtml;
}

// ========== Feature 5: Contact Groups / Saved Lists ==========
function getSavedLists() {
    var stored = localStorage.getItem("mailmergepro_contactgroups");
    return stored ? JSON.parse(stored) : [];
}

function saveSavedListsStorage(lists) {
    var json = JSON.stringify(lists);
    if (json.length > 3 * 1024 * 1024) {
        showStatus(t("sizeWarning"), "warning");
    }
    localStorage.setItem("mailmergepro_contactgroups", json);
}

function loadSavedListsUI() {
    var container = document.getElementById("savedListCards");
    if (!container) return;
    container.innerHTML = "";
    var lists = getSavedLists();
    if (lists.length === 0) {
        container.innerHTML = '<p class="hint">' + t("savedListsEmpty") + '</p>';
        return;
    }
    lists.forEach(function(list, idx) {
        var card = document.createElement("div");
        card.className = "saved-list-card";
        card.innerHTML = '<span class="list-name">' + escapeHtml(list.name) + '</span>' +
            '<span class="list-meta">' + list.rows.length + ' rows</span>';
        var del = document.createElement("button");
        del.className = "btn-icon";
        del.innerHTML = "&times;";
        del.title = "Delete";
        del.addEventListener("click", function(e) {
            e.stopPropagation();
            deleteSavedList(idx);
        });
        card.appendChild(del);
        card.addEventListener("click", function(e) {
            if (e.target.closest(".btn-icon")) return;
            loadSavedList(idx);
        });
        container.appendChild(card);
    });
}

function showSaveListDialog() {
    document.getElementById("listNameInput").value = "";
    document.getElementById("listNameDialog").style.display = "flex";
    document.getElementById("listNameInput").focus();
}

function saveCurrentList() {
    var name = document.getElementById("listNameInput").value.trim();
    if (!name) return;
    if (appState.rows.length === 0) return;
    var lists = getSavedLists();
    lists.push({
        name: name,
        headers: appState.headers.slice(),
        rows: appState.rows.slice(),
        createdAt: new Date().toISOString()
    });
    saveSavedListsStorage(lists);
    document.getElementById("listNameDialog").style.display = "none";
    loadSavedListsUI();
}

function loadSavedList(idx) {
    var lists = getSavedLists();
    var list = lists[idx];
    if (!list) return;
    appState.headers = list.headers;
    appState.rows = list.rows;
    renderDataPreview();
    document.getElementById("btnStep1Next").disabled = false;
    document.getElementById("btnSaveCurrentList").disabled = false;
    document.getElementById("btnMergeLists").disabled = false;
    updateStepBadge(1, String(appState.rows.length));
}

function deleteSavedList(idx) {
    var lists = getSavedLists();
    lists.splice(idx, 1);
    saveSavedListsStorage(lists);
    loadSavedListsUI();
}

function showMergeListDialog() {
    var lists = getSavedLists();
    var sel = document.getElementById("mergeListSelect");
    sel.innerHTML = "";
    if (lists.length === 0) {
        sel.innerHTML = '<option value="">No saved lists</option>';
    } else {
        lists.forEach(function(l, i) {
            var opt = document.createElement("option");
            opt.value = i;
            opt.textContent = l.name + " (" + l.rows.length + " rows)";
            sel.appendChild(opt);
        });
    }
    document.getElementById("mergeListDialog").style.display = "flex";
}

function mergeSelectedList() {
    var sel = document.getElementById("mergeListSelect");
    var idx = parseInt(sel.value);
    var lists = getSavedLists();
    var list = lists[idx];
    document.getElementById("mergeListDialog").style.display = "none";
    if (!list) return;
    // Merge headers
    var mergedHeaders = appState.headers.slice();
    list.headers.forEach(function(h) {
        if (mergedHeaders.indexOf(h) === -1) mergedHeaders.push(h);
    });
    // Normalize rows
    var mergedRows = [];
    appState.rows.forEach(function(row) {
        var newRow = {};
        mergedHeaders.forEach(function(h) { newRow[h] = row[h] !== undefined ? row[h] : ""; });
        mergedRows.push(newRow);
    });
    list.rows.forEach(function(row) {
        var newRow = {};
        mergedHeaders.forEach(function(h) { newRow[h] = row[h] !== undefined ? row[h] : ""; });
        mergedRows.push(newRow);
    });
    appState.headers = mergedHeaders;
    appState.rows = mergedRows;
    renderDataPreview();
    updateStepBadge(1, String(appState.rows.length));
}

// ========== Feature 6: HTML Template Import ==========
function handleHtmlImport(e) {
    var file = e.target.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(evt) {
        var htmlContent = evt.target.result;
        importHtmlContent(htmlContent);
    };
    reader.readAsText(file);
    e.target.value = "";
}

function initHtmlDragDrop() {
    var area = document.getElementById("htmlImportArea");
    if (!area) return;
    area.addEventListener("dragover", function(e) {
        e.preventDefault();
        area.classList.add("drag-over");
    });
    area.addEventListener("dragleave", function() {
        area.classList.remove("drag-over");
    });
    area.addEventListener("drop", function(e) {
        e.preventDefault();
        area.classList.remove("drag-over");
        var file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith(".html") || file.name.endsWith(".htm"))) {
            var reader = new FileReader();
            reader.onload = function(evt) { importHtmlContent(evt.target.result); };
            reader.readAsText(file);
        }
    });
}

function importHtmlContent(htmlContent) {
    var editor = document.getElementById("emailBody");
    editor.innerHTML = htmlContent;
    editor.classList.remove("is-empty");
    // Extract merge fields from imported HTML
    var fields = htmlContent.match(/\{([^}]+)\}/g);
    if (fields) {
        console.log("Merge fields found in imported HTML:", fields);
    }
    showStatus(t("htmlImported"), "info");
}

// ========== Feature 7: Signature Auto-Insert ==========
function showSignatureDialog() {
    var textarea = document.getElementById("signatureTextarea");
    textarea.value = appState.signatureHtml || "";
    document.getElementById("chkAutoSignature").checked = appState.autoSignature;
    document.getElementById("signatureDialog").style.display = "flex";
}

async function fetchOutlookSignature() {
    var statusEl = document.getElementById("signatureFetchStatus");
    statusEl.textContent = "Fetching signature...";
    try {
        var token = await getGraphToken();
        // Create a temporary message and read back the body (which includes signature)
        var draft = await graphFetch(GRAPH_BASE + "/me/messages", token, "POST", {
            subject: "__sig_detect__",
            body: { contentType: "HTML", content: "" }
        });
        var msgId = draft.id;
        // Read the draft back
        var msg = await graphFetch(GRAPH_BASE + "/me/messages/" + encodeURIComponent(msgId), token, "GET");
        var sigContent = msg.body ? msg.body.content : "";
        // Delete the temp draft
        await graphFetch(GRAPH_BASE + "/me/messages/" + encodeURIComponent(msgId), token, "DELETE");
        if (sigContent && sigContent.trim().length > 20) {
            document.getElementById("signatureTextarea").value = sigContent;
            statusEl.textContent = "Signature fetched!";
        } else {
            statusEl.textContent = "No signature found. Paste it manually.";
        }
    } catch (err) {
        statusEl.textContent = "Could not fetch: " + (err.message || String(err));
        console.warn("Signature fetch failed:", err);
    }
}

function saveSignature() {
    var content = document.getElementById("signatureTextarea").value;
    var autoAppend = document.getElementById("chkAutoSignature").checked;
    appState.signatureHtml = content;
    appState.autoSignature = autoAppend;
    localStorage.setItem("mailmergepro_signature", content);
    localStorage.setItem("mailmergepro_autosignature", autoAppend ? "true" : "false");
    document.getElementById("signatureDialog").style.display = "none";
    document.getElementById("chkAutoSignatureInline").checked = autoAppend;
    loadSignaturePreview();
    showStatus(t("signatureSaved"), "info");
}

function loadSignaturePreview() {
    var preview = document.getElementById("signaturePreview");
    if (!preview) return;
    if (appState.signatureHtml) {
        preview.innerHTML = appState.signatureHtml;
        preview.style.display = "block";
    } else {
        preview.style.display = "none";
    }
}

// ========== Feature 8: Rate Limit Dashboard ==========
function getDailySentCount() {
    var stored = localStorage.getItem("mailmergepro_dailysent");
    if (!stored) return 0;
    try {
        var data = JSON.parse(stored);
        var today = new Date().toDateString();
        if (data.date === today) return data.count;
        // Different day, reset
        localStorage.removeItem("mailmergepro_dailysent");
        return 0;
    } catch(e) { return 0; }
}

function updateDailySentCount(count) {
    var today = new Date().toDateString();
    var current = getDailySentCount();
    localStorage.setItem("mailmergepro_dailysent", JSON.stringify({
        date: today,
        count: current + count
    }));
}

function updateRateLimitDisplay() {
    var sentToday = getDailySentCount();
    var limit = 10000;
    var pct = Math.min(100, Math.round(sentToday / limit * 100));
    var el = document.getElementById("rateSentToday");
    if (el) el.textContent = sentToday.toLocaleString();
    var fill = document.getElementById("rateLimitFill");
    if (fill) {
        fill.style.width = pct + "%";
        fill.className = "rate-limit-fill " + (pct < 50 ? "green" : pct < 80 ? "yellow" : "red");
    }
}

function suggestDelay(recipientCount) {
    var el = document.getElementById("rateSuggestion");
    if (!el) return;
    if (recipientCount <= 30) {
        el.textContent = t("suggestedDelay") + ": 1s (low volume)";
    } else if (recipientCount <= 500) {
        el.textContent = t("suggestedDelay") + ": 2s";
    } else if (recipientCount <= 2000) {
        el.textContent = t("suggestedDelay") + ": 3-5s";
    } else {
        el.textContent = t("suggestedDelay") + ": 5-10s (high volume)";
    }
}

// ========== Feature 10: Admin Dashboard ==========
function openDashboard() {
    buildDashboard();
    document.getElementById("dashboardOverlay").classList.add("visible");
}

function closeDashboard() {
    document.getElementById("dashboardOverlay").classList.remove("visible");
}

function buildDashboard() {
    var stored = localStorage.getItem("mailmerge-pro-campaigns");
    var campaigns = stored ? JSON.parse(stored) : [];

    // Stats
    var totalCampaigns = campaigns.length;
    var totalEmails = 0;
    var totalSent = 0;
    campaigns.forEach(function(c) { totalEmails += (c.total || 0); totalSent += (c.sent || 0); });
    var successRate = totalEmails > 0 ? Math.round(totalSent / totalEmails * 100) : 0;
    var avgSize = totalCampaigns > 0 ? Math.round(totalEmails / totalCampaigns) : 0;

    var statsHtml = '' +
        '<div class="dash-stat"><span class="dash-value">' + totalCampaigns + '</span><span class="dash-label">' + t("totalCampaigns") + '</span></div>' +
        '<div class="dash-stat"><span class="dash-value">' + totalSent.toLocaleString() + '</span><span class="dash-label">' + t("totalEmails") + '</span></div>' +
        '<div class="dash-stat"><span class="dash-value">' + successRate + '%</span><span class="dash-label">' + t("successRate") + '</span></div>' +
        '<div class="dash-stat"><span class="dash-value">' + avgSize + '</span><span class="dash-label">' + t("avgCampaignSize") + '</span></div>';
    document.getElementById("dashStats").innerHTML = statsHtml;

    // Monthly chart (text-based CSS bars)
    var monthlyData = {};
    campaigns.forEach(function(c) {
        var dateStr = c.dateISO || c.date;
        var d;
        try { d = new Date(dateStr); } catch(e) { return; }
        if (isNaN(d.getTime())) return;
        var key = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
        if (!monthlyData[key]) monthlyData[key] = 0;
        monthlyData[key] += (c.sent || 0);
    });
    var months = Object.keys(monthlyData).sort().slice(-6);
    var maxMonthly = Math.max.apply(null, months.map(function(m) { return monthlyData[m]; }).concat([1]));
    var chartHtml = "";
    months.forEach(function(m) {
        var pct = Math.round(monthlyData[m] / maxMonthly * 100);
        chartHtml += '<div class="dash-bar-row">' +
            '<span class="dash-bar-label">' + m.slice(5) + '</span>' +
            '<div class="dash-bar-track"><div class="dash-bar-fill" style="width:' + pct + '%;"></div></div>' +
            '<span class="dash-bar-value">' + monthlyData[m] + '</span>' +
            '</div>';
    });
    document.getElementById("dashMonthlyChart").innerHTML = chartHtml || '<p class="hint">No data yet.</p>';

    // Top recipients
    var recipientCounts = {};
    campaigns.forEach(function(c) {
        if (c.recipients && Array.isArray(c.recipients)) {
            c.recipients.forEach(function(r) {
                if (!r) return;
                var email = r.toLowerCase();
                recipientCounts[email] = (recipientCounts[email] || 0) + 1;
            });
        }
    });
    var topRecipients = Object.keys(recipientCounts)
        .map(function(e) { return { email: e, count: recipientCounts[e] }; })
        .sort(function(a, b) { return b.count - a.count; })
        .slice(0, 5);
    var topHtml = "";
    topRecipients.forEach(function(r) {
        topHtml += '<div class="dash-campaign-item"><span>' + escapeHtml(r.email) + '</span><span>' + r.count + 'x</span></div>';
    });
    document.getElementById("dashTopRecipients").innerHTML = topHtml || '<p class="hint">No data yet.</p>';

    // Recent campaigns
    var recentHtml = "";
    campaigns.slice(0, 10).forEach(function(c) {
        var rate = c.total > 0 ? Math.round(c.sent / c.total * 100) : 0;
        recentHtml += '<div class="dash-campaign-item">' +
            '<span>' + escapeHtml(c.name) + '</span>' +
            '<span>' + escapeHtml(c.date) + ' — ' + c.sent + '/' + c.total + ' (' + rate + '%)</span>' +
            '</div>';
    });
    document.getElementById("dashRecentCampaigns").innerHTML = recentHtml || '<p class="hint">No campaigns yet.</p>';
}