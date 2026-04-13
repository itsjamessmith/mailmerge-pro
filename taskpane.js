/* MailMerge-Pro — Outlook Web Add-in JavaScript
 * MSAL.js 2.0 + Microsoft Graph API
 * NO alert/confirm/prompt — all feedback via DOM
 */
"use strict";

// ========== Safe Utilities ==========
function safeLocalStorageSet(key, value) {
    try {
        localStorage.setItem(key, value);
        return true;
    } catch (e) {
        console.error("localStorage write failed:", key, e.message);
        showStatus("⚠️ Storage full. Please delete old templates or contact groups to free space.", "warning");
        return false;
    }
}

function safeJsonParse(str, fallback) {
    if (!str) return fallback;
    try { return JSON.parse(str); } catch (e) { console.warn("JSON parse failed, using fallback:", e.message); return fallback; }
}

function sanitizeHtml(html) {
    if (typeof DOMPurify !== "undefined") return DOMPurify.sanitize(html, { USE_PROFILES: { html: true } });
    // Fallback: strip script/event attributes if DOMPurify unavailable
    var tmp = document.createElement("div");
    tmp.innerHTML = html;
    tmp.querySelectorAll("script,iframe,object,embed,form").forEach(function(el) { el.remove(); });
    tmp.querySelectorAll("*").forEach(function(el) {
        Array.from(el.attributes).forEach(function(attr) {
            if (attr.name.startsWith("on")) el.removeAttribute(attr.name);
        });
    });
    return tmp.innerHTML;
}

// ========== Focus Trap Utility (A2) ==========
let _focusTrapCleanup = null;
function trapFocus(modal) {
    const focusable = modal.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
    const first = focusable[0];
    const last = focusable[focusable.length - 1];

    function handleKeydown(e) {
        if (e.key === 'Escape') {
            modal.style.display = 'none';
            releaseFocus();
            return;
        }
        if (e.key !== 'Tab') return;
        if (e.shiftKey) {
            if (document.activeElement === first) { e.preventDefault(); last.focus(); }
        } else {
            if (document.activeElement === last) { e.preventDefault(); first.focus(); }
        }
    }

    modal.addEventListener('keydown', handleKeydown);
    _focusTrapCleanup = () => modal.removeEventListener('keydown', handleKeydown);
    if (first) first.focus();
}
function releaseFocus() {
    if (_focusTrapCleanup) { _focusTrapCleanup(); _focusTrapCleanup = null; }
}

// ========== Rate Limiting Engine ==========
const rateLimiter = {
    maxPerMinute: 30,
    windowMs: 60000,
    timestamps: [],
    canSend: function() {
        var now = Date.now();
        this.timestamps = this.timestamps.filter(function(t) { return now - t < this.windowMs; }.bind(this));
        return this.timestamps.length < this.maxPerMinute;
    },
    waitTime: function() {
        if (this.canSend()) return 0;
        var oldest = this.timestamps[0];
        return Math.max(0, this.windowMs - (Date.now() - oldest) + 100);
    },
    record: function() {
        this.timestamps.push(Date.now());
    },
    waitUntilReady: function() {
        var self = this;
        return new Promise(function(resolve) {
            var wait = self.waitTime();
            if (wait <= 0) { resolve(); return; }
            console.log("Rate limiter: waiting " + Math.round(wait / 1000) + "s");
            setTimeout(resolve, wait);
        });
    }
};

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
    ,
        onboardingTitle:"Welcome to MailMerge-Pro!",
        onboardingDescription:"Upload your spreadsheet to get started with personalized bulk email.",
        onboardingStep1:"Upload Excel/CSV with recipient data",
        onboardingStep2:"Map columns to email fields",
        onboardingStep3:"Compose with merge fields like {FirstName}",
        onboardingStep4:"Send personalized emails to everyone",
        getStartedBtn:"Get Started",
        insertLinkTitle:"Insert Link",
        linkUrlPlaceholder:"https://example.com",
        cancelBtn:"Cancel",
        insertBtn:"Insert",
        campaignDetailsTitle:"Campaign Details",
        closeBtn:"Close",
        saveBtn:"Save",
        mergeBtn:"Merge",
        signaturePlaceholder:"Paste HTML or plain text signature...",
        authInitializing:"Initializing…",
        campaignNamePlaceholder:"Campaign name (optional)",
        historyOption:"📜 History",
        selectContactsTitle:"Select Contacts",
        loadingContacts:"Loading contacts…",
        selectAllBtn:"Select All",
        useSelectedBtn:"Use Selected",
        searchRecipientsPlaceholder:"🔍 Search recipients...",
        nextBtn:"Next →",
        mapColumnsTitle:"Map Columns",
        mapColumnsDescription:"Match your spreadsheet columns to email fields.",
        labelTo:"To (Email)",
        labelCC:"CC",
        optionalPerRow:"(optional per-row)",
        labelBCC:"BCC",
        labelSubject:"Subject",
        labelAttachments:"Attachments",
        columnWithFilenames:"(column with filenames)",
        filenamesSemicolonNote:"Filenames separated by semicolons",
        backBtn:"← Back",
        composeEmailTitle:"Compose Email",
        labelFromSendAs:"From / Send As",
        aliasNote:"(alias)",
        leaveBlankPlaceholder:"Leave blank for default",
        sendAsPermissionNote:"Requires Send As permission for this address",
        labelSharedMailbox:"Shared Mailbox",
        sendOnBehalfNote:"(send on behalf)",
        sharedMailboxPlaceholder:"e.g. shared@company.com",
        sendViaNote:"Send via /users/{mailbox}/sendMail",
        labelSubjectLine:"Subject Line",
        subjectPlaceholder:"e.g. Hello {FirstName}, your update!",
        useColumnNameNote:"Use {ColumnName} for personalization",
        labelSubjectLineB:"Subject Line (B)",
        altSubjectPlaceholder:"Alternative subject for B group",
        labelGlobalCC:"Global CC",
        allEmailsNote:"(all emails)",
        globalCCPlaceholder:"e.g. manager@company.com",
        labelGlobalBCC:"Global BCC",
        globalBCCPlaceholder:"e.g. audit@company.com",
        labelEmailBody:"Email Body",
        versionALabel:"(Version A)",
        tooltipBold:"Bold",
        tooltipItalic:"Italic",
        tooltipUnderline:"Underline",
        tooltipBulletList:"Bullet List",
        tooltipInsertLink:"Insert Link",
        tooltipFontColor:"Font Color",
        labelEmailBodyB:"Email Body (Version B)",
        sendAsHtmlLabel:"Send as HTML",
        manageSignatureBtn:"✍️ Manage Signature",
        fallbackDefaultsTitle:"Fallback Default Values",
        fallbackDefaultsDesc:"Used when a merge field resolves to empty for a recipient.",
        addFilesBtn:"+ Add Files",
        perRecipientAttLabel:"Per-Recipient Attachments",
        perRecipientAttDesc:"Upload files referenced in your Attachments column. Names must match.",
        uploadFilesBtn:"+ Upload Files",
        reviewSendTitle:"Review & Send",
        prevBtn:"◀ Prev",
        nextPreviewBtn:"Next ▶",
        optionsTitleText:"Options",
        labelDelay:"Delay between emails (seconds)",
        saveDraftsLabel:"Save as Drafts only",
        readReceiptLabel:"Request read receipt",
        highImportanceLabel:"High importance",
        groupByEmailLabel:"Group by email (many-to-one)",
        unsubscribeLabel:"Add unsubscribe link",
        unsubscribePlaceholder:"unsubscribe@company.com",
        testBtn:"🧪 Test",
        sendAllBtn:"🚀 Send All Emails",
        sendShortcutHint:"Ctrl+Enter to send",
        conditionalHint:'Conditional: {{#if Column}}...{{/if}}, {{#ifEquals Column "value"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 Suppression List",
        suppressionDesc:"Emails in this list will be automatically skipped during send.",
        suppressionPlaceholder:"Add email to blocklist...",
        addToBlocklist:"+ Add",
        clearBlocklist:"Clear All",
        noBlockedEmails:"No blocked emails",
        validationTitle:"⚠️ Validation Issues",
        sendAnywayBtn:"Send Anyway",
        tooltipInsertImage:"Insert Image"},
    es: {
        appTitle:"📧 MailMerge-Pro", stepData:"Datos", stepMap:"Mapear", stepCompose:"Redactar", stepSend:"Enviar",
        uploadTitle:"📋 Subir Datos", uploadDesc:"Seleccione un archivo Excel (.xlsx) o CSV.",
        chooseFile:"📁 Elegir Archivo",
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
    ,
        onboardingTitle:"¡Bienvenido a MailMerge-Pro!",
        onboardingDescription:"Suba su hoja de cálculo para comenzar con correos personalizados masivos.",
        onboardingStep1:"Suba Excel/CSV con datos de destinatarios",
        onboardingStep2:"Asocie columnas a campos de correo",
        onboardingStep3:"Redacte con campos de combinación como {FirstName}",
        onboardingStep4:"Envíe correos personalizados a todos",
        getStartedBtn:"Comenzar",
        insertLinkTitle:"Insertar Enlace",
        linkUrlPlaceholder:"https://ejemplo.com",
        cancelBtn:"Cancelar",
        insertBtn:"Insertar",
        campaignDetailsTitle:"Detalles de Campaña",
        closeBtn:"Cerrar",
        saveBtn:"Guardar",
        mergeBtn:"Fusionar",
        signaturePlaceholder:"Pegue HTML o firma en texto plano...",
        authInitializing:"Inicializando…",
        campaignNamePlaceholder:"Nombre de campaña (opcional)",
        historyOption:"📜 Historial",
        selectContactsTitle:"Seleccionar Contactos",
        loadingContacts:"Cargando contactos…",
        selectAllBtn:"Seleccionar Todo",
        useSelectedBtn:"Usar Seleccionados",
        searchRecipientsPlaceholder:"🔍 Buscar destinatarios...",
        nextBtn:"Siguiente →",
        mapColumnsTitle:"Mapear Columnas",
        mapColumnsDescription:"Asocie las columnas de su hoja a campos de correo.",
        labelTo:"Para (Email)",
        labelCC:"CC",
        optionalPerRow:"(opcional por fila)",
        labelBCC:"CCO",
        labelSubject:"Asunto",
        labelAttachments:"Adjuntos",
        columnWithFilenames:"(columna con nombres de archivo)",
        filenamesSemicolonNote:"Nombres de archivo separados por punto y coma",
        backBtn:"← Atrás",
        composeEmailTitle:"Redactar Correo",
        labelFromSendAs:"De / Enviar Como",
        aliasNote:"(alias)",
        leaveBlankPlaceholder:"Dejar en blanco para predeterminado",
        sendAsPermissionNote:"Requiere permiso Enviar Como para esta dirección",
        labelSharedMailbox:"Buzón Compartido",
        sendOnBehalfNote:"(enviar en nombre de)",
        sharedMailboxPlaceholder:"ej. compartido@empresa.com",
        sendViaNote:"Enviar vía /users/{mailbox}/sendMail",
        labelSubjectLine:"Línea de Asunto",
        subjectPlaceholder:"ej. Hola {FirstName}, ¡tu actualización!",
        useColumnNameNote:"Use {ColumnName} para personalizar",
        labelSubjectLineB:"Línea de Asunto (B)",
        altSubjectPlaceholder:"Asunto alternativo para grupo B",
        labelGlobalCC:"CC Global",
        allEmailsNote:"(todos los correos)",
        globalCCPlaceholder:"ej. gerente@empresa.com",
        labelGlobalBCC:"CCO Global",
        globalBCCPlaceholder:"ej. auditoria@empresa.com",
        labelEmailBody:"Cuerpo del Correo",
        versionALabel:"(Versión A)",
        tooltipBold:"Negrita",
        tooltipItalic:"Cursiva",
        tooltipUnderline:"Subrayado",
        tooltipBulletList:"Lista con Viñetas",
        tooltipInsertLink:"Insertar Enlace",
        tooltipFontColor:"Color de Fuente",
        labelEmailBodyB:"Cuerpo del Correo (Versión B)",
        sendAsHtmlLabel:"Enviar como HTML",
        manageSignatureBtn:"✍️ Gestionar Firma",
        fallbackDefaultsTitle:"Valores Predeterminados",
        fallbackDefaultsDesc:"Se usan cuando un campo de combinación está vacío para un destinatario.",
        addFilesBtn:"+ Agregar Archivos",
        perRecipientAttLabel:"Adjuntos por Destinatario",
        perRecipientAttDesc:"Suba archivos referenciados en su columna de Adjuntos. Los nombres deben coincidir.",
        uploadFilesBtn:"+ Subir Archivos",
        reviewSendTitle:"Revisar y Enviar",
        prevBtn:"◀ Anterior",
        nextPreviewBtn:"Siguiente ▶",
        optionsTitleText:"Opciones",
        labelDelay:"Retraso entre correos (segundos)",
        saveDraftsLabel:"Guardar solo como borradores",
        readReceiptLabel:"Solicitar acuse de lectura",
        highImportanceLabel:"Alta importancia",
        groupByEmailLabel:"Agrupar por email (muchos a uno)",
        unsubscribeLabel:"Agregar enlace para darse de baja",
        unsubscribePlaceholder:"baja@empresa.com",
        testBtn:"🧪 Prueba",
        sendAllBtn:"🚀 Enviar Todos",
        sendShortcutHint:"Ctrl+Enter para enviar",
        conditionalHint:'Condicional: {{#if Columna}}...{{/if}}, {{#ifEquals Columna "valor"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 Lista de Supresión",
        suppressionDesc:"Los correos en esta lista se omitirán automáticamente durante el envío.",
        suppressionPlaceholder:"Agregar email a lista de bloqueo...",
        addToBlocklist:"+ Agregar",
        clearBlocklist:"Borrar Todo",
        noBlockedEmails:"Sin correos bloqueados",
        validationTitle:"⚠️ Problemas de Validación",
        sendAnywayBtn:"Enviar de Todas Formas",
        tooltipInsertImage:"Insertar Imagen",
        templateNamePlaceholder:"Ingrese nombre de plantilla...",
        listNamePlaceholder:"Ingrese nombre de lista...",
        templateNamePrompt:"Nombre de Plantilla",
        listNamePrompt:"Nombre de Lista",
        mergeSelectPrompt:"Seleccione una lista para fusionar con los datos actuales:"},
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
    ,
        onboardingTitle:"Bienvenue sur MailMerge-Pro !",
        onboardingDescription:"Téléchargez votre feuille de calcul pour commencer les envois personnalisés.",
        onboardingStep1:"Téléchargez un fichier Excel/CSV avec les destinataires",
        onboardingStep2:"Associez les colonnes aux champs email",
        onboardingStep3:"Rédigez avec des champs de fusion comme {FirstName}",
        onboardingStep4:"Envoyez des emails personnalisés à tous",
        getStartedBtn:"Commencer",
        insertLinkTitle:"Insérer un Lien",
        linkUrlPlaceholder:"https://exemple.com",
        cancelBtn:"Annuler",
        insertBtn:"Insérer",
        campaignDetailsTitle:"Détails de la Campagne",
        closeBtn:"Fermer",
        saveBtn:"Sauver",
        mergeBtn:"Fusionner",
        signaturePlaceholder:"Collez du HTML ou une signature en texte brut...",
        authInitializing:"Initialisation…",
        campaignNamePlaceholder:"Nom de campagne (optionnel)",
        historyOption:"📜 Historique",
        selectContactsTitle:"Sélectionner des Contacts",
        loadingContacts:"Chargement des contacts…",
        selectAllBtn:"Tout Sélectionner",
        useSelectedBtn:"Utiliser la Sélection",
        searchRecipientsPlaceholder:"🔍 Rechercher des destinataires...",
        nextBtn:"Suivant →",
        mapColumnsTitle:"Mapper les Colonnes",
        mapColumnsDescription:"Associez les colonnes de votre feuille aux champs email.",
        labelTo:"À (Email)",
        labelCC:"CC",
        optionalPerRow:"(optionnel par ligne)",
        labelBCC:"CCI",
        labelSubject:"Objet",
        labelAttachments:"Pièces jointes",
        columnWithFilenames:"(colonne avec noms de fichiers)",
        filenamesSemicolonNote:"Noms de fichiers séparés par des points-virgules",
        backBtn:"← Retour",
        composeEmailTitle:"Composer l’Email",
        labelFromSendAs:"De / Envoyer en tant que",
        aliasNote:"(alias)",
        leaveBlankPlaceholder:"Laisser vide pour la valeur par défaut",
        sendAsPermissionNote:"Nécessite l’autorisation Envoyer en tant que",
        labelSharedMailbox:"Boîte aux lettres partagée",
        sendOnBehalfNote:"(envoyer au nom de)",
        sharedMailboxPlaceholder:"ex. partage@entreprise.com",
        sendViaNote:"Envoi via /users/{mailbox}/sendMail",
        labelSubjectLine:"Ligne d’Objet",
        subjectPlaceholder:"ex. Bonjour {FirstName}, votre mise à jour !",
        useColumnNameNote:"Utilisez {ColumnName} pour personnaliser",
        labelSubjectLineB:"Ligne d’Objet (B)",
        altSubjectPlaceholder:"Objet alternatif pour le groupe B",
        labelGlobalCC:"CC Global",
        allEmailsNote:"(tous les emails)",
        globalCCPlaceholder:"ex. responsable@entreprise.com",
        labelGlobalBCC:"CCI Global",
        globalBCCPlaceholder:"ex. audit@entreprise.com",
        labelEmailBody:"Corps de l’Email",
        versionALabel:"(Version A)",
        tooltipBold:"Gras",
        tooltipItalic:"Italique",
        tooltipUnderline:"Souligné",
        tooltipBulletList:"Liste à puces",
        tooltipInsertLink:"Insérer un lien",
        tooltipFontColor:"Couleur de police",
        labelEmailBodyB:"Corps de l’Email (Version B)",
        sendAsHtmlLabel:"Envoyer en HTML",
        manageSignatureBtn:"✍️ Gérer la Signature",
        fallbackDefaultsTitle:"Valeurs par Défaut",
        fallbackDefaultsDesc:"Utilisées quand un champ de fusion est vide pour un destinataire.",
        addFilesBtn:"+ Ajouter des Fichiers",
        perRecipientAttLabel:"Pièces jointes par Destinataire",
        perRecipientAttDesc:"Téléchargez les fichiers référencés dans votre colonne Pièces jointes. Les noms doivent correspondre.",
        uploadFilesBtn:"+ Télécharger",
        reviewSendTitle:"Vérifier et Envoyer",
        prevBtn:"◀ Précédent",
        nextPreviewBtn:"Suivant ▶",
        optionsTitleText:"Options",
        labelDelay:"Délai entre les emails (secondes)",
        saveDraftsLabel:"Sauvegarder en brouillons uniquement",
        readReceiptLabel:"Demander un accusé de réception",
        highImportanceLabel:"Haute importance",
        groupByEmailLabel:"Regrouper par email (plusieurs à un)",
        unsubscribeLabel:"Ajouter un lien de désinscription",
        unsubscribePlaceholder:"desinscription@entreprise.com",
        testBtn:"🧪 Test",
        sendAllBtn:"🚀 Tout Envoyer",
        sendShortcutHint:"Ctrl+Entrée pour envoyer",
        conditionalHint:'Conditionnel : {{#if Colonne}}...{{/if}}, {{#ifEquals Colonne "valeur"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 Liste de Suppression",
        suppressionDesc:"Les emails de cette liste seront automatiquement ignorés lors de l'envoi.",
        suppressionPlaceholder:"Ajouter un email à la liste noire...",
        addToBlocklist:"+ Ajouter",
        clearBlocklist:"Tout Effacer",
        noBlockedEmails:"Aucun email bloqué",
        validationTitle:"⚠️ Problèmes de Validation",
        sendAnywayBtn:"Envoyer Quand Même",
        tooltipInsertImage:"Insérer une Image",
        templateNamePlaceholder:"Entrez le nom du modèle...",
        listNamePlaceholder:"Entrez le nom de la liste...",
        templateNamePrompt:"Nom du Modèle",
        listNamePrompt:"Nom de la Liste",
        mergeSelectPrompt:"Sélectionnez une liste à fusionner avec les données actuelles :",
        monthlyActivity:"Activité Mensuelle",
        dragDropHtml:"Déposez un fichier .html ici",
        trackingNote:"Utilise l'accusé de réception Graph API. Pour un suivi avancé, envisagez une mise à niveau.",
        dailyLimit:"Limite quotidienne recommandée : 10 000",
        perMinute:"Recommandé par minute : 30",
        scheduleInterrupted:"La programmation a été interrompue (page actualisée)."},
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
    ,
        onboardingTitle:"Willkommen bei MailMerge-Pro!",
        onboardingDescription:"Laden Sie Ihre Tabelle hoch, um mit personalisierten Massen-E-Mails zu beginnen.",
        onboardingStep1:"Excel/CSV mit Empfängerdaten hochladen",
        onboardingStep2:"Spalten den E-Mail-Feldern zuordnen",
        onboardingStep3:"Mit Seriendruckfeldern wie {FirstName} verfassen",
        onboardingStep4:"Personalisierte E-Mails an alle senden",
        getStartedBtn:"Loslegen",
        insertLinkTitle:"Link Einfügen",
        linkUrlPlaceholder:"https://beispiel.de",
        cancelBtn:"Abbrechen",
        insertBtn:"Einfügen",
        campaignDetailsTitle:"Kampagnendetails",
        closeBtn:"Schließen",
        saveBtn:"Speichern",
        mergeBtn:"Zusammenführen",
        signaturePlaceholder:"HTML oder Nur-Text-Signatur einfügen...",
        authInitializing:"Initialisierung…",
        campaignNamePlaceholder:"Kampagnenname (optional)",
        historyOption:"📜 Verlauf",
        selectContactsTitle:"Kontakte Auswählen",
        loadingContacts:"Kontakte werden geladen…",
        selectAllBtn:"Alle Auswählen",
        useSelectedBtn:"Ausgewählte Verwenden",
        searchRecipientsPlaceholder:"🔍 Empfänger suchen...",
        nextBtn:"Weiter →",
        mapColumnsTitle:"Spalten Zuordnen",
        mapColumnsDescription:"Ordnen Sie die Spalten Ihrer Tabelle den E-Mail-Feldern zu.",
        labelTo:"An (E-Mail)",
        labelCC:"CC",
        optionalPerRow:"(optional pro Zeile)",
        labelBCC:"BCC",
        labelSubject:"Betreff",
        labelAttachments:"Anhänge",
        columnWithFilenames:"(Spalte mit Dateinamen)",
        filenamesSemicolonNote:"Dateinamen durch Semikolons getrennt",
        backBtn:"← Zurück",
        composeEmailTitle:"E-Mail Verfassen",
        labelFromSendAs:"Von / Senden Als",
        aliasNote:"(Alias)",
        leaveBlankPlaceholder:"Leer lassen für Standard",
        sendAsPermissionNote:"Erfordert Senden-Als-Berechtigung für diese Adresse",
        labelSharedMailbox:"Freigegebenes Postfach",
        sendOnBehalfNote:"(im Auftrag senden)",
        sharedMailboxPlaceholder:"z.B. geteilt@firma.com",
        sendViaNote:"Senden über /users/{mailbox}/sendMail",
        labelSubjectLine:"Betreffzeile",
        subjectPlaceholder:"z.B. Hallo {FirstName}, Ihr Update!",
        useColumnNameNote:"Verwenden Sie {ColumnName} zur Personalisierung",
        labelSubjectLineB:"Betreffzeile (B)",
        altSubjectPlaceholder:"Alternativer Betreff für Gruppe B",
        labelGlobalCC:"Globales CC",
        allEmailsNote:"(alle E-Mails)",
        globalCCPlaceholder:"z.B. manager@firma.com",
        labelGlobalBCC:"Globales BCC",
        globalBCCPlaceholder:"z.B. audit@firma.com",
        labelEmailBody:"E-Mail-Text",
        versionALabel:"(Version A)",
        tooltipBold:"Fett",
        tooltipItalic:"Kursiv",
        tooltipUnderline:"Unterstrichen",
        tooltipBulletList:"Aufzählungsliste",
        tooltipInsertLink:"Link einfügen",
        tooltipFontColor:"Schriftfarbe",
        labelEmailBodyB:"E-Mail-Text (Version B)",
        sendAsHtmlLabel:"Als HTML senden",
        manageSignatureBtn:"✍️ Signatur Verwalten",
        fallbackDefaultsTitle:"Standard-Rückfallwerte",
        fallbackDefaultsDesc:"Wird verwendet, wenn ein Seriendruckfeld für einen Empfänger leer ist.",
        addFilesBtn:"+ Dateien Hinzufügen",
        perRecipientAttLabel:"Empfängerspezifische Anhänge",
        perRecipientAttDesc:"Laden Sie Dateien hoch, die in Ihrer Anhänge-Spalte referenziert sind. Namen müssen übereinstimmen.",
        uploadFilesBtn:"+ Dateien Hochladen",
        reviewSendTitle:"Prüfen & Senden",
        prevBtn:"◀ Zurück",
        nextPreviewBtn:"Weiter ▶",
        optionsTitleText:"Optionen",
        labelDelay:"Verzögerung zwischen E-Mails (Sekunden)",
        saveDraftsLabel:"Nur als Entwürfe speichern",
        readReceiptLabel:"Lesebestätigung anfordern",
        highImportanceLabel:"Hohe Priorität",
        groupByEmailLabel:"Nach E-Mail gruppieren (Viele-zu-Eins)",
        unsubscribeLabel:"Abmeldelink hinzufügen",
        unsubscribePlaceholder:"abmelden@firma.com",
        testBtn:"🧪 Test",
        sendAllBtn:"🚀 Alle Senden",
        sendShortcutHint:"Strg+Enter zum Senden",
        conditionalHint:'Bedingt: {{#if Spalte}}...{{/if}}, {{#ifEquals Spalte "Wert"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 Unterdrückungsliste",
        suppressionDesc:"E-Mails in dieser Liste werden beim Senden automatisch übersprungen.",
        suppressionPlaceholder:"E-Mail zur Sperrliste hinzufügen...",
        addToBlocklist:"+ Hinzufügen",
        clearBlocklist:"Alle Löschen",
        noBlockedEmails:"Keine blockierten E-Mails",
        validationTitle:"⚠️ Validierungsprobleme",
        sendAnywayBtn:"Trotzdem Senden",
        tooltipInsertImage:"Bild Einfügen",
        templateNamePlaceholder:"Vorlagenname eingeben...",
        listNamePlaceholder:"Listennname eingeben...",
        templateNamePrompt:"Vorlagenname",
        listNamePrompt:"Listenname",
        mergeSelectPrompt:"Wählen Sie eine Liste zum Zusammenführen mit den aktuellen Daten:",
        monthlyActivity:"Monatliche Aktivität",
        dragDropHtml:".html-Datei hier ablegen",
        trackingNote:"Verwendet Graph API Lesebestätigung. Für erweitertes Tracking erwägen Sie ein Upgrade.",
        dailyLimit:"Empfohlenes Tageslimit: 10.000",
        perMinute:"Empfohlen pro Minute: 30",
        scheduleInterrupted:"Planung wurde unterbrochen (Seite aktualisiert)."},
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
    ,
        onboardingTitle:"Bem-vindo ao MailMerge-Pro!",
        onboardingDescription:"Carregue sua planilha para começar com emails personalizados em massa.",
        onboardingStep1:"Carregue Excel/CSV com dados dos destinatários",
        onboardingStep2:"Mapeie colunas para campos de email",
        onboardingStep3:"Componha com campos de mesclagem como {FirstName}",
        onboardingStep4:"Envie emails personalizados para todos",
        getStartedBtn:"Começar",
        insertLinkTitle:"Inserir Link",
        linkUrlPlaceholder:"https://exemplo.com",
        cancelBtn:"Cancelar",
        insertBtn:"Inserir",
        campaignDetailsTitle:"Detalhes da Campanha",
        closeBtn:"Fechar",
        saveBtn:"Salvar",
        mergeBtn:"Mesclar",
        signaturePlaceholder:"Cole HTML ou assinatura em texto simples...",
        authInitializing:"Inicializando…",
        campaignNamePlaceholder:"Nome da campanha (opcional)",
        historyOption:"📜 Histórico",
        selectContactsTitle:"Selecionar Contatos",
        loadingContacts:"Carregando contatos…",
        selectAllBtn:"Selecionar Todos",
        useSelectedBtn:"Usar Selecionados",
        searchRecipientsPlaceholder:"🔍 Pesquisar destinatários...",
        nextBtn:"Próximo →",
        mapColumnsTitle:"Mapear Colunas",
        mapColumnsDescription:"Associe as colunas da sua planilha aos campos de email.",
        labelTo:"Para (Email)",
        labelCC:"CC",
        optionalPerRow:"(opcional por linha)",
        labelBCC:"CCO",
        labelSubject:"Assunto",
        labelAttachments:"Anexos",
        columnWithFilenames:"(coluna com nomes de arquivos)",
        filenamesSemicolonNote:"Nomes de arquivos separados por ponto e vírgula",
        backBtn:"← Voltar",
        composeEmailTitle:"Compor Email",
        labelFromSendAs:"De / Enviar Como",
        aliasNote:"(alias)",
        leaveBlankPlaceholder:"Deixar em branco para padrão",
        sendAsPermissionNote:"Requer permissão Enviar Como para este endereço",
        labelSharedMailbox:"Caixa de Correio Compartilhada",
        sendOnBehalfNote:"(enviar em nome de)",
        sharedMailboxPlaceholder:"ex. compartilhado@empresa.com",
        sendViaNote:"Enviar via /users/{mailbox}/sendMail",
        labelSubjectLine:"Linha de Assunto",
        subjectPlaceholder:"ex. Olá {FirstName}, sua atualização!",
        useColumnNameNote:"Use {ColumnName} para personalizar",
        labelSubjectLineB:"Linha de Assunto (B)",
        altSubjectPlaceholder:"Assunto alternativo para grupo B",
        labelGlobalCC:"CC Global",
        allEmailsNote:"(todos os emails)",
        globalCCPlaceholder:"ex. gerente@empresa.com",
        labelGlobalBCC:"CCO Global",
        globalBCCPlaceholder:"ex. auditoria@empresa.com",
        labelEmailBody:"Corpo do Email",
        versionALabel:"(Versão A)",
        tooltipBold:"Negrito",
        tooltipItalic:"Itálico",
        tooltipUnderline:"Sublinhado",
        tooltipBulletList:"Lista com Marcadores",
        tooltipInsertLink:"Inserir Link",
        tooltipFontColor:"Cor da Fonte",
        labelEmailBodyB:"Corpo do Email (Versão B)",
        sendAsHtmlLabel:"Enviar como HTML",
        manageSignatureBtn:"✍️ Gerenciar Assinatura",
        fallbackDefaultsTitle:"Valores Padrão",
        fallbackDefaultsDesc:"Usados quando um campo de mesclagem está vazio para um destinatário.",
        addFilesBtn:"+ Adicionar Arquivos",
        perRecipientAttLabel:"Anexos por Destinatário",
        perRecipientAttDesc:"Carregue arquivos referenciados na sua coluna de Anexos. Os nomes devem coincidir.",
        uploadFilesBtn:"+ Enviar Arquivos",
        reviewSendTitle:"Revisar e Enviar",
        prevBtn:"◀ Anterior",
        nextPreviewBtn:"Próximo ▶",
        optionsTitleText:"Opções",
        labelDelay:"Atraso entre emails (segundos)",
        saveDraftsLabel:"Salvar apenas como rascunhos",
        readReceiptLabel:"Solicitar confirmação de leitura",
        highImportanceLabel:"Alta importância",
        groupByEmailLabel:"Agrupar por email (muitos para um)",
        unsubscribeLabel:"Adicionar link de cancelamento",
        unsubscribePlaceholder:"cancelar@empresa.com",
        testBtn:"🧪 Teste",
        sendAllBtn:"🚀 Enviar Todos",
        sendShortcutHint:"Ctrl+Enter para enviar",
        conditionalHint:'Condicional: {{#if Coluna}}...{{/if}}, {{#ifEquals Coluna "valor"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 Lista de Supressão",
        suppressionDesc:"Emails nesta lista serão automaticamente ignorados durante o envio.",
        suppressionPlaceholder:"Adicionar email à lista de bloqueio...",
        addToBlocklist:"+ Adicionar",
        clearBlocklist:"Limpar Tudo",
        noBlockedEmails:"Nenhum email bloqueado",
        validationTitle:"⚠️ Problemas de Validação",
        sendAnywayBtn:"Enviar Mesmo Assim",
        tooltipInsertImage:"Inserir Imagem",
        templateNamePlaceholder:"Digite o nome do modelo...",
        listNamePlaceholder:"Digite o nome da lista...",
        templateNamePrompt:"Nome do Modelo",
        listNamePrompt:"Nome da Lista",
        mergeSelectPrompt:"Selecione uma lista para mesclar com os dados atuais:",
        monthlyActivity:"Atividade Mensal",
        dragDropHtml:"Solte o arquivo .html aqui",
        trackingNote:"Usa confirmação de leitura da Graph API. Para rastreamento avançado, considere atualizar.",
        dailyLimit:"Limite diário recomendado: 10.000",
        perMinute:"Recomendado por minuto: 30",
        scheduleInterrupted:"O agendamento foi interrompido (página atualizada)."},
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
    ,
        onboardingTitle:"MailMerge-Proへようこそ！",
        onboardingDescription:"スプレッドシートをアップロードして、パーソナライズメールを始めましょう。",
        onboardingStep1:"受信者データ入りのExcel/CSVをアップロード",
        onboardingStep2:"列をメールフィールドにマッピング",
        onboardingStep3:"{FirstName}などの差し込みフィールドで作成",
        onboardingStep4:"全員にパーソナライズメールを送信",
        getStartedBtn:"始める",
        insertLinkTitle:"リンクを挿入",
        linkUrlPlaceholder:"https://example.com",
        cancelBtn:"キャンセル",
        insertBtn:"挿入",
        campaignDetailsTitle:"キャンペーン詳細",
        closeBtn:"閉じる",
        saveBtn:"保存",
        mergeBtn:"結合",
        signaturePlaceholder:"HTMLまたはテキスト署名を貼り付け...",
        authInitializing:"初期化中…",
        campaignNamePlaceholder:"キャンペーン名（任意）",
        historyOption:"📜 履歴",
        selectContactsTitle:"連絡先を選択",
        loadingContacts:"連絡先を読み込み中…",
        selectAllBtn:"すべて選択",
        useSelectedBtn:"選択したものを使用",
        searchRecipientsPlaceholder:"🔍 受信者を検索...",
        nextBtn:"次へ →",
        mapColumnsTitle:"列マッピング",
        mapColumnsDescription:"スプレッドシートの列をメールフィールドに関連付けます。",
        labelTo:"宛先（メール）",
        labelCC:"CC",
        optionalPerRow:"（行ごとに任意）",
        labelBCC:"BCC",
        labelSubject:"件名",
        labelAttachments:"添付ファイル",
        columnWithFilenames:"（ファイル名の列）",
        filenamesSemicolonNote:"ファイル名はセミコロンで区切り",
        backBtn:"← 戻る",
        composeEmailTitle:"メール作成",
        labelFromSendAs:"差出人 / 送信者",
        aliasNote:"（エイリアス）",
        leaveBlankPlaceholder:"デフォルトの場合は空欄",
        sendAsPermissionNote:"このアドレスの「送信者として送信」権限が必要",
        labelSharedMailbox:"共有メールボックス",
        sendOnBehalfNote:"（代理送信）",
        sharedMailboxPlaceholder:"例: shared@company.com",
        sendViaNote:"/users/{mailbox}/sendMail 経由で送信",
        labelSubjectLine:"件名",
        subjectPlaceholder:"例: {FirstName}さん、更新情報です！",
        useColumnNameNote:"{ColumnName}でパーソナライズ",
        labelSubjectLineB:"件名 (B)",
        altSubjectPlaceholder:"Bグループの代替件名",
        labelGlobalCC:"グローバルCC",
        allEmailsNote:"（全メール）",
        globalCCPlaceholder:"例: manager@company.com",
        labelGlobalBCC:"グローバルBCC",
        globalBCCPlaceholder:"例: audit@company.com",
        labelEmailBody:"メール本文",
        versionALabel:"（バージョンA）",
        tooltipBold:"太字",
        tooltipItalic:"斜体",
        tooltipUnderline:"下線",
        tooltipBulletList:"箇条書き",
        tooltipInsertLink:"リンク挿入",
        tooltipFontColor:"フォント色",
        labelEmailBodyB:"メール本文 (バージョンB)",
        sendAsHtmlLabel:"HTMLで送信",
        manageSignatureBtn:"✍️ 署名管理",
        fallbackDefaultsTitle:"フォールバックデフォルト値",
        fallbackDefaultsDesc:"受信者の差し込みフィールドが空の場合に使用されます。",
        addFilesBtn:"+ ファイル追加",
        perRecipientAttLabel:"受信者別添付ファイル",
        perRecipientAttDesc:"添付ファイル列で参照されているファイルをアップロード。名前が一致する必要があります。",
        uploadFilesBtn:"+ ファイルアップロード",
        reviewSendTitle:"確認＆送信",
        prevBtn:"◀ 前",
        nextPreviewBtn:"次 ▶",
        optionsTitleText:"オプション",
        labelDelay:"メール間の遅延（秒）",
        saveDraftsLabel:"下書きとしてのみ保存",
        readReceiptLabel:"開封確認を要求",
        highImportanceLabel:"重要度高",
        groupByEmailLabel:"メールでグループ化（多対一）",
        unsubscribeLabel:"配信停止リンクを追加",
        unsubscribePlaceholder:"unsubscribe@company.com",
        testBtn:"🧪 テスト",
        sendAllBtn:"🚀 全て送信",
        sendShortcutHint:"Ctrl+Enterで送信",
        conditionalHint:'条件付き: {{#if 列名}}...{{/if}}, {{#ifEquals 列名 "値"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 抑制リスト",
        suppressionDesc:"このリストのメールアドレスは送信時に自動的にスキップされます。",
        suppressionPlaceholder:"ブロックリストにメールを追加...",
        addToBlocklist:"+ 追加",
        clearBlocklist:"すべてクリア",
        noBlockedEmails:"ブロックされたメールはありません",
        validationTitle:"⚠️ バリデーション問題",
        sendAnywayBtn:"それでも送信",
        tooltipInsertImage:"画像を挿入",
        templateNamePlaceholder:"テンプレート名を入力...",
        listNamePlaceholder:"リスト名を入力...",
        templateNamePrompt:"テンプレート名",
        listNamePrompt:"リスト名",
        mergeSelectPrompt:"現在のデータと結合するリストを選択してください：",
        monthlyActivity:"月間アクティビティ",
        dragDropHtml:".htmlファイルをここにドロップ",
        trackingNote:"Graph API開封確認を使用。高度なトラッキングにはホスト版へのアップグレードをご検討ください。",
        dailyLimit:"推奨日次制限：10,000",
        perMinute:"推奨1分あたり：30",
        scheduleInterrupted:"スケジュールが中断されました（ページが更新されました）。"},
    zh: {
        appTitle:"📧 MailMerge-Pro",
        stepData:"数据",
        stepMap:"映射",
        stepCompose:"撰写",
        stepSend:"发送",
        uploadTitle:"📋 上传收件人数据",
        uploadDesc:"选择 Excel (.xlsx) 或 CSV 文件。",
        chooseFile:"📁 选择文件",
        noFileSelected:"未选择文件",
        orDivider:"— 或 —",
        importContacts:"👤 导入联系人",
        savedLists:"📑 已保存列表",
        saveCurrentList:"💾 保存当前列表",
        mergeLists:"🔗 合并列表",
        savedListsEmpty:"没有已保存的列表。",
        sizeWarning:"⚠️ 数据超过3MB。",
        mapTitle:"🔗 列映射",
        mapDesc:"将列与邮件字段关联。",
        toLabel:"收件人（邮箱）",
        ccLabel:"抄送",
        bccLabel:"密送",
        subjectLabel:"主题",
        attachmentsLabel:"附件",
        composeTitle:"✏️ 撰写邮件",
        fromLabel:"发件人 / 代发",
        sharedMailboxLabel:"共享邮箱",
        subjectLineLabel:"主题行",
        globalCCLabel:"全局抄送",
        globalBCCLabel:"全局密送",
        emailBodyLabel:"邮件正文",
        sendAsHtml:"以HTML发送",
        templatesTitle:"📝 模板",
        saveAsTemplate:"💾 保存为模板",
        abTestEnable:"🔬 A/B测试",
        versionA:"版本A",
        versionB:"版本B",
        splitRatio:"分割比例",
        importHtml:"📄 导入HTML",
        insertSignature:"✍️ 签名",
        signatureTitle:"✍️ 签名",
        autoAppendSig:"自动添加签名",
        fetchSignature:"从Outlook获取",
        pasteManually:"或粘贴您的签名：",
        signatureSaved:"签名已保存！",
        fallbackDefaults:"🔄 回退默认值",
        fallbackHint:"当收件人的合并字段为空时使用。",
        globalAttachments:"📎 全局附件",
        perRecipientAttachments:"📎 每个收件人的附件",
        reviewTitle:"🚀 审查并发送",
        optionsTitle:"⚙️ 选项",
        delayLabel:"邮件间延迟（秒）",
        draftOnly:"📝 仅保存为草稿",
        readReceipt:"📬 请求已读回执",
        highImportance:"❗ 高重要性",
        groupByEmail:"📊 按邮件分组",
        unsubscribe:"🚫 添加取消订阅链接",
        addTracking:"添加阅读跟踪",
        trackingNote:"使用Graph API已读回执。如需高级跟踪（打开率、点击率），请考虑升级到托管版本。",
        scheduleTitle:"⏰ 定时发送",
        scheduleBtn:"定时",
        cancelSchedule:"取消定时",
        scheduleWarning:"⚠️ Outlook必须保持打开以便定时发送。",
        schedulePast:"定时时间已过——立即发送。",
        scheduleSet:"发送倒计时",
        scheduleInterrupted:"定时已中断（页面已刷新）。",
        rateLimitTitle:"📊 发送限制",
        sentToday:"今日发送数",
        dailyLimit:"建议每日限制：10,000",
        perMinute:"建议每分钟：30",
        suggestedDelay:"建议延迟",
        sendAll:"🚀 发送所有邮件",
        test:"🧪 测试",
        back:"← 返回",
        next:"下一步 →",
        prev:"◀ 上一个",
        nextPreview:"下一个 ▶",
        ctrlEnterHint:"Ctrl+Enter 发送",
        dashboard:"📊 仪表盘",
        dashboardTitle:"📊 管理仪表盘",
        totalCampaigns:"总活动数",
        totalEmails:"总发送邮件数",
        successRate:"成功率",
        avgCampaignSize:"平均活动规模",
        topRecipients:"热门收件人",
        monthlyActivity:"月度活动",
        recentCampaigns:"最近活动",
        closeDashboard:"关闭",
        signIn:"登录",
        signOut:"退出",
        getStarted:"开始使用",
        welcome:"👋 欢迎使用 MailMerge-Pro！",
        welcomeDesc:"上传您的电子表格，开始发送个性化批量邮件。",
        recipients:"收件人",
        columns:"列",
        required:"*",
        optional:"（可选）",
        none:"（无）",
        history:"📜 历史记录",
        campaignDetails:"📜 活动详情",
        close:"关闭",
        insertLink:"插入链接",
        cancel:"取消",
        insert:"插入",
        searchRecipients:"🔍 搜索收件人...",
        selectContacts:"👤 选择联系人",
        selectAll:"全选",
        useSelected:"使用已选",
        addFiles:"+ 添加文件",
        uploadFiles:"+ 上传文件",
        templateNamePrompt:"模板名称",
        templateNamePlaceholder:"输入模板名称...",
        save:"保存",
        listNamePrompt:"列表名称",
        listNamePlaceholder:"输入列表名称...",
        mergeSelectPrompt:"选择要与当前数据合并的列表：",
        merge:"合并",
        noRecipients:"没有收件人。",
        enterSubject:"⚠️ 请输入主题。",
        enterBody:"⚠️ 请输入邮件正文。",
        uploadFirst:"⚠️ 请先上传数据文件或导入联系人。",
        selectToColumn:"⚠️ 请选择收件人（邮箱）列。",
        authFailed:"❌ 认证失败",
        sendComplete:"邮件合并完成",
        error:"错误",
        sent:"已发送",
        draft:"草稿",
        dragDropHtml:"拖放 .html 文件或点击导入HTML",
        htmlImported:"HTML模板已导入！",
        scheduledFor:"定时于",
        sending:"发送中",
        notSignedIn:"未登录",
        onboardingTitle:"欢迎使用 MailMerge-Pro！",
        onboardingDescription:"上传您的电子表格，开始发送个性化批量邮件。",
        onboardingStep1:"上传包含收件人数据的 Excel/CSV",
        onboardingStep2:"将列映射到邮件字段",
        onboardingStep3:"使用 {FirstName} 等合并字段撰写",
        onboardingStep4:"向每个人发送个性化邮件",
        getStartedBtn:"开始使用",
        insertLinkTitle:"插入链接",
        linkUrlPlaceholder:"https://example.com",
        cancelBtn:"取消",
        insertBtn:"插入",
        campaignDetailsTitle:"活动详情",
        closeBtn:"关闭",
        saveBtn:"保存",
        mergeBtn:"合并",
        signaturePlaceholder:"粘贴 HTML 或纯文本签名...",
        authInitializing:"初始化中…",
        campaignNamePlaceholder:"活动名称（可选）",
        historyOption:"📜 历史记录",
        selectContactsTitle:"选择联系人",
        loadingContacts:"正在加载联系人…",
        selectAllBtn:"全选",
        useSelectedBtn:"使用已选",
        searchRecipientsPlaceholder:"🔍 搜索收件人...",
        nextBtn:"下一步 →",
        mapColumnsTitle:"列映射",
        mapColumnsDescription:"将电子表格列与邮件字段匹配。",
        labelTo:"收件人（邮箱）",
        labelCC:"抄送",
        optionalPerRow:"（每行可选）",
        labelBCC:"密送",
        labelSubject:"主题",
        labelAttachments:"附件",
        columnWithFilenames:"（文件名列）",
        filenamesSemicolonNote:"文件名用分号分隔",
        backBtn:"← 返回",
        composeEmailTitle:"撰写邮件",
        labelFromSendAs:"发件人 / 代发",
        aliasNote:"（别名）",
        leaveBlankPlaceholder:"留空使用默认值",
        sendAsPermissionNote:"需要此地址的“代发”权限",
        labelSharedMailbox:"共享邮箱",
        sendOnBehalfNote:"（代表发送）",
        sharedMailboxPlaceholder:"例如 shared@company.com",
        sendViaNote:"通过 /users/{mailbox}/sendMail 发送",
        labelSubjectLine:"主题行",
        subjectPlaceholder:"例如 你好 {FirstName}，您的更新！",
        useColumnNameNote:"使用 {ColumnName} 进行个性化",
        labelSubjectLineB:"主题行 (B)",
        altSubjectPlaceholder:"B 组的备选主题",
        labelGlobalCC:"全局抄送",
        allEmailsNote:"（所有邮件）",
        globalCCPlaceholder:"例如 manager@company.com",
        labelGlobalBCC:"全局密送",
        globalBCCPlaceholder:"例如 audit@company.com",
        labelEmailBody:"邮件正文",
        versionALabel:"（版本A）",
        tooltipBold:"加粗",
        tooltipItalic:"斜体",
        tooltipUnderline:"下划线",
        tooltipBulletList:"项目符号列表",
        tooltipInsertLink:"插入链接",
        tooltipFontColor:"字体颜色",
        labelEmailBodyB:"邮件正文（版本B）",
        sendAsHtmlLabel:"以 HTML 发送",
        manageSignatureBtn:"✍️ 管理签名",
        fallbackDefaultsTitle:"回退默认值",
        fallbackDefaultsDesc:"当收件人的合并字段为空时使用。",
        addFilesBtn:"+ 添加文件",
        perRecipientAttLabel:"每个收件人的附件",
        perRecipientAttDesc:"上传附件列中引用的文件。名称必须匹配。",
        uploadFilesBtn:"+ 上传文件",
        reviewSendTitle:"审查并发送",
        prevBtn:"◀ 上一个",
        nextPreviewBtn:"下一个 ▶",
        optionsTitleText:"选项",
        labelDelay:"邮件间延迟（秒）",
        saveDraftsLabel:"仅保存为草稿",
        readReceiptLabel:"请求已读回执",
        highImportanceLabel:"高重要性",
        groupByEmailLabel:"按邮件分组（多对一）",
        unsubscribeLabel:"添加取消订阅链接",
        unsubscribePlaceholder:"unsubscribe@company.com",
        testBtn:"🧪 测试",
        sendAllBtn:"🚀 发送所有邮件",
        sendShortcutHint:"Ctrl+Enter 发送",
        conditionalHint:'条件内容: {{#if 列名}}...{{/if}}, {{#ifEquals 列名 "值"}}...{{/ifEquals}}',
        suppressionTitle:"🚫 抑制列表",
        suppressionDesc:"此列表中的邮件在发送时将自动跳过。",
        suppressionPlaceholder:"添加邮箱到黑名单...",
        addToBlocklist:"+ 添加",
        clearBlocklist:"清除全部",
        noBlockedEmails:"没有被屏蔽的邮箱",
        validationTitle:"⚠️ 验证问题",
        sendAnywayBtn:"仍然发送",
        tooltipInsertImage:"插入图片"
    }
};

function t(key) {
    return (translations[currentLang] && translations[currentLang][key]) || translations.en[key] || key;
}

function setLanguage(lang) {
    currentLang = lang;
    safeLocalStorageSet("mailmergepro_language", lang);
    applyTranslations();
}

function applyTranslations() {
    document.querySelectorAll("[data-i18n]").forEach(function(el) {
        var key = el.getAttribute("data-i18n");
        var val = t(key);
        if (val) el.textContent = val;
    });
    document.querySelectorAll("[data-i18n-placeholder]").forEach(function(el) {
        var key = el.getAttribute("data-i18n-placeholder");
        var val = t(key);
        if (val) el.placeholder = val;
    });
    document.querySelectorAll("[data-i18n-title]").forEach(function(el) {
        var key = el.getAttribute("data-i18n-title");
        var val = t(key);
        if (val) el.title = val;
    });
}

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
    filteredRows: null,
    previewPage: 0,
    previewPageSize: 50,
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
    // Delayed auth check — fetch email from Graph API if MSAL account is incomplete
    setTimeout(async () => {
        if (!appState.userEmail && msalInstance) {
            try {
                const token = await getGraphToken();
                if (token) {
                    const resp = await fetch(GRAPH_BASE + "/me?$select=mail,userPrincipalName,displayName", {
                        headers: { "Authorization": "Bearer " + token }
                    });
                    if (resp.ok) {
                        const me = await resp.json();
                        const email = me.mail || me.userPrincipalName || me.displayName || "Signed in";
                        console.log("Graph /me resolved:", email);
                        appState.userEmail = email;
                        const authEl = document.getElementById("authStatus");
                        authEl.textContent = "✅ " + email;
                        authEl.classList.add("signed-in");
                        document.getElementById("btnSignIn").style.display = "none";
                        document.getElementById("btnSignOut").style.display = "inline-block";
                    }
                }
            } catch (e) { console.warn("Delayed Graph /me check failed:", e.message); }
        }
    }, 2000);
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
    safeLocalStorageSet("mailmergepro_darkmode", String(isDark));
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
    safeLocalStorageSet("mailmerge-pro-onboarded", "true");
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
        
        // handleRedirectPromise with timeout — can hang after redirect in Outlook
        try {
            const redirectPromise = msalInstance.handleRedirectPromise();
            const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), 5000));
            const resp = await Promise.race([redirectPromise, timeoutPromise]);
            if (resp && resp.account) {
                console.log("handleRedirectPromise: signed in via redirect:", resp.account.username);
                updateAuthUI(resp.account);
                return;
            }
        } catch (e) {
            console.warn("handleRedirectPromise error/timeout:", e.message || e);
        }
        
        // Check if already signed in (from cached session)
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            console.log("Already signed in:", accounts[0].username);
            updateAuthUI(accounts[0]);
        } else {
            authEl.textContent = t("notSignedIn");
            authEl.classList.remove("error");
            btnIn.style.display = "inline-block";
            btnOut.style.display = "none";
        }
    } catch (err) {
        console.error("MSAL init error:", err);
        authEl.textContent = t("notSignedIn");
        authEl.classList.remove("error");
        btnIn.style.display = "inline-block";
        btnOut.style.display = "none";
    }
}

function updateAuthUI(account) {
    const authEl = document.getElementById("authStatus");
    const btnIn = document.getElementById("btnSignIn");
    const btnOut = document.getElementById("btnSignOut");
    if (account) {
        // MSAL account may have username, name, or localAccountId — try all
        const displayName = account.username || account.name || account.localAccountId || account.homeAccountId || "Signed in";
        console.log("updateAuthUI: signed in as", displayName, "| account keys:", Object.keys(account).join(","));
        authEl.textContent = "✅ " + displayName;
        authEl.classList.add("signed-in");
        authEl.classList.remove("error");
        btnIn.style.display = "none";
        btnOut.style.display = "inline-block";
        appState.userEmail = account.username || account.name || "";
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
        // Clear any stuck MSAL interaction state
        try {
            const keys = Object.keys(sessionStorage);
            keys.forEach(k => { if (k.indexOf("interaction") !== -1) sessionStorage.removeItem(k); });
        } catch (_) {}

        // Try silent token first if there's an active account
        const activeAccount = msalInstance.getActiveAccount();
        if (activeAccount) {
            try {
                const result = await msalInstance.acquireTokenSilent({ ...loginRequest, account: activeAccount });
                updateAuthUI(activeAccount);
                return result.accessToken;
            } catch (_) { /* silent failed, proceed to interactive */ }
        }

        // Detect if running inside Outlook task pane (iframe/popup context)
        const isInIframe = (window !== window.parent) || (typeof Office !== "undefined" && Office.context);
        
        if (isInIframe) {
            // Use redirect flow — popups are blocked inside Outlook task pane
            console.log("signIn: detected task pane/iframe, using redirect flow");
            authEl.textContent = "Redirecting to sign in...";
            authEl.classList.remove("error");
            await msalInstance.acquireTokenRedirect(loginRequest);
            return; // Page will redirect — execution stops here
        } else {
            // Use popup flow — works in standalone browser
            const result = await msalInstance.acquireTokenPopup(loginRequest);
            console.log("signIn success (popup):", result.account.username);
            updateAuthUI(result.account);
            return result.accessToken;
        }
    } catch (err) {
        console.error("signIn error:", err);
        try {
            const keys = Object.keys(sessionStorage);
            keys.forEach(k => { if (k.indexOf("interaction") !== -1) sessionStorage.removeItem(k); });
        } catch (_) {}
        const msg = err.message || String(err);
        if (msg.includes("interaction_in_progress")) {
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
    document.getElementById("btnLinkCancel").addEventListener("click", () => { document.getElementById("linkDialog").style.display = "none"; releaseFocus(); });
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
        releaseFocus();
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
    document.getElementById("btnTemplateNameCancel").addEventListener("click", () => { document.getElementById("templateNameDialog").style.display = "none"; releaseFocus(); });
    document.getElementById("btnTemplateNameSave").addEventListener("click", saveCurrentTemplate);
    document.getElementById("templateNameInput").addEventListener("keydown", (e) => { if (e.key === "Enter") saveCurrentTemplate(); });
    loadTemplatesUI();

    // Feature 5: Contact Groups / Saved Lists
    document.getElementById("savedListsHeader").addEventListener("click", () => toggleCollapsible("savedListsHeader", "savedListsBody"));
    document.getElementById("btnSaveCurrentList").addEventListener("click", showSaveListDialog);
    document.getElementById("btnMergeLists").addEventListener("click", showMergeListDialog);
    document.getElementById("btnListNameCancel").addEventListener("click", () => { document.getElementById("listNameDialog").style.display = "none"; releaseFocus(); });
    document.getElementById("btnListNameSave").addEventListener("click", saveCurrentList);
    document.getElementById("listNameInput").addEventListener("keydown", (e) => { if (e.key === "Enter") saveCurrentList(); });
    document.getElementById("btnMergeListCancel").addEventListener("click", () => { document.getElementById("mergeListDialog").style.display = "none"; releaseFocus(); });
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
    document.getElementById("btnSignatureCancel").addEventListener("click", () => { document.getElementById("signatureDialog").style.display = "none"; releaseFocus(); });
    document.getElementById("btnSignatureSave").addEventListener("click", saveSignature);
    document.getElementById("chkAutoSignatureInline").checked = appState.autoSignature;
    document.getElementById("chkAutoSignatureInline").addEventListener("change", (e) => {
        appState.autoSignature = e.target.checked;
        safeLocalStorageSet("mailmergepro_autosignature", e.target.checked ? "true" : "false");
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

    // A3: Keyboard navigation for collapsible headers
    document.querySelectorAll(".collapsible-header[tabindex]").forEach(header => {
        header.addEventListener("keydown", (e) => {
            if (e.key === "Enter" || e.key === " ") { e.preventDefault(); header.click(); }
        });
    });
    // A3: Keyboard navigation for A/B tabs
    document.querySelectorAll(".ab-tab").forEach(tab => {
        tab.addEventListener("keydown", (e) => {
            if (e.key === "Enter" || e.key === " ") { e.preventDefault(); tab.click(); }
        });
    });

    // A5: Event delegation for attachment lists and template/list cards
    document.getElementById("globalAttachmentList").addEventListener("click", (e) => {
        const removeBtn = e.target.closest(".att-remove");
        if (removeBtn) {
            const name = removeBtn.dataset.name;
            if (name) { appState.globalAttachments.delete(name); renderGlobalAttachmentList(); }
        }
    });
    document.getElementById("perRecipientAttachmentList").addEventListener("click", (e) => {
        const removeBtn = e.target.closest(".att-remove");
        if (removeBtn) {
            const key = removeBtn.dataset.key;
            if (key) { appState.perRecipientFiles.delete(key); renderPerRecipientAttachmentList(); checkMissingAttachments(); }
        }
    });
    document.getElementById("templateCards").addEventListener("click", (e) => {
        const delBtn = e.target.closest(".tmpl-delete");
        if (delBtn) {
            e.stopPropagation();
            const idx = parseInt(delBtn.dataset.idx);
            if (!isNaN(idx)) deleteTemplate(idx);
            return;
        }
        const card = e.target.closest(".template-card");
        if (card && card.dataset.tmplIdx) {
            const builtIn = getBuiltInTemplates();
            const custom = getSavedTemplates();
            const all = builtIn.concat(custom);
            const tmpl = all[parseInt(card.dataset.tmplIdx)];
            if (tmpl) loadTemplate(tmpl);
        }
    });
    document.getElementById("savedListCards").addEventListener("click", (e) => {
        const delBtn = e.target.closest(".list-delete");
        if (delBtn) {
            e.stopPropagation();
            const idx = parseInt(delBtn.dataset.idx);
            if (!isNaN(idx)) deleteSavedList(idx);
            return;
        }
        const card = e.target.closest(".saved-list-card");
        if (card && card.dataset.listIdx) {
            loadSavedList(parseInt(card.dataset.listIdx));
        }
    });

    // B3: Suppression list
    document.getElementById("suppressionHeader").addEventListener("click", () => {
        const content = document.getElementById("suppressionContent");
        content.style.display = content.style.display === "none" ? "block" : "none";
    });
    document.getElementById("suppressionHeader").addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); document.getElementById("suppressionHeader").click(); }
    });
    document.getElementById("btnAddSuppression").addEventListener("click", () => {
        const input = document.getElementById("suppressionInput");
        const email = input.value.trim();
        if (email) { addToSuppressionList(email); input.value = ""; }
    });
    document.getElementById("suppressionInput").addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); document.getElementById("btnAddSuppression").click(); }
    });
    document.getElementById("btnClearSuppression").addEventListener("click", () => {
        localStorage.removeItem("mmPro_suppressionList");
        renderSuppressionList();
    });
    document.getElementById("suppressionList").addEventListener("click", (e) => {
        const removeBtn = e.target.closest(".att-remove");
        if (removeBtn) {
            const email = removeBtn.dataset.email;
            if (email) removeFromSuppressionList(email);
        }
    });
    renderSuppressionList();

    // B2: Validation modal
    document.getElementById("btnValidationCancel").addEventListener("click", () => {
        document.getElementById("validationResultsModal").style.display = "none";
        releaseFocus();
    });

    // B4: Inline image insert
    document.getElementById("btnInsertImage").addEventListener("click", () => {
        document.getElementById("inlineImageInput").click();
    });
    document.getElementById("inlineImageInput").addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file || !file.type.startsWith("image/")) return;
        if (file.size > 2 * 1024 * 1024) {
            showStatus("Image must be under 2MB", "error");
            return;
        }
        const reader = new FileReader();
        reader.onload = () => {
            const img = '<img src="' + reader.result + '" alt="' + sanitizeHtml(file.name) + '" style="max-width:100%;height:auto;"/>';
            document.getElementById("emailBody").focus();
            document.execCommand("insertHTML", false, img);
        };
        reader.readAsDataURL(file);
        e.target.value = "";
    });

    // Dark mode toggle (bound programmatically for CSP compliance)
    document.getElementById("darkModeToggle").addEventListener("click", toggleDarkMode);

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
    const modal = document.getElementById("linkDialog");
    modal.style.display = "flex";
    trapFocus(modal);
}
function insertLink() {
    const url = document.getElementById("linkUrlInput").value.trim();
    document.getElementById("linkDialog").style.display = "none";
    releaseFocus();
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
    const campaigns = safeJsonParse(stored, []);
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
    const campaigns = safeJsonParse(stored, []);
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
    safeLocalStorageSet("mailmerge-pro-campaigns", JSON.stringify(campaigns));
    loadCampaignHistory();
}
function showCampaignDetails() {
    const sel = document.getElementById("pastCampaigns");
    const idx = sel.value;
    if (idx === "") return;
    const stored = localStorage.getItem("mailmerge-pro-campaigns");
    const campaigns = safeJsonParse(stored, []);
    const c = campaigns[parseInt(idx)];
    if (!c) return;
    document.getElementById("campaignDetailContent").innerHTML =
        "<p><strong>Name:</strong> " + escapeHtml(c.name) + "</p>" +
        "<p><strong>Date:</strong> " + escapeHtml(c.date) + "</p>" +
        "<p><strong>Total:</strong> " + c.total + "</p>" +
        "<p><strong>Sent:</strong> " + c.sent + "</p>" +
        "<p><strong>Errors:</strong> " + c.errors + "</p>";
    const campaignModal = document.getElementById("campaignDetailModal");
    campaignModal.style.display = "flex";
    trapFocus(campaignModal);
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
            appState.filteredRows = null;
            appState.previewPage = 0;
            document.getElementById("recipientSearch").value = "";
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
    const sourceRows = appState.filteredRows || appState.rows;
    const totalRows = sourceRows.length;
    const pageSize = appState.previewPageSize;
    const maxPage = Math.max(0, Math.ceil(totalRows / pageSize) - 1);
    if (appState.previewPage > maxPage) appState.previewPage = maxPage;
    const page = appState.previewPage;
    const start = page * pageSize;
    const end = Math.min(start + pageSize, totalRows);
    const pageRows = sourceRows.slice(start, end);

    document.getElementById("dataStats").textContent =
        "\u2705 " + appState.rows.length + " recipients, " + appState.headers.length + " columns: " + appState.headers.join(", ");
    let t = "<table><tr>";
    appState.headers.forEach(h => { t += "<th>" + escapeHtml(h) + "</th>"; });
    t += "</tr>";
    pageRows.forEach((row, i) => {
        t += '<tr data-idx="' + (start + i) + '">';
        appState.headers.forEach(h => { t += "<td>" + escapeHtml(String(row[h] || "")) + "</td>"; });
        t += "</tr>";
    });
    t += "</table>";
    // Pagination controls
    t += '<div class="pagination-controls" style="display:flex;align-items:center;justify-content:space-between;margin-top:8px;font-size:12px;">';
    t += '<span>Showing ' + (totalRows === 0 ? 0 : start + 1) + '-' + end + ' of ' + totalRows + (appState.filteredRows ? ' (filtered)' : '') + '</span>';
    t += '<span>';
    t += '<button class="btn btn-secondary btn-small" onclick="prevDataPage()" ' + (page <= 0 ? 'disabled' : '') + '>\u25C0 Previous</button> ';
    t += '<button class="btn btn-secondary btn-small" onclick="nextDataPage()" ' + (page >= maxPage ? 'disabled' : '') + '>Next \u25B6</button>';
    t += '</span></div>';
    document.getElementById("previewTable").innerHTML = t;
    document.getElementById("dataPreview").style.display = "block";
}

function prevDataPage() {
    if (appState.previewPage > 0) { appState.previewPage--; renderDataPreview(); }
}
function nextDataPage() {
    const sourceRows = appState.filteredRows || appState.rows;
    const maxPage = Math.max(0, Math.ceil(sourceRows.length / appState.previewPageSize) - 1);
    if (appState.previewPage < maxPage) { appState.previewPage++; renderDataPreview(); }
}

function filterRecipients() {
    const q = document.getElementById("recipientSearch").value.toLowerCase().trim();
    if (!q) {
        appState.filteredRows = null;
    } else {
        appState.filteredRows = appState.rows.filter(row => {
            return appState.headers.some(h => String(row[h] || "").toLowerCase().includes(q));
        });
    }
    appState.previewPage = 0;
    renderDataPreview();
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
    appState.filteredRows = null;
    appState.previewPage = 0;
    document.getElementById("recipientSearch").value = "";
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
            '<button class="att-remove" data-name="' + escapeHtml(name) + '" title="Remove">&times;</button>';
        container.appendChild(div);
    }
}
function handlePerRecipientAttachmentUpload(e) {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    console.log("Per-recipient files:", files.map(f => f.name));
    files.forEach(file => {
        // Store the File object for lazy reading; read content on-demand during send
        appState.perRecipientFiles.set(file.name.toLowerCase(), {
            name: file.name,
            contentBytes: null,
            contentType: file.type || "application/octet-stream",
            file: file,
            size: file.size
        });
        renderPerRecipientAttachmentList();
        checkMissingAttachments();
    });
    e.target.value = "";
}
function renderPerRecipientAttachmentList() {
    const container = document.getElementById("perRecipientAttachmentList");
    container.innerHTML = "";
    for (const [key, att] of appState.perRecipientFiles) {
        const sizeKB = att.contentBytes
            ? Math.round(att.contentBytes.length * 3 / 4 / 1024)
            : Math.round((att.size || 0) / 1024);
        const div = document.createElement("div");
        div.className = "attachment-item";
        div.innerHTML = '<span class="att-name" title="' + escapeHtml(att.name) + '">' + escapeHtml(att.name) + '</span>' +
            '<span class="att-size">' + sizeKB + ' KB</span>' +
            '<button class="att-remove" data-key="' + escapeHtml(key) + '" title="Remove">&times;</button>';
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
            ? mergeFieldsWithGroup(String(row[appState.mapping.subject] || subject), groupRows, false)
            : mergeFieldsWithGroup(subject, groupRows, false);
        mergedBody = mergeFieldsWithGroup(bodyHtml, groupRows, sendAsHtml);
    } else {
        row = appState.rows[idx];
        mergedSubj = appState.mapping.subject
            ? mergeFields(String(row[appState.mapping.subject] || subject), row, false)
            : mergeFields(subject, row, false);
        mergedBody = mergeFields(bodyHtml, row, sendAsHtml);
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

// ========== Conditional Content Processing (B1) ==========
function processConditionals(text, row) {
    // {{#ifEquals Col "val"}}...{{/ifEquals}}
    text = text.replace(/\{\{#ifEquals\s+(\w+)\s+"([^"]+)"\}\}([\s\S]*?)\{\{\/ifEquals\}\}/g,
        (_, col, val, content) => (row[col] || '').toString().trim() === val ? content : '');
    // {{#ifNotEquals Col "val"}}...{{/ifNotEquals}}
    text = text.replace(/\{\{#ifNotEquals\s+(\w+)\s+"([^"]+)"\}\}([\s\S]*?)\{\{\/ifNotEquals\}\}/g,
        (_, col, val, content) => (row[col] || '').toString().trim() !== val ? content : '');
    // {{#if Col}}...{{/if}}
    text = text.replace(/\{\{#if\s+(\w+)\}\}([\s\S]*?)\{\{\/if\}\}/g,
        (_, col, content) => (row[col] && row[col].toString().trim()) ? content : '');
    // {{#ifNot Col}}...{{/ifNot}}
    text = text.replace(/\{\{#ifNot\s+(\w+)\}\}([\s\S]*?)\{\{\/ifNot\}\}/g,
        (_, col, content) => (!row[col] || !row[col].toString().trim()) ? content : '');
    return text;
}

// ========== Merge Engine ==========
function mergeFields(template, row, shouldEscapeHtml) {
    if (!template) return "";
    let result = processConditionals(template, row);
    for (const key of appState.headers) {
        const regex = new RegExp("\\{" + escapeRegExp(key) + "\\}", "g");
        let value = String(row[key] === undefined || row[key] === null ? "" : row[key]);
        if (!value && appState.defaults[key]) value = appState.defaults[key];
        if (shouldEscapeHtml) {
            value = value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
        }
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

function mergeFieldsWithGroup(template, rows, shouldEscapeHtml) {
    if (!template) return "";
    let result = template;
    // Handle {#rows}...{/rows} repeating sections
    result = result.replace(/\{#rows\}([\s\S]*?)\{\/rows\}/g, function(match, inner) {
        return rows.map(function(row) { return mergeFields(inner, row, shouldEscapeHtml); }).join("");
    });
    // Replace remaining fields with first row values
    result = mergeFields(result, rows[0], shouldEscapeHtml);
    return result;
}

// ========== Graph API Helpers ==========
async function graphFetch(url, token, method, body, _retryCount) {
    var retryCount = _retryCount || 0;
    var maxRetries = 3;
    var controller = new AbortController();
    var timeoutId = setTimeout(function() { controller.abort(); }, 30000);
    var options = {
        method: method || "GET",
        headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
        signal: controller.signal
    };
    if (body) options.body = JSON.stringify(body);
    console.log("Graph " + method + " " + url + (retryCount > 0 ? " (retry " + retryCount + ")" : ""));
    try {
        var response = await fetch(url, options);
        clearTimeout(timeoutId);
        if (!response.ok) {
            var errMsg = "HTTP " + response.status + " " + response.statusText;
            try { var eb = await response.json(); if (eb.error && eb.error.message) errMsg += ": " + eb.error.message; } catch(_){}
            if (response.status === 401) throw new Error("SESSION_EXPIRED:" + errMsg);
            if (response.status === 403) throw new Error("Permission denied. Ask admin to grant required permissions.");
            if (response.status === 429) {
                var retryAfter = parseInt(response.headers.get("Retry-After")) || 10;
                if (retryCount < maxRetries) {
                    console.warn("Throttled (429). Retry-After: " + retryAfter + "s");
                    await sleep(retryAfter * 1000);
                    return graphFetch(url, token, method, body, retryCount + 1);
                }
                throw new Error("THROTTLED:" + errMsg);
            }
            // Retry on server errors (500, 502, 503, 504)
            if (response.status >= 500 && retryCount < maxRetries) {
                var backoff = Math.pow(2, retryCount) * 1000;
                console.warn("Server error " + response.status + ". Retrying in " + (backoff/1000) + "s...");
                await sleep(backoff);
                return graphFetch(url, token, method, body, retryCount + 1);
            }
            throw new Error(errMsg);
        }
        var ct = response.headers.get("content-type");
        if (ct && ct.includes("application/json")) return response.json();
        return null;
    } catch (fetchErr) {
        clearTimeout(timeoutId);
        // Retry on network/timeout errors
        if ((fetchErr.name === "AbortError" || fetchErr.message === "Failed to fetch") && retryCount < maxRetries) {
            var backoffMs = Math.pow(2, retryCount) * 1000;
            console.warn("Network error: " + fetchErr.message + ". Retrying in " + (backoffMs/1000) + "s...");
            await sleep(backoffMs);
            return graphFetch(url, token, method, body, retryCount + 1);
        }
        if (fetchErr.name === "AbortError") throw new Error("Request timed out after 30 seconds");
        throw fetchErr;
    }
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

// Large attachment upload via Graph upload session (>3MB)
async function uploadLargeAttachment(baseUrl, token, messageId, att) {
    const contentBytes = att.contentBytes;
    const sizeInBytes = Math.ceil(contentBytes.length * 3 / 4);
    // Create upload session
    const sessionBody = {
        AttachmentItem: {
            "@odata.type": "microsoft.graph.fileAttachment",
            name: att.name,
            size: sizeInBytes,
            contentType: att.contentType || "application/octet-stream"
        }
    };
    const session = await graphFetch(
        baseUrl + "/messages/" + encodeURIComponent(messageId) + "/attachments/createUploadSession",
        token, "POST", sessionBody
    );
    const uploadUrl = session.uploadUrl;
    // Decode base64 to binary
    const binaryStr = atob(contentBytes);
    const bytes = new Uint8Array(binaryStr.length);
    for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
    // Upload in chunks of ~3.1MB
    const chunkSize = 3276800;
    let offset = 0;
    while (offset < bytes.length) {
        const end = Math.min(offset + chunkSize, bytes.length);
        const chunk = bytes.slice(offset, end);
        const resp = await fetch(uploadUrl, {
            method: "PUT",
            headers: {
                "Content-Type": "application/octet-stream",
                "Content-Length": String(chunk.length),
                "Content-Range": "bytes " + offset + "-" + (end - 1) + "/" + bytes.length
            },
            body: chunk
        });
        if (!resp.ok && resp.status !== 200 && resp.status !== 201) {
            let errMsg = "Upload chunk failed: HTTP " + resp.status;
            try { const eb = await resp.json(); if (eb.error && eb.error.message) errMsg += ": " + eb.error.message; } catch(_){}
            throw new Error(errMsg);
        }
        offset = end;
    }
    console.log("Large attachment uploaded:", att.name, sizeInBytes, "bytes");
}

async function collectAttachmentsForRow(row) {
    const attachments = [];
    // Add global attachments (already cached with content)
    for (const [, att] of appState.globalAttachments) attachments.push(att);
    // Add per-recipient attachments from spreadsheet column (lazy read)
    if (appState.mapping.attachments) {
        const val = String(row[appState.mapping.attachments] || "").trim();
        if (val) {
            const parts = val.split(";");
            for (const f of parts) {
                const rawName = f.trim();
                if (!rawName) continue;
                const fileName = rawName.replace(/^.*[\\\/]/, "").toLowerCase();
                let att = appState.perRecipientFiles.get(rawName.toLowerCase());
                if (!att) att = appState.perRecipientFiles.get(fileName);
                if (att) {
                    // Lazy load: read file content on-demand if not yet loaded
                    if (!att.contentBytes && att.file) {
                        att.contentBytes = await readFileAsBase64(att.file);
                    }
                    attachments.push({ name: att.name, contentBytes: att.contentBytes, contentType: att.contentType });
                    console.log("Per-recipient attachment matched:", rawName, "->", att.name);
                } else {
                    console.warn("Per-recipient attachment not found:", rawName, "| Tried key:", fileName, "| Available:", Array.from(appState.perRecipientFiles.keys()).join(", "));
                }
            }
        }
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
        const rawSizeBytes = Math.ceil((att.contentBytes || "").length * 3 / 4);
        if (rawSizeBytes < 3000000) {
            await graphFetch(baseUrl + "/messages/" + encodeURIComponent(msgId) + "/attachments", token, "POST", buildGraphAttachment(att));
        } else {
            await uploadLargeAttachment(baseUrl, token, msgId, att);
        }
        console.log("Added attachment:", att.name, "(" + rawSizeBytes + " bytes)");
    }
    if (!draftOnly) {
        await graphFetch(baseUrl + "/messages/" + encodeURIComponent(msgId) + "/send", token, "POST", null);
        console.log("Sent:", msgId);
    } else {
        console.log("Draft saved:", msgId);
    }
}

// ========== Email Validation (B2) ==========
function validateRecipients(data, toColumn) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const issues = [];
    const seen = new Map();

    data.forEach((row, i) => {
        const email = (row[toColumn] || '').toString().trim();
        if (!email) {
            issues.push({ row: i + 2, email: email, issue: 'Empty email' });
        } else if (!emailRegex.test(email)) {
            issues.push({ row: i + 2, email: email, issue: 'Invalid format' });
        }
        if (email && seen.has(email.toLowerCase())) {
            issues.push({ row: i + 2, email: email, issue: 'Duplicate (first at row ' + seen.get(email.toLowerCase()) + ')' });
        } else if (email) {
            seen.set(email.toLowerCase(), i + 2);
        }
    });

    return issues;
}

function showValidationIssues(issues) {
    return new Promise((resolve) => {
        const modal = document.getElementById("validationResultsModal");
        const content = document.getElementById("validationResultsContent");
        let html = '<p>' + issues.length + ' issue(s) found:</p>';
        html += '<table class="results-table"><tr><th>Row</th><th>Email</th><th>Issue</th></tr>';
        issues.slice(0, 50).forEach(issue => {
            html += '<tr><td>' + issue.row + '</td><td>' + escapeHtml(issue.email) + '</td><td>' + escapeHtml(issue.issue) + '</td></tr>';
        });
        if (issues.length > 50) html += '<tr><td colspan="3">...and ' + (issues.length - 50) + ' more</td></tr>';
        html += '</table>';
        content.innerHTML = html;
        modal.style.display = "flex";
        trapFocus(modal);

        const sendBtn = document.getElementById("btnValidationSendAnyway");
        const cancelBtn = document.getElementById("btnValidationCancel");
        const cleanup = () => {
            sendBtn.removeEventListener("click", onSend);
            cancelBtn.removeEventListener("click", onCancel);
            modal.style.display = "none";
            releaseFocus();
        };
        const onSend = () => { cleanup(); resolve(true); };
        const onCancel = () => { cleanup(); resolve(false); };
        sendBtn.addEventListener("click", onSend);
        cancelBtn.addEventListener("click", onCancel);
    });
}

// ========== Suppression / Blocklist (B3) ==========
function getSuppressionList() {
    return safeJsonParse(localStorage.getItem('mmPro_suppressionList'), []);
}
function addToSuppressionList(email) {
    const list = getSuppressionList();
    const normalized = email.toLowerCase().trim();
    if (normalized && !list.includes(normalized)) {
        list.push(normalized);
        safeLocalStorageSet('mmPro_suppressionList', JSON.stringify(list));
    }
    renderSuppressionList();
}
function removeFromSuppressionList(email) {
    let list = getSuppressionList();
    list = list.filter(e => e !== email.toLowerCase().trim());
    safeLocalStorageSet('mmPro_suppressionList', JSON.stringify(list));
    renderSuppressionList();
}
function renderSuppressionList() {
    const container = document.getElementById('suppressionList');
    if (!container) return;
    const list = getSuppressionList();
    if (!list.length) { container.innerHTML = '<p class="text-muted" data-i18n="noBlockedEmails">No blocked emails</p>'; }
    else {
        container.innerHTML = list.map(e =>
            '<div class="suppression-item"><span>' + escapeHtml(e) + '</span><button class="btn-icon att-remove" data-email="' + escapeHtml(e) + '">✕</button></div>'
        ).join('');
    }
    const countEl = document.getElementById('suppressionCount');
    if (countEl) countEl.textContent = list.length + ' blocked';
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

        // B2: Validate recipients before sending
        if (!testMode) {
            const validationIssues = validateRecipients(appState.rows, appState.mapping.to);
            if (validationIssues.length > 0) {
                const proceed = await showValidationIssues(validationIssues);
                if (!proceed) return;
            }
        }

        // B3: Get suppression list
        const suppressionList = getSuppressionList();

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
            const mSubj = appState.mapping.subject ? mergeFields(String(row[appState.mapping.subject] || subject), row, false) : mergeFields(subject, row, false);
            const mBody = sendAsHtml ? mergeFields(bodyContent + signatureBlock, row, true) : mergeFields(plainBody + signatureBlock, row, false);
            let ccL = ""; if (appState.mapping.cc && row[appState.mapping.cc]) ccL = String(row[appState.mapping.cc]);
            if (globalCC) ccL = ccL ? ccL + ";" + globalCC : globalCC;
            let bccL = ""; if (appState.mapping.bcc && row[appState.mapping.bcc]) bccL = String(row[appState.mapping.bcc]);
            if (globalBCC) bccL = bccL ? bccL + ";" + globalBCC : globalBCC;
            const atts = await collectAttachmentsForRow(row);
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

        // Checkpoint: save progress so send can resume on page refresh
        function saveCheckpoint() {
            safeLocalStorageSet("mailmergepro_checkpoint", JSON.stringify({
                sent: sent, errors: errors, lastIndex: 0, results: appState.results.slice(-20),
                timestamp: Date.now()
            }));
        }

        for (let i = 0; i < total; i++) {
            const item = sendItems[i];
            const toAddr = item.to;
            if (!toAddr) {
                errors++;
                appState.results.push({ row: i + 2, to: "(empty)", status: "Error", error: "No email address" });
                continue;
            }
            // B3: Skip suppressed emails
            if (suppressionList.includes(toAddr.toLowerCase().trim())) {
                appState.results.push({ row: i + 2, to: toAddr, status: "Skipped", error: "Suppressed (blocklist)" });
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
                ? (isGroup ? mergeFieldsWithGroup(String(row[appState.mapping.subject] || activeSubject), item.rows, false) : mergeFields(String(row[appState.mapping.subject] || activeSubject), row, false))
                : (isGroup ? mergeFieldsWithGroup(activeSubject, item.rows, false) : mergeFields(activeSubject, row, false));
            const mBody = isGroup
                ? (sendAsHtml ? mergeFieldsWithGroup(activeBodyHtml, item.rows, true) : mergeFieldsWithGroup(activeBodyPlain, item.rows, false))
                : (sendAsHtml ? mergeFields(activeBodyHtml, row, true) : mergeFields(activeBodyPlain, row, false));
            let ccL = ""; if (appState.mapping.cc && row[appState.mapping.cc]) ccL = String(row[appState.mapping.cc]);
            if (globalCC) ccL = ccL ? ccL + ";" + globalCC : globalCC;
            let bccL = ""; if (appState.mapping.bcc && row[appState.mapping.bcc]) bccL = String(row[appState.mapping.bcc]);
            if (globalBCC) bccL = bccL ? bccL + ";" + globalBCC : globalBCC;
            // Collect attachments from ALL rows in the group, deduped by name
            let atts;
            if (item.rows.length > 1) {
                const allAttachments = new Map();
                for (const groupRow of item.rows) {
                    const rowAttachments = await collectAttachmentsForRow(groupRow);
                    rowAttachments.forEach(att => {
                        if (!allAttachments.has(att.name)) allAttachments.set(att.name, att);
                    });
                }
                atts = Array.from(allAttachments.values());
            } else {
                atts = await collectAttachmentsForRow(row);
            }

            // Enforce rate limit before each send
            await rateLimiter.waitUntilReady();
            updateProgress(i, total, modeLabel + " " + (i+1) + " of " + total + " \u2014 " + escapeHtml(toAddr));

            try {
                await sendOneEmail(token, toAddr, ccL, bccL, mSubj, mBody, sendAsHtml, fromAlias, atts, draftOnly, opts);
                rateLimiter.record();
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
            // Checkpoint every 10 emails
            if (i > 0 && i % 10 === 0) saveCheckpoint();
            if (i < total - 1 && delay > 0) await sleep(delay * 1000);
        }

        // Clear checkpoint on completion
        localStorage.removeItem("mailmergepro_checkpoint");
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
    el.innerHTML = "<p>" + escapeHtml(message) + "</p>";
    el.setAttribute("role", "alert");
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
    return safeJsonParse(stored, []);
}

function saveTemplatesStorage(templates) {
    safeLocalStorageSet("mailmergepro_templates", JSON.stringify(templates));
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
        card.dataset.tmplIdx = idx;
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
            del.className = "btn-icon tmpl-delete";
            del.title = "Delete";
            del.dataset.idx = idx - builtIn.length;
            del.innerHTML = "&times;";
            card.appendChild(del);
        }
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
    const modal = document.getElementById("templateNameDialog");
    modal.style.display = "flex";
    trapFocus(modal);
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
    releaseFocus();
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
    safeLocalStorageSet("mailmergepro_scheduled", JSON.stringify({
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
    return safeJsonParse(stored, []);
}

function saveSavedListsStorage(lists) {
    var json = JSON.stringify(lists);
    if (json.length > 3 * 1024 * 1024) {
        showStatus(t("sizeWarning"), "warning");
    }
    safeLocalStorageSet("mailmergepro_contactgroups", json);
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
        card.dataset.listIdx = idx;
        card.innerHTML = '<span class="list-name">' + escapeHtml(list.name) + '</span>' +
            '<span class="list-meta">' + list.rows.length + ' rows</span>';
        var del = document.createElement("button");
        del.className = "btn-icon list-delete";
        del.innerHTML = "&times;";
        del.title = "Delete";
        del.dataset.idx = idx;
        card.appendChild(del);
        container.appendChild(card);
    });
}

function showSaveListDialog() {
    document.getElementById("listNameInput").value = "";
    const modal = document.getElementById("listNameDialog");
    modal.style.display = "flex";
    trapFocus(modal);
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
    releaseFocus();
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
    const mergeModal = document.getElementById("mergeListDialog");
    mergeModal.style.display = "flex";
    trapFocus(mergeModal);
}

function mergeSelectedList() {
    var sel = document.getElementById("mergeListSelect");
    var idx = parseInt(sel.value);
    var lists = getSavedLists();
    var list = lists[idx];
    document.getElementById("mergeListDialog").style.display = "none";
    releaseFocus();
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
    editor.innerHTML = sanitizeHtml(htmlContent);
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
    const sigModal = document.getElementById("signatureDialog");
    sigModal.style.display = "flex";
    trapFocus(sigModal);
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
    safeLocalStorageSet("mailmergepro_signature", content);
    safeLocalStorageSet("mailmergepro_autosignature", autoAppend ? "true" : "false");
    document.getElementById("signatureDialog").style.display = "none";
    releaseFocus();
    document.getElementById("chkAutoSignatureInline").checked = autoAppend;
    loadSignaturePreview();
    showStatus(t("signatureSaved"), "info");
}

function loadSignaturePreview() {
    var preview = document.getElementById("signaturePreview");
    if (!preview) return;
    if (appState.signatureHtml) {
        preview.innerHTML = sanitizeHtml(appState.signatureHtml);
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
    safeLocalStorageSet("mailmergepro_dailysent", JSON.stringify({
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
    var campaigns = safeJsonParse(stored, []);

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