/**
 * ============================================================
 *  AGENT PROSPECTION GSA PRADO — Google Apps Script
 *  Version 1.1 — Mai 2026
 * ============================================================
 *
 *  INSTALLATION :
 *  1. Ouvrez votre Google Sheet "Target list"
 *  2. Extensions → Apps Script
 *  3. Copiez-collez ce fichier (remplacez le contenu existant)
 *  4. Enregistrez (Ctrl+S)
 *  5. Menu "🚀 Prospection → ⚙️ Configurer les clés API"
 *     → Saisissez vos clés une seule fois (stockées de façon sécurisée)
 *  6. Rechargez le Sheet → le menu est prêt
 *
 *  Les clés API ne sont PAS dans ce fichier — elles sont stockées
 *  dans les Propriétés du script Apps Script (chiffrées par Google).
 *  Ce fichier peut donc être hébergé sur GitHub sans risque.
 * ============================================================
 */

// ── CONFIGURATION — colonnes uniquement (pas de clés ici) ─────────────────────
const CONFIG = {
  COL_NOM:      1,
  COL_ODOO:     2,
  COL_EMAIL:    3,
  COL_PHONE:    4,
  COL_ZELIQ_ST: 5,
  FIRST_ROW:    2,
};

/**
 * Récupère les clés API depuis les Propriétés du script (stockage sécurisé).
 */
function getSecrets() {
  const props = PropertiesService.getScriptProperties();
  const secrets = {
    ODOO_URL:      props.getProperty("ODOO_URL")      || "https://gsa-prado.odoo.com",
    ODOO_DB:       props.getProperty("ODOO_DB")       || "",
    ODOO_USER:     props.getProperty("ODOO_USER")     || "",
    ODOO_API_KEY:  props.getProperty("ODOO_API_KEY")  || "",
    ZELIQ_API_KEY: props.getProperty("ZELIQ_API_KEY") || "",
    ZELIQ_BASE:    "https://api.zeliq.com/api",
  };
  if (!secrets.ODOO_API_KEY || !secrets.ZELIQ_API_KEY) {
    throw new Error(
      "Clés API non configurées !\n\nMenu 🚀 Prospection → ⚙️ Configurer les clés API"
    );
  }
  return secrets;
}

// ── COULEURS ──────────────────────────────────────────────────────────────────
const COLORS = {
  HEADER:  "#1A3A5C",
  GREEN:   "#D5F5E3",
  BLUE:    "#D6EAF8",
  RED:     "#FADBD8",
  GRAY:    "#F2F3F4",
  YELLOW:  "#FEF9E7",
  WHITE:   "#FFFFFF",
  CREATED: "#A9DFBF",
};

// ── MENU ──────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Prospection")
    .addItem("⚡ Vérifier Odoo — toutes les lignes",      "checkAllOdoo")
    .addItem("🔵 Vérifier Odoo — ligne sélectionnée",    "checkSelectedOdoo")
    .addSeparator()
    .addItem("📧 Enrichir Zeliq — ligne sélectionnée",   "enrichSelectedZeliq")
    .addItem("📧 Enrichir Zeliq — toutes les lignes",    "enrichAllZeliq")
    .addSeparator()
    .addItem("➕ Créer lead Odoo — ligne sélectionnée",  "createLeadSelected")
    .addSeparator()
    .addItem("🏗️  Initialiser les en-têtes",              "initHeaders")
    .addItem("⚙️  Configurer les clés API",               "configureSecrets")
    .addSeparator()
    .addItem("📖  Guide d'utilisation",                   "showHelp")
    .addToUi();
}

// ── INITIALISATION DES EN-TÊTES ───────────────────────────────────────────────
function initHeaders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = {
    [CONFIG.COL_ODOO]:     "STATUT ODOO CRM",
    [CONFIG.COL_EMAIL]:    "📧 EMAIL (Zeliq)",
    [CONFIG.COL_PHONE]:    "📞 TÉLÉPHONE (Zeliq)",
    [CONFIG.COL_ZELIQ_ST]: "⚡ STATUT ENRICHISSEMENT",
  };
  const widths = { 2: 220, 3: 240, 4: 160, 5: 200 };

  Object.entries(headers).forEach(([col, label]) => {
    const cell = sheet.getRange(1, parseInt(col));
    cell.setValue(label)
        .setBackground(COLORS.HEADER)
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setWrap(true);
    sheet.setColumnWidth(parseInt(col), widths[col] || 180);
  });
  sheet.setRowHeight(1, 40);
  toast("✅ En-têtes initialisés !");
}

// ══════════════════════════════════════════════════════════════════════════════
//  ODOO — VÉRIFICATION
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Authentifie et retourne l'UID Odoo.
 * Utilise le cache pour éviter de s'authentifier à chaque appel.
 */
function odooGetUid() {
  const cache = CacheService.getScriptCache();
  let uid = cache.get("odoo_uid");
  if (uid) return parseInt(uid);

  const S = getSecrets();
  const payload = JSON.stringify({
    jsonrpc: "2.0", method: "call", id: 1,
    params: {
      service: "common", method: "authenticate",
      args: [S.ODOO_DB, S.ODOO_USER, S.ODOO_API_KEY, {}]
    }
  });
  const resp = UrlFetchApp.fetch(S.ODOO_URL + "/jsonrpc", {
    method: "post",
    contentType: "application/json",
    payload: payload,
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  uid = data.result;
  if (!uid) throw new Error("Odoo : authentification échouée. Vérifiez vos identifiants et la config (menu ⚙️).");
  cache.put("odoo_uid", uid.toString(), 1500);
  return uid;
}

/**
 * Génère des variantes de recherche pour tolérer les fautes d'orthographe.
 * Ex: "Dupessay" → ["Dupessay", "Dupes", "Dupess"]
 * Ex: "FLB Group GT2i" → ["FLB Group GT2i", "FLB", "GT2i"]
 */
function searchVariants(name) {
  const variants = [name];

  // Ajouter chaque mot individuel de plus de 3 caractères
  name.split(/[\s\/\-&]+/).forEach(word => {
    const w = word.trim();
    if (w.length > 3 && !variants.includes(w)) variants.push(w);
  });

  // Ajouter un préfixe de 5 caractères du premier mot (utile pour les noms mal orthographiés)
  const firstWord = name.split(/[\s\/\-&]+/)[0].trim();
  if (firstWord.length >= 5) {
    const prefix = firstWord.substring(0, 5);
    if (!variants.includes(prefix)) variants.push(prefix);
  }

  return variants;
}

/**
 * Recherche des opportunités CRM pour un nom d'entreprise.
 * Utilise plusieurs variantes pour tolérer les fautes d'orthographe.
 */
function odooSearchLeads(companyName) {
  const uid = odooGetUid();
  const clean = cleanName(companyName);
  if (!clean) return [];

  const variants = searchVariants(clean);

  // Construire un domaine OR sur toutes les variantes et tous les champs
  const conditions = [];
  variants.forEach(v => {
    conditions.push(["partner_name",    "ilike", v]);
    conditions.push(["partner_id.name", "ilike", v]);
    conditions.push(["name",            "ilike", v]);
  });

  // Odoo OR domain : ["|", "|", ..., "|", cond1, cond2, ..., condN]
  // Pour N conditions, il faut N-1 opérateurs "|"
  const ors = Array(conditions.length - 1).fill("|");
  const domain = [...ors, ...conditions];

  const searchPayload = JSON.stringify({
    jsonrpc: "2.0", method: "call", id: 2,
    params: {
      service: "object", method: "execute_kw",
      args: [getSecrets().ODOO_DB, uid, getSecrets().ODOO_API_KEY,
             "crm.lead", "search", [domain], { limit: 5 }]
    }
  });

  const searchResp = UrlFetchApp.fetch(getSecrets().ODOO_URL + "/jsonrpc", {
    method: "post", contentType: "application/json",
    payload: searchPayload, muteHttpExceptions: true,
  });
  const ids = JSON.parse(searchResp.getContentText()).result || [];
  if (!ids.length) return [];

  const readPayload = JSON.stringify({
    jsonrpc: "2.0", method: "call", id: 3,
    params: {
      service: "object", method: "execute_kw",
      args: [getSecrets().ODOO_DB, uid, getSecrets().ODOO_API_KEY,
             "crm.lead", "read", [ids],
             { fields: ["name", "stage_id", "probability", "user_id", "active"] }]
    }
  });

  const readResp = UrlFetchApp.fetch(getSecrets().ODOO_URL + "/jsonrpc", {
    method: "post", contentType: "application/json",
    payload: readPayload, muteHttpExceptions: true,
  });
  return JSON.parse(readResp.getContentText()).result || [];
}

/**
 * Formate le statut Odoo pour affichage dans la cellule.
 */
function formatOdooStatus(leads) {
  if (!leads.length) return { text: "⚪ Aucune opportunité", color: COLORS.GRAY };

  const lines = [];
  let bestColor = COLORS.GRAY;

  leads.slice(0, 3).forEach(l => {
    const stage  = Array.isArray(l.stage_id) ? l.stage_id[1] : "?";
    const user   = Array.isArray(l.user_id)  ? l.user_id[1]  : "?";
    const proba  = l.probability || 0;
    const active = l.active !== false;

    let emoji, color;
    if (!active)       { emoji = "🔴"; color = COLORS.RED;  }
    else if (proba === 100) { emoji = "🟢"; color = COLORS.GREEN; }
    else               { emoji = "🔵"; color = COLORS.BLUE; }

    lines.push(`${emoji} ${l.name}\n   ${stage} | ${user}`);
    if (color === COLORS.GREEN) bestColor = COLORS.GREEN;
    else if (color === COLORS.BLUE && bestColor !== COLORS.GREEN) bestColor = COLORS.BLUE;
    else if (color === COLORS.RED  && bestColor === COLORS.GRAY)  bestColor = COLORS.RED;
  });

  const suffix = leads.length > 1 ? `\n(${leads.length} opp.)` : "";
  return { text: lines.join("\n\n") + suffix, color: bestColor };
}

/**
 * Écrit le statut Odoo dans la cellule correspondante.
 */
function writeOdooStatus(sheet, row, leads) {
  const { text, color } = formatOdooStatus(leads);
  const cell = sheet.getRange(row, CONFIG.COL_ODOO);
  cell.setValue(text)
      .setBackground(color)
      .setFontSize(9)
      .setWrap(true)
      .setVerticalAlignment("middle");
  sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 60));
  return color;
}

// ── ACTION : Vérifier toutes les lignes ───────────────────────────────────────
function checkAllOdoo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let checked = 0, errors = 0;

  toast("🔍 Vérification Odoo en cours…");

  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;

    try {
      const leads = odooSearchLeads(nom);
      writeOdooStatus(sheet, row, leads);
      checked++;
      SpreadsheetApp.flush();
      Utilities.sleep(300); // éviter le rate limiting
    } catch (e) {
      sheet.getRange(row, CONFIG.COL_ODOO)
           .setValue("❌ Erreur : " + e.message)
           .setBackground(COLORS.RED);
      errors++;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ ${checked} vérifiées${errors ? `, ❌ ${errors} erreurs` : ""}`,
    "Odoo — Terminé", 5
  );
}

// ── ACTION : Vérifier la ligne sélectionnée ───────────────────────────────────
function checkSelectedOdoo() {
  const sheet  = SpreadsheetApp.getActiveSheet();
  const row    = sheet.getActiveCell().getRow();

  if (row < CONFIG.FIRST_ROW) {
    showAlert("Sélectionnez une ligne de données (pas l'en-tête).");
    return;
  }

  const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  if (!nom) { showAlert("La cellule Nom est vide."); return; }

  toast("🔍 Vérification Odoo pour " + nom + "…");

  try {
    const leads = odooSearchLeads(nom);
    const color = writeOdooStatus(sheet, row, leads);
    const msg   = leads.length
      ? `${leads.length} opportunité(s) trouvée(s)`
      : "Aucune opportunité dans Odoo";
    toast("✅ " + msg);
  } catch (e) {
    showAlert("Erreur Odoo : " + e.message);
  }
}

// ── ACTION : Créer un lead Odoo ───────────────────────────────────────────────
function createLeadSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();

  if (row < CONFIG.FIRST_ROW) {
    showAlert("Sélectionnez une ligne de données.");
    return;
  }

  const nom    = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  const email  = sheet.getRange(row, CONFIG.COL_EMAIL).getValue();
  const phone  = sheet.getRange(row, CONFIG.COL_PHONE).getValue();

  if (!nom) { showAlert("La cellule Nom est vide."); return; }

  // Confirmation
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Créer un lead Odoo",
    `Créer une opportunité pour "${nom}" dans Odoo CRM ?`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  try {
    const uid = odooGetUid();
    const vals = {
      name:         "Prospection Assurance Transport — " + nom,
      partner_name: nom,
      description:  "Lead créé depuis Google Sheets GSA Prado\nDate : " + new Date().toLocaleDateString("fr-FR"),
      type:         "opportunity",
    };
    if (email && email.includes("@")) vals.email_from = email;
    if (phone && phone.length > 5)    vals.phone      = phone;

    const payload = JSON.stringify({
      jsonrpc: "2.0", method: "call", id: 4,
      params: {
        service: "object", method: "execute_kw",
        args: [getSecrets().ODOO_DB, uid, getSecrets().ODOO_API_KEY,
               "crm.lead", "create", [vals]]
      }
    });

    const resp2 = UrlFetchApp.fetch(getSecrets().ODOO_URL + "/jsonrpc", {
      method: "post", contentType: "application/json",
      payload: payload, muteHttpExceptions: true,
    });
    const leadId = JSON.parse(resp2.getContentText()).result;
    const url = getSecrets().ODOO_URL + "/odoo/crm/" + leadId;

    sheet.getRange(row, CONFIG.COL_ODOO)
         .setValue("🆕 Lead créé (ID:" + leadId + ")\n" + url)
         .setBackground(COLORS.CREATED)
         .setFontWeight("bold")
         .setWrap(true);

    ui.alert("✅ Lead créé !", "ID : " + leadId + "\n" + url, ui.ButtonSet.OK);
  } catch (e) {
    showAlert("Erreur création lead : " + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
//  ZELIQ — ENRICHISSEMENT
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Appelle l'API Zeliq pour enrichir téléphone et email.
 * linkedin_url est optionnel mais améliore les résultats.
 */
function zeliqEnrich(companyName, linkedinUrl) {
  const headers = {
    "x-api-key":    getSecrets().ZELIQ_API_KEY,
    "Content-Type": "application/json",
  };

  const payloadBase = { callback_url: "https://webhook.site/zeliq-gsaprado" };
  if (linkedinUrl && linkedinUrl.includes("linkedin.com")) {
    payloadBase.linkedin_url = linkedinUrl;
  }

  let phone = "", email = "";

  // Téléphone
  try {
    const r = UrlFetchApp.fetch(getSecrets().ZELIQ_BASE + "/contact/enrich/phone", {
      method: "post",
      headers: headers,
      payload: JSON.stringify(payloadBase),
      muteHttpExceptions: true,
    });
    const d = JSON.parse(r.getContentText());
    phone = (d.contact && d.contact.most_probable_phone) ? d.contact.most_probable_phone : "";
  } catch (e) { /* silencieux */ }

  // Email
  try {
    const r = UrlFetchApp.fetch(getSecrets().ZELIQ_BASE + "/contact/enrich/email", {
      method: "post",
      headers: headers,
      payload: JSON.stringify(payloadBase),
      muteHttpExceptions: true,
    });
    const d = JSON.parse(r.getContentText());
    email = (d.contact && d.contact.most_probable_email) ? d.contact.most_probable_email : "";
  } catch (e) { /* silencieux */ }

  return { phone, email };
}

/**
 * Écrit les résultats Zeliq dans les cellules.
 */
function writeZeliqResult(sheet, row, phone, email) {
  // Email
  sheet.getRange(row, CONFIG.COL_EMAIL)
       .setValue(email || "Non trouvé")
       .setBackground(email ? COLORS.GREEN : COLORS.RED)
       .setFontWeight(email ? "bold" : "normal")
       .setFontSize(9).setWrap(true);

  // Téléphone
  sheet.getRange(row, CONFIG.COL_PHONE)
       .setValue(phone || "Non trouvé")
       .setBackground(phone ? COLORS.GREEN : COLORS.RED)
       .setFontWeight(phone ? "bold" : "normal")
       .setFontSize(9).setWrap(true);

  // Statut
  let status, color;
  if (phone && email) { status = "✅ Complet";   color = COLORS.GREEN; }
  else if (phone || email) { status = "⚠️ Partiel"; color = COLORS.YELLOW; }
  else { status = "❌ Non trouvé"; color = COLORS.RED; }

  sheet.getRange(row, CONFIG.COL_ZELIQ_ST)
       .setValue(status).setBackground(color).setFontSize(9).setWrap(true);
}

// ── ACTION : Enrichir la ligne sélectionnée ───────────────────────────────────
function enrichSelectedZeliq() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();

  if (row < CONFIG.FIRST_ROW) {
    showAlert("Sélectionnez une ligne de données.");
    return;
  }

  const nom     = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  if (!nom) { showAlert("La cellule Nom est vide."); return; }

  // Demander URL LinkedIn si non renseignée
  const ui = SpreadsheetApp.getUi();
  let linkedinUrl = "";
  const resp = ui.prompt(
    "Enrichissement Zeliq — " + nom,
    "URL LinkedIn du décisionnaire (optionnel, laissez vide pour ignorer) :",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  linkedinUrl = resp.getResponseText().trim();

  toast("⚡ Enrichissement Zeliq pour " + nom + "…");

  const { phone, email } = zeliqEnrich(nom, linkedinUrl);
  writeZeliqResult(sheet, row, phone, email);

  const msg = `📞 ${phone || "—"}  |  📧 ${email || "—"}`;
  toast("✅ " + msg);
}

// ── ACTION : Enrichir toutes les lignes ───────────────────────────────────────
function enrichAllZeliq() {
  const ui    = SpreadsheetApp.getUi();
  const conf  = ui.alert(
    "Enrichissement Zeliq — toutes les lignes",
    "Cela va consommer des crédits Zeliq pour chaque ligne. Continuer ?",
    ui.ButtonSet.OK_CANCEL
  );
  if (conf !== ui.Button.OK) return;

  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let done = 0;

  toast("⚡ Enrichissement Zeliq en cours…");

  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;

    // Skip si déjà enrichi
    const existEmail = sheet.getRange(row, CONFIG.COL_EMAIL).getValue();
    if (existEmail && existEmail.includes("@")) continue;

    try {
      const { phone, email } = zeliqEnrich(nom, "");
      writeZeliqResult(sheet, row, phone, email);
      done++;
      SpreadsheetApp.flush();
      Utilities.sleep(600);
    } catch (e) {
      sheet.getRange(row, CONFIG.COL_ZELIQ_ST)
           .setValue("❌ Erreur : " + e.message)
           .setBackground(COLORS.RED);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("✅ " + done + " lignes traitées", "Zeliq — Terminé", 5);
}

// ══════════════════════════════════════════════════════════════════════════════
//  UTILITAIRES
// ══════════════════════════════════════════════════════════════════════════════

function cleanName(raw) {
  if (!raw) return "";
  let name = raw.toString().split("\n")[0].trim();
  name = name.replace(/\s*\(.*?\)/g, "").trim();
  [" SAS", " SA", " SARL", " SASU", " SNC", " INC.", " INC", " LLC"].forEach(s => {
    if (name.toUpperCase().endsWith(s)) name = name.slice(0, -s.length).trim();
  });
  return name;
}

function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Agent Prospection", 3);
}

function showAlert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}


// ── CONFIGURATION DES CLÉS API ────────────────────────────────────────────────
function configureSecrets() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const fields = [
    { key: "ODOO_URL",      label: "URL Odoo",         def: "https://gsa-prado.odoo.com" },
    { key: "ODOO_DB",       label: "Base de données Odoo", def: "unikerp-gsaprado-prod-13772120" },
    { key: "ODOO_USER",     label: "Email Odoo",        def: "georges-eric.michel@gsaprado.fr" },
    { key: "ODOO_API_KEY",  label: "Clé API Odoo",      def: "" },
    { key: "ZELIQ_API_KEY", label: "Clé API Zeliq",     def: "" },
  ];

  for (const f of fields) {
    const current = props.getProperty(f.key) || f.def;
    const resp = ui.prompt(
      "⚙️ Configuration — " + f.label,
      "Valeur actuelle : " + (current ? current.substring(0, 20) + "..." : "(vide)") +
      "\n\nNouvelle valeur (laissez vide pour conserver) :",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const val = resp.getResponseText().trim();
    if (val) props.setProperty(f.key, val);
    else if (!current) props.setProperty(f.key, f.def);
  }

  // Vider le cache d'auth Odoo pour forcer une reconnexion
  CacheService.getScriptCache().remove("odoo_uid");
  toast("✅ Clés API enregistrées de façon sécurisée !");
}

// ── GUIDE ─────────────────────────────────────────────────────────────────────
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; font-size: 13px; padding: 16px; }
      h2 { color: #1A3A5C; }
      h3 { color: #2E6DA4; margin-top: 16px; }
      code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
      ul { padding-left: 20px; }
      li { margin-bottom: 6px; }
    </style>
    <h2>🚀 Agent Prospection GSA Prado</h2>
    <h3>Workflow recommandé</h3>
    <ol>
      <li>Ajoutez le nom d'une société en <strong>colonne A</strong></li>
      <li>Menu <code>🚀 Prospection → Vérifier Odoo — toutes les lignes</code><br>
          → La colonne B se remplit avec le statut CRM</li>
      <li>Pour les lignes <strong>⚪ Aucune opportunité</strong> :<br>
          Sélectionnez la ligne → <code>Enrichir Zeliq — ligne sélectionnée</code><br>
          → Collez l'URL LinkedIn du décisionnaire si demandé</li>
      <li>Si l'enrichissement est satisfaisant :<br>
          <code>Créer lead Odoo — ligne sélectionnée</code></li>
    </ol>
    <h3>Colonnes gérées par ce script</h3>
    <ul>
      <li><strong>A</strong> — Nom entreprise (votre saisie)</li>
      <li><strong>B</strong> — Statut Odoo CRM</li>
      <li><strong>C</strong> — Email enrichi (Zeliq)</li>
      <li><strong>D</strong> — Téléphone enrichi (Zeliq)</li>
      <li><strong>E</strong> — Statut enrichissement</li>
    </ul>
    <h3>Couleurs des statuts Odoo</h3>
    <ul>
      <li>🟢 Vert — opportunité gagnée</li>
      <li>🔵 Bleu — en cours</li>
      <li>🔴 Rouge — perdue / archivée</li>
      <li>⚪ Gris — aucune opportunité</li>
      <li>🆕 Vert foncé — lead venant d'être créé</li>
    </ul>
  `)
  .setTitle("Guide d'utilisation")
  .setWidth(480)
  .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, "Guide d'utilisation");
}
