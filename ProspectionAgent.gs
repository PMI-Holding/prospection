/**
 * ============================================================
 *  AGENT PROSPECTION GSA PRADO — Google Apps Script
 *  Version 2.0 — Mai 2026
 * ============================================================
 *
 *  COLONNES (v2.0) :
 *  A — Nom entreprise (saisie manuelle)
 *  B — Secteur / Activité NAF       ← Pappers
 *  C — Ville / Siège                ← Pappers
 *  D — CA (M€)                      ← Pappers
 *  E — Résultat net (M€)            ← Pappers
 *  F — Évolution CA                 ← Pappers
 *  G — Effectif                     ← Pappers
 *  H — 🥇 Décideur final (DG/PDG)   ← Pappers
 *  I — 🥈 Décideur opérationnel     ← Pappers
 *  J — 🥉 Prescripteur              ← Pappers
 *  K — LinkedIn entreprise          ← Lien de recherche
 *  L — LinkedIn décisionnaire       ← Saisie manuelle (pour Zeliq)
 *  M — Signal / Déclencheur BODACC  ← Pappers
 *  N — Actualité & sources          ← Google News RSS
 *  O — Accroche commerciale         ← Généré depuis signaux
 *  P — Statut Odoo CRM              ← Script (checkOdoo)
 *  Q — Email (Zeliq)                ← Script (enrichZeliq)
 *  R — Téléphone (Zeliq)            ← Script (enrichZeliq)
 *  S — Statut enrichissement Zeliq  ← Script (enrichZeliq)
 *
 *  ⚠️ Pappers et Zeliq = enrichissement TOUJOURS manuel (crédits payants)
 * ============================================================
 */

// ── CONFIGURATION ─────────────────────────────────────────────────────────────
const CONFIG = {
  COL_NOM:      1,   // A — Nom entreprise
  COL_SECTEUR:  2,   // B — Secteur / Activité NAF
  COL_VILLE:    3,   // C — Ville / Siège
  COL_CA:       4,   // D — CA (M€)
  COL_RESULTAT: 5,   // E — Résultat net (M€)
  COL_EVOL_CA:  6,   // F — Évolution CA
  COL_EFFECTIF: 7,   // G — Effectif
  COL_DG:       8,   // H — 🥇 Décideur final (DG/PDG)
  COL_DAF:      9,   // I — 🥈 Décideur opérationnel (DAF/CFO/Risk)
  COL_PRESCR:   10,  // J — 🥉 Prescripteur (Dir. Logistique/Supply Chain)
  COL_LI_ENT:   11,  // K — LinkedIn entreprise
  COL_LI_DEC:   12,  // L — LinkedIn décisionnaire (saisie manuelle → Zeliq)
  COL_SIGNAL:   13,  // M — Signal / Déclencheur BODACC
  COL_ACTU:     14,  // N — Actualité & sources
  COL_ACCROCHE: 15,  // O — Accroche commerciale
  COL_ODOO:     16,  // P — Statut Odoo CRM
  COL_EMAIL:    17,  // Q — Email (Zeliq)
  COL_PHONE:    18,  // R — Téléphone (Zeliq)
  COL_ZELIQ_ST: 19,  // S — Statut enrichissement Zeliq
  FIRST_ROW:    2,
};

function getSecrets() {
  const props = PropertiesService.getScriptProperties();
  return {
    ODOO_URL:        props.getProperty("ODOO_URL")        || "https://gsa-prado.odoo.com",
    ODOO_DB:         props.getProperty("ODOO_DB")         || "",
    ODOO_USER:       props.getProperty("ODOO_USER")       || "",
    ODOO_API_KEY:    props.getProperty("ODOO_API_KEY")    || "",
    ZELIQ_API_KEY:   props.getProperty("ZELIQ_API_KEY")   || "",
    ZELIQ_BASE:      "https://api.zeliq.com/api",
  };
}

// ── COULEURS ──────────────────────────────────────────────────────────────────
const COLORS = {
  HEADER:   "#1A3A5C",
  GREEN:    "#D5F5E3",
  BLUE:     "#D6EAF8",
  RED:      "#FADBD8",
  GRAY:     "#F2F3F4",
  YELLOW:   "#FEF9E7",
  ORANGE:   "#FDEBD0",
  PURPLE:   "#EDE7F6",
  LINKEDIN: "#E8F4FD",
  WHITE:    "#FFFFFF",
  CREATED:  "#A9DFBF",
  PAPPERS:  "#EAF2FF",
  SIGNAL:   "#FFF3CD",
};

// ── MENU ──────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 Prospection")
    .addItem("🏢 Enrichir données — ligne sélectionnée",  "enrichSelectedData")
    .addItem("🏢 Enrichir données — toutes les lignes",   "enrichAllData")
    .addSeparator()
    .addItem("📰 Actualités — ligne sélectionnée",        "enrichSelectedNews")
    .addItem("📰 Actualités — toutes les lignes",         "enrichAllNews")
    .addSeparator()
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
  const headers = [
    { col: CONFIG.COL_NOM,      label: "🏭 ENTREPRISE",                     width: 200 },
    { col: CONFIG.COL_SECTEUR,  label: "📂 SECTEUR / NAF",                  width: 180 },
    { col: CONFIG.COL_VILLE,    label: "📍 VILLE / SIÈGE",                  width: 130 },
    { col: CONFIG.COL_CA,       label: "💰 CA (M€)",                        width: 90  },
    { col: CONFIG.COL_RESULTAT, label: "📊 RÉS. NET\n(M€)",                 width: 90  },
    { col: CONFIG.COL_EVOL_CA,  label: "📈 ÉVOL. CA",                       width: 80  },
    { col: CONFIG.COL_EFFECTIF, label: "👥 EFFECTIF",                       width: 80  },
    { col: CONFIG.COL_DG,       label: "🥇 DÉCIDEUR FINAL\n(DG / PDG)",     width: 180 },
    { col: CONFIG.COL_DAF,      label: "🥈 DÉCIDEUR OPS\n(DAF/CFO/Risk)",   width: 180 },
    { col: CONFIG.COL_PRESCR,   label: "🥉 PRESCRIPTEUR\n(Logistique/SC)",  width: 180 },
    { col: CONFIG.COL_LI_ENT,   label: "🔗 LINKEDIN ENTREPRISE",            width: 210 },
    { col: CONFIG.COL_LI_DEC,   label: "🔗 LINKEDIN DÉCISIONNAIRE\n(→ Zeliq)", width: 210 },
    { col: CONFIG.COL_SIGNAL,   label: "⚡ SIGNAL BODACC",                  width: 220 },
    { col: CONFIG.COL_ACTU,     label: "📰 ACTUALITÉ & SOURCES",            width: 300 },
    { col: CONFIG.COL_ACCROCHE, label: "💬 ACCROCHE COMMERCIALE",           width: 320 },
    { col: CONFIG.COL_ODOO,     label: "🏢 STATUT ODOO CRM",                width: 220 },
    { col: CONFIG.COL_EMAIL,    label: "📧 EMAIL\n(Zeliq)",                  width: 200 },
    { col: CONFIG.COL_PHONE,    label: "📞 TÉLÉPHONE\n(Zeliq)",              width: 140 },
    { col: CONFIG.COL_ZELIQ_ST, label: "⚡ STATUT ZELIQ",                   width: 130 },
  ];

  headers.forEach(({ col, label, width }) => {
    const cell = sheet.getRange(1, col);
    cell.setValue(label)
        .setBackground(COLORS.HEADER)
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setFontSize(9)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
    sheet.setColumnWidth(col, width);
  });
  sheet.setRowHeight(1, 50);
  toast("✅ En-têtes v2.0 initialisés (19 colonnes) !");
}

// ══════════════════════════════════════════════════════════════════════════════
//  ENRICHISSEMENT ENTREPRISES — API Sirene / recherche-entreprises.api.gouv.fr
//  Gratuit, sans clé API, données officielles INSEE
// ══════════════════════════════════════════════════════════════════════════════

const SIRENE_BASE = "https://recherche-entreprises.api.gouv.fr";

// Niveaux décisionnaires pour l'assurance transport
const DIRECTOR_LEVELS = {
  1: ["directeur general", "dg", "pdg", "president directeur general", "president",
      "gerant", "ceo", "managing director", "directeur executif", "administrateur",
      "directeur general delegue"],
  2: ["directeur administratif et financier", "daf", "directeur financier", "cfo",
      "risk manager", "directeur des risques", "responsable assurance",
      "directeur des assurances", "risk management", "secretaire general",
      "directeur general adjoint", "dga", "directeur financier et administratif"],
  3: ["directeur logistique", "directeur supply chain", "directeur des operations",
      "coo", "directeur transport", "responsable logistique", "responsable transport",
      "responsable supply chain", "directeur achats", "directeur commercial",
      "directeur industriel", "directeur de site"],
};

function normalizeStr(s) {
  return (s || "").toLowerCase()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9 ]/g, " ");
}

function classifyDirectorLevel(titre) {
  const t = normalizeStr(titre);
  for (const level of [1, 2, 3]) {
    if (DIRECTOR_LEVELS[level].some(kw => t.includes(kw))) return level;
  }
  return 4;
}

function classifyDirectors(dirigeants) {
  const groups = { 1: [], 2: [], 3: [] };
  (dirigeants || []).forEach(d => {
    const titre = d.titre || d.qualite || d.fonction || "";
    const nom   = [d.prenom, d.nom].filter(Boolean).join(" ").trim()
                || (d.nom_complet || "").trim();
    const level = classifyDirectorLevel(titre);
    if (level <= 3 && (nom || titre)) groups[level].push({ nom, titre: titre.trim() });
  });
  const EMOJIS = { 1: "🥇", 2: "🥈", 3: "🥉" };
  const format = (level) => {
    const persons = groups[level];
    if (!persons.length) return "";
    return persons.slice(0, 2).map(p =>
      p.titre ? EMOJIS[level] + " " + p.nom + "\n   " + p.titre : EMOJIS[level] + " " + p.nom
    ).join("\n");
  };
  return { dg: format(1), daf: format(2), prescr: format(3) };
}

/**
 * Interroge l'API Sirene gratuite par nom d'entreprise.
 * Retourne le premier résultat ou null.
 */
function rechercheEntreprises(companyName) {
  const url = SIRENE_BASE + "/search?q=" + encodeURIComponent(cleanName(companyName))
            + "&page=1&per_page=1";
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200)
    throw new Error("API Sirene HTTP " + resp.getResponseCode());
  const data = JSON.parse(resp.getContentText());
  const results = data.results || [];
  return results[0] || null;
}

/**
 * Extrait les signaux contextuels depuis les données Sirene.
 * Limité à ce que l'API fournit : date de création, forme juridique.
 */
function extractSignals(company) {
  const signals = [];
  const dateCreation = company.date_creation || "";

  if (dateCreation) {
    const created = new Date(dateCreation);
    const twoYearsAgo = new Date();
    twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);
    if (created >= twoYearsAgo) {
      signals.push("🆕 Création récente (" + created.toLocaleDateString("fr-FR", { month: "long", year: "numeric" }) + ")");
    }
  }

  // Effectif élevé = flux importants à assurer
  const tranche = company.tranche_effectif_salarie || "";
  if (["41","42","51","52","53"].includes(tranche)) {
    signals.push("🏭 Grande entreprise (" + (company.libelle_tranche_effectif || tranche) + ")");
  }

  return signals.length ? signals.join("\n") : "Aucun signal détecté\n(pour des signaux BODACC : acquisitions, changements de dirigeants, consultez Societe.com)";
}

function generateAccroche(signals, companyName, effectif) {
  const name = cleanName(companyName);
  if (signals.includes("🆕")) {
    return name + " vient de démarrer son activité — c'est le bon moment pour mettre en place des couvertures assurance transport adaptées à votre croissance, avant que les volumes augmentent.";
  }
  if (signals.includes("🏭")) {
    return "Avec " + effectif + ", " + name + " génère probablement des flux de marchandises significatifs. Vos couvertures assurance transport sont-elles calibrées pour ce niveau d'activité ?";
  }
  return "GSA Prado, courtier spécialisé en assurance transport et marchandises, serait heureux d'auditer gratuitement les couvertures de " + name + " et d'identifier des pistes d'optimisation.";
}

function writeCompanyData(sheet, row, company, originalName) {
  const siege    = company.siege || {};
  const { dg, daf, prescr } = classifyDirectors(company.dirigeants || []);
  const effectif = company.libelle_tranche_effectif || company.tranche_effectif_salarie || "N/D";
  const signals  = extractSignals(company);
  const accroche = generateAccroche(signals, originalName, effectif);

  // Secteur : on préfère les données du siège si disponibles
  const codeNaf    = siege.activite_principale || company.activite_principale || "";
  const libelleNaf = siege.libelle_activite_principale || company.libelle_activite_principale || "";
  const secteur    = [libelleNaf, codeNaf].filter(Boolean).join(" — ");
  const ville      = [siege.libelle_commune, siege.code_postal].filter(Boolean).join(" ") || "N/D";
  const liUrl      = "https://www.linkedin.com/search/results/companies/?keywords="
                   + encodeURIComponent(cleanName(originalName));
  const hasSignal  = signals.includes("🆕") || signals.includes("🏭");

  const cells = [
    { col: CONFIG.COL_SECTEUR,  val: secteur || "N/D",        bg: COLORS.PAPPERS  },
    { col: CONFIG.COL_VILLE,    val: ville,                    bg: COLORS.PAPPERS  },
    { col: CONFIG.COL_CA,       val: "N/D (source: Sirene)",   bg: COLORS.GRAY     },
    { col: CONFIG.COL_RESULTAT, val: "N/D (source: Sirene)",   bg: COLORS.GRAY     },
    { col: CONFIG.COL_EVOL_CA,  val: "N/D",                    bg: COLORS.GRAY     },
    { col: CONFIG.COL_EFFECTIF, val: effectif,                 bg: COLORS.PAPPERS  },
    { col: CONFIG.COL_DG,       val: dg    || "Non identifié", bg: dg    ? COLORS.BLUE   : COLORS.GRAY },
    { col: CONFIG.COL_DAF,      val: daf   || "Non identifié", bg: daf   ? COLORS.PURPLE : COLORS.GRAY },
    { col: CONFIG.COL_PRESCR,   val: prescr|| "Non identifié", bg: prescr? COLORS.GREEN  : COLORS.GRAY },
    { col: CONFIG.COL_LI_ENT,   val: liUrl,                    bg: COLORS.LINKEDIN },
    { col: CONFIG.COL_SIGNAL,   val: signals,                  bg: hasSignal ? COLORS.SIGNAL : COLORS.GRAY },
    { col: CONFIG.COL_ACCROCHE, val: accroche,                 bg: COLORS.YELLOW   },
  ];

  cells.forEach(({ col, val, bg }) => {
    sheet.getRange(row, col)
         .setValue(val).setBackground(bg)
         .setFontSize(9).setWrap(true).setVerticalAlignment("middle");
  });
  sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 80));
}

// ── ACTION : Enrichir la ligne sélectionnée ───────────────────────────────────
function enrichSelectedData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();
  if (row < CONFIG.FIRST_ROW) { showAlert("Sélectionnez une ligne de données (pas l'en-tête)."); return; }
  const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  if (!nom) { showAlert("La cellule Nom (colonne A) est vide."); return; }

  toast('🔍 Recherche Sirene pour "' + nom + '"…');
  try {
    const company = rechercheEntreprises(nom);
    if (!company) {
      sheet.getRange(row, CONFIG.COL_SECTEUR)
           .setValue("❌ Entreprise non trouvée (API Sirene)")
           .setBackground(COLORS.RED);
      showAlert('"' + nom + '" introuvable dans la base Sirene.\nVérifiez l\'orthographe du nom.');
      return;
    }
    toast("✅ SIREN " + company.siren + " trouvé — écriture des données…");
    writeCompanyData(sheet, row, company, nom);
    SpreadsheetApp.flush();
    toast('✅ Enrichissement terminé pour "' + nom + '"');
  } catch (e) {
    sheet.getRange(row, CONFIG.COL_SECTEUR)
         .setValue("❌ Erreur : " + e.message).setBackground(COLORS.RED);
    showAlert("Erreur API Sirene : " + e.message);
  }
}

// ── ACTION : Enrichir toutes les lignes ───────────────────────────────────────
function enrichAllData() {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let done = 0, skipped = 0, errors = 0;

  toast("🔍 Enrichissement Sirene en cours…");

  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;
    const existing = sheet.getRange(row, CONFIG.COL_SECTEUR).getValue();
    if (existing && !existing.toString().startsWith("❌")) { skipped++; continue; }

    toast("[" + (row - 1) + "/" + (lastRow - 1) + "] Sirene : " + nom + "…");
    try {
      const company = rechercheEntreprises(nom);
      if (!company) {
        sheet.getRange(row, CONFIG.COL_SECTEUR).setValue("❌ Non trouvé (Sirene)").setBackground(COLORS.RED);
        errors++;
      } else {
        writeCompanyData(sheet, row, company, nom);
        done++;
      }
    } catch (e) {
      sheet.getRange(row, CONFIG.COL_SECTEUR)
           .setValue("❌ " + e.message).setBackground(COLORS.RED);
      errors++;
    }
    SpreadsheetApp.flush();
    Utilities.sleep(300);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ " + done + " enrichies | ⏭ " + skipped + " déjà faites | ❌ " + errors + " erreurs",
    "Sirene — Terminé", 8
  );
}

// ══════════════════════════════════════════════════════════════════════════════
//  GOOGLE NEWS RSS — ACTUALITÉS
// ══════════════════════════════════════════════════════════════════════════════

// Articles plus vieux que MAX_NEWS_AGE_DAYS sont ignorés
const MAX_NEWS_AGE_DAYS = 730; // 2 ans

/**
 * Récupère les actualités Google News et retourne un tableau d'articles structurés.
 * Chaque article : { title, source, dateLabel, url }
 */
function fetchGoogleNews(companyName) {
  try {
    const query  = encodeURIComponent('"' + cleanName(companyName) + '"');
    const url    = "https://news.google.com/rss/search?q=" + query + "&hl=fr&gl=FR&ceid=FR:fr";
    const resp   = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) return [];

    const xml    = resp.getContentText();
    const items  = xml.match(/<item>([\s\S]*?)<\/item>/g) || [];
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - MAX_NEWS_AGE_DAYS);

    const articles = [];
    for (const item of items) {
      if (articles.length >= 5) break;

      const rawTitle = (item.match(/<title><!\[CDATA\[(.*?)\]\]><\/title>/)
                     || item.match(/<title>(.*?)<\/title>/) || [])[1] || "";
      const link     = (item.match(/<link>(.*?)<\/link>/) || [])[1] || "";
      const dateStr  = (item.match(/<pubDate>(.*?)<\/pubDate>/) || [])[1] || "";
      if (!rawTitle) continue;

      let pubDate = null;
      if (dateStr) { try { pubDate = new Date(dateStr); } catch(e) {} }
      if (pubDate && pubDate < cutoff) continue;

      const parts  = rawTitle.split(" - ");
      const source = parts.length > 1 ? parts.pop().trim() : "";
      const title  = parts.join(" - ").trim() || rawTitle;

      let dateLabel = "";
      if (pubDate) {
        try {
          dateLabel = " [" + pubDate.toLocaleDateString("fr-FR", {
            day: "numeric", month: "short", year: "numeric"
          }) + "]";
        } catch(e) {}
      }

      articles.push({ title, source, dateLabel, url: link });
    }
    return articles;
  } catch(e) {
    return [];
  }
}

/**
 * Construit un RichTextValue pour la cellule : titres cliquables, URL masquées.
 * Format affiché : "• Titre [date] — Source"
 */
function buildNewsRichText(articles) {
  const lines = articles.map(a =>
    "• " + a.title + a.dateLabel + (a.source ? " — " + a.source : "")
  );
  const fullText = lines.join("\n\n");
  const builder  = SpreadsheetApp.newRichTextValue().setText(fullText);

  let offset = 0;
  articles.forEach((a, i) => {
    if (a.url) {
      const start = offset + 2; // saute "• "
      const end   = start + a.title.length;
      builder.setLinkUrl(start, end, a.url);
    }
    offset += lines[i].length + (i < articles.length - 1 ? 2 : 0); // +2 pour "\n\n"
  });

  return builder.build();
}

function writeNewsCell(range, articles) {
  if (!articles.length) {
    range.setValue("Aucune actualité récente trouvée (< 2 ans)")
         .setBackground(COLORS.GRAY);
    return;
  }
  range.setRichTextValue(buildNewsRichText(articles))
       .setBackground(COLORS.PAPPERS);
  range.setFontSize(9).setWrap(true).setVerticalAlignment("middle");
}

// ── ACTION : Actualités — ligne sélectionnée ──────────────────────────────────
function enrichSelectedNews() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();
  if (row < CONFIG.FIRST_ROW) { showAlert("Sélectionnez une ligne de données."); return; }

  const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  if (!nom) { showAlert("La cellule Nom (colonne A) est vide."); return; }

  toast('📰 Recherche d\'actualités pour "' + nom + '"…');
  const articles = fetchGoogleNews(nom);
  writeNewsCell(sheet.getRange(row, CONFIG.COL_ACTU), articles);
  sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 80));
  SpreadsheetApp.flush();
  toast("✅ " + articles.length + " article(s) trouvé(s) pour \"" + nom + "\"");
}

// ── ACTION : Actualités — toutes les lignes ───────────────────────────────────
function enrichAllNews() {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let done = 0;
  toast("📰 Recherche d'actualités en cours…");

  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;
    toast("[" + (row - 1) + "/" + (lastRow - 1) + "] Actualités : " + nom + "…");
    const articles = fetchGoogleNews(nom);
    writeNewsCell(sheet.getRange(row, CONFIG.COL_ACTU), articles);
    sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 80));
    done++;
    SpreadsheetApp.flush();
    Utilities.sleep(500);
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("✅ " + done + " lignes traitées", "Actualités — Terminé", 5);
}

// ══════════════════════════════════════════════════════════════════════════════
//  ODOO — VÉRIFICATION
// ══════════════════════════════════════════════════════════════════════════════

function odooGetUid() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("odoo_uid");
  if (cached) return parseInt(cached);
  const S = getSecrets();
  if (!S.ODOO_API_KEY) throw new Error("Clé API Odoo non configurée (menu ⚙️).");
  const resp = UrlFetchApp.fetch(S.ODOO_URL + "/jsonrpc", {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      jsonrpc: "2.0", method: "call", id: 1,
      params: { service: "common", method: "authenticate",
                args: [S.ODOO_DB, S.ODOO_USER, S.ODOO_API_KEY, {}] }
    }),
    muteHttpExceptions: true,
  });
  const uid = JSON.parse(resp.getContentText()).result;
  if (!uid) throw new Error("Odoo : authentification échouée. Vérifiez vos identifiants.");
  cache.put("odoo_uid", uid.toString(), 1500);
  return uid;
}

// Mots courants français à ignorer pour le fallback de recherche
const STOP_WORDS = new Set(["de","du","des","le","la","les","et","en","au","aux",
  "un","une","sur","par","pour","dans","avec","chez","sans","sous","vers","entre"]);

function odooFetch_(S, uid, domain) {
  const ids = JSON.parse(UrlFetchApp.fetch(S.ODOO_URL + "/jsonrpc", {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      jsonrpc: "2.0", method: "call", id: 2,
      params: { service: "object", method: "execute_kw",
                args: [S.ODOO_DB, uid, S.ODOO_API_KEY, "crm.lead", "search", [domain], { limit: 5 }] }
    }),
    muteHttpExceptions: true,
  }).getContentText()).result || [];
  if (!ids.length) return [];
  return JSON.parse(UrlFetchApp.fetch(S.ODOO_URL + "/jsonrpc", {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      jsonrpc: "2.0", method: "call", id: 3,
      params: { service: "object", method: "execute_kw",
                args: [S.ODOO_DB, uid, S.ODOO_API_KEY, "crm.lead", "read", [ids],
                       { fields: ["name", "stage_id", "probability", "user_id", "active"] }] }
    }),
    muteHttpExceptions: true,
  }).getContentText()).result || [];
}

function odooSearchLeads(companyName) {
  const uid = odooGetUid();
  const S   = getSecrets();
  const clean = cleanName(companyName);
  if (!clean) return [];

  // Passe 1 : nom complet (strict)
  const domain1 = ["|", "|",
    ["partner_name",    "ilike", clean],
    ["partner_id.name", "ilike", clean],
    ["name",            "ilike", clean],
  ];
  const results1 = odooFetch_(S, uid, domain1);
  if (results1.length) return results1;

  // Passe 2 (fallback) : phrase des 2 premiers mots significatifs (≥4 chars, hors stop words)
  // Ex: "Comptoir Export des matières premières" → "Comptoir Export"
  const sigWords = clean.split(/[\s\/\-&,]+/)
    .filter(w => w.length >= 4 && !STOP_WORDS.has(w.toLowerCase()));
  if (sigWords.length < 2) return [];
  const phrase = sigWords.slice(0, 2).join(" ");
  if (phrase === clean) return []; // éviter de refaire la même recherche

  const domain2 = ["|", "|",
    ["partner_name",    "ilike", phrase],
    ["partner_id.name", "ilike", phrase],
    ["name",            "ilike", phrase],
  ];
  return odooFetch_(S, uid, domain2);
}


function formatOdooStatus(leads) {
  if (!leads.length) return { text: "⚪ Aucune opportunité", color: COLORS.GRAY };
  const lines = [];
  let best = COLORS.GRAY;
  leads.slice(0, 3).forEach(l => {
    const stage = Array.isArray(l.stage_id) ? l.stage_id[1] : "?";
    const user  = Array.isArray(l.user_id)  ? l.user_id[1]  : "?";
    const proba = l.probability || 0;
    const active = l.active !== false;
    let emoji, color;
    if (!active)       { emoji = "🔴"; color = COLORS.RED;   }
    else if (proba===100){ emoji = "🟢"; color = COLORS.GREEN; }
    else               { emoji = "🔵"; color = COLORS.BLUE;  }
    lines.push(emoji + " " + l.name + "\n   " + stage + " | " + user);
    if (color===COLORS.GREEN) best=COLORS.GREEN;
    else if (color===COLORS.BLUE && best!==COLORS.GREEN) best=COLORS.BLUE;
    else if (color===COLORS.RED  && best===COLORS.GRAY)  best=COLORS.RED;
  });
  const suffix = leads.length > 1 ? "\n(" + leads.length + " opp.)" : "";
  return { text: lines.join("\n\n") + suffix, color: best };
}

function writeOdooStatus(sheet, row, leads) {
  const { text, color } = formatOdooStatus(leads);
  sheet.getRange(row, CONFIG.COL_ODOO)
       .setValue(text).setBackground(color)
       .setFontSize(9).setWrap(true).setVerticalAlignment("middle");
  sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 60));
  return color;
}

function checkAllOdoo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let checked = 0, errors = 0;
  toast("🔍 Vérification Odoo en cours…");
  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;
    try {
      writeOdooStatus(sheet, row, odooSearchLeads(nom));
      checked++;
      SpreadsheetApp.flush();
      Utilities.sleep(300);
    } catch (e) {
      sheet.getRange(row, CONFIG.COL_ODOO)
           .setValue("❌ " + e.message).setBackground(COLORS.RED);
      errors++;
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ " + checked + " vérifiées" + (errors ? ", ❌ " + errors + " erreurs" : ""),
    "Odoo — Terminé", 5
  );
}

function checkSelectedOdoo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();
  if (row < CONFIG.FIRST_ROW) { showAlert("Sélectionnez une ligne de données."); return; }
  const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  if (!nom) { showAlert("La cellule Nom est vide."); return; }
  toast("🔍 Vérification Odoo pour " + nom + "…");
  try {
    const leads = odooSearchLeads(nom);
    writeOdooStatus(sheet, row, leads);
    toast("✅ " + (leads.length ? leads.length + " opportunité(s)" : "Aucune opportunité"));
  } catch (e) { showAlert("Erreur Odoo : " + e.message); }
}

function createLeadSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();
  if (row < CONFIG.FIRST_ROW) { showAlert("Sélectionnez une ligne de données."); return; }

  const nom      = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  const email    = sheet.getRange(row, CONFIG.COL_EMAIL).getValue();
  const phone    = sheet.getRange(row, CONFIG.COL_PHONE).getValue();
  const signal   = sheet.getRange(row, CONFIG.COL_SIGNAL).getValue();
  const accroche = sheet.getRange(row, CONFIG.COL_ACCROCHE).getValue();
  const dg       = sheet.getRange(row, CONFIG.COL_DG).getValue();
  if (!nom) { showAlert("La cellule Nom est vide."); return; }

  const ui   = SpreadsheetApp.getUi();
  const conf = ui.alert(
    "Créer un lead Odoo",
    "Créer une opportunité pour \"" + nom + "\" ?\n\nSignal : " + (signal || "").substring(0, 100),
    ui.ButtonSet.OK_CANCEL
  );
  if (conf !== ui.Button.OK) return;

  try {
    const uid = odooGetUid();
    const S   = getSecrets();
    const desc = [
      signal   ? "SIGNAL BODACC :\n" + signal     : "",
      accroche ? "\nACCROCHE :\n"    + accroche   : "",
      dg       ? "\nDÉCIDEUR :\n"    + dg         : "",
    ].filter(Boolean).join("\n")
    + "\n\nDate : " + new Date().toLocaleDateString("fr-FR");

    const vals = {
      name:         "Prospection Assurance Transport — " + nom,
      partner_name: nom,
      description:  desc,
      type:         "opportunity",
    };
    if (email && email.includes("@")) vals.email_from = email;
    if (phone && phone.length > 5)    vals.phone      = phone;

    const leadId = JSON.parse(UrlFetchApp.fetch(S.ODOO_URL + "/jsonrpc", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({
        jsonrpc: "2.0", method: "call", id: 4,
        params: { service: "object", method: "execute_kw",
                  args: [S.ODOO_DB, uid, S.ODOO_API_KEY, "crm.lead", "create", [vals]] }
      }),
      muteHttpExceptions: true,
    }).getContentText()).result;

    const url = S.ODOO_URL + "/odoo/crm/" + leadId;
    sheet.getRange(row, CONFIG.COL_ODOO)
         .setValue("🆕 Lead créé (ID:" + leadId + ")\n" + url)
         .setBackground(COLORS.CREATED).setFontWeight("bold").setWrap(true);
    ui.alert("✅ Lead créé !", "ID : " + leadId + "\n" + url, ui.ButtonSet.OK);
  } catch (e) { showAlert("Erreur création lead : " + e.message); }
}

// ══════════════════════════════════════════════════════════════════════════════
//  ZELIQ — ENRICHISSEMENT (déclenché MANUELLEMENT uniquement)
// ══════════════════════════════════════════════════════════════════════════════

function zeliqEnrich(linkedinUrl) {
  const S = getSecrets();
  if (!S.ZELIQ_API_KEY) throw new Error("Clé API Zeliq non configurée (menu ⚙️).");
  const headers     = { "x-api-key": S.ZELIQ_API_KEY, "Content-Type": "application/json" };
  const payloadBase = { callback_url: "https://webhook.site/zeliq-gsaprado" };
  if (linkedinUrl && linkedinUrl.includes("linkedin.com")) payloadBase.linkedin_url = linkedinUrl;

  let phone = "", email = "";
  try {
    const r = UrlFetchApp.fetch(S.ZELIQ_BASE + "/contact/enrich/phone", {
      method: "post", headers, payload: JSON.stringify(payloadBase), muteHttpExceptions: true,
    });
    phone = (JSON.parse(r.getContentText()).contact || {}).most_probable_phone || "";
  } catch(e) {}
  try {
    const r = UrlFetchApp.fetch(S.ZELIQ_BASE + "/contact/enrich/email", {
      method: "post", headers, payload: JSON.stringify(payloadBase), muteHttpExceptions: true,
    });
    email = (JSON.parse(r.getContentText()).contact || {}).most_probable_email || "";
  } catch(e) {}
  return { phone, email };
}

function writeZeliqResult(sheet, row, phone, email) {
  sheet.getRange(row, CONFIG.COL_EMAIL)
       .setValue(email || "Non trouvé")
       .setBackground(email ? COLORS.GREEN : COLORS.RED)
       .setFontWeight(email ? "bold" : "normal").setFontSize(9).setWrap(true);
  sheet.getRange(row, CONFIG.COL_PHONE)
       .setValue(phone || "Non trouvé")
       .setBackground(phone ? COLORS.GREEN : COLORS.RED)
       .setFontWeight(phone ? "bold" : "normal").setFontSize(9).setWrap(true);
  let status, color;
  if (phone && email)      { status = "✅ Complet";    color = COLORS.GREEN;  }
  else if (phone || email) { status = "⚠️ Partiel";   color = COLORS.YELLOW; }
  else                     { status = "❌ Non trouvé"; color = COLORS.RED;    }
  sheet.getRange(row, CONFIG.COL_ZELIQ_ST)
       .setValue(status).setBackground(color).setFontSize(9).setWrap(true);
}

function enrichSelectedZeliq() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row   = sheet.getActiveCell().getRow();
  if (row < CONFIG.FIRST_ROW) { showAlert("Sélectionnez une ligne de données."); return; }
  const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
  if (!nom) { showAlert("La cellule Nom est vide."); return; }

  // Lire l'URL LinkedIn depuis la colonne L (saisie manuelle)
  let linkedinUrl = (sheet.getRange(row, CONFIG.COL_LI_DEC).getValue() || "").toString().trim();

  if (!linkedinUrl || !linkedinUrl.includes("linkedin.com")) {
    const ui   = SpreadsheetApp.getUi();
    const resp = ui.prompt(
      "Enrichissement Zeliq — " + nom,
      "URL LinkedIn du décisionnaire (colonne L) :\n(Optionnel — laissez vide pour ignorer)",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const input = resp.getResponseText().trim();
    if (input && input.includes("linkedin.com")) {
      linkedinUrl = input;
      sheet.getRange(row, CONFIG.COL_LI_DEC)
           .setValue(linkedinUrl).setBackground(COLORS.LINKEDIN).setFontSize(9);
    }
  }

  toast("⚡ Enrichissement Zeliq pour " + nom + "…");
  try {
    const { phone, email } = zeliqEnrich(linkedinUrl);
    writeZeliqResult(sheet, row, phone, email);
    toast("✅ 📞 " + (phone || "—") + "  |  📧 " + (email || "—"));
  } catch (e) { showAlert("Erreur Zeliq : " + e.message); }
}

function enrichAllZeliq() {
  const ui   = SpreadsheetApp.getUi();
  const conf = ui.alert(
    "📧 Enrichissement Zeliq — toutes les lignes",
    "⚠️ Cette action consomme des crédits Zeliq pour chaque ligne non encore enrichie. Continuer ?",
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
    const existEmail = sheet.getRange(row, CONFIG.COL_EMAIL).getValue();
    if (existEmail && existEmail.toString().includes("@")) continue;

    const linkedinUrl = (sheet.getRange(row, CONFIG.COL_LI_DEC).getValue() || "").toString().trim();
    try {
      const { phone, email } = zeliqEnrich(linkedinUrl);
      writeZeliqResult(sheet, row, phone, email);
      done++;
      SpreadsheetApp.flush();
      Utilities.sleep(600);
    } catch (e) {
      sheet.getRange(row, CONFIG.COL_ZELIQ_ST)
           .setValue("❌ " + e.message).setBackground(COLORS.RED);
    }
  }
  SpreadsheetApp.getActiveSpreadsheet()
    .toast("✅ " + done + " lignes enrichies", "Zeliq — Terminé", 5);
}

// ══════════════════════════════════════════════════════════════════════════════
//  UTILITAIRES
// ══════════════════════════════════════════════════════════════════════════════

function cleanName(raw) {
  if (!raw) return "";
  let name = raw.toString().split("\n")[0].trim();
  name = name.replace(/\s*\(.*?\)/g, "").trim();
  [" SAS", " SA", " SARL", " SASU", " SNC", " INC.", " INC", " LLC",
   " GROUP", " GROUPE", " HOLDING", " HOLDINGS"].forEach(s => {
    if (name.toUpperCase().endsWith(s)) name = name.slice(0, -s.length).trim();
  });
  return name;
}

function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Agent Prospection", 4);
}

function showAlert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}

// ── CONFIGURATION DES CLÉS API ────────────────────────────────────────────────
function configureSecrets() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const fields = [
    { key: "ODOO_URL",        label: "URL Odoo",               def: "https://gsa-prado.odoo.com"          },
    { key: "ODOO_DB",         label: "Base de données Odoo",   def: "unikerp-gsaprado-prod-13772120"       },
    { key: "ODOO_USER",       label: "Email Odoo",             def: "georges-eric.michel@gsaprado.fr"      },
    { key: "ODOO_API_KEY",    label: "Clé API Odoo",           def: ""                                     },
    { key: "ZELIQ_API_KEY",   label: "Clé API Zeliq",          def: ""                                     },
  ];
  for (const f of fields) {
    const current = props.getProperty(f.key) || f.def;
    const resp = ui.prompt(
      "⚙️ Configuration — " + f.label,
      "Valeur actuelle : " + (current ? current.substring(0, 30) + (current.length > 30 ? "…" : "") : "(vide)")
      + "\n\nNouvelle valeur (laissez vide pour conserver) :",
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    const val = resp.getResponseText().trim();
    if (val) props.setProperty(f.key, val);
    else if (!current) props.setProperty(f.key, f.def);
  }
  CacheService.getScriptCache().remove("odoo_uid");
  toast("✅ Clés API enregistrées de façon sécurisée !");
}

// ── GUIDE D'UTILISATION ───────────────────────────────────────────────────────
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body{font-family:Arial,sans-serif;font-size:12px;padding:16px;color:#333;}
      h2{color:#1A3A5C;border-bottom:2px solid #1A3A5C;padding-bottom:6px;}
      h3{color:#2E6DA4;margin-top:14px;}
      code{background:#f4f4f4;padding:2px 6px;border-radius:3px;font-size:11px;font-family:monospace;}
      table{border-collapse:collapse;width:100%;font-size:11px;margin-top:8px;}
      th{background:#1A3A5C;color:white;padding:5px 8px;text-align:left;}
      td{padding:4px 8px;border-bottom:1px solid #eee;vertical-align:top;}
      tr:hover td{background:#f9f9f9;}
      .p{background:#EAF2FF;padding:1px 5px;border-radius:8px;font-size:10px;}
      .z{background:#D5F5E3;padding:1px 5px;border-radius:8px;font-size:10px;}
      .o{background:#FEF9E7;padding:1px 5px;border-radius:8px;font-size:10px;}
      .m{background:#F2F3F4;padding:1px 5px;border-radius:8px;font-size:10px;}
      ol li{margin-bottom:6px;}
    </style>
    <h2>🚀 Agent Prospection GSA Prado v2.0</h2>
    <h3>📋 Workflow recommandé</h3>
    <ol>
      <li>Saisissez le <strong>nom de l'entreprise en colonne A</strong></li>
      <li>Sélectionnez la ligne → <code>🏢 Enrichir Pappers — ligne sélectionnée</code><br>
          → Remplit B à O : données légales, finances, dirigeants, signaux BODACC, actualités, accroche</li>
      <li><code>⚡ Vérifier Odoo — ligne sélectionnée</code> → remplit colonne P</li>
      <li>Si pas d'opportunité Odoo :
        <ul>
          <li>Collez l'URL LinkedIn du décisionnaire en <strong>colonne L</strong></li>
          <li><code>📧 Enrichir Zeliq — ligne sélectionnée</code> → remplit Q (email) et R (tél.)</li>
          <li><code>➕ Créer lead Odoo — ligne sélectionnée</code></li>
        </ul>
      </li>
    </ol>
    <h3>📊 Colonnes v2.0</h3>
    <table>
      <tr><th>Col.</th><th>Contenu</th><th>Source</th></tr>
      <tr><td>A</td><td>🏭 Nom entreprise</td><td><span class="m">Manuel</span></td></tr>
      <tr><td>B</td><td>📂 Secteur / NAF</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>C</td><td>📍 Ville / Siège</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>D</td><td>💰 CA (M€)</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>E</td><td>📊 Résultat net (M€)</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>F</td><td>📈 Évolution CA</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>G</td><td>👥 Effectif</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>H</td><td>🥇 Décideur final — DG / PDG</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>I</td><td>🥈 Décideur opérationnel — DAF / CFO / Risk</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>J</td><td>🥉 Prescripteur — Dir. Logistique / Supply Chain</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>K</td><td>🔗 LinkedIn entreprise (lien de recherche)</td><td><span class="p">Généré</span></td></tr>
      <tr><td>L</td><td>🔗 LinkedIn décisionnaire → utilisé par Zeliq</td><td><span class="m">Manuel</span></td></tr>
      <tr><td>M</td><td>⚡ Signal BODACC (déclencheur contextuel)</td><td><span class="p">Pappers</span></td></tr>
      <tr><td>N</td><td>📰 Actualité & sources</td><td><span class="p">Google News</span></td></tr>
      <tr><td>O</td><td>💬 Accroche commerciale recommandée</td><td><span class="p">Généré</span></td></tr>
      <tr><td>P</td><td>🏢 Statut Odoo CRM</td><td><span class="o">Odoo</span></td></tr>
      <tr><td>Q</td><td>📧 Email</td><td><span class="z">Zeliq</span></td></tr>
      <tr><td>R</td><td>📞 Téléphone</td><td><span class="z">Zeliq</span></td></tr>
      <tr><td>S</td><td>⚡ Statut Zeliq</td><td><span class="z">Zeliq</span></td></tr>
    </table>
    <h3>🎯 Niveaux décisionnaires (assurance transport)</h3>
    <table>
      <tr><th>Niveau</th><th>Rôles</th><th>Rôle assurance</th></tr>
      <tr><td>🥇 Décideur final</td><td>DG, PDG, Président, CEO, Gérant</td><td>Signe les contrats</td></tr>
      <tr><td>🥈 Décideur opérationnel</td><td>DAF, CFO, Risk Manager, DGA, Secrétaire Général</td><td>Choisit les couvertures</td></tr>
      <tr><td>🥉 Prescripteur</td><td>Dir. Logistique, Supply Chain, Transport, COO</td><td>Exprime le besoin</td></tr>
    </table>
    <h3>⚡ Signaux BODACC = opportunités de démarchage contextuel</h3>
    <table>
      <tr><th>Signal</th><th>Accroche recommandée</th></tr>
      <tr><td>👤 Nouveau dirigeant</td><td>Félicitations pour la prise de poste → revue des contrats</td></tr>
      <tr><td>🤝 Acquisition/Cession</td><td>Nouveau périmètre à couvrir</td></tr>
      <tr><td>🆕 Création</td><td>Mise en place des premières couvertures</td></tr>
      <tr><td>📈 Forte croissance</td><td>Risque de sous-assurance</td></tr>
      <tr><td>🏗️ Apport d'actifs</td><td>Révision du périmètre assuré</td></tr>
      <tr><td>📍 Déménagement</td><td>Nouveaux flux logistiques à couvrir</td></tr>
    </table>
    <p style="color:#999;font-size:10px;margin-top:16px;">
      ⚠️ Pappers et Zeliq consomment des crédits payants — toujours déclenchés manuellement, jamais automatiquement.
    </p>
  `)
  .setTitle("Guide d'utilisation v2.0")
  .setWidth(580)
  .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, "Guide d'utilisation v2.0");
}
