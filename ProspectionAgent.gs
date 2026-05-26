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
    PAPPERS_API_KEY: props.getProperty("PAPPERS_API_KEY") || "",
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
//  ENRICHISSEMENT ENTREPRISES — API Pappers
//  Plan gratuit : 100 requêtes/mois — inscription sur pappers.fr/api
//  Données : NAF, siège, effectif, CA, résultat, dirigeants, BODACC
// ══════════════════════════════════════════════════════════════════════════════

const PAPPERS_BASE = "https://api.pappers.fr/v2";

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
 * Enrichit une entreprise depuis Pappers.
 *
 * MODE GRATUIT (sans clé API) :
 *   Utilise l'endpoint Autocomplete Pappers (100 req/j par IP, sans clé).
 *   → Remplit : secteur NAF, ville, effectif.  Dirigeants et financiers = N/D.
 *
 * MODE COMPLET (avec clé API Pappers) :
 *   Utilise Recherche (0.1 crédit) + Fiche entreprise (1 crédit).
 *   → Remplit tout : dirigeants, CA, résultat, BODACC.
 *   Clé à configurer via menu ⚙️ (inscription gratuite sur pappers.fr/api,
 *   puis achat de crédits selon besoin).
 */
function rechercheEntreprises(companyName) {
  const S     = getSecrets();
  const clean = cleanName(companyName);
  const opts  = { muteHttpExceptions: true };

  if (S.PAPPERS_API_KEY) {
    return rechercheEntreprisesComplet_(clean, S.PAPPERS_API_KEY, opts);
  }
  return rechercheEntreprisesAutocomplete_(clean, opts);
}

// ── Mode gratuit : autocomplete (100 req/j par IP, sans clé) ─────────────────
function rechercheEntreprisesAutocomplete_(clean, opts) {
  const url  = PAPPERS_BASE + "/autocomplete?q=" + encodeURIComponent(clean) + "&longueur=3";
  const resp = UrlFetchApp.fetch(url, opts);
  const code = resp.getResponseCode();

  if (code === 429) {
    throw new Error(
      "Limite autocomplete Pappers atteinte (100 req/j).\n" +
      "→ Réessayez demain, ou configurez une clé API Pappers via ⚙️ pour un accès illimité."
    );
  }
  if (code !== 200) throw new Error("API Pappers inaccessible (HTTP " + code + ").");

  const data    = JSON.parse(resp.getContentText());
  const results = data.resultats_entreprises || data.resultats || [];
  if (!results.length) return null;

  return normalizePappersBasic_(results[0]);
}

function normalizePappersBasic_(p) {
  const siege = p.siege || {};
  return {
    siren:                       p.siren || "",
    nom_complet:                 p.nom_entreprise || "",
    activite_principale:         p.code_naf || "",
    libelle_activite_principale: p.libelle_code_naf || "",
    siege: {
      libelle_commune: siege.ville || "",
      code_postal:     siege.code_postal || "",
    },
    tranche_effectif_salarie:    p.tranche_effectif || p.effectif || "",
    libelle_tranche_effectif:    p.tranche_effectif || "",
    date_creation:               p.date_creation || "",
    dirigeants:                  [],
    chiffre_affaires:            null,
    resultat_net:                null,
    evolution_ca:                "N/D (clé API Pappers requise)",
    annee_fin:                   "",
    publications:                [],
  };
}

// ── Mode complet : recherche + fiche (avec clé API) ──────────────────────────
function rechercheEntreprisesComplet_(clean, apiKey, opts) {
  const token = encodeURIComponent(apiKey);

  // Étape 1 — Recherche (0.1 crédit)
  const searchUrl = PAPPERS_BASE + "/recherche?q=" + encodeURIComponent(clean) +
                    "&api_token=" + token + "&par_page=3";
  let resp = UrlFetchApp.fetch(searchUrl, opts);
  let code = resp.getResponseCode();

  if (code === 401) throw new Error("Clé API Pappers invalide — vérifiez dans ⚙️ Configurer les clés API.");
  if (code !== 200) throw new Error("API Pappers inaccessible (HTTP " + code + ").");

  const results = JSON.parse(resp.getContentText()).resultats || [];
  if (!results.length) return null;

  const siren = results[0].siren;
  if (!siren) return normalizePappers_(results[0]);

  // Étape 2 — Fiche complète (1 crédit) : dirigeants + financiers + BODACC
  const detailUrl = PAPPERS_BASE + "/entreprise?siren=" + siren + "&api_token=" + token;
  resp = UrlFetchApp.fetch(detailUrl, opts);
  if (resp.getResponseCode() === 200) {
    return normalizePappers_(JSON.parse(resp.getContentText()));
  }
  return normalizePappers_(results[0]);
}

/** Normalise une fiche complète Pappers. */
function normalizePappers_(p) {
  const siege = p.siege || {};
  const fin   = p.finances || [];
  const fin0  = fin[0] || {};
  const fin1  = fin[1] || {};

  let evolCa = "N/D";
  if (fin0.chiffre_affaires && fin1.chiffre_affaires) {
    const pct = ((fin0.chiffre_affaires - fin1.chiffre_affaires) / Math.abs(fin1.chiffre_affaires) * 100).toFixed(1);
    evolCa = (pct > 0 ? "+" : "") + pct + "% (" + (fin0.annee || "") + ")";
  }

  return {
    siren:                       p.siren || "",
    nom_complet:                 p.nom_entreprise || "",
    activite_principale:         p.code_naf || siege.code_naf || "",
    libelle_activite_principale: p.libelle_code_naf || "",
    siege: {
      libelle_commune: siege.ville || siege.libelle_commune || "",
      code_postal:     siege.code_postal || "",
    },
    tranche_effectif_salarie:    p.tranche_effectif || "",
    libelle_tranche_effectif:    p.libelle_tranche_effectif || p.effectif || "",
    date_creation:               p.date_creation || "",
    dirigeants: (p.dirigeants || []).map(d => ({
      prenom:  d.prenom  || "",
      nom:     d.nom     || "",
      titre:   d.qualite || d.titre || "",
      qualite: d.qualite || d.titre || "",
    })),
    chiffre_affaires: fin0.chiffre_affaires || p.chiffre_affaires || null,
    resultat_net:     fin0.resultat_net     || p.resultat_net     || null,
    evolution_ca:     evolCa,
    annee_fin:        fin0.annee || "",
    publications:     p.publications || [],
  };
}

/**
 * Extrait les signaux contextuels depuis les publications BODACC Pappers
 * et les données d'entreprise (date de création, effectif).
 */
function extractSignals(company) {
  const signals     = [];
  const twoYearsAgo = new Date();
  twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);

  const PUB_LABELS = {
    "Vente": "🤝 Cession / Vente",
    "Achat": "🛒 Acquisition",
    "Apport partiel d'actifs": "🏗️ Apport d'actifs",
    "Création": "🆕 Création",
    "Dissolution": "⚠️ Dissolution",
    "Redressement judiciaire": "🚨 Redressement judiciaire",
    "Liquidation judiciaire": "🚨 Liquidation judiciaire",
    "Déménagement": "📍 Déménagement / Nouveau siège",
  };

  // Publications BODACC récentes (< 2 ans)
  for (const pub of (company.publications || []).slice(0, 8)) {
    const type  = pub.type || pub.type_publication || "";
    const date  = pub.date ? new Date(pub.date) : null;
    if (date && date < twoYearsAgo) continue;
    const label = PUB_LABELS[type];
    if (label) {
      const dateStr = date
        ? " (" + date.toLocaleDateString("fr-FR", { month: "long", year: "numeric" }) + ")"
        : "";
      signals.push(label + dateStr);
    }
  }

  // Création récente (si pas déjà dans les publications)
  if (!signals.some(s => s.includes("🆕"))) {
    const dateCreation = company.date_creation || "";
    if (dateCreation) {
      const created = new Date(dateCreation);
      if (created >= twoYearsAgo) {
        signals.push("🆕 Création récente (" + created.toLocaleDateString("fr-FR", { month: "long", year: "numeric" }) + ")");
      }
    }
  }

  return signals.length
    ? signals.join("\n")
    : "Aucun signal BODACC récent (< 2 ans)";
}

// ── PROFIL ENTREPRISE — Classification transporteur vs chargeur ───────────────

// Codes NAF section H (Transport et entreposage)
const NAF_CARRIERS = new Set([
  "4941A","4941B","4941C", // Transport routier de fret
  "4942Z",                  // Déménagement
  "4950Z",                  // Transport par conduites
  "5010Z","5020Z",          // Transport maritime / fluvial
  "5030Z","5040Z",          // Transport aérien de passagers / fret
  "7712Z",                  // Location de camions
]);
const NAF_LOGISTICS = new Set([
  "5210A","5210B",          // Entreposage (frigorifique / non)
  "5221Z","5222Z","5223Z",  // Services auxiliaires terrestres / maritimes / aériens
  "5224A","5224B",          // Manutention portuaire / non portuaire
]);
const NAF_FORWARDING = new Set([
  "5229A",                  // Messagerie, fret express
  "5229B",                  // Affrètement et organisation des transports
]);

// Mots-clés dans le nom d'entreprise révélant un professionnel du transport
const CARRIER_KEYWORDS = [
  "transport","transports","logistique","logistics","logistic",
  "fret","freight","transitaire","commissionnaire",
  "affretement","affreteur","shipping","cargo","express",
  "messagerie","demenagement","livraison","groupage",
];

/**
 * Détecte le profil de l'entreprise pour personnaliser l'accroche.
 * Retourne : "carrier_transport" | "carrier_forwarding" | "carrier_logistics" |
 *            "shipper_industrial" | "shipper_trade" | "shipper_generic"
 */
function detectCompanyProfile(company, companyName) {
  const naf  = company.activite_principale || "";
  const name = normalizeStr(companyName);
  const lib  = normalizeStr(company.libelle_activite_principale || "");

  if (NAF_CARRIERS.has(naf))    return "carrier_transport";
  if (NAF_FORWARDING.has(naf))  return "carrier_forwarding";
  if (NAF_LOGISTICS.has(naf))   return "carrier_logistics";

  if (CARRIER_KEYWORDS.some(kw => name.includes(kw) || lib.includes(kw)))
    return "carrier_transport";

  const nafSect = naf.charAt(0);
  if (nafSect === "C") return "shipper_industrial";  // Industrie manufacturière
  if (nafSect === "G") return "shipper_trade";        // Commerce de gros/détail
  if (nafSect === "A") return "shipper_industrial";   // Agriculture
  return "shipper_generic";
}

/**
 * Extrait un hook contextuel depuis les articles d'actualité.
 * Retourne { theme, label } ou null.
 */
function extractNewsHook(newsArticles) {
  if (!newsArticles || !newsArticles.length) return null;
  const article = newsArticles[0];
  const title   = normalizeStr(article.title);
  const ref     = [article.source, article.dateLabel].filter(Boolean).join(", ");

  const THEMES = [
    { kws: ["acquisition","rachat","fusion","cession","reprise"],
      label: "Suite à votre actualité M&A" + (ref ? " (" + ref + ")" : "") + ", votre périmètre s'est élargi — avez-vous revu vos couvertures transport en conséquence ?" },
    { kws: ["international","export","etranger","overseas","mondial","import"],
      label: "Votre déploiement à l'international" + (ref ? " (" + ref + ")" : "") + " génère des risques spécifiques : transit douanier, Incoterms, P&I." },
    { kws: ["ouverture","nouveau site","entrepot","logistique","demenagement","expansion"],
      label: "Votre développement" + (ref ? " (" + ref + ")" : "") + " crée de nouveaux flux logistiques à sécuriser avant que les volumes augmentent." },
    { kws: ["croissance","developpement","hausse","record","chiffre affaires","performance"],
      label: "Votre forte croissance" + (ref ? " (" + ref + ")" : "") + " s'accompagne souvent d'un risque de sous-assurance — vos couvertures ont-elles suivi ?" },
    { kws: ["incendie","accident","sinistre","perte","vol","cambriolage"],
      label: "Face aux risques du secteur" + (ref ? " (" + ref + ")" : "") + ", la solidité de vos couvertures transport est un enjeu critique." },
    { kws: ["investissement","financement","levee de fonds","capital"],
      label: "Votre développement capitalistique" + (ref ? " (" + ref + ")" : "") + " est le bon moment pour sécuriser vos actifs marchandises." },
  ];

  for (const { kws, label } of THEMES) {
    if (kws.some(kw => title.includes(kw))) return label;
  }
  // Accroche générique sur l'actualité
  return "J'ai suivi l'actualité de " + (article.source ? "votre entreprise dans " + article.source + (article.dateLabel || "") : "votre entreprise") + " et souhaitais prendre contact.";
}

/**
 * Génère une accroche commerciale personnalisée pour GSA Prado.
 *
 * Structure :
 *   [Bonjour Prénom,]  — si interlocuteur connu via LinkedIn
 *   [Hook actualité ou BODACC]
 *   [Produits adaptés au profil : chargeur / transporteur / logisticien]
 *   [Positionnement GSA Prado]
 *   [CTA adapté au niveau de l'interlocuteur : DG / DAF / Dir. Logistique]
 *
 * @param {Object} company       — objet entreprise normalisé
 * @param {string} originalName  — nom tel que saisi
 * @param {Array}  newsArticles  — articles Google News
 * @param {Object} personInfo    — {name, titre, level, source} ou null
 */
function generateAccroche(company, originalName, newsArticles, personInfo) {
  const name    = cleanName(originalName);
  const profile = detectCompanyProfile(company, originalName);
  const naf     = company.libelle_activite_principale || "";
  const pubs    = company.publications || [];
  const GSA     = "GSA Prado — 1er courtier indépendant de la région Sud, partenaire Gallagher (105 pays)";

  // ── Salutation personnalisée ──────────────────────────────────────────────────
  const firstName = personInfo && personInfo.name
    ? personInfo.name.trim().split(/\s+/)[0]
    : null;
  const greeting = firstName ? "Bonjour " + firstName + ",\n\n" : "";

  // ── Hook (actualité > BODACC) ─────────────────────────────────────────────────
  const newsHook = extractNewsHook(newsArticles);
  let bodaccHook = "";
  if (!newsHook && pubs.length) {
    const twoYearsAgo = new Date();
    twoYearsAgo.setFullYear(twoYearsAgo.getFullYear() - 2);
    for (const pub of pubs.slice(0, 5)) {
      const type = pub.type || pub.type_publication || "";
      const date = pub.date ? new Date(pub.date) : null;
      if (date && date < twoYearsAgo) continue;
      const ds = date ? " (" + date.toLocaleDateString("fr-FR", { month: "long", year: "numeric" }) + ")" : "";
      if (type === "Achat" || type.includes("cquisition"))
        { bodaccHook = "Suite à votre acquisition (BODACC" + ds + "), votre périmètre transport s'est élargi."; break; }
      if (type === "Apport partiel d'actifs")
        { bodaccHook = "Votre apport d'actifs récent (BODACC" + ds + ") modifie votre exposition aux risques transport."; break; }
      if (type === "Création")
        { bodaccHook = "Félicitations pour le démarrage de " + name + ds + " ! C'est le bon moment pour poser les bonnes bases en assurance transport."; break; }
      if (type === "Déménagement")
        { bodaccHook = "Votre déménagement" + ds + " génère de nouveaux flux logistiques à sécuriser."; break; }
    }
  }
  const hook = newsHook || bodaccHook;
  const hookPart = hook ? hook + "\n\n" : "";

  // ── Intro si pas de hook ──────────────────────────────────────────────────────
  const DEFAULT_INTRO = {
    carrier_transport:  "Je souhaitais prendre contact au sujet de vos risques de responsabilité transport.\n\n",
    carrier_forwarding: "Je souhaitais évoquer vos obligations de responsabilité professionnelle.\n\n",
    carrier_logistics:  "Je souhaitais évoquer la protection de vos activités logistiques et d'entreposage.\n\n",
    shipper_industrial: "Je souhaitais prendre contact au sujet de la protection de vos marchandises.\n\n",
    shipper_trade:      "Je souhaitais évoquer la protection de vos stocks et flux de marchandises.\n\n",
    shipper_generic:    "Je souhaitais prendre contact au sujet de vos risques transport et marchandises.\n\n",
  };
  const intro = hookPart || DEFAULT_INTRO[profile] || DEFAULT_INTRO.shipper_generic;

  // ── Produits selon profil ────────────────────────────────────────────────────
  let products = "", valueProp = "";

  if (profile === "carrier_transport") {
    products  = "En tant que " + (naf || "transporteur") + ", vos expositions clés :\n";
    products += "• RC Transporteur (dommages aux marchandises confiées)\n";
    products += "• Police tiers chargeur\n";
    products += "• RC affrètement / sous-traitance";
    valueProp = GSA + " — accompagne de nombreux transporteurs avec une gestion des sinistres intégrée.";

  } else if (profile === "carrier_forwarding") {
    products  = "En tant que " + (naf || "commissionnaire / affréteur") + ", votre responsabilité envers vos clients est étendue :\n";
    products += "• RC pro commissionnaire de transport\n";
    products += "• RC affrètement\n";
    products += "• Police tiers chargeur";
    valueProp = GSA + " — couvre l'ensemble de vos risques, y compris vos flux internationaux.";

  } else if (profile === "carrier_logistics") {
    products  = "En tant que " + (naf || "logisticien / entrepositaire") + ", votre responsabilité couvre stockage ET flux :\n";
    products += "• RC dépositaire (marchandises confiées)\n";
    products += "• Stock & transit\n";
    products += "• RC Transporteur (si vous effectuez des livraisons)";
    valueProp = GSA + " — propose une approche sur mesure avec gestion des sinistres intégrée.";

  } else if (profile === "shipper_industrial") {
    products  = "En tant qu'industriel" + (naf ? " (" + naf + ")" : "") + ", vos marchandises sont exposées à chaque étape :\n";
    products += "• Dommages sur matières premières et produits finis (Ad valorem)\n";
    products += "• Stock & transit (entrepôts + flux)\n";
    products += "• P&I sur vos flux import / export";
    valueProp = GSA + " — spécialiste des risques transport — propose un audit gratuit de vos couvertures.";

  } else if (profile === "shipper_trade") {
    products  = "En tant que " + (naf || "négociant / distributeur") + ", vos marchandises sont exposées :\n";
    products += "• Lors du stockage (entrepôts, plateformes)\n";
    products += "• En transit (livraisons fournisseurs et clients)\n";
    products += "• Sur vos flux import / export";
    valueProp = GSA + " — vous propose des solutions Ad valorem et stock & transit adaptées à votre secteur.";

  } else {
    products  = GSA + " — accompagne les entreprises de tout secteur :\n";
    products += "• Assurance Ad valorem, stock & transit\n";
    products += "• RC Transporteur, commissionnaire, affrètement\n";
    products += "• P&I et couvertures internationales";
    valueProp = "";
  }

  // ── CTA personnalisé selon le niveau de l'interlocuteur ─────────────────────
  const personLevel = personInfo && personInfo.level;
  let cta = "";
  if (personLevel === 1) {
    // DG / PDG — angle stratégique
    const titre = personInfo.titre ? " (" + personInfo.titre + ")" : "";
    cta = "En tant que dirigeant" + titre + ", un sinistre non couvert peut impacter directement la continuité de votre activité. "
        + "Un point de 20 min suffit pour évaluer votre exposition réelle.";
  } else if (personLevel === 2) {
    // DAF / CFO / Risk Manager — angle financier
    const titre = personInfo.titre ? " (" + personInfo.titre + ")" : "";
    cta = "En tant que responsable financier" + titre + ", vous savez que le coût d'un sinistre non couvert dépasse celui de la prime. "
        + "Notre approche vous permettra d'optimiser vos couvertures sans sur-assurance.";
  } else if (personLevel === 3) {
    // Dir. Logistique / Supply Chain — angle opérationnel
    const titre = personInfo.titre ? " (" + personInfo.titre + ")" : "";
    cta = "Votre retour opérationnel sur vos flux" + titre + " serait précieux pour calibrer une solution vraiment adaptée. "
        + "Un échange de 20 min suffit — avec gestion des sinistres intégrée chez GSA Prado.";
  } else if (personInfo && personInfo.titre) {
    cta = "Seriez-vous disponible pour un échange de 20 min afin de vérifier que vos couvertures actuelles sont adaptées à votre activité ?";
  } else {
    cta = "Seriez-vous disponible pour un échange de 20 min ?";
  }

  // ── Assemblage final ─────────────────────────────────────────────────────────
  const parts = [greeting + intro + products, valueProp, cta].filter(Boolean);
  return parts.join("\n\n");
}

// ── LINKEDIN — Récupération du profil de l'interlocuteur ────────────────────────

/**
 * Tente de récupérer le nom et le titre depuis une page LinkedIn publique.
 * Utilise l'User-Agent LinkedInBot pour obtenir les balises OpenGraph.
 * Fallback : extraction depuis le slug de l'URL.
 * Retourne { name, titre, level, source } ou null.
 */
function fetchLinkedInProfile_(url) {
  if (!url || !url.includes("linkedin.com/in/")) return parseLinkedInUrl_(url);
  try {
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        "User-Agent":      "LinkedInBot/1.0 (+https://www.linkedin.com/help/linkedin/answer/a521883)",
        "Accept":          "text/html,application/xhtml+xml",
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
      }
    });
    if (resp.getResponseCode() !== 200) return parseLinkedInUrl_(url);

    const html = resp.getContentText();
    // LinkedIn sert og:title = "Prénom NOM - Titre chez Société | LinkedIn"
    const ogTitle = (
      html.match(/property="og:title"\s+content="([^"]+)"/) ||
      html.match(/content="([^"]+)"\s+property="og:title"/) ||
      [])[1] || "";

    if (!ogTitle || /log\s?in|sign\s?up|rejoindre/i.test(ogTitle)) {
      return parseLinkedInUrl_(url);
    }

    // Parse "Jean Dupont - Directeur Financier chez Acme | LinkedIn"
    const clean   = ogTitle.replace(/\s*\|\s*LinkedIn\s*$/i, "").trim();
    const dashIdx = clean.indexOf(" - ");
    const name    = (dashIdx > 0 ? clean.substring(0, dashIdx) : clean).trim();
    const roleStr = dashIdx > 0 ? clean.substring(dashIdx + 3).split(/ chez /i)[0].trim() : "";

    const roleNorm = normalizeStr(roleStr);
    let level = null;
    if (DIRECTOR_LEVELS[1].some(kw => roleNorm.includes(kw))) level = 1;
    else if (DIRECTOR_LEVELS[2].some(kw => roleNorm.includes(kw))) level = 2;
    else if (DIRECTOR_LEVELS[3].some(kw => roleNorm.includes(kw))) level = 3;

    return { name: name || null, titre: roleStr || null, level, source: "linkedin_og" };
  } catch(e) {
    return parseLinkedInUrl_(url);
  }
}

/**
 * Extrait nom et titre approximatifs depuis le slug de l'URL LinkedIn.
 * Exemple : /in/jean-dupont-daf-1a2b → { name: "Jean Dupont", level: 2, ... }
 */
function parseLinkedInUrl_(url) {
  const match = (url || "").match(/linkedin\.com\/in\/([^\/\?\#]+)/);
  if (!match) return null;

  const slug  = match[1].toLowerCase();
  const parts = slug.split("-");

  // Retire le suffixe de hash numérique de fin (ex: "1a2b3c")
  while (parts.length > 2 && /^[a-z0-9]{4,}$/.test(parts[parts.length - 1]) && /\d/.test(parts[parts.length - 1])) {
    parts.pop();
  }

  // Détection du niveau depuis des mots-clés dans le slug
  const SLUG_LEVELS = [
    { kw: ["dg","pdg","president","ceo","gerant","dirigeant"],                 level: 1, titre: "Directeur Général" },
    { kw: ["daf","cfo","financier","risk","assurance","controleur"],           level: 2, titre: "DAF / Risk Manager" },
    { kw: ["logistique","supply","transport","coo","operations","expeditions"], level: 3, titre: "Dir. Logistique / SC" },
  ];
  let detectedLevel = null, detectedTitre = null;
  for (const { kw, level, titre } of SLUG_LEVELS) {
    if (kw.some(k => slug.includes(k))) { detectedLevel = level; detectedTitre = titre; break; }
  }

  // Extraction du nom (premières parties alphabétiques, hors mots-clés titre)
  const SKIP = new Set(["dg","pdg","daf","cfo","coo","directeur","president","manager",
    "risk","logistique","transport","supply","chain","financier","general",
    "admin","commercial","gerant","dirigeant","assurance","controleur"]);
  const nameParts = [];
  for (const part of parts) {
    if (/^[a-z]{2,}$/.test(part) && !SKIP.has(part)) {
      nameParts.push(part.charAt(0).toUpperCase() + part.slice(1));
      if (nameParts.length >= 2) break;
    }
  }

  return {
    name:   nameParts.join(" ") || null,
    titre:  detectedTitre,
    level:  detectedLevel,
    source: "url_parse",
  };
}

/**
 * Récupère les informations sur l'interlocuteur depuis l'URL LinkedIn (col L).
 * Complète le niveau détecté en croisant avec les colonnes dirigeants (H/I/J).
 */
function getPersonInfo_(linkedinUrl, sheet, row) {
  if (!linkedinUrl || !linkedinUrl.includes("linkedin.com")) return null;
  const info = fetchLinkedInProfile_(linkedinUrl);
  if (!info) return null;

  // Si le niveau n'a pas été détecté, on croise avec les colonnes H/I/J
  if (!info.level && info.name) {
    const lastName  = normalizeStr(info.name.split(/\s+/).slice(-1)[0]);
    const dgText    = normalizeStr(sheet.getRange(row, CONFIG.COL_DG).getValue() || "");
    const dafText   = normalizeStr(sheet.getRange(row, CONFIG.COL_DAF).getValue() || "");
    const prescrTxt = normalizeStr(sheet.getRange(row, CONFIG.COL_PRESCR).getValue() || "");
    if (lastName.length >= 3 && dgText.includes(lastName))    info.level = 1;
    else if (lastName.length >= 3 && dafText.includes(lastName))   info.level = 2;
    else if (lastName.length >= 3 && prescrTxt.includes(lastName)) info.level = 3;
  }
  return info;
}

/**
 * Reconstruit un objet entreprise minimal depuis les cellules déjà remplies du Sheet.
 * Utilisé pour régénérer l'accroche sans rappeler l'API.
 */
function reconstructCompanyFromSheet_(sheet, row) {
  const secteur = (sheet.getRange(row, CONFIG.COL_SECTEUR).getValue() || "").toString();
  const nafMatch = secteur.match(/\b([A-Z]\d{3}[A-Z])\b/);
  return {
    activite_principale:         nafMatch ? nafMatch[1] : "",
    libelle_activite_principale: secteur.split(" — ")[0].trim() || secteur,
    siege:            { libelle_commune: "", code_postal: "" },
    publications:     [],
    chiffre_affaires: null,
    resultat_net:     null,
    evolution_ca:     "N/D",
    date_creation:    "",
    dirigeants:       [],
  };
}

function writeCompanyData(sheet, row, company, originalName, newsArticles) {
  const siege    = company.siege || {};
  const { dg, daf, prescr } = classifyDirectors(company.dirigeants || []);
  const effectif = company.libelle_tranche_effectif || company.tranche_effectif_salarie || "N/D";
  const signals  = extractSignals(company);
  const accroche = generateAccroche(company, originalName, newsArticles || []);

  const codeNaf    = siege.activite_principale || company.activite_principale || "";
  const libelleNaf = siege.libelle_activite_principale || company.libelle_activite_principale || "";
  const secteur    = [libelleNaf, codeNaf].filter(Boolean).join(" — ");
  const ville      = [siege.libelle_commune, siege.code_postal].filter(Boolean).join(" ") || "N/D";
  const liUrl      = "https://www.linkedin.com/search/results/companies/?keywords="
                   + encodeURIComponent(cleanName(originalName));
  const hasSignal  = signals.includes("🆕") || signals.includes("🤝") || signals.includes("🏗️") || signals.includes("🛒") || signals.includes("📍");

  // Formatage CA / résultat (Pappers en euros)
  const formatMoney = (val) => {
    if (val == null) return "N/D";
    if (Math.abs(val) >= 1e6) return (val / 1e6).toFixed(2) + " M€";
    if (Math.abs(val) >= 1e3) return (val / 1e3).toFixed(0) + " K€";
    return val + " €";
  };
  const ca      = formatMoney(company.chiffre_affaires);
  const res     = formatMoney(company.resultat_net);
  const evolCa  = company.evolution_ca || "N/D";
  const hasCA   = company.chiffre_affaires != null;

  const cells = [
    { col: CONFIG.COL_SECTEUR,  val: secteur || "N/D", bg: COLORS.PAPPERS  },
    { col: CONFIG.COL_VILLE,    val: ville,             bg: COLORS.PAPPERS  },
    { col: CONFIG.COL_CA,       val: ca,                bg: hasCA ? COLORS.PAPPERS : COLORS.GRAY },
    { col: CONFIG.COL_RESULTAT, val: res,               bg: hasCA ? COLORS.PAPPERS : COLORS.GRAY },
    { col: CONFIG.COL_EVOL_CA,  val: evolCa,            bg: hasCA ? COLORS.PAPPERS : COLORS.GRAY },
    { col: CONFIG.COL_EFFECTIF, val: effectif,          bg: COLORS.PAPPERS  },
    { col: CONFIG.COL_DG,       val: dg    || "Non identifié", bg: dg    ? COLORS.BLUE   : COLORS.GRAY },
    { col: CONFIG.COL_DAF,      val: daf   || "Non identifié", bg: daf   ? COLORS.PURPLE : COLORS.GRAY },
    { col: CONFIG.COL_PRESCR,   val: prescr|| "Non identifié", bg: prescr? COLORS.GREEN  : COLORS.GRAY },
    { col: CONFIG.COL_LI_ENT,   val: liUrl,             bg: COLORS.LINKEDIN },
    { col: CONFIG.COL_SIGNAL,   val: signals,           bg: hasSignal ? COLORS.SIGNAL : COLORS.GRAY },
    { col: CONFIG.COL_ACCROCHE, val: accroche,          bg: COLORS.YELLOW   },
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

  toast('🔍 Recherche Pappers pour "' + nom + '"…');
  try {
    const company = rechercheEntreprises(nom);
    if (!company) {
      sheet.getRange(row, CONFIG.COL_SECTEUR)
           .setValue("❌ Entreprise non trouvée (Pappers)")
           .setBackground(COLORS.RED);
      showAlert('"' + nom + '" introuvable dans la base Pappers.\nVérifiez l\'orthographe du nom.');
      return;
    }
    toast("✅ SIREN " + company.siren + " trouvé — actualités en cours…");
    const articles = fetchGoogleNews(nom);
    writeCompanyData(sheet, row, company, nom, articles);
    writeNewsCell(sheet.getRange(row, CONFIG.COL_ACTU), articles);
    SpreadsheetApp.flush();
    toast('✅ Enrichissement terminé pour "' + nom + '" (' + articles.length + ' article(s))');
  } catch (e) {
    sheet.getRange(row, CONFIG.COL_SECTEUR)
         .setValue("❌ Erreur : " + e.message).setBackground(COLORS.RED);
    showAlert("Erreur enrichissement : " + e.message);
  }
}

// ── ACTION : Enrichir toutes les lignes ───────────────────────────────────────
function enrichAllData() {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let done = 0, skipped = 0, errors = 0;

  toast("🔍 Enrichissement Pappers en cours…");

  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;
    const existing = sheet.getRange(row, CONFIG.COL_SECTEUR).getValue();
    if (existing && !existing.toString().startsWith("❌")) { skipped++; continue; }

    toast("[" + (row - 1) + "/" + (lastRow - 1) + "] " + nom + "…");
    try {
      const company = rechercheEntreprises(nom);
      if (!company) {
        sheet.getRange(row, CONFIG.COL_SECTEUR).setValue("❌ Non trouvé (Pappers)").setBackground(COLORS.RED);
        errors++;
      } else {
        const articles = fetchGoogleNews(nom);
        writeCompanyData(sheet, row, company, nom, articles);
        writeNewsCell(sheet.getRange(row, CONFIG.COL_ACTU), articles);
        done++;
      }
    } catch (e) {
      sheet.getRange(row, CONFIG.COL_SECTEUR)
           .setValue("❌ " + e.message).setBackground(COLORS.RED);
      errors++;
    }
    SpreadsheetApp.flush();
    Utilities.sleep(500);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ " + done + " enrichies | ⏭ " + skipped + " déjà faites | ❌ " + errors + " erreurs",
    "Pappers — Terminé", 8
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

  // 1 — Actualités Google News
  toast('📰 Actualités pour "' + nom + '"…');
  const articles = fetchGoogleNews(nom);
  writeNewsCell(sheet.getRange(row, CONFIG.COL_ACTU), articles);

  // 2 — Profil LinkedIn de l'interlocuteur (colonne L) si renseigné
  const linkedinUrl = (sheet.getRange(row, CONFIG.COL_LI_DEC).getValue() || "").toString().trim();
  let personInfo = null;
  if (linkedinUrl && linkedinUrl.includes("linkedin.com")) {
    toast("🔗 Récupération du profil LinkedIn…");
    personInfo = getPersonInfo_(linkedinUrl, sheet, row);
  }

  // 3 — Régénération de l'accroche (col O) avec nouvelles actualités + personne
  const company  = reconstructCompanyFromSheet_(sheet, row);
  const accroche = generateAccroche(company, nom, articles, personInfo);
  sheet.getRange(row, CONFIG.COL_ACCROCHE)
       .setValue(accroche).setBackground(COLORS.YELLOW)
       .setFontSize(9).setWrap(true).setVerticalAlignment("middle");

  sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 100));
  SpreadsheetApp.flush();

  const personStr = personInfo && personInfo.name
    ? " | " + personInfo.name + (personInfo.titre ? " — " + personInfo.titre : "")
    : "";
  toast("✅ " + articles.length + " article(s)" + personStr + " — accroche mise à jour");
}

// ── ACTION : Actualités — toutes les lignes ───────────────────────────────────
function enrichAllNews() {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  let done = 0;
  toast("📰 Actualités + accroches en cours…");

  for (let row = CONFIG.FIRST_ROW; row <= lastRow; row++) {
    const nom = sheet.getRange(row, CONFIG.COL_NOM).getValue();
    if (!nom) continue;

    toast("[" + (row - 1) + "/" + (lastRow - 1) + "] " + nom + "…");
    const articles = fetchGoogleNews(nom);
    writeNewsCell(sheet.getRange(row, CONFIG.COL_ACTU), articles);

    // Profil LinkedIn si renseigné
    const liUrl = (sheet.getRange(row, CONFIG.COL_LI_DEC).getValue() || "").toString().trim();
    const personInfo = liUrl && liUrl.includes("linkedin.com")
      ? getPersonInfo_(liUrl, sheet, row)
      : null;

    // Régénération accroche
    const company  = reconstructCompanyFromSheet_(sheet, row);
    const accroche = generateAccroche(company, nom, articles, personInfo);
    sheet.getRange(row, CONFIG.COL_ACCROCHE)
         .setValue(accroche).setBackground(COLORS.YELLOW)
         .setFontSize(9).setWrap(true).setVerticalAlignment("middle");

    sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 100));
    done++;
    SpreadsheetApp.flush();
    Utilities.sleep(600);
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast("✅ " + done + " lignes traitées (actualités + accroches)", "Actualités — Terminé", 5);
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
    { key: "PAPPERS_API_KEY", label: "Clé API Pappers (pappers.fr/api — gratuit 100 req/mois)", def: "" },
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
