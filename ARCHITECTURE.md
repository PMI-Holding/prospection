# Agent Prospection GSA Prado — Architecture

## Contexte métier
Georges-Eric Michel, chargé de clientèle chez GSA Prado (courtage assurance transport),
prospecte des entreprises susceptibles d'avoir besoin de couvertures assurance transport/marchandises.

## Stack technique

```
Google Sheets (Target list)          ← Point d'entrée : Georges ajoute les noms ici
       │
       │  Extensions → Apps Script
       ▼
ProspectionAgent.gs                  ← Fichier unique, tout le code est ici
       │
       ├── Pappers API               ← Enrichissement légal/financier (manuel, crédits payants)
       ├── Zeliq API                 ← Enrichissement email + téléphone (manuel, crédits payants)
       ├── Odoo XML-RPC API          ← Vérification + création leads CRM
       └── Google Search (UrlFetch)  ← Actualités entreprises
```

## Fichiers du repo

```
prospection/
├── ProspectionAgent.gs      ← FICHIER PRINCIPAL — tout le code Apps Script
├── appsscript.json          ← Config Apps Script (timezone, oauthScopes)
├── .clasp.json              ← Config clasp (scriptId pour le déploiement)
├── .gitignore               ← Exclut xlsx, py, pdf, xls
└── ARCHITECTURE.md          ← Ce fichier
```

## Déploiement

Le script est déployé via **clasp** :
```bash
clasp push    # pousse ProspectionAgent.gs → Apps Script (lié au Google Sheet)
git push      # pousse vers GitHub (source de vérité)
```

Script ID Apps Script : `1LiOD-UTH-ohmkLwA0N3IFLUiIgyd7ndcK5cELLaLBvzFz9USnPaRPAHo`
GitHub repo : `https://github.com/PMI-Holding/prospection`

## Clés API (jamais dans le code)

Stockées dans **PropertiesService.getScriptProperties()** d'Apps Script.
Saisies via menu 🚀 Prospection → ⚙️ Configurer les clés API.

| Clé | Service | Usage |
|-----|---------|-------|
| `ODOO_URL` | Odoo | `https://gsa-prado.odoo.com` |
| `ODOO_DB` | Odoo | `unikerp-gsaprado-prod-13772120` |
| `ODOO_USER` | Odoo | `georges-eric.michel@gsaprado.fr` |
| `ODOO_API_KEY` | Odoo | Clé API utilisateur Odoo |
| `ZELIQ_API_KEY` | Zeliq | Clé API enrichissement contacts |
| `PAPPERS_API_KEY` | Pappers | Clé API données légales (à ajouter) |

Récupérées via `getSecrets()` dans le code.

## Structure du Google Sheet "Target list"

**ID Sheet** : `1JEjyqMUbLFNdKVNMiSFT_zzWMS64xeVM0aLFTos5bM0`

### Colonnes actuelles (v1.1)

| Col | Lettre | Contenu | Géré par |
|-----|--------|---------|----------|
| 1 | A | Nom entreprise | Utilisateur (saisie manuelle) |
| 2 | B | Statut Odoo CRM | Script (checkOdoo) |
| 3 | C | Email (Zeliq) | Script (enrichZeliq) |
| 4 | D | Téléphone (Zeliq) | Script (enrichZeliq) |
| 5 | E | Statut enrichissement Zeliq | Script (enrichZeliq) |

### Colonnes cibles (v2.0 — à implémenter)

| Col | Lettre | Contenu | Géré par |
|-----|--------|---------|----------|
| 1 | A | Nom entreprise | Utilisateur |
| 2 | B | Secteur / Activité (NAF) | Pappers |
| 3 | C | Ville / Siège | Pappers |
| 4 | D | CA (M€) | Pappers |
| 5 | E | Résultat net (M€) | Pappers |
| 6 | F | Évolution CA | Pappers |
| 7 | G | Effectif | Pappers |
| 8 | H | 🥇 Décideur final (DG/PDG) | Pappers |
| 9 | I | 🥈 Décideur opérationnel (DAF/CFO/Risk Manager) | Pappers |
| 10 | J | 🥉 Prescripteur (Dir. Logistique/Supply Chain) | Pappers |
| 11 | K | LinkedIn entreprise | Pappers / manuel |
| 12 | L | LinkedIn décisionnaire (saisie manuelle) | Utilisateur |
| 13 | M | Signal / Déclencheur contextuel | Pappers BODACC |
| 14 | N | Actualité & sources | Pappers + Google News |
| 15 | O | Accroche commerciale recommandée | Script (généré) |
| 16 | P | Statut Odoo CRM | Script (checkOdoo) |
| 17 | Q | Email (Zeliq) | Script (enrichZeliq) |
| 18 | R | Téléphone (Zeliq) | Script (enrichZeliq) |
| 19 | S | Statut enrichissement Zeliq | Script (enrichZeliq) |

## Menu Apps Script (🚀 Prospection)

```
🚀 Prospection
├── 🏢 Enrichir Pappers — ligne sélectionnée    (enrichSelectedPappers)
├── 🏢 Enrichir Pappers — toutes les lignes     (enrichAllPappers)
├── ─────────────────────────────────────────
├── ⚡ Vérifier Odoo — toutes les lignes        (checkAllOdoo)
├── 🔵 Vérifier Odoo — ligne sélectionnée       (checkSelectedOdoo)
├── ─────────────────────────────────────────
├── 📧 Enrichir Zeliq — ligne sélectionnée      (enrichSelectedZeliq)
├── 📧 Enrichir Zeliq — toutes les lignes       (enrichAllZeliq)
├── ─────────────────────────────────────────
├── ➕ Créer lead Odoo — ligne sélectionnée     (createLeadSelected)
├── ─────────────────────────────────────────
├── 🏗️  Initialiser les en-têtes                (initHeaders)
├── ⚙️  Configurer les clés API                 (configureSecrets)
└── 📖  Guide d'utilisation                     (showHelp)
```

## APIs utilisées

### Pappers API
- Base URL : `https://api.pappers.fr/v2`
- Auth : `?api_token=PAPPERS_API_KEY`
- Endpoints utilisés :
  - `GET /entreprise?siren={siren}` — infos légales, finances, dirigeants
  - `GET /recherche?q={nom}` — recherche par nom
  - `GET /entreprise?siren={siren}&extrait_kbis=true` — publications BODACC
- **Crédits payants** → enrichissement uniquement manuel (jamais automatique)

### Zeliq API
- Base URL : `https://api.zeliq.com/api`
- Auth : header `x-api-key`
- Endpoints :
  - `POST /contact/enrich/phone` — enrichissement téléphone via LinkedIn URL
  - `POST /contact/enrich/email` — enrichissement email via LinkedIn URL ou nom+société
- **Crédits payants** → enrichissement uniquement manuel (jamais automatique)
- Input optionnel : `linkedin_url` (améliore la précision)

### Odoo API (JSON-RPC)
- URL : `https://gsa-prado.odoo.com/jsonrpc`
- Auth : authenticate → retourne uid, mis en cache 25 min
- Méthodes :
  - `crm.lead / search + read` — vérification opportunités existantes
  - `crm.lead / create` — création lead
- Recherche fuzzy : génère des variantes du nom (tolérance fautes d'orthographe)

## Règles importantes

1. **Enrichissement Pappers et Zeliq = TOUJOURS manuel** — jamais déclenché automatiquement
   (crédits payants, confirmation utilisateur obligatoire)

2. **Clés API = jamais dans le code** — toujours via `getSecrets()`

3. **`clasp push` après chaque modification** — sinon le Sheet n'est pas mis à jour

4. **Recherche Odoo fuzzy** — la fonction `searchVariants()` génère des variantes
   pour tolérer les fautes d'orthographe (ex: "Dupessay" trouve "Dupessey")

5. **Niveaux décisionnaires assurance** :
   - 🥇 Décideur final : DG, PDG, Président → signe les contrats
   - 🥈 Décideur opérationnel : DAF, CFO, Risk Manager → choisit les couvertures
   - 🥉 Prescripteur : Dir. Logistique, Supply Chain → exprime le besoin

6. **Signaux contextuels (BODACC)** à surveiller :
   - Nouveau dirigeant nommé → revue contrats probable
   - Acquisition / fusion → nouveau périmètre à couvrir
   - Ouverture de site → nouveaux flux à assurer
   - Forte croissance CA → risque de sous-assurance
   - Changement de forme juridique → révision couvertures

## Workflow utilisateur

```
1. Georges ajoute un nom en colonne A du Google Sheet
2. Menu → 🏢 Enrichir Pappers (ligne) → remplit B à O automatiquement
3. Menu → ⚡ Vérifier Odoo (ligne) → remplit colonne P
4. Si pas d'opportunité Odoo :
   a. Optionnel : coller URL LinkedIn décisionnaire en colonne L
   b. Menu → 📧 Enrichir Zeliq (ligne) → remplit Q et R
   c. Menu → ➕ Créer lead Odoo → crée l'opportunité dans le CRM
```
