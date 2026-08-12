# Journal des modifications — branche Cursor

> **Pour Claude / agents IA** : lire ce fichier en priorité avant de modifier le
> projet. Il décrit les changements métier et techniques faits en collaboration
> avec le gérant ChickNSter, le contexte, et les fichiers touchés.

## Convention de travail

- **Branche Git** : `Cursor` — toutes les modifications faites via Cursor Cloud
  Agent doivent être poussées sur cette branche (pas directement sur
  `claude/chicknster-reconciliation-system-460chx` sans merge explicite).
- **Ce journal** : chaque session de modification significative ajoute une entrée
  datée ci-dessous (règle métier, fichiers modifiés, raison, impact).
- **Parité Python / JS** : le moteur existe en double —
  `reconciliation/` (Streamlit) et `docs/reconcile.js` (navigateur). Toute règle
  métier doit être appliquée **dans les deux**.

---

## Contexte projet (rappel)

Application de **réconciliation caisse** ChickNSter : 4 fichiers Excel/jour.

| Fichier type | Export typique | Rôle |
|--------------|----------------|------|
| POS | `mahaal_sales_history*.xlsx` | Maître — toutes les transactions caisse |
| Glovo | `orderDetails*.xlsx` | Marketplace Glovo |
| NAPS | `MON_RELEVE_NAPSPRO*.xlsx` | Relevé TPE carte |
| Site | `orders*.xlsx` | Commandes site web |

La **période analysée** est toujours celle des dates présentes dans le POS ;
Glovo/Site hors période sont ignorés.

---

## Historique des modifications

### 2026-08-12 — Montant Glovo : colonne W − colonne AE

**Demandé par** : gérant / utilisateur (session Cursor).

**Problème** : le rapprochement Glovo comparait le POS au **Subtotal seul**
(colonne Excel **W**). Sur les exports `orderDetails`, le montant saisi au POS
correspond au subtotal **après déduction des remises financées par le
restaurant** (colonne **AE** — « Discount Funded by you »).

**Règle appliquée** :

```
montant_glovo = Subtotal (W) − Discount Funded by you (AE)
```

- Champ interne Python : `amount` (calculé dans `load_glovo`)
- Champ interne JS : `amount` (calculé dans `loadGlovo`)
- `subtotal` et `discount_funded` restent chargés pour traçabilité
- Si AE est vide → traité comme 0

**Fichiers modifiés** :

| Fichier | Changement |
|---------|------------|
| `reconciliation/loaders.py` | Charge AE, calcule `amount` |
| `reconciliation/reconcile.py` | Utilise `g["amount"]` au lieu de `g["subtotal"]` |
| `docs/reconcile.js` | Idem + libellé financier « Glovo (W − AE) » |
| `app.py` | Texte d'aide utilisateur |
| `README.md` | Table des sources |

**Impact observé** (fichiers test 5–10 août 2026) :

- 105 commandes avec remise AE > 0 (−1 401 DH au total)
- Total livrées : 63 414 DH (W) → **62 013 DH** (W − AE)
- Anomalies : 14 → 13 (commande Glovo Cash 140 DH retrouvée avec montant net)

**Commits** : `fe88541` sur `cursor/glovo-amount-w-minus-ae-32b0`, intégré
dans branche `Cursor`.

---

### 2026-08-12 — Réconciliation financière Glovo : écart Cash vs Online

**Demandé par** : gérant / utilisateur (session Cursor).

**Problème** : la ligne Glovo agrégée (−602 DH) ne permettait pas de savoir si
l'écart venait des commandes **Cash** ou **Online** (Bank Transfer au POS).

**Règle appliquée** : trois lignes dans le tableau financier :

| Ligne | Côté POS | Côté Glovo |
|-------|----------|------------|
| Glovo — Online | Tickets Glovo « Bank Transfer » | Livrées `Online` (W − AE) |
| Glovo — Cash | Tickets Glovo « Cash » | Livrées `Cash` (W − AE) |
| Glovo — Total | Tous tickets Glovo | Toutes livrées (W − AE) |

Écart = source − POS (identique aux autres lignes).

**Fichiers modifiés** :

| Fichier | Changement |
|---------|------------|
| `docs/reconcile.js` | `computeFinancial` : détail Online + Cash + total |
| `docs/app.js` | Styles lignes sous-total Glovo |
| `docs/index.html` | Texte d'aide |
| `docs/style.css` | `.fin-glovo-sub`, `.fin-total` |

**Note** : réconciliation financière **navigateur uniquement** (pas encore dans
Streamlit Python).

---

### 2026-08-12 — Validation des anomalies (hors calcul)

**Demandé par** : gérant / utilisateur (session Cursor).

**Problème** : une fois une anomalie traitée en caisse, elle restait dans les
compteurs et dans l'écart financier.

**Règle appliquée** :

- Bouton **✅ Valider** sur chaque anomalie (Haute + Moyenne).
- Validée → masquée de la liste active, section **Anomalies validées** avec
  **↩ Annuler**.
- **Vue d'ensemble** : recompte anomalies / gravités sans les validées.
- **Réconciliation financière** : ajustement des montants selon le type
  d'anomalie (ex. commande absente du POS → retire le montant côté source ;
  ticket POS orphelin → retire côté POS ; doublon → retire le surplus ; etc.).
- **POS annoté** : statut du ticket mis à jour si toutes ses anomalies sont
  validées.
- Export Excel : colonne « Validée », montants financiers ajustés.

**Fichiers modifiés** :

| Fichier | Changement |
|---------|------------|
| `docs/reconcile.js` | `anomaly.id`, `financialAdjustment`, `applyFinancialAdjustments` |
| `docs/app.js` | UI validation, recompte, export |
| `docs/index.html` | Section validées + texte d'aide |
| `docs/style.css` | Boutons valider / annuler |

**Note** : les validations sont **réinitialisées** à chaque nouvelle
réconciliation (nouveau clic « Lancer »). Pas de persistance entre sessions.

---

### 2026-08-12 — Détail des écarts financiers (anomalies sources)

**Demandé par** : gérant / utilisateur (session Cursor).

**Problème** : l'écart Glovo Online / Cash ne montrait pas **quelles anomalies**
le composaient (ex. −495 DH Online, −107 DH Cash).

**Règle appliquée** :

- Bouton **🔍 Voir l'écart (N)** sur chaque ligne de réconciliation financière.
- Panneau détaillé : liste des anomalies en attente qui **contribuent** à cet
  écart, avec **impact en DH** (+ = source > POS, − = POS > source).
- Total expliqué vs écart affiché ; validation possible depuis ce panneau.

**Fichiers modifiés** : `docs/reconcile.js` (`lineKey`, `getFinancialContributors`,
`ecartContributionForAnomaly`), `docs/app.js`, `docs/index.html`, `docs/style.css`.

---

### Antérieur (branche `claude/chicknster-reconciliation-system-460chx`)

Modifications déjà présentes **avant** la session Cursor du 12 août — ne pas
réintroduire l'ancienne logique sans vérifier :

| Commit | Résumé |
|--------|--------|
| `7e749f5` | Glovo : montant = col W seul (remplace une logique commission) — **partiellement remplacé par W−AE** |
| `955ff76` | Détection doublons / corrections (même n° tapé plusieurs fois) — **JS uniquement** (`detectDuplicates`) |
| `fd70570` | Réconciliation financière + fichier/ligne/date dans anomalies — **JS** |
| `a450507` | Fautes de frappe site (chiffre en trop/en moins) |
| `674313e` | Appariement Glovo global + section Infos séparée |
| `8ac4ab9` | Rattachement conjoint tickets sans numéro |
| `7cfc644` | Commande site non livrée au POS = pas une anomalie |
| `48c8626` | Fautes de frappe numéro site (5 chiffres) |
| `e21b13f` | Période d'analyse = dates du POS |

---

## Points métier encore ouverts (session analyse 12 août)

Analyse sur fichiers réels (non versionnés) — pour référence future :

1. **Canal « Autre »** (66 tickets) : formats `9sp`, `2sp`, `Hibaster` non
   reconnus par `classify.py` (regex attend `sp`/`emp` en **début** de ticket
   name). Impact sur classification sur place.
2. **NAPS 7 août** : écarts 89 DH (`9sur`), 119 DH (`Emp9`) POS vs 124 DH NAPS.
3. **Site** : commande `57767` (205 DH) absente du POS ; typos `58158`/`58159`,
   `57559`/`57599`.

---

## Où chercher dans le code

| Sujet | Python | JavaScript |
|-------|--------|------------|
| Lecture Excel | `reconciliation/loaders.py` | `docs/reconcile.js` (`loadPOS`, `loadGlovo`, …) |
| Classification ticket POS | `reconciliation/classify.py` | `docs/reconcile.js` (`classifyTicket`) |
| Rapprochement | `reconciliation/reconcile.py` | `docs/reconcile.js` (`reconcileSite`, `reconcileGlovo`, …) |
| Export Excel | `reconciliation/report.py` | `docs/app.js` |
| UI Streamlit | `app.py` | — |
| UI navigateur | — | `docs/index.html`, `docs/app.js` |
| Déploiement Pages | — | `.github/workflows/deploy-pages.yml` |

---

## Template pour les prochaines entrées

```markdown
### YYYY-MM-DD — Titre court

**Demandé par** : …
**Problème** : …
**Règle appliquée** : …
**Fichiers modifiés** : …
**Impact** : …
**Commits** : …
```
