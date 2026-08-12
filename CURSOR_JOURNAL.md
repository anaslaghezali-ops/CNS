# Journal des modifications — branche Cursor

> **Pour Claude / agents IA** : lire ce fichier en priorité avant de modifier le
> projet. Il décrit les changements métier et techniques faits en collaboration
> avec le gérant ChickNSter, le contexte, et les fichiers touchés.

## Convention de travail

- **Branche Git** : `Cursor` **uniquement** — toutes les modifications poussées sur
  `origin/Cursor`. Ne pas créer des branches `cursor/...` (préfixe cloud agent) :
  le gérant travaille sur une seule branche.
- **Ce journal** : chaque session de modification significative ajoute une entrée
  datée ci-dessous (règle métier, fichiers modifiés, raison, impact).
- **Parité Python / JS** : l’app utilisée en production est **`docs/`** (navigateur).
  Le package `reconciliation/` (Python) sert aux **scripts de diagnostic** et doit
  rester aligné sur `docs/reconcile.js` pour les règles métier.

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

**Note** : réconciliation financière dans **`docs/`** (navigateur).

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

### 2026-08-12 — Glovo : commandes annulées rapprochées au POS

**Demandé par** : gérant (ex. ticket POS `369` Cash 95 DH — commande Glovo annulée
`101736516017` non reconnue → fausse orpheline).

**Problème** : seules les commandes **Delivered** participaient au rapprochement.
Si le caissier tape la commande puis Glovo l'annule, le ticket POS restait « sans
commande correspondante ».

**Règle appliquée** :

- Après le rapprochement des **livrées**, les commandes **Cancelled** sont aussi
  comparées aux tickets Glovo POS encore libres (même fenêtre temps + montant + paiement).
- Match → **info** « Commande Glovo annulée — présente au POS » (pas une anomalie).
- Annulée **sans** ticket POS → pas d'alerte (normal).
- **Réconciliation financière** : annulées **tapées au POS** (`matched_pos`) comptées
  côté source Glovo (W − AE), en plus des livrées.

**Fichiers modifiés** : `reconciliation/reconcile.py`, `docs/reconcile.js`.

---

### 2026-08-12 — Glovo JS : exiger le montant en phase 1 (ticket 658 / 199 DH)

**Problème** : fausse orpheline ticket POS `658` (199 DH, 16:15) alors que Glovo
ligne 158 = commande `101736427596` (Online, 199 DH W−AE, reçue 16:12).

**Cause** : appariement global JS phase 1 classait par **temps seul** ; le ticket
`725` (120 DH à 16:11, +1 min) était préféré à `658` (199 DH à 16:15, +3 min).
La commande 199 DH était « consommée » par le mauvais ticket → 658 orphelin.

**Correction** : phase 1 (et saisie tardive / annulées) **exige** montant identique
(comme le moteur Python). Phase 2 (mauvais paiement) privilégie le montant sans
l'exiger.

**Fichiers modifiés** : `docs/reconcile.js` (`assignGlobalList`).

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

### 2026-08-12 — Distinction tickets POS même `ticket_name` (ex. deux « 725 »)

**Demandé par** : gérant (session Cursor).

**Problème** : deux lignes POS avec le même numéro saisi (`ticket_name` = 725) partageaient
les anomalies dans la table POS et les exports — impossible de voir que le 725 à 16:11
(120 DH, sans Glovo) est en anomalie alors que le 725 à 16:18 (108 DH) est OK.

**Règle appliquée** :

- Chaque anomalie liée au POS porte `pos_ticket_no` = colonne **Ticket No.** (identifiant
  unique de la ligne caisse), en plus de `ticket_name`.
- L’annotation POS (`statut`, colonne anomalies) groupe par **`ticket_no`**, pas par
  `ticket_name`.
- L’UI affiche `725 · #12345 · 16:11` pour distinguer visuellement.

**Fichiers modifiés** :

| Fichier | Changement |
|---------|------------|
| `docs/reconcile.js` | `posFields`, `posAnomaly`, `annotate` par ticket_no |
| `docs/app.js` | `fmtTicketLabel`, `anomalyMatchesPos`, export N° POS |
| `reconciliation/reconcile.py` | `_pos_ticket_kw`, `_annotate_pos` par ticket_no |

**Impact** : les deux tickets 725 du 07/08 ont chacun leur statut et leurs anomalies
isolées.

---

### 2026-08-12 — Classification `9sp` / `2sp` → Sur place

**Demandé par** : gérant (session Cursor).

**Problème** : tickets comme `9sp`, `2sp` (chiffre **avant** sp/emp) classés « Autre »
— ~9 % des tickets hors des contrôles Glovo/Site/sur place.

**Règle appliquée** : regex sur place étendue à `^\d*(sp|emp)\d*$` (compact, sans espaces)
→ reconnaît `sp`, `emp9`, `9sp`, `2sp`, `Sp 3`, etc.

**Workflow gérant** (noté, pas codé) : une fois Site + Glovo + TPE à 0 écart,
le reste (cash sur place, enveloppes caissiers) est vérifié **manuellement**.

**Fichiers modifiés** : `reconciliation/classify.py`, `docs/reconcile.js`, `README.md`

---

### 2026-08-12 — Glovo « absente » alors que ticket POS existe (écart W vs W−AE)

**Cas** : commande `101735404708` (ligne **87** du fichier Glovo) — Online 275 DH
(W−AE), 06/08 15:04 — signalée absente du POS.

**Cause** : pas la ligne POS 622 (ticket **365**, 95 DH → autre commande Glovo
`101735406241`). La commande 275 DH est sur la ligne POS **623** : ticket **80**,
300 DH, 15:04 — le caissier a tapé le **subtotal brut (W)** sans déduire la remise
AE (−25 DH). L’outil exige W−AE au POS → fausse « absente » + fausse orpheline.

**Correctif** : nouvelle passe Glovo « Écart de montant » (paiement + fenêtre OK,
montant différent) avec note si POS = W brut.

**Fichiers** : `docs/reconcile.js`, `reconciliation/reconcile.py`

---

### 2026-08-12 — Retrait Streamlit (app navigateur seule)

**Demandé par** : gérant — plus d’usage de Streamlit.

**Supprimé** : `app.py`, `reconciliation/report.py`, dépendance `streamlit` dans
`requirements.txt`.

**Conservé** : `docs/` (app production), `reconciliation/` (parité + scripts),
export Excel via `docs/app.js`.

**CI** : GitHub Pages déploie aussi sur push branche `Cursor`.

---

### 2026-08-12 — Contrôle produits (POS J / Glovo AY) + doublons clarifiés

**Demandé par** : gérant.

**Produits** : chargement `Designations (Reference)` (POS) et `Order Items` (Glovo).
Score de similarité affiché sur les anomalies liées — **contrôle secondaire**, pas
clé d'appariement. Suggestion par contenu sur commandes Glovo « absentes ».

**Doublons / correction** : le message indique quel paiement **retenir** selon la
commande Glovo (ex. ticket 878 → Glovo Cash → retenir v1 Cash, annuler v2 Bank
Transfer). L'écart financier retire la copie en trop du bon bucket de paiement.

**Fichiers** : `docs/reconcile.js`, `reconciliation/loaders.py`

---

### 2026-08-12 — Ticket 15 : pas un doublon (n° Glovo réutilisé pour Cash)

**Cas** : deux lignes POS « 15 » (122,5 DH Online + Cash séparé) classées comme
correction — en réalité la ligne 144 = Glovo `101739184927` (135 W−AE), la ligne
Cash = **autre vente** au même numéro (erreur de saisie).

**Règle** : doublon/correction seulement si même commande (montant, produits,
une seule commande Glovo). Sinon → anomalie **« Numéro Glovo réutilisé »** +
contrôle enveloppe cash.

**Logique anti-vol** (gérant) : le Cash POS doit sortir dans l'enveloppe ; le
risque principal est l'**Online / Bank Transfer** sans commande Glovo réelle.

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Doublon : préciser Cash vs Bank Transfer dans le surplus POS

**Demandé par** : gérant (ex. ticket 167 : +98 DH — en Cash ou BT ?).

**Changement** : message « Ticket en double (correction) » indique le détail par
mode de paiement, ex. `+98 DH → 98 DH en « Bank Transfer » — surplus uniquement
en Bank Transfer (pas en Cash)`.

**Fichiers** : `docs/reconcile.js`

---

## Points métier encore ouverts (session analyse 12 août)

Analyse sur fichiers réels (non versionnés) — pour référence future :

1. **Canal « Autre »** : formats encore non reconnus (ex. `Hibaster`). Les formats
   `9sp` / `2sp` sont maintenant classés sur place (voir entrée ci-dessus).
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
| Export Excel | — | `docs/app.js` |
| UI | — | `docs/index.html`, `docs/app.js` |
| Déploiement Pages | — | `.github/workflows/deploy-pages.yml` |

---

### 2026-08-12 — Écart financier Glovo Online / Site sans explication (bouton « Voir l'écart (0) »)

**Demandé par** : gérant (capture : Glovo Online −188 DH mais 0 ligne d'explication).

**Problèmes identifiés** :
1. **Syntaxe JS** : `flagMisusedGlovoNumber` sans déclaration de fonction (fichier non chargeable en Node).
2. **Site −398 DH** : tickets POS Site pour commandes existantes mais **non DELIVERED** (statut vide /
   « Fermée ») — comptés au POS, pas dans les livrées col L, **sans anomalie**.
3. **Glovo Cash** : commande **annulée** appariée au POS en phase 1 cancelled sans `matched_pos` →
   comptée au POS mais pas côté source Glovo Cash.
4. **Doublons** : surplus Bank Transfer (ex. ticket 878) pouvait générer une 2ᵉ ligne « sans
   appariement » en double du doublon.

**Corrections** (`docs/reconcile.js`) :
- Restauration de `flagMisusedGlovoNumber`.
- `markGlovoPosMatch` sur tous les appariements Glovo ; `matched_pos = true` si commande annulée.
- `appendFinancialGapAnomalies` : site non livré, Glovo Online/Cash sans appariement financier.
- `financialAdjustment` pour les nouveaux types + exclusion doublon dans le scan d'écart.

**Script** : `scripts/diagnose_financial.js` (données uploads).

---

### 2026-08-12 — Site emporter : Bank Transfer interdit au POS

**Demandé par** : gérant — col H « Delivery Method » (livraison / à emporter).

**Règle** : commande site **à emporter** → POS doit être Cash ou Credit card, **pas** Bank Transfer
(livraison site reste Bank Transfer).

**Fichiers** : `docs/reconcile.js`, `reconciliation/loaders.py`, `reconciliation/reconcile.py`

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
