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

### 2026-08-12 — Ticket 167 : doublon/correction (98 BT erroné → 110 Cash)

**Cas** : ticket 167 saisi 2 fois — v1 98 DH Bank Transfer (aucune commande Glovo) +
v2 110 DH Cash (commande Glovo `101734838230`). Classé « orphelin » au lieu de doublon.

**Cause** : `isLikelyCorrectionCluster` exigeait le même montant sur toutes les lignes non nulles,
ou une similarité produits qui échouait (format Glovo `Order Items` avec crochets).

**Règle** : si **une seule** commande Glovo matche une ligne du cluster, les autres lignes avec
**mauvais paiement** (ex. BT vs Cash attendu) = correction ; Cash 0 DH = vente distincte (ticket 15).

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Site emporter : pas d’anomalie « statut non livré »

**Demandé par** : gérant (ex. commande `57388` emporter Fermée + Cash au POS).

**Règle** : les commandes **à emporter** (col H) n’ont pas `DELIVERED` col L — statut
**Fermée** = normale. Elles doivent être au POS en Cash ou Credit card.

**Changements** :
- `siteOrderCountsInReconciliation` : livraison = DELIVERED · emporter = Fermée
- Réconciliation site + total financier incluent les emporter Fermées
- Plus d’anomalie « statut non livré » pour les emporter

**Fichiers** : `docs/reconcile.js`, `reconciliation/reconcile.py`

---

### 2026-08-12 — Réconciliation par journée (onglets + export Excel)

**Demandé par** : gérant — analyser jour par jour (ex. fichiers du 5 au 10 août = 6 journées)
tout en gardant le récapitulatif général.

**Fonctionnement** :
- Après réconciliation, onglets **📊 Général** + **📅 &lt;date&gt;** pour chaque journée POS détectée
  (affichés seulement s’il y a **plus d’une** journée).
- Vue journée = même contenu que le général (métriques, réconciliation financière, anomalies,
  infos, POS annoté) mais POS filtré sur la date ; Glovo / Site / NAPS alignés sur cette journée.
- Validation d’anomalies : état global partagé (`state.validated`) — un ID validé en vue jour
  est validé partout.
- Export Excel : feuilles globales + par jour (`05/08 Résumé`, `05/08 Réconcil`, `05/08 Anomalies`, etc.).

**Fichiers** : `docs/reconcile.js` (`runDailyBreakdown`, `listPosDates`), `docs/app.js`,
`docs/index.html`, `docs/style.css`

---

### 2026-08-12 — TPE NAPS : section « Totaux alignés » (appariement sans écart)

**Demandé par** : gérant — cas 110 DH POS CB vs 95+15 DH NAPS : pas d'écart financier,
mais 3 anomalies d'appariement. Objectif : voir immédiatement que c'est OK si total POS CB = total NAPS.

**Règle** : si sur une journée `sum(POS Credit card) = sum(NAPS)` mais l'appariement
transaction par transaction échoue, marquer les anomalies `naps_totals_ok` et afficher
une section verte **✅ TPE — Totaux alignés** avec le détail des deux côtés.

**UI** :
- Section entre réconciliation financière et anomalies
- Compteurs : ces lignes ne comptent plus dans « Anomalies » mais dans « TPE totaux OK »
- Table anomalies : sous-section avec badge ✅ Totaux OK
- Ligne financière NAPS : note si totaux alignés

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `docs/style.css`

---

### 2026-08-12 — Totaux POS par mode de paiement (split canal)

**Demandé par** : gérant — vue synthétique des totaux POS Cash / CB / Bank Transfer
avec ventilation par canal sous chaque mode :
- Cash → Glovo · SP&EMP · Site
- CB → SP&EMP · Site
- Bank Transfer → Glovo · Site

**UI** : cartes sous la vue d'ensemble + export Excel (`Totaux paiement` / `JJ/MM Paiements`).

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `docs/style.css`

---

### 2026-08-12 — Cash à collecter (écarts Cash Glovo / Site emporter)

**Demandé par** : gérant — savoir combien récupérer des caissiers en plus du Cash POS
(ex. Glovo Cash +5 DH = 5 DH à collecter au-delà du POS).

**Règle** : `Cash à collecter = écarts positifs (source − POS)` sur Glovo Cash et Site emporter Cash.
`Cash réel attendu = total Cash POS + net à collecter`.

**UI** : section orange sous totaux paiement + métrique en tête si &gt; 0. Export Excel.

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `docs/style.css`

---

### 2026-08-12 — Cash à collecter : POS CB > NAPS (TPE)

**Demandé par** : gérant — si le POS Credit card &gt; relevé NAPS, l'écart est du cash à collecter.
Le total se recalcule après validation d'anomalies (ligne TPE ajustée comme la réconciliation financière).

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`

---

### 2026-08-12 — Paiements POS fractionnés (Cash, Cash / Cash + CB)

**Demandé par** : gérant — « Cash, Cash, Cash » = fractionnement cash (pas une anomalie).
« Cash, Credit card » = split cash + carte : la ligne NAPS non appariée du jour = part CB,
le reste = cash.

**Règles** :
- `parsePaymentTypes` sur la colonne Payment type
- Plus d’anomalie « mode inattendu » si tous les modes sont Cash ou Credit card
- Split Cash+CB : appariement NAPS (montant &lt; total ticket) → `_naps_split_cc` sur le ticket POS
- Totaux financiers ventilés (part CB / part Cash)

**Fichiers** : `docs/reconcile.js`, `reconciliation/reconcile.py`

---

### 2026-08-12 — Glovo Total POS = Online + Cash

**Problème** : ligne « Glovo Total » POS = somme brute tous tickets Glovo (incl. CB),
alors que Online + Cash ne comptaient que BT et Cash → totaux POS incohérents.

**Fix** : Total POS = Bank Transfer Glovo + Cash Glovo (comme la source).
Note si tickets Glovo en CB/autre hors décomposition Online/Cash.

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Canal « Autre » → SP&EMP + liste tickets hors Glovo/Site

**Demandé par** : gérant — tickets ni Glovo ni Site = sur place (pas de source externe).
Lister les libellés libres et « À rattacher ».

**Règles** :
- Classification : défaut `Sur place / Emporter` (plus de canal « Autre »)
- Section **SP&EMP hors Glovo/Site** : libellés libres + tickets à rattacher
- Export Excel `SP&EMP hors Glovo-Site`

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `reconciliation/classify.py`

---

### 2026-08-12 — Info Glovo « écart cash » vs écart financier 0 DH

**Problème** : info « Glovo Cash 25 vs POS 26 → écart -1 » alors que réconciliation
financière Cash = 0 DH.

**Explication** : l'info compare le **nombre** de commandes / tickets ; la réconciliation
financière compare les **montants** en DH (peuvent être égaux avec un ticket en plus ou en moins).

**Fix** : message info explicite + montants Glovo vs POS dans le détail.

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Bipeur SP&EMP vs faux Glovo (ex. ticket « 14 » CB + NAPS)

**Cas** : ticket POS `14` classé Glovo (1–3 chiffres) sans commande Glovo, mais 230 DH CB
sur NAPS sans doublon → bipeur sur place, pas vol.

**Règle** : si orphelin Glovo + numéro 1–3 chiffres + CB + **une seule** ligne NAPS au même
montant sur la journée → reclasse **SP&EMP** (`spemp_bipeur`), info au lieu d’anomalie Glovo.

**Fichiers** : `docs/reconcile.js`, `docs/app.js`

---

### 2026-08-12 — Montant affiché sur toutes les anomalies

**Demandé par** : gérant — toujours voir le montant (ex. SP&EMP Bank Transfer).

**Changements** :
- Colonne **Montant** dans les tableaux anomalies / validées / infos
- `posAnomaly` remplit `amount_pos` depuis le ticket si absent
- Détail enrichi automatiquement avec les DH si non mentionnés

**Fichiers** : `docs/reconcile.js`, `docs/app.js`

---

### 2026-08-12 — Fix faux +80 DH « Cash à collecter » Site emporter (Cash + CB)

**Problème** : réconciliation Site = 0 DH écart mais « Cash à collecter » +80 DH
(Site emporter 190 vs POS Cash 110). Les 80 DH étaient déjà saisis en **Credit card** Site.

**Règle** : emporter comptoir = comparer fichier site à POS **Cash + Credit card** (pas Cash seul).
Livraison site reste Bank Transfer (hors cette ligne).

**Fichiers** : `docs/reconcile.js`, `docs/index.html`

---

### 2026-08-12 — Tickets non rattachés Cash / CB+NAPS → SP&EMP

**Demandé par** : gérant — ticket non rattaché payé Cash = sur place/emporter ;
ticket CB non rattaché avec ligne NAPS correspondante = sur place/emporter.

**Règles** :
- Après passe à rattacher : Cash seul → `spemp_unattached_cash`
- CB avec **une seule** ligne NAPS (jour + montant) → `spemp_unattached_naps`
- Idem pour orphelins Glovo/Site sans commande (pas seulement « Ticket » vide)
- Info « Ticket non rattaché → SP&EMP » au lieu d’anomalie

**Fichiers** : `docs/reconcile.js`, `docs/app.js`

---

### 2026-08-12 — Cash à collecter ventilé par utilisateur POS (col. F)

**Demandé par** : gérant — savoir **chez qui** récupérer le cash (qui a fait l’erreur).

**Règles** : chaque écart cash à collecter est rattaché au **User** du ticket POS (colonne F) :
- Glovo Cash : sous-saisie / commande absente / sur-saisie par ticket
- Site emporter : écart comptoir par ticket site
- TPE CB &gt; NAPS : tickets CB sans ligne NAPS

**UI** : tableaux par utilisateur + détail tickets · Excel `Cash par utilisateur` / `Cash tickets détail`.

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `docs/style.css`

---

### 2026-08-12 — Affectation manuelle cash « Non attribué » → User

**Demandé par** : gérant — une fois la preuve trouvée, affecter une ligne Non attribué
à un utilisateur POS ; le montant se reporte dans le tableau par user.

**UI** : section « À clarifier », liste déroulante des Users POS, badge Manuel + Annuler.
**Excel** : colonne Affectation sur « Cash tickets détail ».

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `docs/style.css`

---

### 2026-08-12 — Fix Glovo : mauvais paiement exige même montant + fenêtre 20 min

**Problème** : commande Online 104 DH rattachée au ticket Cash 55 DH (anomalie paiement)
alors que le ticket BT 115 DH était la bonne commande (4 min plus tard).

**Cause** : phase 2 « mode incorrect » appariement sans exiger le montant — le ticket
le plus proche en temps (mauvais montant) volait la commande.

**Fix** : phase 2 exige W−AE = total POS · fenêtre après réception **10 → 20 min**.

**Fichiers** : `docs/reconcile.js`, `reconciliation/reconcile.py`, `README.md`

---

### 2026-08-12 — Anomalies Glovo : heure de réception dans le détail

**Demandé par** : gérant — sur écart de montant / mode incorrect, afficher l’heure
Glovo comme l’heure POS (« reçue à 21:22 » vs « ticket 16 à 21:34 »).

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Journée POS 03:00 → 02:59 (pas minuit)

**Demandé par** : gérant — commande Glovo 23:59 tapée après minuit = même journée caisse.

**Règle** : date POS = `posBusinessDateKey` (ticket − 3 h). Onglets journaliers,
filtre Glovo/Site sur la période POS, ventilation cash/NAPS POS côté caisse.

**Fichiers** : `docs/reconcile.js`, `docs/index.html`

---

### 2026-08-12 — Cash à collecter TPE : user depuis anomalie NAPS

**Problème** : +164 DH à collecter sans nom de caissier ; l’anomalie « Paiement POS
absent du TPE » (ticket Sp5) a un User POS.

**Fix** : ventilation par user via les anomalies NAPS (User col. F) · encart
« À récupérer par utilisateur » en tête de section.

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/style.css`

---

### 2026-08-12 — Cash à collecter : nom du caissier dans le bandeau résumé

**Demandé par** : gérant — +164 DH visible mais pas « chez qui » (ex. Hiba / Sp5).

**Fix** : sous « À collecter des caissiers », chips **Hiba +164 DH** directement dans
le bandeau ; correction ordre `site`/`naps` dans `enrichFinCashCollect`.

**Fichiers** : `docs/app.js`, `docs/style.css`

---

### 2026-08-12 — Site emporter : caissier sur la ligne +130 DH

**Demandé par** : gérant — +130 DH Site emporter (POS comptoir 0 DH) sans nom de caissier
(ex. 19 juil. : commande 54767, User Hiba).

**Fix** : attribution Site emporter via ticket POS (tous canaux), anomalies Site,
appariement montant+heure si n° absent · ligne détail « À récupérer : Hiba +130 DH ».

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/style.css`

---

### 2026-08-12 — Suggestion Glovo : heure du ticket POS proposé

**Demandé par** : gérant — suggestion « ticket 50 ligne 1301 » sans l’heure POS.

**Fix** : `à HH:MM` ajouté dans le texte de suggestion (ex. `ligne 1301 à 21:42`).

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Glovo : numéro mal saisi (SP/EMP au lieu du n° commande)

**Demandé par** : gérant — commande 101718699078 (190 DH Online 21:55) absente +
ticket Sp235 (BT 190 DH 21:56) classé sur place ; écart nombre 48 vs 47.

**Règle** : phase 2b — appariement montant + paiement + fenêtre sur tickets SP/EMP/libre
(avant écart montant) · anomalie « Numéro Glovo mal saisi » · suggestion absente
priorise montant+paiement+heure avant produits.

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Glovo : écart de nombre = erreur (strict)

**Demandé par** : gérant — nombre commandes Glovo doit égaler tickets POS Glovo ;
sinon erreur de saisie du nom de ticket.

**Règle** : anomalies **haute** si total livrées ≠ tickets canal Glovo, ou Online≠BT,
Cash≠Cash · message explicite (mauvais SP/EMP, etc.) + détail appariements manquants.

**Fichiers** : `docs/reconcile.js`, `reconciliation/reconcile.py`

---

### 2026-08-12 — Saisie POS dans l'heure + n° Glovo mal saisi (correctif)

**Demandé par** : gérant — commande 101718699078 (Online 190 DH, 21:55) restait
« absente du POS » avec une suggestion à **23:09** (+74 min), alors qu'elle correspond
au ticket **Sp235** tapé à **21:56** pour exactement 190 DH.

**Règles** :
1. `MAX_LATE_ENTRY_MIN = 60` — impossible de taper une commande Glovo/Site plus d'une
   heure après réception : plus de rapprochement (saisie tardive, annulée, doublon)
   ni de **suggestion** au-delà de 60 min (avant : 120 min).
2. Phase **2b** avant les phases tardives / écart de montant : commande livrée non
   appariée + ticket **SP/EMP ou libre** au **même montant** et **même mode de paiement**
   dans l'heure ⇒ anomalie **haute** « Numéro Glovo mal saisi au POS », ticket reclassé
   **Glovo** (2 passes : 20 min puis 60 min). Les tickets « à rattacher » restent à la
   passe commune Glovo+Site.
3. `AMOUNT_MISMATCH_MAX_RATIO = 0.3` — un « Écart de montant » n'est retenu que si le
   POS est proche du montant réel (ou = subtotal brut W) ; sinon la commande reste
   « absente du POS » au lieu d'être appariée à tort.

**Effet** : le ticket compte désormais en **Glovo Bank Transfer** (et non SP&EMP) →
l'écart financier Glovo Online de 190 DH et l'écart de nombre 48/47 se résorbent.

**Fichiers** : `docs/reconcile.js`, `docs/index.html`, `reconciliation/reconcile.py`

---

### 2026-08-12 — Suppression du graphique des canaux

**Demandé par** : gérant — graphique à barres (Glovo / SP&EMP / Site) inutile.

**Fix** : remplacé par une ligne texte « Tickets POS par canal : Glovo 75 · … »
(`#chart` → `#channels`), CSS `.chart` / `.bar` supprimé. L'export Excel garde le
détail par canal.

**Fichiers** : `docs/index.html`, `docs/app.js`, `docs/style.css`

---

### 2026-08-12 — Passe finale : recouper les anomalies entre elles

**Demandé par** : gérant — « une fois qu'il termine, il doit refaire un tour sur les
anomalies pour voir si il n'y a pas des anomalies connectées entre elles ». Cas resté
non résolu : commande Glovo 101718699078 (Online 190 DH, 21:55) « absente du POS » +
ticket Sp235 (BT 190 DH, 21:56) « mode de paiement inattendu » = même opération.

**Causes du non-appariement** :
1. Les phases Glovo ne comparaient que les tickets **déjà classés Glovo** (n° 1–3
   chiffres) ; un ticket nommé « Sp235 » n'était dans aucun pool.
2. Le « même montant » exigeait ≤ 0,5 DH : les centimes de Glovo (W−AE) contre un
   total POS arrondi pouvaient dépasser cette tolérance.

**Règles ajoutées** :
- `SAME_AMOUNT_TOL = 1` DH (`sameAmountRounded`) pour « même montant » à l'arrondi près.
- `linkRelatedAnomalies(pos, glovo, site, anomalies)` — passe finale, **avant** les
  contrôles de nombre : croise les anomalies « commande absente du POS » (Glovo/Site)
  avec les anomalies POS restantes (paiement inattendu SP/EMP, ticket sans numéro,
  ticket orphelin Glovo/Site, ticket non rattaché). Critères : **même montant** (±1 DH),
  **≤ 60 min**, et **mode de paiement attendu OU produits ≥ 30 % compatibles**.
  Les 2 anomalies sont **fusionnées** en une seule « Numéro de ticket mal saisi (Glovo/Site) »
  (haute), le ticket est reclassé Glovo/Site et la commande marquée appariée.
  Si le mode de paiement diffère aussi, un ⚠️ le signale dans le détail.

**Impact** : le ticket compte dans le bon canal (ex. +190 DH en Glovo Bank Transfer,
−190 DH en SP&EMP), les écarts de nombre et l'écart financier se résorbent.
Non-régression vérifiée sur POS+NAPS+Site réels (174 anomalies, 0 fusion parasite).

**Fichiers** : `docs/reconcile.js`

---

### 2026-08-12 — Rapprochement tolérant + version affichée

**Problème** : le gérant voyait toujours « Commande Glovo absente du POS » (101718699078)
+ « Mode de paiement inattendu » (Sp235). L'absence de **suggestion** dans le détail
prouvait que l'écart de montant réel dépassait la tolérance (0,5 puis 1 DH), les deux
montants étant pourtant affichés « 190 DH » (arrondi à l'entier partout dans l'UI).

**Règles** :
- Passe finale : un ticket **Bank Transfer** hors canal Glovo/Site (sur place, emporter,
  sans numéro) ne peut venir que d'une commande Glovo/Site ⇒ le montant peut être
  **approché** (≤ 30 %) au lieu d'exact ; le détail affiche alors les deux montants
  au centime et l'écart. Idem si les produits sont ≥ 30 % compatibles.
- Le coût d'appariement pénalise l'écart de montant (`+ diff * 5`) : un montant exact
  gagne toujours contre un montant approché.
- Anomalie « absente du POS » sans candidat : **diagnostic** avec le ticket le plus
  proche au bon mode de paiement et les montants **exacts au centime** (« 190,00 DH au
  POS vs 189,00 DH côté Glovo, écart 1,00 DH »).

**Cache navigateur** : `?v=20260812-5` sur `reconcile.js` / `app.js` / `style.css` et
**version affichée en pied de page** (`CNS.BUILD`) pour vérifier le code réellement chargé.

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`

---

### 2026-08-12 — 2ᵉ passe visible dans l'UI (rapport de rapprochement)

**Demandé par** : gérant — « après la 1ʳᵉ réconciliation, y a-t-il une 2ᵉ tentative de
rapprochement des anomalies qui ont exactement les mêmes montants ? ». Cas Glovo
101718699078 / Sp235 toujours non rapproché sur ses données, sans explication visible.

**Ajouts** :
- `run()` renvoie **`link_report`** = `{ merged, unresolved }`, alimenté par les **deux**
  passes : 1ʳᵉ passe (`assignWrongTicketName` dans `reconcileGlovo`) et 2ᵉ passe
  (`linkRelatedAnomalies`).
- La 2ᵉ passe évalue **tous** les tickets POS à moins d'une heure (plus seulement ceux
  déjà porteurs d'une anomalie) et enregistre pour chacun : heure, écart temps, montant
  POS **au centime**, écart de montant, paiement, canal, similitude produits, anomalie
  POS existante, **verdict** (« rapprochable » ou raison du refus).
- Nouvelle section UI **« 🔗 2ᵉ passe — rapprochement des anomalies entre elles »** :
  tableau des fusions (avec la passe d'origine) + par commande non résolue, la liste
  des candidats et le motif de rejet.
- Rappel : un ticket **sans anomalie** n'est jamais rapproché automatiquement (ce serait
  reclasser une vente comptoir normale) mais il apparaît au rapport.

**Fichiers** : `docs/reconcile.js`, `docs/app.js`, `docs/index.html`, `docs/style.css`
(version `2026-08-12 · 6`)

---

### 2026-08-12 — 🐞 CAUSE RACINE : état d'appariement partagé entre les runs

**Symptôme** : sur l'**onglet journalier** (19 juil.), la commande Glovo 101718699078
restait « absente du POS » et Sp235 gardait « Mode de paiement inattendu », alors que la
**vue générale** rapprochait correctement les deux. Le rapport de 2ᵉ passe a donné le
verdict décisif : candidat Sp235, 190,00 DH, 0,00 DH d'écart, 1 min →
**« déjà rattaché à la commande 101718699078 »**.

**Cause** : `app.js` appelle `run(pos, …)` (vue générale) **puis**
`runDailyBreakdown(pos, …)`, qui rappelait `run()` sur **les mêmes objets** de ligne.
`addChannel()` réinitialisait `p.channel` mais **pas** `p.matched_glovo_order` /
`g.matched_pos_ticket_no` : le ticket arrivait « déjà rattaché » dans le run du jour,
donc écarté de la 1ʳᵉ passe (nom de ticket) et de la 2ᵉ passe. Effet miroir : les runs
journaliers écrasaient aussi l'état de la vue générale (dernier jour gagnant).

**Fix** :
1. `resetReconcileState(pos, glovo)` en tête de `run()` — efface `matched_glovo_order`,
   `matched_pos_ticket_no`, `matched_pos`, `_naps_split_*`, `statut`, `anomalies`.
2. `runDailyBreakdown()` réconcilie chaque journée sur des **copies** (`cloneRows`)
   de POS / Glovo / NAPS / Site (`null` préservé = fichier absent), pour que les onglets
   n'altèrent plus la vue générale ni les autres journées.

**Vérifié** : vue générale et onglet 19 juil. donnent le même verdict (Sp235 → Glovo,
« Numéro Glovo mal saisi au POS », plus d'anomalie SP/EMP) ; résultats **idempotents**
sur 2 exécutions consécutives ; non-régression POS+NAPS+Site (174 anomalies, 26 jours).

**Fichiers** : `docs/reconcile.js` (version `2026-08-12 · 7`)

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
