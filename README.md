# 🐔 ChickNSter — Système de réconciliation

Application **navigateur** pour réconcilier les ventes du restaurant **ChickNSter** :
on dépose les 4 fichiers Excel d'une journée et le système vérifie que
**tout ce qui est encaissé correspond bien aux commandes**, avec les bons
montants et les bons modes de paiement.

**Aucune installation** : tout le calcul se fait dans le navigateur (les fichiers
ne partent pas sur Internet).

## Utilisation

1. Ouvrir [`docs/index.html`](docs/index.html) en local, ou le site GitHub Pages.
2. Déposer les 4 fichiers Excel (le **POS** est obligatoire, les autres facultatifs).
3. Cliquer sur **Lancer la réconciliation**.
4. Consulter les anomalies, valider celles traitées, télécharger le rapport Excel.

Redirection depuis la racine du dépôt : [`index.html`](index.html) → `docs/`.

## Version en ligne (GitHub Pages)

**https://anaslaghezali-ops.github.io/CNS/** (après activation de Pages)

1. GitHub → dépôt → **Settings** → **Pages**
2. **Source** → **GitHub Actions**
3. Le workflow déploie le dossier `docs/` (branche `Cursor` ou branche principale selon config CI)

---

## Les 4 sources

| Fichier | Rôle | Clé de rapprochement avec le POS |
|---------|------|----------------------------------|
| **POS (Mahaal)** | Fichier maître : toutes les transactions caisse | — |
| **Glovo** | Commandes de la marketplace Glovo | Ticket name `1-3 chiffres` · montant (col W − col AE) · heure (≤ 20 min après réception) |
| **NAPS (TPE)** | Relevé des paiements par carte | Date (col I) + montant (col N) → paiements `Credit card` |
| **Site** | Commandes du site web | Identifiant (col A) = ticket name |

## Règles de réconciliation

**Classification des tickets POS** (d'après le *Ticket name*) :

| Ticket name | Canal | Paiements attendus |
|-------------|-------|--------------------|
| `1-3 chiffres` (ex. `137`) | Glovo | Bank Transfer / Cash |
| `5 chiffres` (ex. `58379`) | Site | Bank Transfer |
| `sp…` / `emp…` (ex. `sp`, `9sp`, `emp2`) | Sur place / À emporter | Cash / Credit card |
| *(vide)* | À rattacher | recherche par heure |

**Correspondances de paiement :**
- Glovo `Online` → POS `Bank Transfer` · Glovo `Cash` → POS `Cash`
- NAPS → toujours `Credit card` au POS
- Site → toujours `Bank Transfer` au POS

## Anomalies détectées

- 🔴 Commande **livrée absente** du POS (non tapée)
- 🔴 **Mauvais mode de paiement** au POS
- 🟠 **Numéro de commande mal saisi** (faute de frappe sur le n° de ticket)
- 🟠 **Écart de montant** (Glovo W−AE vs POS, Site, etc.)
- 🟠 Ticket POS **non rattachable** à une source
- 🟠 Journée **non couverte** par le relevé NAPS (décalage de télécollecte)
- 🔵 Écart **agrégé** de paiement (Glovo) · **saisie tardive** (> 20 min)

## Structure du projet

```
CNS/
├── docs/                      # Application navigateur (moteur + UI + export Excel)
│   ├── index.html
│   ├── app.js
│   ├── reconcile.js           # Moteur de réconciliation (JavaScript)
│   └── style.css
├── reconciliation/            # Moteur Python (parité métier, scripts diagnostic)
│   ├── loaders.py
│   ├── classify.py
│   └── reconcile.py
├── scripts/                   # Outils CLI (ex. diagnostic commande Glovo)
├── requirements.txt           # pandas + openpyxl (scripts Python uniquement)
├── index.html                 # Redirection vers docs/
└── README.md
```

> **Note sur les fichiers Glovo :** certains exports Glovo contiennent un
> `styles.xml` invalide qui empêche leur ouverture. Le lecteur du navigateur
> (SheetJS) et le loader Python **réparent** automatiquement à la lecture.

## Confidentialité

Les fichiers réels du restaurant ne sont **pas** versionnés (voir
`.gitignore`). Le traitement se fait localement, en mémoire.

## Modifications Cursor (agent IA)

Les changements faits via **Cursor Cloud Agent** sont poussés sur la branche
`Cursor` et documentés dans [`CURSOR_JOURNAL.md`](CURSOR_JOURNAL.md) (règles
métier, fichiers touchés, contexte pour les sessions suivantes).
