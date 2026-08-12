# 🐔 ChickNSter — Système de réconciliation

Application web pour réconcilier les ventes du restaurant **ChickNSter** :
on dépose les 4 fichiers Excel d'une journée et le système vérifie que
**tout ce qui est encaissé correspond bien aux commandes**, avec les bons
montants et les bons modes de paiement.

## 🌐 Version en ligne (sans installation)

Le dossier [`docs/`](docs/) contient une version **100 % navigateur** : rien
à installer, tout le calcul se fait sur ta machine (aucun fichier n'est
envoyé sur Internet). Une fois GitHub Pages activé (voir plus bas), elle est
accessible à :

**https://anaslaghezali-ops.github.io/CNS/**

### Activer GitHub Pages (une seule fois)

1. Sur GitHub, ouvre le dépôt → onglet **Settings**.
2. Menu de gauche → **Pages**.
3. Section **Build and deployment** → **Source** → choisis **GitHub Actions**.
4. C'est tout : le site se déploie automatiquement (et se met à jour à chaque
   nouveau commit sur la branche).

---

## Les 4 sources

| Fichier | Rôle | Clé de rapprochement avec le POS |
|---------|------|----------------------------------|
| **POS (Mahaal)** | Fichier maître : toutes les transactions caisse | — |
| **Glovo** | Commandes de la marketplace Glovo | Ticket name `1-3 chiffres` · montant (col W − col AE) · heure (≤ 10 min après réception) |
| **NAPS (TPE)** | Relevé des paiements par carte | Date (col I) + montant (col N) → paiements `Credit card` |
| **Site** | Commandes du site web | Identifiant (col A) = ticket name |

## Règles de réconciliation

**Classification des tickets POS** (d'après le *Ticket name*) :

| Ticket name | Canal | Paiements attendus |
|-------------|-------|--------------------|
| `1-3 chiffres` (ex. `137`) | Glovo | Bank Transfer / Cash |
| `5 chiffres` (ex. `58379`) | Site | Bank Transfer |
| `sp…` / `emp…` | Sur place / À emporter | Cash / Credit card |
| *(vide)* | À rattacher | recherche par heure |

**Correspondances de paiement :**
- Glovo `Online` → POS `Bank Transfer` · Glovo `Cash` → POS `Cash`
- NAPS → toujours `Credit card` au POS
- Site → toujours `Bank Transfer` au POS

## Anomalies détectées

- 🔴 Commande **livrée absente** du POS (non tapée)
- 🔴 **Mauvais mode de paiement** au POS
- 🟠 **Numéro de commande mal saisi** (faute de frappe sur le n° de ticket)
- 🟠 **Écart de montant**
- 🟠 Ticket POS **non rattachable** à une source
- 🟠 Journée **non couverte** par le relevé NAPS (décalage de télécollecte)
- 🔵 Écart **agrégé** de paiement (Glovo) · **saisie tardive** (> 10 min)

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

Puis, dans le navigateur :
1. Déposer les 4 fichiers Excel (le **POS** est obligatoire, les autres sont facultatifs).
2. Cliquer sur **Lancer la réconciliation**.
3. Consulter les écarts à l'écran et **télécharger le rapport Excel**.

Le rapport Excel contient 3 feuilles : **Résumé**, **Anomalies**, et
**POS annoté** (le fichier POS avec le canal détecté, le statut et le
détail des incohérences, lignes surlignées).

## Structure du projet

```
CNS/
├── app.py                     # Interface web Streamlit
├── reconciliation/
│   ├── loaders.py             # Lecture robuste des 4 fichiers Excel
│   ├── classify.py            # Classification des tickets POS par canal
│   ├── reconcile.py           # Moteur de réconciliation
│   └── report.py              # Génération du rapport Excel
├── requirements.txt
└── README.md
```

> **Note sur les fichiers Glovo :** certains exports Glovo contiennent un
> `styles.xml` invalide qui empêche leur ouverture. Le système le **répare
> automatiquement** à la lecture.

## Confidentialité

Les fichiers réels du restaurant ne sont **pas** versionnés (voir
`.gitignore`). Le traitement se fait localement, en mémoire.
