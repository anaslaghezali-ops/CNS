"""
ChickNSter — Système de réconciliation.

Interface web : dépose les 4 fichiers Excel (POS, Glovo, NAPS, Site),
lance la réconciliation, visualise les écarts et télécharge le rapport.

Lancement :  streamlit run app.py
"""

from __future__ import annotations

import datetime as dt

import pandas as pd
import streamlit as st

from reconciliation import load_pos, load_glovo, load_naps, load_site, run_reconciliation
from reconciliation.report import build_excel_report

st.set_page_config(page_title="Réconciliation ChickNSter", page_icon="🐔", layout="wide")

SEV_ORDER = {"haute": 0, "moyenne": 1, "info": 2}
SEV_BADGE = {"haute": "🔴 Haute", "moyenne": "🟠 Moyenne", "info": "🔵 Info"}


# --------------------------------------------------------------------------- #
# En-tête
# --------------------------------------------------------------------------- #
st.markdown(
    """
    <div style="background:#C0392B;padding:18px 24px;border-radius:10px;margin-bottom:8px">
      <h1 style="color:#fff;margin:0">🐔 ChickNSter — Réconciliation</h1>
      <p style="color:#fde2e1;margin:4px 0 0">
        POS &nbsp;·&nbsp; Glovo &nbsp;·&nbsp; NAPS (TPE) &nbsp;·&nbsp; Site
      </p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("ℹ️ Comment ça marche ?", expanded=False):
    st.markdown(
        """
        1. **Dépose les 4 fichiers Excel** ci-dessous (le POS est obligatoire ;
           les 3 sources sont facultatives — tu peux en réconcilier une seule).
        2. Clique sur **Lancer la réconciliation**.
        3. Consulte les **écarts et commandes erronées**, puis **télécharge le
           rapport Excel** (résumé + anomalies + POS annoté).

        **Règles appliquées :**
        - Tickets POS `1-3 chiffres` → **Glovo** · `5 chiffres` → **Site** ·
          `sp…/emp…` → **Sur place / Emporter**.
        - Glovo : `Online`→`Bank Transfer`, `Cash`→`Cash` · montant col W − col AE ·
          saisie ≤ 10 min après réception.
        - NAPS : toutes en `Credit card`, rapprochées par date + montant.
        - Site : identifiant = ticket name · une commande **livrée** doit être
          présente au POS en `Bank Transfer` (sinon on cherche une faute de
          frappe sur le n°). Une commande **non livrée** peut être présente
          (tapée puis annulée) — ce n'est pas une anomalie.
        - Sur place / emporter : `Cash` ou `Credit card`.
        """
    )


# --------------------------------------------------------------------------- #
# Dépôt des fichiers
# --------------------------------------------------------------------------- #
st.subheader("📂 Fichiers")
c1, c2, c3, c4 = st.columns(4)
with c1:
    f_pos = st.file_uploader("POS (Mahaal) — obligatoire", type=["xlsx", "xls"], key="pos")
with c2:
    f_glovo = st.file_uploader("Glovo", type=["xlsx", "xls"], key="glovo")
with c3:
    f_naps = st.file_uploader("NAPS (TPE)", type=["xlsx", "xls"], key="naps")
with c4:
    f_site = st.file_uploader("Site", type=["xlsx", "xls"], key="site")

run = st.button("🚀 Lancer la réconciliation", type="primary", use_container_width=True)


# --------------------------------------------------------------------------- #
# Traitement
# --------------------------------------------------------------------------- #
def _safe_load(loader, file, label):
    if file is None:
        return None
    try:
        return loader(file)
    except Exception as exc:  # noqa: BLE001
        st.error(f"❌ Erreur de lecture du fichier **{label}** : {exc}")
        st.stop()


if run:
    if f_pos is None:
        st.warning("⚠️ Le fichier **POS** est obligatoire.")
        st.stop()

    with st.spinner("Réconciliation en cours…"):
        pos = _safe_load(load_pos, f_pos, "POS")
        glovo = _safe_load(load_glovo, f_glovo, "Glovo")
        naps = _safe_load(load_naps, f_naps, "NAPS")
        site = _safe_load(load_site, f_site, "Site")

        anomalies, pos_annotated, summary = run_reconciliation(pos, glovo, naps, site)
        report_bytes = build_excel_report(anomalies, pos_annotated, summary)

    st.session_state["result"] = {
        "anomalies": anomalies,
        "pos_annotated": pos_annotated,
        "summary": summary,
        "report": report_bytes,
    }


# --------------------------------------------------------------------------- #
# Affichage des résultats
# --------------------------------------------------------------------------- #
if "result" in st.session_state:
    res = st.session_state["result"]
    summary = res["summary"]
    anomalies = res["anomalies"]

    st.divider()
    st.subheader("📊 Vue d'ensemble")

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Transactions POS", summary["pos_transactions"])
    m2.metric("Total POS (DH)", f"{summary['pos_total']:,.0f}")
    m3.metric("Anomalies", summary["n_anomalies"])
    m4.metric("🔴 Haute", summary["severity"].get("haute", 0))
    m5.metric("🟠 Moyenne", summary["severity"].get("moyenne", 0))
    m6.metric("🔵 Infos", summary.get("n_infos", 0))

    _period = (f"📅 Période analysée (d'après le POS) : **{summary.get('pos_date_min','')}** "
               f"→ **{summary.get('pos_date_max','')}**.")
    _ex = []
    if summary.get("glovo_excluded"):
        _ex.append(f"{summary['glovo_excluded']} commande(s) Glovo")
    if summary.get("site_excluded"):
        _ex.append(f"{summary['site_excluded']} commande(s) Site")
    if _ex:
        _period += " " + " et ".join(_ex) + " hors de cette période ont été ignorée(s)."
    st.info(_period)

    # Répartition par canal
    ch = summary["channels"]
    if ch:
        st.caption("Répartition des tickets POS par canal détecté")
        st.bar_chart(pd.Series(ch, name="Tickets"))

    # Téléchargement
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        "⬇️ Télécharger le rapport Excel (résumé + anomalies + POS annoté)",
        data=res["report"],
        file_name=f"reconciliation_chicknster_{stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    def _table(items):
        return pd.DataFrame([{
            "Gravité": SEV_BADGE[a["severity"]], "Source": a["source"],
            "Type": a["type"], "Ticket": a.get("ticket_name", ""), "Détail": a["detail"],
        } for a in sorted(items, key=lambda x: SEV_ORDER.get(x["severity"], 9))])

    real = [a for a in anomalies if a["severity"] != "info"]
    infos = [a for a in anomalies if a["severity"] == "info"]

    # Anomalies (Haute + Moyenne)
    st.divider()
    st.subheader("🚨 Anomalies")
    st.caption("Éléments à corriger (Haute + Moyenne). Les rattachements et notes sont en « Infos ».")
    if not real:
        st.success("✅ Aucune anomalie détectée — tout est réconcilié !")
    else:
        st.caption(f"{len(real)} anomalie(s)")
        st.dataframe(_table(real), use_container_width=True, hide_index=True)

    # Infos (rattachements & notes) — pas des anomalies
    st.divider()
    st.subheader("🔵 Infos (rattachements & notes)")
    st.caption("Commandes retrouvées sous un ticket sans numéro, saisies tardives, "
               "tickets sans numéro non rattachés, écarts globaux. Rien à corriger.")
    if infos:
        st.caption(f"{len(infos)} info(s)")
        st.dataframe(_table(infos), use_container_width=True, hide_index=True)
    else:
        st.write("—")

    # POS annoté
    st.divider()
    st.subheader("🧾 POS annoté")
    pos_ann = res["pos_annotated"]
    only_anom = st.checkbox("N'afficher que les tickets en anomalie", value=False)
    view_cols = ["ticket_no", "date", "hour", "user", "ticket_name",
                 "channel_detected", "total", "payment_type",
                 "statut_reconciliation", "anomalies"]
    view_cols = [c for c in view_cols if c in pos_ann.columns]
    view = pos_ann[view_cols]
    if only_anom:
        view = view[view["statut_reconciliation"].str.contains("Anomalie", na=False)]
    st.dataframe(view, use_container_width=True, hide_index=True)

else:
    st.info("👆 Dépose au moins le fichier POS puis lance la réconciliation.")
