"""
Génération du rapport Excel de réconciliation ChickNSter.

Produit un classeur à 3 feuilles :
  - « Résumé »      : chiffres clés + compte des anomalies
  - « Anomalies »   : liste détaillée, triée par gravité
  - « POS annoté »  : le fichier POS d'origine + canal détecté + statut,
                      les lignes en anomalie étant surlignées.
"""

from __future__ import annotations

import io

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


SEVERITY_ORDER = {"haute": 0, "moyenne": 1, "info": 2}
SEVERITY_LABEL = {"haute": "🔴 Haute", "moyenne": "🟠 Moyenne", "info": "🔵 Info"}

FILL_HEADER = PatternFill("solid", fgColor="C0392B")   # rouge ChickNSter
FILL_ANOMALY = PatternFill("solid", fgColor="FDE2E1")   # rouge très clair
FILL_OK = PatternFill("solid", fgColor="E8F6EF")        # vert très clair
FILL_HAUTE = PatternFill("solid", fgColor="F5B7B1")
FILL_MOYENNE = PatternFill("solid", fgColor="FAD7A0")
FILL_INFO = PatternFill("solid", fgColor="AED6F1")


def _anomalies_dataframe(anomalies: list[dict], empty_msg: str) -> pd.DataFrame:
    rows = []
    for a in sorted(anomalies, key=lambda x: SEVERITY_ORDER.get(x["severity"], 9)):
        rows.append({
            "Gravité": SEVERITY_LABEL.get(a["severity"], a["severity"]),
            "Source": a["source"],
            "Type": a["type"],
            "Ticket POS": a.get("ticket_name", ""),
            "Réf. source": a.get("source_ref", ""),
            "Montant POS": a.get("amount_pos"),
            "Montant source": a.get("amount_source"),
            "Paiement POS": a.get("payment_pos", ""),
            "Paiement attendu": a.get("payment_source", ""),
            "Détail": a["detail"],
        })
    if not rows:
        rows.append({"Gravité": "—", "Source": "", "Type": "",
                     "Ticket POS": "", "Réf. source": "", "Montant POS": None,
                     "Montant source": None, "Paiement POS": "",
                     "Paiement attendu": "", "Détail": empty_msg})
    return pd.DataFrame(rows)


def _summary_dataframe(summary: dict) -> pd.DataFrame:
    rows = [
        ("Période analysée (POS)",
         f"{summary.get('pos_date_min', '')} → {summary.get('pos_date_max', '')}"),
        ("Transactions POS", summary["pos_transactions"]),
        ("Total POS (DH)", round(summary["pos_total"], 2)),
        ("Commandes Glovo (livrées, période)", summary["glovo_orders"]),
        ("Commandes Glovo hors période (ignorées)", summary.get("glovo_excluded", 0)),
        ("Transactions NAPS", summary["naps_transactions"]),
        ("Commandes Site (période)", summary["site_orders"]),
        ("Commandes Site hors période (ignorées)", summary.get("site_excluded", 0)),
        ("", ""),
        ("Anomalies (Haute + Moyenne)", summary["n_anomalies"]),
        ("  dont gravité haute", summary["severity"].get("haute", 0)),
        ("  dont gravité moyenne", summary["severity"].get("moyenne", 0)),
        ("Infos (rattachements & notes)", summary.get("n_infos", summary["severity"].get("info", 0))),
        ("", ""),
    ]
    for channel, n in summary["channels"].items():
        rows.append((f"POS — {channel}", n))
    return pd.DataFrame(rows, columns=["Indicateur", "Valeur"])


# Colonnes du POS annoté à exporter (dans l'ordre)
_POS_EXPORT_COLS = [
    ("ticket_no", "Ticket No."),
    ("date", "Date"),
    ("hour", "Heure"),
    ("user", "Caissier"),
    ("ticket_name", "Ticket name"),
    ("channel_detected", "Canal détecté"),
    ("total", "Total"),
    ("payment_type", "Mode de paiement"),
    ("statut_reconciliation", "Statut"),
    ("anomalies", "Anomalies détectées"),
]


def _pos_export_dataframe(pos_annotated: pd.DataFrame) -> pd.DataFrame:
    cols = [(src, dst) for src, dst in _POS_EXPORT_COLS if src in pos_annotated.columns]
    df = pos_annotated[[src for src, _ in cols]].copy()
    df.columns = [dst for _, dst in cols]
    return df


def build_excel_report(anomalies, pos_annotated, summary) -> bytes:
    """Construit le rapport Excel complet et le renvoie en bytes."""
    buf = io.BytesIO()
    real = [a for a in anomalies if a["severity"] != "info"]
    infos = [a for a in anomalies if a["severity"] == "info"]
    df_summary = _summary_dataframe(summary)
    df_anom = _anomalies_dataframe(real, "Aucune anomalie — tout est réconcilié.")
    df_infos = _anomalies_dataframe(infos, "Aucune info.")
    df_pos = _pos_export_dataframe(pos_annotated)

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Résumé", index=False)
        df_anom.to_excel(writer, sheet_name="Anomalies", index=False)
        df_infos.to_excel(writer, sheet_name="Infos", index=False)
        df_pos.to_excel(writer, sheet_name="POS annoté", index=False)

        _style_sheet(writer.sheets["Résumé"], df_summary)
        _style_anomalies(writer.sheets["Anomalies"], df_anom)
        _style_anomalies(writer.sheets["Infos"], df_infos)
        _style_pos(writer.sheets["POS annoté"], df_pos)

    buf.seek(0)
    return buf.getvalue()


def _style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = FILL_HEADER
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def _autofit(ws, df, max_width=60):
    for i, col in enumerate(df.columns, 1):
        lengths = [len(str(col))] + [len(str(v)) for v in df.iloc[:, i - 1].tolist()]
        ws.column_dimensions[get_column_letter(i)].width = min(max(lengths) + 2, max_width)


def _style_sheet(ws, df):
    _style_header(ws, len(df.columns))
    _autofit(ws, df)


def _style_anomalies(ws, df):
    _style_header(ws, len(df.columns))
    _autofit(ws, df)
    for r in range(2, len(df) + 2):
        grav = str(ws.cell(row=r, column=1).value)
        fill = None
        if "Haute" in grav:
            fill = FILL_HAUTE
        elif "Moyenne" in grav:
            fill = FILL_MOYENNE
        elif "Info" in grav:
            fill = FILL_INFO
        if fill:
            ws.cell(row=r, column=1).fill = fill


def _style_pos(ws, df):
    _style_header(ws, len(df.columns))
    _autofit(ws, df)
    status_col = list(df.columns).index("Statut") + 1 if "Statut" in df.columns else None
    if status_col:
        for r in range(2, len(df) + 2):
            val = str(ws.cell(row=r, column=status_col).value)
            fill = FILL_ANOMALY if "Anomalie" in val else FILL_OK
            ws.cell(row=r, column=status_col).fill = fill
