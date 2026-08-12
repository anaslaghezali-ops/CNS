"""
Moteur de réconciliation ChickNSter.

Compare le fichier POS (maître) aux 3 sources (Glovo, NAPS, Site) et produit :
  - une liste d'anomalies structurées,
  - un DataFrame POS annoté (statut de réconciliation par ticket),
  - un résumé chiffré.

Chaque anomalie est un dict :
  {
    source, severity, type, ticket_name, pos_datetime,
    source_ref, detail, amount_pos, amount_source,
    payment_pos, payment_source
  }
severity ∈ {"haute", "moyenne", "info"}
"""

from __future__ import annotations

import pandas as pd

from .classify import (
    CHANNEL_GLOVO,
    CHANNEL_SITE,
    CHANNEL_DINEIN,
    CHANNEL_UNASSIGNED,
    add_channel_column,
)

# Correspondance des modes de paiement source -> POS
GLOVO_PAYMENT_MAP = {"Online": "Bank Transfer", "Cash": "Cash"}
SITE_EXPECTED_PAYMENT = "Bank Transfer"
DINEIN_ALLOWED_PAYMENTS = {"Cash", "Credit card"}

# Fenêtre temporelle : le caissier tape la commande Glovo entre l'heure de
# réception (possiblement à la minute exacte) et 10 min après au maximum.
GLOVO_WINDOW_BEFORE_MIN = 1    # simple tolérance d'arrondi à la minute
GLOVO_WINDOW_AFTER_MIN = 10    # maximum observé entre réception et saisie POS

AMOUNT_TOLERANCE = 0.5  # écart de montant toléré (arrondis)


def _anomaly(source, severity, type_, detail, *, ticket_name="", pos_datetime=None,
             source_ref="", amount_pos=None, amount_source=None,
             payment_pos="", payment_source=""):
    return {
        "source": source,
        "severity": severity,
        "type": type_,
        "ticket_name": ticket_name,
        "pos_datetime": pos_datetime,
        "source_ref": source_ref,
        "detail": detail,
        "amount_pos": amount_pos,
        "amount_source": amount_source,
        "payment_pos": payment_pos,
        "payment_source": payment_source,
    }


# --------------------------------------------------------------------------- #
# SITE
# --------------------------------------------------------------------------- #

def reconcile_site(pos_df: pd.DataFrame, site_df: pd.DataFrame):
    """Rapproche les commandes du site avec le POS (par identifiant)."""
    anomalies = []
    pos_by_name = {r["ticket_name"]: r for _, r in pos_df.iterrows()}
    site_ids = set(site_df["identifiant"].astype(str))
    matched_pos_names = set()

    for _, s in site_df.iterrows():
        sid = str(s["identifiant"])
        delivered = str(s.get("delivery_status", "")).upper() == "DELIVERED"
        pos_row = pos_by_name.get(sid)

        if delivered:
            if pos_row is None:
                anomalies.append(_anomaly(
                    "Site", "haute", "Commande livrée absente du POS",
                    f"Commande site {sid} livrée mais introuvable dans le POS.",
                    source_ref=sid, amount_source=s["order_total"],
                ))
            else:
                matched_pos_names.add(sid)
                # Montant
                if pd.notna(pos_row["total"]) and abs(pos_row["total"] - s["order_total"]) > AMOUNT_TOLERANCE:
                    anomalies.append(_anomaly(
                        "Site", "haute", "Écart de montant",
                        f"Commande site {sid} : {s['order_total']} DH (site) "
                        f"vs {pos_row['total']} DH (POS).",
                        ticket_name=sid, pos_datetime=pos_row.get("datetime"),
                        source_ref=sid, amount_pos=pos_row["total"],
                        amount_source=s["order_total"],
                    ))
                # Paiement
                if pos_row["payment_type"] != SITE_EXPECTED_PAYMENT:
                    anomalies.append(_anomaly(
                        "Site", "haute", "Mode de paiement incorrect",
                        f"Commande site {sid} : attendu '{SITE_EXPECTED_PAYMENT}', "
                        f"trouvé '{pos_row['payment_type']}' au POS.",
                        ticket_name=sid, pos_datetime=pos_row.get("datetime"),
                        source_ref=sid, payment_pos=pos_row["payment_type"],
                        payment_source=SITE_EXPECTED_PAYMENT,
                    ))
        else:
            # Commande non livrée (refusée) : ne doit PAS être au POS
            if pos_row is not None:
                matched_pos_names.add(sid)
                anomalies.append(_anomaly(
                    "Site", "haute", "Commande non livrée mais tapée au POS",
                    f"Commande site {sid} refusée/non livrée "
                    f"(statut '{s.get('last_status','')}') mais présente au POS.",
                    ticket_name=sid, pos_datetime=pos_row.get("datetime"),
                    source_ref=sid, amount_pos=pos_row["total"],
                ))

    # Tickets POS classés Site mais absents du fichier site
    pos_site = pos_df[pos_df["channel_detected"] == CHANNEL_SITE]
    for _, p in pos_site.iterrows():
        if p["ticket_name"] not in site_ids:
            anomalies.append(_anomaly(
                "Site", "moyenne", "Ticket Site au POS sans commande correspondante",
                f"Ticket POS {p['ticket_name']} ressemble à une commande site "
                f"mais n'existe pas dans le fichier site.",
                ticket_name=p["ticket_name"], pos_datetime=p.get("datetime"),
                amount_pos=p["total"], payment_pos=p["payment_type"],
            ))

    return anomalies, matched_pos_names


# --------------------------------------------------------------------------- #
# NAPS (TPE) — toutes les transactions = Credit card au POS
# --------------------------------------------------------------------------- #

def reconcile_naps(pos_df: pd.DataFrame, naps_df: pd.DataFrame):
    """Rapproche les transactions TPE avec les paiements 'Credit card' du POS."""
    anomalies = []
    pos_cc = pos_df[pos_df["payment_type"] == "Credit card"].copy()
    pos_cc["date"] = pos_cc["datetime"].dt.date

    if naps_df.empty:
        return anomalies

    naps_dates = set(naps_df["date"].dropna())
    naps_min, naps_max = min(naps_dates), max(naps_dates)

    # 1) Alerte de couverture : jours POS hors du relevé NAPS
    pos_dates = set(pos_cc["date"].dropna())
    uncovered = sorted(d for d in pos_dates if not (naps_min <= d <= naps_max))
    for d in uncovered:
        n = (pos_cc["date"] == d).sum()
        total = pos_cc.loc[pos_cc["date"] == d, "total"].sum()
        anomalies.append(_anomaly(
            "NAPS", "moyenne", "Journée non couverte par le relevé NAPS",
            f"{n} paiement(s) 'Credit card' du {d} ({total:.0f} DH) : "
            f"le relevé NAPS fourni couvre du {naps_min} au {naps_max} "
            f"(décalage de télécollecte probable).",
            source_ref=str(d),
        ))

    # 2) Rapprochement par (date, montant) sur les jours couverts — nombre + montant
    covered_dates = pos_dates & set(d for d in pos_dates if naps_min <= d <= naps_max)
    for d in sorted(covered_dates):
        pos_amounts = pos_cc.loc[pos_cc["date"] == d, "total"].round(2)
        naps_amounts = naps_df.loc[naps_df["date"] == d, "montant"].round(2)

        pos_counts = pos_amounts.value_counts().to_dict()
        naps_counts = naps_amounts.value_counts().to_dict()
        all_amounts = set(pos_counts) | set(naps_counts)

        for amt in sorted(all_amounts):
            pc = pos_counts.get(amt, 0)
            nc = naps_counts.get(amt, 0)
            if pc > nc:
                anomalies.append(_anomaly(
                    "NAPS", "haute", "Paiement POS absent du TPE",
                    f"{pc - nc} paiement(s) 'Credit card' de {amt:.0f} DH le {d} "
                    f"au POS sans équivalent dans le relevé NAPS.",
                    source_ref=str(d), amount_pos=amt,
                ))
            elif nc > pc:
                anomalies.append(_anomaly(
                    "NAPS", "haute", "Transaction TPE absente du POS",
                    f"{nc - pc} transaction(s) NAPS de {amt:.0f} DH le {d} "
                    f"sans équivalent 'Credit card' au POS.",
                    source_ref=str(d), amount_source=amt,
                ))

    return anomalies


# --------------------------------------------------------------------------- #
# GLOVO — rapprochement temporel (le POS est tapé après la réception)
# --------------------------------------------------------------------------- #

def _glovo_nearest(g, pos_glovo, used, lo, hi, payment=None, require_amount=True):
    """
    Meilleur ticket POS pour une commande Glovo :
      - dans la fenêtre temporelle [lo, hi],
      - du bon mode de paiement si `payment` est fourni,
      - de montant identique (col W = total POS) si `require_amount`,
      - le plus proche dans le temps en cas d'ex-æquo.
    """
    g_amount = g.get("subtotal")
    best_pidx, best_gap = None, None
    for pidx, p in pos_glovo.iterrows():
        if pidx in used:
            continue
        dt = p["datetime"]
        if pd.isna(dt) or not (lo <= dt <= hi):
            continue
        if payment is not None and p["payment_type"] != payment:
            continue
        if require_amount and pd.notna(g_amount) and pd.notna(p["total"]):
            if abs(p["total"] - g_amount) > AMOUNT_TOLERANCE:
                continue
        gap = abs((dt - g["received_at"]).total_seconds())
        if best_gap is None or gap < best_gap:
            best_gap, best_pidx = gap, pidx
    return best_pidx


def reconcile_glovo(pos_df: pd.DataFrame, glovo_df: pd.DataFrame):
    """
    Rapproche les commandes Glovo livrées avec les tickets POS Glovo.

    L'appariement se fait sur **montant (col W) + heure** (bien plus fiable que
    l'heure seule : 100 % des montants Glovo existent côté POS), en 2 passes :
      Passe 1 — ticket POS de même montant ET déjà au bon mode de paiement.
      Passe 2 — ticket POS de même montant mais paiement différent → révèle
                les VRAIES erreurs de mode de paiement.
    Les commandes sans ticket de même montant dans la fenêtre sont signalées
    comme absentes du POS ; l'écart agrégé par mode de paiement (fiable) est
    calculé en parallèle.
    """
    anomalies = []
    delivered = glovo_df[glovo_df["status"].str.lower() == "delivered"].copy()
    delivered = delivered.sort_values("received_at")
    pos_glovo = pos_df[pos_df["channel_detected"] == CHANNEL_GLOVO].copy()
    pos_glovo = pos_glovo.sort_values("datetime")

    before = pd.Timedelta(minutes=GLOVO_WINDOW_BEFORE_MIN)
    after = pd.Timedelta(minutes=GLOVO_WINDOW_AFTER_MIN)
    used = set()
    matches = []  # (glovo_idx, pos_idx)

    # ---- Réconciliation agrégée (fiable) par mode de paiement ----
    anomalies += _glovo_aggregate(delivered, pos_glovo)

    # ---- Passe 1 : même montant + bon mode de paiement ----
    remaining = []
    for gidx, g in delivered.iterrows():
        if pd.isna(g["received_at"]):
            remaining.append(gidx)
            continue
        lo, hi = g["received_at"] - before, g["received_at"] + after
        expected = GLOVO_PAYMENT_MAP.get(g["payment_type"])
        best = _glovo_nearest(g, pos_glovo, used, lo, hi, payment=expected)
        if best is not None:
            used.add(best)
            matches.append((gidx, best))
        else:
            remaining.append(gidx)

    # ---- Passe 2 : même montant, paiement quelconque → erreur de paiement ----
    remaining2 = []
    for gidx in remaining:
        g = delivered.loc[gidx]
        if pd.isna(g["received_at"]):
            anomalies.append(_anomaly(
                "Glovo", "moyenne", "Commande Glovo sans heure de réception",
                f"Commande Glovo {g['order_id']} sans heure exploitable — "
                f"rapprochement manuel nécessaire.",
                source_ref=str(g["order_id"]),
            ))
            continue
        lo, hi = g["received_at"] - before, g["received_at"] + after
        best = _glovo_nearest(g, pos_glovo, used, lo, hi, payment=None)
        if best is not None:
            used.add(best)
            matches.append((gidx, best))
            p = pos_glovo.loc[best]
            expected = GLOVO_PAYMENT_MAP.get(g["payment_type"])
            anomalies.append(_anomaly(
                "Glovo", "haute", "Mode de paiement incorrect",
                f"Commande Glovo {g['order_id']} ({g['payment_type']}, {g['subtotal']:.0f} DH) : "
                f"attendu '{expected}' au POS, trouvé '{p['payment_type']}' "
                f"(ticket {p['ticket_name']} à {p['datetime']:%H:%M}).",
                ticket_name=p["ticket_name"], pos_datetime=p.get("datetime"),
                source_ref=str(g["order_id"]),
                payment_pos=p["payment_type"], payment_source=expected,
            ))
        else:
            remaining2.append(gidx)

    # ---- Passe 3 : rattrapage hors fenêtre (saisie tardive) ----
    # Même montant + bon mode de paiement, mais au-delà des 10 min.
    for gidx in remaining2:
        g = delivered.loc[gidx]
        expected = GLOVO_PAYMENT_MAP.get(g["payment_type"])
        big = pd.Timedelta(days=1)
        best = _glovo_nearest(g, pos_glovo, used,
                              g["received_at"] - big, g["received_at"] + big,
                              payment=expected)
        if best is not None:
            used.add(best)
            matches.append((gidx, best))
            p = pos_glovo.loc[best]
            delay = (p["datetime"] - g["received_at"]).total_seconds() / 60
            anomalies.append(_anomaly(
                "Glovo", "info", "Saisie tardive (hors fenêtre 10 min)",
                f"Commande Glovo {g['order_id']} ({g['subtotal']:.0f} DH) reçue à "
                f"{g['received_at']:%H:%M}, tapée au POS à {p['datetime']:%H:%M} "
                f"(ticket {p['ticket_name']}, +{delay:.0f} min) — présente mais tardive.",
                ticket_name=p["ticket_name"], pos_datetime=p.get("datetime"),
                source_ref=str(g["order_id"]),
                amount_pos=p["total"], amount_source=g["subtotal"],
            ))
        else:
            anomalies.append(_anomaly(
                "Glovo", "haute", "Commande Glovo absente du POS",
                f"Commande Glovo {g['order_id']} reçue à "
                f"{g['received_at']:%Y-%m-%d %H:%M} ({g['payment_type']}, "
                f"{g['subtotal']:.0f} DH) non retrouvée au POS "
                f"(aucun ticket de même montant et mode de paiement).",
                source_ref=str(g["order_id"]), amount_source=g.get("subtotal"),
                payment_source=GLOVO_PAYMENT_MAP.get(g["payment_type"], "?"),
            ))

    # ---- Tickets POS Glovo non appariés ----
    for pidx, p in pos_glovo.iterrows():
        if pidx not in used:
            anomalies.append(_anomaly(
                "Glovo", "moyenne", "Ticket Glovo au POS sans commande correspondante",
                f"Ticket POS {p['ticket_name']} ({p['datetime']:%H:%M}, "
                f"{p['payment_type']}, {p['total']:.0f} DH) classé Glovo mais sans "
                f"commande Glovo de même montant dans la fenêtre temporelle.",
                ticket_name=p["ticket_name"], pos_datetime=p.get("datetime"),
                amount_pos=p["total"], payment_pos=p["payment_type"],
            ))

    return anomalies, dict(matches)


def _glovo_aggregate(delivered, pos_glovo):
    """Compare, par mode de paiement, le nombre de commandes Glovo vs POS."""
    anomalies = []
    glovo_online = int((delivered["payment_type"] == "Online").sum())
    glovo_cash = int((delivered["payment_type"] == "Cash").sum())
    pos_bt = int((pos_glovo["payment_type"] == "Bank Transfer").sum())
    pos_cash = int((pos_glovo["payment_type"] == "Cash").sum())

    if glovo_online != pos_bt:
        anomalies.append(_anomaly(
            "Glovo", "info", "Écart global paiement en ligne",
            f"Glovo 'Online' : {glovo_online} vs POS 'Bank Transfer' (tickets Glovo) : "
            f"{pos_bt} → écart de {glovo_online - pos_bt}.",
        ))
    if glovo_cash != pos_cash:
        anomalies.append(_anomaly(
            "Glovo", "info", "Écart global paiement cash",
            f"Glovo 'Cash' : {glovo_cash} vs POS 'Cash' (tickets Glovo) : "
            f"{pos_cash} → écart de {glovo_cash - pos_cash}.",
        ))
    return anomalies


# --------------------------------------------------------------------------- #
# SUR PLACE / À EMPORTER + tickets à rattacher
# --------------------------------------------------------------------------- #

def reconcile_dinein(pos_df: pd.DataFrame):
    """Contrôle les paiements des commandes sur place / à emporter."""
    anomalies = []
    dinein = pos_df[pos_df["channel_detected"] == CHANNEL_DINEIN]
    for _, p in dinein.iterrows():
        if p["payment_type"] not in DINEIN_ALLOWED_PAYMENTS:
            anomalies.append(_anomaly(
                "Sur place", "moyenne", "Mode de paiement inattendu (sur place/emporter)",
                f"Ticket {p['ticket_name']} sur place/emporter payé "
                f"'{p['payment_type']}' (attendu Credit card ou Bank Transfer).",
                ticket_name=p["ticket_name"], pos_datetime=p.get("datetime"),
                amount_pos=p["total"], payment_pos=p["payment_type"],
            ))
    return anomalies


def flag_unassigned(pos_df: pd.DataFrame, glovo_df: pd.DataFrame | None = None):
    """Signale les tickets POS sans nom, avec suggestion de rattachement Glovo."""
    anomalies = []
    unassigned = pos_df[pos_df["channel_detected"] == CHANNEL_UNASSIGNED]
    for _, p in unassigned.iterrows():
        hint = ""
        if glovo_df is not None and pd.notna(p.get("datetime")):
            delivered = glovo_df[glovo_df["status"].str.lower() == "delivered"]
            near = delivered.assign(
                gap=(p["datetime"] - delivered["received_at"]).abs()
            ).nsmallest(1, "gap")
            if not near.empty:
                r = near.iloc[0]
                hint = (f" Suggestion : commande Glovo {r['order_id']} "
                        f"reçue à {r['received_at']:%H:%M}.")
        anomalies.append(_anomaly(
            "POS", "moyenne", "Ticket sans nom (à rattacher)",
            f"Ticket {p['ticket_no']} à {p['datetime']:%H:%M} "
            f"({p['payment_type']}, {p['total']} DH) sans ticket name.{hint}",
            ticket_name="(vide)", pos_datetime=p.get("datetime"),
            amount_pos=p["total"], payment_pos=p["payment_type"],
        ))
    return anomalies


# --------------------------------------------------------------------------- #
# ORCHESTRATION
# --------------------------------------------------------------------------- #

def run_reconciliation(pos_df, glovo_df=None, naps_df=None, site_df=None):
    """
    Lance la réconciliation complète.
    Renvoie (anomalies: list[dict], pos_annotated: DataFrame, summary: dict).
    """
    site_ids = set(site_df["identifiant"].astype(str)) if site_df is not None else set()
    pos = add_channel_column(pos_df, site_ids)

    # Le POS définit la période d'analyse : on ignore les commandes Glovo/Site
    # dont la date n'est pas présente dans le POS (jours non couverts).
    pos_dates = set(pos["datetime"].dropna().dt.date)
    glovo_excluded = site_excluded = 0
    if glovo_df is not None and pos_dates:
        mask = glovo_df["received_at"].dt.date.isin(pos_dates) | glovo_df["received_at"].isna()
        glovo_excluded = int((~mask).sum())
        glovo_df = glovo_df[mask].reset_index(drop=True)
    if site_df is not None and pos_dates:
        mask = site_df["created_at"].dt.date.isin(pos_dates) | site_df["created_at"].isna()
        site_excluded = int((~mask).sum())
        site_df = site_df[mask].reset_index(drop=True)

    all_anomalies = []
    glovo_matches = {}

    if site_df is not None:
        a, _ = reconcile_site(pos, site_df)
        all_anomalies += a
    if naps_df is not None:
        all_anomalies += reconcile_naps(pos, naps_df)
    if glovo_df is not None:
        a, glovo_matches = reconcile_glovo(pos, glovo_df)
        all_anomalies += a
    all_anomalies += reconcile_dinein(pos)
    all_anomalies += flag_unassigned(pos, glovo_df)

    pos_annotated = _annotate_pos(pos, all_anomalies)
    summary = _build_summary(pos, all_anomalies, glovo_df, naps_df, site_df)
    dates = sorted(pos_dates)
    summary["pos_date_min"] = str(dates[0]) if dates else ""
    summary["pos_date_max"] = str(dates[-1]) if dates else ""
    summary["glovo_excluded"] = glovo_excluded
    summary["site_excluded"] = site_excluded
    return all_anomalies, pos_annotated, summary


def _annotate_pos(pos_df, anomalies):
    """Ajoute au POS le statut de réconciliation et le détail des anomalies."""
    df = pos_df.copy()
    by_ticket: dict[str, list[str]] = {}
    for a in anomalies:
        tn = a.get("ticket_name")
        if tn and tn not in ("", "(vide)"):
            by_ticket.setdefault(str(tn), []).append(f"[{a['type']}] {a['detail']}")

    def status(row):
        issues = by_ticket.get(str(row["ticket_name"]), [])
        return "⚠️ Anomalie" if issues else "✅ OK"

    def detail(row):
        return " | ".join(by_ticket.get(str(row["ticket_name"]), []))

    df["statut_reconciliation"] = df.apply(status, axis=1)
    df["anomalies"] = df.apply(detail, axis=1)
    return df


def _build_summary(pos_df, anomalies, glovo_df, naps_df, site_df):
    sev = {"haute": 0, "moyenne": 0, "info": 0}
    for a in anomalies:
        sev[a["severity"]] = sev.get(a["severity"], 0) + 1

    channel_counts = pos_df["channel_detected"].value_counts().to_dict()
    return {
        "pos_transactions": len(pos_df),
        "pos_total": float(pd.to_numeric(pos_df["total"], errors="coerce").sum()),
        "channels": channel_counts,
        "n_anomalies": len(anomalies),
        "severity": sev,
        "glovo_orders": 0 if glovo_df is None else int(
            (glovo_df["status"].str.lower() == "delivered").sum()),
        "naps_transactions": 0 if naps_df is None else len(naps_df),
        "site_orders": 0 if site_df is None else len(site_df),
    }
