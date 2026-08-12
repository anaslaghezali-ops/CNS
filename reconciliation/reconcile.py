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


def _anomaly(source, severity, type_, detail, *, ticket_name="", pos_ticket_no="",
             pos_datetime=None,
             source_ref="", amount_pos=None, amount_source=None,
             payment_pos="", payment_source=""):
    return {
        "source": source,
        "severity": severity,
        "type": type_,
        "ticket_name": ticket_name,
        "pos_ticket_no": pos_ticket_no,
        "pos_datetime": pos_datetime,
        "source_ref": source_ref,
        "detail": detail,
        "amount_pos": amount_pos,
        "amount_source": amount_source,
        "payment_pos": payment_pos,
        "payment_source": payment_source,
    }


def _pos_ticket_kw(p):
    """Lie une anomalie à UNE ligne POS (ticket_no unique, pas seulement ticket_name)."""
    return {
        "ticket_name": p["ticket_name"],
        "pos_ticket_no": str(p["ticket_no"]),
        "pos_datetime": p.get("datetime"),
    }


# --------------------------------------------------------------------------- #
# SITE
# --------------------------------------------------------------------------- #

SITE_TYPO_WINDOW_MIN = 20  # fenêtre pour détecter une faute de frappe sur le n°


def reconcile_site(pos_df: pd.DataFrame, site_df: pd.DataFrame):
    """
    Rapproche le site par NUMÉRO (exact) et par faute de frappe (orphelin
    5 chiffres). Les tickets SANS numéro sont laissés à la passe commune.
    Renvoie (anomalies, missing) — commandes livrées non encore rattachées.
    """
    anomalies, missing = [], []
    pos_by_name = {r["ticket_name"]: r for _, r in pos_df.iterrows()}
    site_ids = set(site_df["identifiant"].astype(str))
    orphans = [(i, p) for i, p in pos_df[pos_df["channel_detected"] == CHANNEL_SITE].iterrows()
               if p["ticket_name"] not in site_ids]
    used_orphan = set()

    unmatched_delivered = []
    for _, s in site_df.iterrows():
        sid = str(s["identifiant"])
        delivered = str(s.get("delivery_status", "")).upper() == "DELIVERED"
        pos_row = pos_by_name.get(sid)
        if delivered:
            if pos_row is None:
                unmatched_delivered.append(s)
                continue
            if pd.notna(pos_row["total"]) and abs(pos_row["total"] - s["order_total"]) > AMOUNT_TOLERANCE:
                anomalies.append(_anomaly(
                    "Site", "haute", "Écart de montant",
                    f"Commande site {sid} : {s['order_total']} DH (site) "
                    f"vs {pos_row['total']} DH (POS).",
                    **{_pos_ticket_kw(pos_row), "source_ref": sid,
                       "amount_pos": pos_row["total"], "amount_source": s["order_total"]},
                ))
            if pos_row["payment_type"] != SITE_EXPECTED_PAYMENT:
                anomalies.append(_anomaly(
                    "Site", "haute", "Mode de paiement incorrect",
                    f"Commande site {sid} : attendu '{SITE_EXPECTED_PAYMENT}', "
                    f"trouvé '{pos_row['payment_type']}' au POS.",
                    **{_pos_ticket_kw(pos_row), "source_ref": sid,
                       "payment_pos": pos_row["payment_type"],
                       "payment_source": SITE_EXPECTED_PAYMENT},
                ))
        # Une commande non livrée (refusée/annulée) PEUT être présente au POS.

    # Faute de frappe : orphelin « site-like » (5 chiffres) de même montant/heure.
    for s in unmatched_delivered:
        sid = str(s["identifiant"])
        best_i, best_row, best_gap = None, None, None
        for i, p in orphans:
            if i in used_orphan:
                continue
            if pd.isna(p["total"]) or abs(p["total"] - s["order_total"]) > AMOUNT_TOLERANCE:
                continue
            if pd.isna(p.get("datetime")) or pd.isna(s.get("created_at")):
                continue
            gap = abs((p["datetime"] - s["created_at"]).total_seconds()) / 60
            if gap > SITE_TYPO_WINDOW_MIN:
                continue
            if best_gap is None or gap < best_gap:
                best_i, best_row, best_gap = i, p, gap
        if best_row is not None:
            used_orphan.add(best_i)
            pos_df.loc[best_i, "channel_detected"] = CHANNEL_SITE
            pay_note = ""
            if best_row["payment_type"] != SITE_EXPECTED_PAYMENT:
                pay_note = (f" ⚠️ De plus, son paiement est '{best_row['payment_type']}' "
                            f"au lieu de '{SITE_EXPECTED_PAYMENT}'.")
            anomalies.append(_anomaly(
                "Site", "moyenne", "Numéro de commande mal saisi (faute de frappe)",
                f"Commande site {sid} livrée : introuvable sous ce numéro, mais le ticket "
                f"POS {best_row['ticket_name']} correspond (même montant "
                f"{s['order_total']:.0f} DH, +{best_gap:.0f} min, et {best_row['ticket_name']} "
                f"n'existe pas dans le fichier site). Le caissier a probablement tapé "
                f"{best_row['ticket_name']} au lieu de {sid}." + pay_note,
                **{_pos_ticket_kw(best_row), "source_ref": sid,
                   "amount_pos": best_row["total"], "amount_source": s["order_total"],
                   "payment_pos": best_row["payment_type"],
                   "payment_source": SITE_EXPECTED_PAYMENT},
            ))
        else:
            missing.append(s)  # -> passe commune, puis « absente »

    # Orphelins « site-like » non expliqués.
    for i, p in orphans:
        if i in used_orphan:
            continue
        anomalies.append(_anomaly(
            "Site", "moyenne", "Ticket Site au POS sans commande correspondante",
            f"Ticket POS {p['ticket_name']} ressemble à une commande site "
            f"mais n'existe pas dans le fichier site.",
            **{_pos_ticket_kw(p), "amount_pos": p["total"], "payment_pos": p["payment_type"]},
        ))
    return anomalies, missing


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

def _glovo_nearest(g, pool, used, lo, hi, payment=None,
                   require_amount=False, prefer_amount=False):
    """
    Meilleur ticket du pool pour une commande Glovo.
      payment        : n'accepter que ce mode de paiement (None = indifférent)
      require_amount : n'accepter qu'un montant identique
      prefer_amount  : à défaut d'exiger, privilégier un montant identique
    Priorités : montant identique > ticket Glovo numéroté > proximité temporelle.
    `pool` est une liste de (index, row).
    """
    g_amount = g.get("amount")
    best_idx, best_score = None, None
    for idx, p in pool:
        if idx in used:
            continue
        dt = p["datetime"]
        if pd.isna(dt) or not (lo <= dt <= hi):
            continue
        if payment is not None and p["payment_type"] != payment:
            continue
        amount_match = (pd.notna(g_amount) and pd.notna(p["total"])
                        and abs(p["total"] - g_amount) <= AMOUNT_TOLERANCE)
        if require_amount and not amount_match:
            continue
        gap = abs((dt - g["received_at"]).total_seconds()) / 60
        score = gap + (100000 if (prefer_amount and not amount_match) else 0) \
                    + (1000 if p["channel_detected"] == CHANNEL_UNASSIGNED else 0)
        if best_score is None or score < best_score:
            best_score, best_idx = score, idx
    return best_idx


def reconcile_glovo(pos_df: pd.DataFrame, glovo_df: pd.DataFrame):
    """
    Rapproche Glovo avec les tickets POS NUMÉROTÉS (canal Glovo). Les commandes
    encore introuvables sont renvoyées (passe commune des tickets sans numéro).
    Renvoie (anomalies, missing). Appariement par HEURE + paiement (le montant
    Glovo peut différer du POS : promos/frais), en 4 passes.
    """
    anomalies, missing = [], []
    delivered = glovo_df[glovo_df["status"].str.lower() == "delivered"].copy()
    delivered = delivered.sort_values("received_at")

    pool = [(i, p) for i, p in pos_df[pos_df["channel_detected"] == CHANNEL_GLOVO].iterrows()]
    pool.sort(key=lambda t: (t[1]["datetime"] if pd.notna(t[1]["datetime"])
                             else pd.Timestamp.min))

    before = pd.Timedelta(minutes=GLOVO_WINDOW_BEFORE_MIN)
    after = pd.Timedelta(minutes=GLOVO_WINDOW_AFTER_MIN)
    wide = pd.Timedelta(minutes=120)
    used = set()

    remaining = []
    for gidx, g in delivered.iterrows():
        if pd.isna(g["received_at"]):
            remaining.append(gidx)
            continue
        lo, hi = g["received_at"] - before, g["received_at"] + after
        best = _glovo_nearest(g, pool, used, lo, hi,
                              payment=GLOVO_PAYMENT_MAP.get(g["payment_type"]), require_amount=True)
        if best is not None:
            used.add(best)
        else:
            remaining.append(gidx)

    remaining2 = []
    for gidx in remaining:
        g = delivered.loc[gidx]
        if pd.isna(g["received_at"]):
            anomalies.append(_anomaly(
                "Glovo", "moyenne", "Commande Glovo sans heure de réception",
                f"Commande Glovo {g['order_id']} sans heure exploitable — "
                f"rapprochement manuel nécessaire.", source_ref=str(g["order_id"]),
            ))
            continue
        lo, hi = g["received_at"] - before, g["received_at"] + after
        best = _glovo_nearest(g, pool, used, lo, hi,
                              payment=GLOVO_PAYMENT_MAP.get(g["payment_type"]), prefer_amount=True)
        if best is not None:
            used.add(best)
        else:
            remaining2.append(gidx)

    remaining3 = []
    for gidx in remaining2:
        g = delivered.loc[gidx]
        lo, hi = g["received_at"] - before, g["received_at"] + after
        best = _glovo_nearest(g, pool, used, lo, hi, payment=None, prefer_amount=True)
        if best is not None:
            p = pos_df.loc[best]
            used.add(best)
            exp = GLOVO_PAYMENT_MAP.get(g["payment_type"])
            anomalies.append(_anomaly(
                "Glovo", "haute", "Mode de paiement incorrect",
                f"Commande Glovo {g['order_id']} ({g['payment_type']}, {g['amount']:.0f} DH) : "
                f"attendu '{exp}' au POS, trouvé '{p['payment_type']}' "
                f"(ticket {p['ticket_name']} à {p['datetime']:%H:%M}).",
                **{_pos_ticket_kw(p), "source_ref": str(g["order_id"]),
                   "payment_pos": p["payment_type"], "payment_source": exp},
            ))
        else:
            remaining3.append(gidx)

    for gidx in remaining3:
        g = delivered.loc[gidx]
        lo, hi = g["received_at"] - wide, g["received_at"] + wide
        best = _glovo_nearest(g, pool, used, lo, hi,
                              payment=GLOVO_PAYMENT_MAP.get(g["payment_type"]), require_amount=True)
        if best is not None:
            p = pos_df.loc[best]
            used.add(best)
            delay = (p["datetime"] - g["received_at"]).total_seconds() / 60
            anomalies.append(_anomaly(
                "Glovo", "info", "Saisie tardive (hors fenêtre 10 min)",
                f"Commande Glovo {g['order_id']} ({g['amount']:.0f} DH) reçue à "
                f"{g['received_at']:%H:%M}, tapée au POS à {p['datetime']:%H:%M} "
                f"(ticket {p['ticket_name']}, {delay:+.0f} min) — présente mais tardive.",
                **{_pos_ticket_kw(p), "source_ref": str(g["order_id"]),
                   "amount_pos": p["total"], "amount_source": g["amount"]},
            ))
        else:
            missing.append(g)  # -> passe commune (ticket sans numéro) puis « absente »

    # Commandes ANNULÉES : si tapées au POS avant annulation, rapprocher (pas une orpheline).
    cancelled = glovo_df[glovo_df["status"].str.lower() == "cancelled"].copy()
    cancelled = cancelled.sort_values("received_at")
    for gidx, g in cancelled.iterrows():
        if pd.isna(g["received_at"]):
            continue
        lo, hi = g["received_at"] - before, g["received_at"] + after
        best = _glovo_nearest(
            g, pool, used, lo, hi,
            payment=GLOVO_PAYMENT_MAP.get(g["payment_type"]), require_amount=True,
        )
        if best is None:
            lo, hi = g["received_at"] - wide, g["received_at"] + wide
            best = _glovo_nearest(
                g, pool, used, lo, hi,
                payment=GLOVO_PAYMENT_MAP.get(g["payment_type"]), require_amount=True,
            )
        if best is not None:
            used.add(best)
            p = pos_df.loc[best]
            glovo_df.loc[gidx, "matched_pos"] = True
            delay = (p["datetime"] - g["received_at"]).total_seconds() / 60
            anomalies.append(_anomaly(
                "Glovo", "info", "Commande Glovo annulée — présente au POS",
                f"Commande Glovo {g['order_id']} annulée ({g['payment_type']}, "
                f"{g['amount']:.0f} DH) reçue à {g['received_at']:%H:%M}, tapée au POS "
                f"(ticket {p['ticket_name']} à {p['datetime']:%H:%M}, {delay:+.0f} min) — "
                f"commande annulée sur Glovo mais ticket caisse présent.",
                **{_pos_ticket_kw(p), "source_ref": str(g["order_id"]),
                   "amount_pos": p["total"], "amount_source": g["amount"],
                   "payment_pos": p["payment_type"],
                   "payment_source": GLOVO_PAYMENT_MAP.get(g["payment_type"])},
            ))

    # Tickets classés Glovo non appariés.
    for idx, p in pool:
        if idx not in used:
            anomalies.append(_anomaly(
                "Glovo", "moyenne", "Ticket Glovo au POS sans commande correspondante",
                f"Ticket POS {p['ticket_name']} ({p['datetime']:%H:%M}, "
                f"{p['payment_type']}, {p['total']:.0f} DH) classé Glovo mais sans "
                f"commande Glovo de même montant dans la fenêtre temporelle.",
                **{_pos_ticket_kw(p), "amount_pos": p["total"], "payment_pos": p["payment_type"]},
            ))
    return anomalies, missing


def reconcile_unassigned(pos_df, missing_site, missing_glovo):
    """
    Passe COMMUNE : les tickets sans numéro (« Ticket »/vide) sont attribués
    JOINTEMENT aux commandes Glovo ET Site encore manquantes — à la plus proche
    en temps, de même montant et de paiement compatible. Un ticket ne peut
    appartenir qu'à une seule commande, quel que soit le canal.
    """
    anomalies = []
    tickets = [(i, p) for i, p in
               pos_df[pos_df["channel_detected"] == CHANNEL_UNASSIGNED].iterrows()]
    used_t = set()

    demands = []
    for s in missing_site:
        demands.append({"src": "Site", "ref": s.get("created_at"),
                        "amount": s["order_total"], "pay": SITE_EXPECTED_PAYMENT,
                        "id": str(s["identifiant"]), "o": s})
    for g in missing_glovo:
        demands.append({"src": "Glovo", "ref": g.get("received_at"),
                        "amount": g["amount"], "pay": GLOVO_PAYMENT_MAP.get(g["payment_type"]),
                        "id": str(g["order_id"]), "o": g})
    demands.sort(key=lambda d: d["ref"] if pd.notna(d["ref"]) else pd.Timestamp.min)

    before = pd.Timedelta(minutes=GLOVO_WINDOW_BEFORE_MIN)
    after = pd.Timedelta(minutes=GLOVO_WINDOW_AFTER_MIN)

    for d in demands:
        best_i, best_row, best_gap = None, None, None
        if pd.notna(d["ref"]):
            for i, t in tickets:
                if i in used_t or pd.isna(t.get("datetime")):
                    continue
                if t["payment_type"] != d["pay"]:
                    continue
                if pd.isna(t["total"]) or pd.isna(d["amount"]) or abs(t["total"] - d["amount"]) > AMOUNT_TOLERANCE:
                    continue
                gap = abs((t["datetime"] - d["ref"]).total_seconds()) / 60
                if d["src"] == "Site":
                    ok = gap <= SITE_TYPO_WINDOW_MIN
                else:
                    ok = (d["ref"] - before) <= t["datetime"] <= (d["ref"] + after)
                if ok and (best_gap is None or gap < best_gap):
                    best_i, best_row, best_gap = i, t, gap
        if best_row is not None:
            # Commande retrouvée sous un ticket sans numéro : ce n'est PAS une
            # anomalie (la commande existe), juste une info de rattachement.
            used_t.add(best_i)
            pos_df.loc[best_i, "channel_detected"] = (
                CHANNEL_SITE if d["src"] == "Site" else CHANNEL_GLOVO)
            extra = f"{d['o']['payment_type']}, " if d["src"] == "Glovo" else ""
            anomalies.append(_anomaly(
                d["src"], "info", "Commande rattachée (ticket sans numéro)",
                f"Commande {d['src']} {d['id']} ({extra}{d['amount']:.0f} DH) retrouvée au "
                f"POS sous le ticket sans numéro {best_row['ticket_no']} (+{best_gap:.0f} min). "
                f"Présente — simple oubli de numéro.",
                **{_pos_ticket_kw(best_row), "ticket_name": best_row["ticket_name"] or "(vide)",
                   "source_ref": d["id"], "amount_pos": best_row["total"],
                   "amount_source": d["amount"], "payment_pos": best_row["payment_type"]},
            ))
        elif d["src"] == "Site":
            anomalies.append(_anomaly(
                "Site", "haute", "Commande livrée absente du POS",
                f"Commande site {d['id']} livrée mais introuvable dans le POS.",
                source_ref=d["id"], amount_source=d["amount"],
            ))
        else:
            g = d["o"]
            anomalies.append(_anomaly(
                "Glovo", "haute", "Commande Glovo absente du POS",
                f"Commande Glovo {d['id']} reçue à {g['received_at']:%Y-%m-%d %H:%M} "
                f"({g['payment_type']}, {d['amount']:.0f} DH) non retrouvée au POS "
                f"(aucun ticket au bon mode de paiement).",
                source_ref=d["id"], amount_source=d["amount"],
                payment_source=d["pay"] or "?",
            ))

    # Tickets sans numéro restants -> à rattacher.
    for i, p in tickets:
        if i in used_t:
            continue
        anomalies.append(_anomaly(
            "POS", "info", "Ticket sans numéro (à rattacher)",
            f"Ticket {p['ticket_no']} du {p['datetime']:%Y-%m-%d} à {p['datetime']:%H:%M} "
            f"({p['payment_type']}, {p['total']} DH) sans numéro — non rattaché à "
            f"une commande Glovo ni Site.",
            **{_pos_ticket_kw(p), "ticket_name": p["ticket_name"] or "(vide)",
               "amount_pos": p["total"], "payment_pos": p["payment_type"]},
        ))
    return anomalies


def glovo_aggregate(pos_df, glovo_df):
    """Écart global par mode de paiement (info), sur canaux finaux."""
    anomalies = []
    delivered = glovo_df[glovo_df["status"].str.lower() == "delivered"]
    pos_glovo = pos_df[pos_df["channel_detected"] == CHANNEL_GLOVO]
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
                f"'{p['payment_type']}' (attendu Cash ou Credit card).",
                **{_pos_ticket_kw(p), "amount_pos": p["total"], "payment_pos": p["payment_type"]},
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

    # 1) Rapprochements par NUMÉRO (canaux disjoints).
    # 2) Passe COMMUNE : tickets sans numéro attribués jointement (Glovo + Site).
    all_anomalies = []
    missing_site, missing_glovo = [], []

    if site_df is not None:
        a, missing_site = reconcile_site(pos, site_df)
        all_anomalies += a
    if naps_df is not None:
        all_anomalies += reconcile_naps(pos, naps_df)
    if glovo_df is not None:
        a, missing_glovo = reconcile_glovo(pos, glovo_df)
        all_anomalies += a
    all_anomalies += reconcile_dinein(pos)
    all_anomalies += reconcile_unassigned(pos, missing_site, missing_glovo)
    if glovo_df is not None:
        all_anomalies += glovo_aggregate(pos, glovo_df)

    pos_annotated = _annotate_pos(pos, all_anomalies)
    summary = _build_summary(pos, all_anomalies, glovo_df, naps_df, site_df)
    dates = sorted(pos_dates)
    summary["pos_date_min"] = str(dates[0]) if dates else ""
    summary["pos_date_max"] = str(dates[-1]) if dates else ""
    summary["glovo_excluded"] = glovo_excluded
    summary["site_excluded"] = site_excluded
    return all_anomalies, pos_annotated, summary


def _annotate_pos(pos_df, anomalies):
    """Ajoute au POS le statut de réconciliation et le détail (par ticket_no, pas ticket_name)."""
    df = pos_df.copy()
    by_ticket_no: dict[str, list[str]] = {}
    sev_ticket_no: dict[str, set] = {}
    for a in anomalies:
        no = a.get("pos_ticket_no") or ""
        if not no:
            continue
        by_ticket_no.setdefault(str(no), []).append(f"[{a['type']}] {a['detail']}")
        sev_ticket_no.setdefault(str(no), set()).add(a["severity"])

    def status(row):
        s = sev_ticket_no.get(str(row["ticket_no"]), set())
        if "haute" in s or "moyenne" in s:
            return "⚠️ Anomalie"
        if "info" in s:
            return "ℹ️ Info"
        return "✅ OK"

    def detail(row):
        return " | ".join(by_ticket_no.get(str(row["ticket_no"]), []))

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
        "n_anomalies": sev["haute"] + sev["moyenne"],  # « info » = pas des anomalies
        "n_infos": sev["info"],
        "severity": sev,
        "glovo_orders": 0 if glovo_df is None else int(
            (glovo_df["status"].str.lower() == "delivered").sum()),
        "naps_transactions": 0 if naps_df is None else len(naps_df),
        "site_orders": 0 if site_df is None else len(site_df),
    }
