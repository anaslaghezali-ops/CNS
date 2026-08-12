"""
Classification des tickets du POS par canal de vente, à partir du
libellé « Ticket name ».

Règles (définies avec le gérant ChickNSter) :
  - 1 à 3 chiffres                → Glovo
  - identifiant à 5 chiffres      → Site (correspond à l'ID du fichier site)
  - « sp… » / « emp… »            → Sur place / À emporter
  - vide                          → À rattacher (ticket name oublié)
  - autre                         → Autre / manuel
"""

from __future__ import annotations

import re

import pandas as pd


CHANNEL_GLOVO = "Glovo"
CHANNEL_SITE = "Site"
CHANNEL_DINEIN = "Sur place / Emporter"
CHANNEL_UNASSIGNED = "À rattacher"
CHANNEL_OTHER = "Autre"

# Modes de paiement attendus par canal (côté POS)
EXPECTED_PAYMENTS = {
    CHANNEL_GLOVO: {"Bank Transfer", "Cash"},
    CHANNEL_SITE: {"Bank Transfer"},
    CHANNEL_DINEIN: {"Cash", "Credit card"},
}

_RE_1_3_DIGITS = re.compile(r"^\d{1,3}$")
_RE_5_DIGITS = re.compile(r"^\d{5}$")
# « sp » ou « emp », seul ou suivi d'un numéro, avec ou sans espace :
# Sp, Sp3, Sp 3, Emp, Emp2, Emp 11 …
_RE_SP_EMP = re.compile(r"^(sp|emp)\s*\d*$", re.IGNORECASE)


def classify_ticket_name(ticket_name: str, site_ids: set[str] | None = None) -> str:
    """Renvoie le canal déduit d'un ticket name."""
    name = "" if ticket_name is None else str(ticket_name).strip()

    if name == "" or name.lower() == "nan":
        return CHANNEL_UNASSIGNED

    # Priorité : un identifiant présent dans le fichier Site est toujours Site.
    if site_ids and name in site_ids:
        return CHANNEL_SITE

    # « Ticket » (placeholder) = le caissier a oublié de saisir le numéro
    # (ou sp/emp) → à rattacher, comme un ticket sans nom.
    if name.lower() == "ticket":
        return CHANNEL_UNASSIGNED

    # Normaliser (retirer les espaces internes) pour reconnaître « Sp 3 » = « Sp3 ».
    compact = name.replace(" ", "")
    if _RE_SP_EMP.match(compact):
        return CHANNEL_DINEIN

    if _RE_1_3_DIGITS.match(name):
        return CHANNEL_GLOVO

    if _RE_5_DIGITS.match(name):
        return CHANNEL_SITE

    return CHANNEL_OTHER


def add_channel_column(pos_df: pd.DataFrame, site_ids: set[str] | None = None) -> pd.DataFrame:
    """Ajoute une colonne 'channel_detected' au DataFrame POS."""
    df = pos_df.copy()
    df["channel_detected"] = df["ticket_name"].apply(
        lambda n: classify_ticket_name(n, site_ids)
    )
    return df
