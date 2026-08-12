"""
Chargement robuste des 4 fichiers Excel de ChickNSter.

Chaque loader :
  - répare automatiquement les fichiers Excel mal formés (export Glovo),
  - localise dynamiquement la ligne d'en-tête,
  - renvoie un DataFrame aux colonnes normalisées (noms internes stables).

Les noms de colonnes internes sont volontairement en snake_case et
indépendants de la langue du fichier source, pour que le reste du moteur
ne dépende jamais d'un libellé exact.
"""

from __future__ import annotations

import io
import re
import zipfile

import pandas as pd


# --------------------------------------------------------------------------- #
# Lecture Excel tolérante aux fichiers corrompus (styles.xml invalide)
# --------------------------------------------------------------------------- #

# openpyxl refuse certaines valeurs d'alignement écrites par des exports tiers
# (ex. Glovo écrit vertical="left", qui n'est pas une valeur verticale valide).
_INVALID_ALIGNMENTS = {
    'vertical="left"': 'vertical="center"',
    'vertical="right"': 'vertical="center"',
    'horizontal="top"': 'horizontal="center"',
    'horizontal="bottom"': 'horizontal="center"',
}


def _repair_xlsx_bytes(data: bytes) -> bytes:
    """Corrige les valeurs d'alignement invalides dans xl/styles.xml."""
    buf_in = io.BytesIO(data)
    with zipfile.ZipFile(buf_in, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    if "xl/styles.xml" in items:
        styles = items["xl/styles.xml"].decode("utf-8", errors="replace")
        for bad, good in _INVALID_ALIGNMENTS.items():
            styles = styles.replace(bad, good)
        items["xl/styles.xml"] = styles.encode("utf-8")

    buf_out = io.BytesIO()
    with zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, payload in items.items():
            zout.writestr(name, payload)
    return buf_out.getvalue()


def _to_bytes(source) -> bytes:
    """Accepte un chemin, des bytes, ou un objet fichier (upload Streamlit)."""
    if isinstance(source, bytes):
        return source
    if hasattr(source, "read"):
        pos = source.tell() if hasattr(source, "tell") else None
        data = source.read()
        if pos is not None:
            try:
                source.seek(pos)
            except Exception:
                pass
        return data
    with open(source, "rb") as fh:
        return fh.read()


def read_excel_safe(source, **kwargs) -> pd.DataFrame:
    """Lit un Excel, en réparant automatiquement s'il est mal formé."""
    data = _to_bytes(source)
    try:
        return pd.read_excel(io.BytesIO(data), **kwargs)
    except Exception:
        repaired = _repair_xlsx_bytes(data)
        return pd.read_excel(io.BytesIO(repaired), **kwargs)


def _find_header_row(source, tokens, max_scan: int = 15) -> int:
    """Renvoie l'index (0-based) de la première ligne contenant un des tokens."""
    raw = read_excel_safe(source, header=None, nrows=max_scan)
    tokens_low = [t.lower() for t in tokens]
    for idx, row in raw.iterrows():
        cells = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if any(tok in cells for tok in tokens_low):
            return idx
    raise ValueError(
        f"En-tête introuvable (tokens recherchés : {tokens}). "
        "Le fichier n'a pas le format attendu."
    )


def _clean_str(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


# --------------------------------------------------------------------------- #
# POS (Mahaal Sales History) — fichier maître
# --------------------------------------------------------------------------- #

def load_pos(source) -> pd.DataFrame:
    """
    Colonnes normalisées :
      ticket_no, date, hour, location, user, customer, channel,
      ticket_name, designations, subtotal, discount, refund, tax,
      total, payment_type
    """
    header_row = _find_header_row(source, ["Ticket No.", "Ticket name"])
    df = read_excel_safe(source, header=header_row)

    # Supprimer une éventuelle première colonne vide (marge de gauche du POS)
    if df.columns[0] is None or str(df.columns[0]).startswith("Unnamed"):
        first = df.iloc[:, 0]
        if first.isna().all():
            df = df.iloc[:, 1:]

    rename = {
        "Ticket No.": "ticket_no",
        "Date": "date",
        "Hour": "hour",
        "Location": "location",
        "User": "user",
        "Customer": "customer",
        "Channel": "channel",
        "Ticket name": "ticket_name",
        "Designations (Reference)": "designations",
        "Sub total": "subtotal",
        "Discount": "discount",
        "Refund": "refund",
        "Tax": "tax",
        "Total": "total",
        "Payment type": "payment_type",
    }
    df = df.rename(columns=rename)

    # Ne garder que les lignes de transaction réelles (ticket_no renseigné)
    df = df[df["ticket_no"].apply(lambda x: _clean_str(x) != "")].copy()

    df["ticket_name"] = df["ticket_name"].apply(_clean_str)
    df["payment_type"] = df["payment_type"].apply(_clean_str)
    df["total"] = pd.to_numeric(df["total"], errors="coerce")
    df["datetime"] = _parse_pos_datetime(df["date"], df["hour"])

    df = df.reset_index(drop=True)
    return df


def _parse_pos_datetime(date_series, hour_series) -> pd.Series:
    """Combine 'Aug 11, 2026' + '23:46' en datetime."""
    combined = (
        date_series.astype(str).str.strip()
        + " "
        + hour_series.astype(str).str.strip()
    )
    return pd.to_datetime(combined, format="%b %d, %Y %H:%M", errors="coerce")


# --------------------------------------------------------------------------- #
# Glovo (orderDetails)
# --------------------------------------------------------------------------- #

def load_glovo(source) -> pd.DataFrame:
    """
    Colonnes normalisées :
      order_id, payment_type, received_at, status, earnings,
      subtotal, discount_funded, amount (subtotal − discount_funded)
    """
    # La vraie ligne d'en-tête est la 2e (la 1re regroupe des catégories).
    header_row = _find_header_row(source, ["Order ID", "Payment type"])
    df = read_excel_safe(source, header=header_row)

    rename = {
        "Order ID": "order_id",
        "Payment type": "payment_type",
        "Order status": "status",
        "Order received at": "received_at",
        "Estimated earnings": "earnings",
        "Subtotal": "subtotal",
        "Discount Funded by you": "discount_funded",
    }
    df = df.rename(columns=rename)

    keep = [
        "order_id", "payment_type", "status", "received_at", "earnings",
        "subtotal", "discount_funded",
    ]
    df = df[[c for c in keep if c in df.columns]].copy()

    df = df[df["order_id"].apply(lambda x: _clean_str(x) != "")].copy()
    df["payment_type"] = df["payment_type"].apply(_clean_str)
    df["status"] = df["status"].apply(_clean_str)
    df["received_at"] = pd.to_datetime(df["received_at"], errors="coerce")
    df["earnings"] = pd.to_numeric(df["earnings"], errors="coerce")
    if "subtotal" in df.columns:
        df["subtotal"] = pd.to_numeric(df["subtotal"], errors="coerce")
    if "discount_funded" in df.columns:
        df["discount_funded"] = pd.to_numeric(df["discount_funded"], errors="coerce").fillna(0)
    else:
        df["discount_funded"] = 0.0
    # Montant de rapprochement Glovo — voir CURSOR_JOURNAL.md (2026-08-12 : W − AE).
    df["amount"] = df["subtotal"] - df["discount_funded"] if "subtotal" in df.columns else pd.NA

    df = df.reset_index(drop=True)
    return df


# --------------------------------------------------------------------------- #
# NAPS (relevé TPE)
# --------------------------------------------------------------------------- #

def load_naps(source) -> pd.DataFrame:
    """
    Colonnes normalisées :
      date_transaction, montant, auth_code, card_type, order_no
    """
    header_row = _find_header_row(source, ["Date de transaction", "Montant"])
    df = read_excel_safe(source, header=header_row)

    rename = {
        "Date de transaction": "date_transaction",
        "Montant": "montant",
        "Code d'autorisation": "auth_code",
        "Type de carte": "card_type",
        "N° de commande": "order_no",
    }
    df = df.rename(columns=rename)

    df = df[df["montant"].apply(lambda x: _clean_str(x) != "")].copy()
    df["date_transaction"] = pd.to_datetime(df["date_transaction"], errors="coerce")
    df["date"] = df["date_transaction"].dt.date
    df["montant"] = pd.to_numeric(df["montant"], errors="coerce")

    df = df.reset_index(drop=True)
    return df


# --------------------------------------------------------------------------- #
# Site (orders)
# --------------------------------------------------------------------------- #

def load_site(source) -> pd.DataFrame:
    """
    Colonnes normalisées :
      identifiant, created_at, last_status, order_total,
      delivery_status, payment_mode, payment_status
    """
    header_row = _find_header_row(source, ["identifiant", "Order Total"])
    df = read_excel_safe(source, header=header_row)

    rename = {
        "identifiant": "identifiant",
        "Created At": "created_at",
        "Last Status": "last_status",
        "Order Total": "order_total",
        "Order Value": "order_value",
        "Mode de paiement": "payment_mode",
        "Statut du paiement": "payment_status",
        "Delivery Provider Status": "delivery_status",
    }
    df = df.rename(columns=rename)

    keep = [
        "identifiant", "created_at", "last_status", "order_total",
        "order_value", "payment_mode", "payment_status", "delivery_status",
    ]
    df = df[[c for c in keep if c in df.columns]].copy()

    df = df[df["identifiant"].apply(lambda x: _clean_str(x) != "")].copy()
    df["identifiant"] = df["identifiant"].apply(_clean_str)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")
    df["order_total"] = pd.to_numeric(df["order_total"], errors="coerce")
    df["last_status"] = df["last_status"].apply(_clean_str)
    df["delivery_status"] = df["delivery_status"].apply(_clean_str)

    df = df.reset_index(drop=True)
    return df
