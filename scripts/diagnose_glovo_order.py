#!/usr/bin/env python3
"""Diagnose why a Glovo order was not matched to POS."""
import sys
from pathlib import Path

import pandas as pd

from reconciliation.classify import classify_ticket_name, add_channel_column
from reconciliation.loaders import load_pos, load_glovo, load_site, read_excel_safe, _find_header_row
from reconciliation.reconcile import run_reconciliation

UP = Path("/home/ubuntu/.cursor/projects/workspace/uploads")
ORDER = sys.argv[1] if len(sys.argv) > 1 else "101735404708"
EXCEL_ROW = int(sys.argv[2]) if len(sys.argv) > 2 else 622

pos_path = UP / "mahaal_sales_history__10__3b98.xlsx"
glovo_path = UP / "orderDetails__9__2548.xlsx"
site_path = UP / "orders_1786548042_d43a.xlsx"

pos_df = load_pos(pos_path)
glovo_df = load_glovo(glovo_path)
site_df = load_site(site_path)
site_ids = set(site_df["identifiant"].astype(str))

print("=== GLOVO ORDER", ORDER, "===")
g = glovo_df[glovo_df["order_id"].astype(str) == ORDER]
if g.empty:
    print("NOT FOUND")
else:
    r = g.iloc[0]
    print(f"  status={r['status']} payment={r['payment_type']}")
    print(f"  received_at={r['received_at']}")
    print(f"  subtotal={r['subtotal']} discount_funded={r['discount_funded']} amount={r['amount']}")
    print(f"  excel_row~{g.index[0] + 3}")  # approximate

header = _find_header_row(pos_path, ["Ticket No.", "Ticket name"])
aoa = read_excel_safe(pos_path, header=None)
headers = [str(x).strip() if pd.notna(x) else "" for x in aoa.iloc[header].values]

print("\n=== POS EXCEL ROW", EXCEL_ROW, "===")
for i in range(header + 1, len(aoa)):
    if i + 1 != EXCEL_ROW:
        continue
    vals = dict(zip(headers, aoa.iloc[i].values))
    for k, v in vals.items():
        if str(k).strip():
            print(f"  {k}: {v}")
    tn = str(vals.get("Ticket name", "")).strip()
    print(f"  classify: {classify_ticket_name(tn, site_ids)}")
    print(f"  ticket_name in site_ids: {tn in site_ids}")
    break
else:
    print("  ROW NOT FOUND")

pos_df = add_channel_column(pos_df, site_ids)
d = pd.Timestamp("2026-08-06").date()
aug6 = pos_df[pos_df["datetime"].dt.date == d]
print("\n=== POS Aug 6, total ~275 ===")
for _, p in aug6.iterrows():
    if pd.isna(p["total"]) or abs(p["total"] - 275) > 0.5:
        continue
    print(
        f"  excel~{header+2+p.name} hour={p['hour']} name={p['ticket_name']} "
        f"total={p['total']} pay={p['payment_type']} ch={p['channel_detected']} no={p['ticket_no']}"
    )

anoms, _, _ = run_reconciliation(pos_df, glovo_df, None, site_df)
print("\n=== ANOMALIES for order ===")
for a in anoms:
    if a.get("source_ref") == ORDER or ORDER in a.get("detail", ""):
        print(f"  [{a['severity']}] {a['type']}: {a['detail'][:120]}...")
