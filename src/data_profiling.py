"""
data_profiling.py — profile RAW shipments and scan events.
Outputs:
  outputs/kpi_reports/column_profile_shipments.csv
  outputs/kpi_reports/column_profile_scans.csv
  outputs/kpi_reports/data_quality_report.md
"""
import json
import os
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
REF = os.path.join(ROOT, "data", "reference")
OUT = os.path.join(ROOT, "outputs", "kpi_reports")
os.makedirs(OUT, exist_ok=True)

ship = pd.read_parquet(os.path.join(RAW, "shipments_raw.parquet"))
scans = pd.read_parquet(os.path.join(RAW, "scan_events_raw.parquet"))
crosswalk = pd.read_csv(os.path.join(REF, "hub_crosswalk.csv"))
routes = pd.read_csv(os.path.join(REF, "dim_route.csv"))
vehicles = pd.read_csv(os.path.join(REF, "dim_vehicle.csv"))
hubs = pd.read_csv(os.path.join(REF, "dim_hub.csv"))

known_hub_tokens = set(crosswalk["raw_hub_id"]) | set(hubs["hub_id"])


def profile(df, name):
    rows = []
    for col in df.columns:
        s = df[col]
        n_missing = int(s.isna().sum())
        # empty-string missing for object cols
        if s.dtype == object:
            n_missing += int((s.astype(str).str.strip() == "").sum())
        rows.append({
            "table": name, "column": col, "dtype": str(s.dtype),
            "total_records": len(df), "missing_values": n_missing,
            "missing_pct": round(100 * n_missing / len(df), 3),
            "unique_values": int(s.nunique(dropna=True)),
            "duplicate_values": int(len(df) - s.nunique(dropna=True)) if col.endswith("_id") else "",
            "sample_values": "; ".join(map(str, s.dropna().unique()[:3])),
        })
    return pd.DataFrame(rows)


prof_s = profile(ship, "shipments_raw")
prof_c = profile(scans, "scan_events_raw")
prof_s.to_csv(os.path.join(OUT, "column_profile_shipments.csv"), index=False)
prof_c.to_csv(os.path.join(OUT, "column_profile_scans.csv"), index=False)

# ---- targeted integrity checks -------------------------------------------
checks = {}

# duplicates
dup_ids = ship["shipment_id"].duplicated(keep=False)
checks["duplicate_shipment_id_rows"] = int(dup_ids.sum())
checks["duplicate_shipment_ids_distinct"] = int(ship["shipment_id"].duplicated().sum())
checks["duplicate_scan_event_id_rows"] = int(scans["scan_event_id"].duplicated(keep=False).sum())

# costs
cost = pd.to_numeric(ship["shipment_cost_inr"], errors="coerce")
checks["negative_cost_rows"] = int((cost < 0).sum())
checks["zero_cost_rows"] = int((cost == 0).sum())
checks["invalid_cost_rows"] = int(cost.isna().sum() - ship["shipment_cost_inr"].isna().sum())

# timestamps
def parse_ts(s):
    return pd.to_datetime(s, errors="coerce", utc=True, format="mixed")

prom = parse_ts(ship["promised_delivery_ts"])
act = parse_ts(ship["actual_delivery_ts"])
checks["invalid_promised_ts_rows"] = int(prom.isna().sum() - ship["promised_delivery_ts"].isna().sum())
checks["invalid_actual_ts_rows"] = int(act.isna().sum() - ship["actual_delivery_ts"].isna().sum())
checks["missing_promised_ts_rows"] = int(ship["promised_delivery_ts"].isna().sum())
checks["missing_actual_ts_rows"] = int(ship["actual_delivery_ts"].isna().sum())

# hubs
for col in ("origin_hub", "destination_hub"):
    bad = ~ship[col].isin(known_hub_tokens) & ship[col].notna()
    checks[f"unrecognised_{col}_rows"] = int(bad.sum())
    checks[f"messy_variant_{col}_rows"] = int(
        ship[col].isin(set(crosswalk["raw_hub_id"]) - set(hubs["hub_id"])).sum())

# route / vehicle refs
checks["invalid_route_ref_rows"] = int((~ship["route_id"].isin(set(routes["route_id"]))).sum())
checks["invalid_vehicle_ref_rows"] = int(((~ship["vehicle_id"].isin(set(vehicles["vehicle_id"]))) & ship["vehicle_id"].notna()).sum())
checks["missing_vehicle_rows"] = int(ship["vehicle_id"].isna().sum())
checks["missing_category_rows"] = int(ship["category"].isna().sum())

# scans
scan_ts = parse_ts(scans["scan_ts"])
checks["invalid_scan_ts_rows"] = int(scan_ts.isna().sum())
checks["orphan_scan_rows"] = int((~scans["shipment_id"].isin(set(ship["shipment_id"])) ).sum())

# out-of-order scans: within a shipment, file order vs chronological order
scans2 = scans.copy()
scans2["ts_parsed"] = scan_ts
scans2["file_order"] = np.arange(len(scans2))
valid = scans2.dropna(subset=["ts_parsed"])
grp = valid.sort_values(["shipment_id", "file_order"]).groupby("shipment_id")["ts_parsed"]
ooo_shipments = 0
ooo_events = 0
for _, g in grp:
    diffs = g.diff().dropna()
    bad = (diffs < pd.Timedelta(0)).sum()
    if bad:
        ooo_shipments += 1
        ooo_events += int(bad)
checks["out_of_order_scan_shipments"] = int(ooo_shipments)
checks["out_of_order_scan_events"] = int(ooo_events)

# negative dwell: per shipment, Loading->Departure and Arrival->Delivered pairs
piv = valid.pivot_table(index="shipment_id", columns="scan_type",
                        values="ts_parsed", aggfunc="first")
neg_origin = neg_dest = 0
if "Loading" in piv and "Departure" in piv:
    neg_origin = int(((piv["Departure"] - piv["Loading"]) < pd.Timedelta(0)).sum())
if "Arrival" in piv and "Delivered" in piv:
    neg_dest = int(((piv["Delivered"] - piv["Arrival"]) < pd.Timedelta(0)).sum())
checks["negative_dwell_origin_hub_shipments"] = neg_origin
checks["negative_dwell_dest_hub_shipments"] = neg_dest

with open(os.path.join(OUT, "data_quality_checks.json"), "w") as f:
    json.dump(checks, f, indent=2)

# ---- markdown report ------------------------------------------------------
lines = ["# Data Quality Report — RAW data (synthetic)\n",
         "Source: `data/raw/shipments_raw.parquet` (304,500 rows) and "
         "`data/raw/scan_events_raw.parquet` (1,200,000 rows). "
         "All data is synthetic, generated with seed 42.\n",
         "## Integrity checks\n",
         "| Check | Count |",
         "|---|---|"]
for k, v in checks.items():
    lines.append(f"| {k} | {v:,} |")
lines.append("\n## Column profiles\n")
lines.append("Full per-column profiles: `column_profile_shipments.csv` and "
             "`column_profile_scans.csv`.\n")
lines.append("## Key observations\n")
lines.append(f"- Duplicate shipment rows: {checks['duplicate_shipment_id_rows']:,} "
             f"({checks['duplicate_shipment_ids_distinct']:,} extra rows beyond the first copy).")
lines.append(f"- Negative costs: {checks['negative_cost_rows']:,} rows; zero costs: "
             f"{checks['zero_cost_rows']:,} rows — excluded from the cost KPI denominator.")
lines.append(f"- Invalid promised timestamps: {checks['invalid_promised_ts_rows']:,}; "
             f"invalid actual timestamps: {checks['invalid_actual_ts_rows']:,} — excluded from OTD.")
lines.append(f"- Orphan scan events: {checks['orphan_scan_rows']:,} rows referencing "
             "shipments that do not exist.")
lines.append(f"- Out-of-order scan events: {checks['out_of_order_scan_events']:,} events in "
             f"{checks['out_of_order_scan_shipments']:,} shipments — re-sequenced in cleaning.")
lines.append(f"- Negative dwell pairs: origin hub {neg_origin:,}, destination hub {neg_dest:,} "
             "shipments — flagged, excluded from dwell KPI.")
with open(os.path.join(OUT, "data_quality_report.md"), "w") as f:
    f.write("\n".join(lines))

print(json.dumps(checks, indent=2))
print("DONE")
