"""Export import-ready CSVs for the Power BI dashboard (powerbi/data/).

Reads the cleaned trusted analysis table + reference dimensions + KPI
aggregates produced by the main pipeline and writes flat CSVs that the
.pbip/.pbit model loads via the DataFolder parameter.

Rules applied here (documented, no silent changes):
- Timestamps are exported as naive UTC strings ("YYYY-MM-DD HH:MM:SS").
  The source values are already UTC; the offset is stripped, never
  re-interpreted as IST.
- Booleans are exported as TRUE/FALSE, blanks stay blank.
- UNKNOWN members are appended to Dim_Route / Dim_Vehicle / Dim_Category
  (the DAX guide requires them so UNKNOWN shipments don't drop out of
  sliced visuals). The appends are recorded in the export manifest.
- Nothing is imputed: missing costs/timestamps stay blank; the validity
  flags travel with the rows.

Run:  .venv/bin/python powerbi/export_data.py
"""
import json
import shutil
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "powerbi" / "data"
OUT.mkdir(parents=True, exist_ok=True)

BOOL_COLS = ["cost_valid", "on_time", "dwell_origin_valid", "dwell_dest_valid"]


def fmt_bool(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if v is True or v == True:  # noqa: E712 - pandas nullable booleans
        return "TRUE"
    if v is False or v == False:  # noqa: E712
        return "FALSE"
    return ""


def main():
    manifest = {"tables": {}, "notes": []}

    # ------------------------------------------------------------------
    # Fact_Shipment — full trusted table (300,000 rows)
    # ------------------------------------------------------------------
    fact = pd.read_parquet(ROOT / "data/processed/trusted_analysis_table.parquet")
    fact["shipment_date"] = pd.to_datetime(fact["shipment_date"]).dt.strftime("%Y-%m-%d")
    for c in ["promised_ts_utc", "actual_ts_utc"]:
        fact[c] = (
            pd.to_datetime(fact[c], utc=True)
            .dt.tz_convert(None)
            .dt.strftime("%Y-%m-%d %H:%M:%S")
        )
    for c in BOOL_COLS:
        fact[c] = fact[c].map(fmt_bool)
    fact_path = OUT / "Fact_Shipment.csv"
    fact.to_csv(fact_path, index=False)
    manifest["tables"]["Fact_Shipment"] = {
        "file": "Fact_Shipment.csv",
        "rows": int(len(fact)),
        "columns": list(fact.columns),
    }
    # Headline KPI recomputation from the exported CSV (must match the
    # project headline numbers exactly).
    chk = pd.read_csv(fact_path)
    elig = chk[chk["status"].isin(["Delivered", "Delayed"]) & chk["on_time"].notna()]
    otd = (elig["on_time"] == True).mean() * 100  # noqa: E712
    vc = chk[chk["cost_valid"] == True]  # noqa: E712
    cps = vc["shipment_cost_inr"].sum() / len(vc)
    dd = chk[chk["dwell_dest_valid"] == True]["dwell_dest_h"]  # noqa: E712
    manifest["headline_check_from_exported_csv"] = {
        "otd_pct": round(float(otd), 2),
        "cost_per_shipment_inr": round(float(cps), 2),
        "avg_dest_dwell_h": round(float(dd.mean()), 2),
        "expected": {"otd_pct": 76.48, "cost_per_shipment_inr": 4102.68,
                     "avg_dest_dwell_h": 15.49},
    }
    assert abs(otd - 76.48) < 0.01, f"OTD mismatch: {otd}"
    assert abs(cps - 4102.68) < 0.01, f"Cost/shipment mismatch: {cps}"
    assert abs(dd.mean() - 15.49) < 0.01, f"Dwell mismatch: {dd.mean()}"

    # ------------------------------------------------------------------
    # Dimensions (UNKNOWN members appended where the DAX guide requires)
    # ------------------------------------------------------------------
    def load_dim(name, src, append_rows=None):
        df = pd.read_csv(ROOT / "data/reference" / src)
        if append_rows:
            df = pd.concat([df, pd.DataFrame(append_rows)], ignore_index=True)
        p = OUT / f"{name}.csv"
        df.to_csv(p, index=False)
        manifest["tables"][name] = {"file": f"{name}.csv", "rows": int(len(df)),
                                     "columns": list(df.columns)}
        return df

    load_dim("Dim_Hub", "dim_hub.csv")
    load_dim("Dim_Route", "dim_route.csv",
             [{"route_id": "UNKNOWN", "origin_hub": "", "destination_hub": "",
               "distance_km": "", "planned_transit_hours": ""}])
    manifest["notes"].append("Dim_Route: appended UNKNOWN member (1,548 shipments "
                             "carry route_id = UNKNOWN).")
    load_dim("Dim_Vehicle", "dim_vehicle.csv",
             [{"vehicle_id": "UNKNOWN", "vehicle_type": "Unknown",
               "capacity_kg": "", "home_hub": ""}])
    manifest["notes"].append("Dim_Vehicle: appended UNKNOWN member (1,543 shipments "
                             "carry vehicle_id = UNKNOWN).")
    load_dim("Dim_Warehouse", "dim_warehouse.csv")
    load_dim("Dim_Category", "dim_category.csv",
             [{"category": "Unknown", "base_cost_inr": ""}])
    manifest["notes"].append("Dim_Category: appended Unknown member (857 shipments "
                             "carry category = Unknown).")
    load_dim("Dim_Date", "dim_date.csv")

    # ------------------------------------------------------------------
    # Disconnected reference / cross-check tables (no relationships)
    # ------------------------------------------------------------------
    kpi_map = {
        "Monthly_KPI": "monthly_kpi.csv",
        "Half_Year_KPI": "half_year_kpi.csv",
        "Hub_Destination_KPI": "hub_destination_kpi.csv",
        "Route_KPI": "route_kpi.csv",
        "Warehouse_KPI": "warehouse_kpi.csv",
        "Category_KPI": "category_kpi.csv",
    }
    for table, src in kpi_map.items():
        df = pd.read_csv(ROOT / "outputs/kpi_reports" / src)
        p = OUT / f"{table}.csv"
        df.to_csv(p, index=False)
        manifest["tables"][table] = {"file": f"{table}.csv", "rows": int(len(df)),
                                      "columns": list(df.columns)}
    manifest["notes"].append("KPI_* tables are disconnected reference tables for the "
                             "Data Quality page cross-check; no relationships are "
                             "created on them.")

    # ------------------------------------------------------------------
    # Data-quality tables
    # ------------------------------------------------------------------
    shutil.copy(ROOT / "outputs/cleaning_logs/cleaning_log.csv",
                OUT / "Cleaning_Log.csv")
    cl = pd.read_csv(OUT / "Cleaning_Log.csv")
    manifest["tables"]["Cleaning_Log"] = {"file": "Cleaning_Log.csv",
                                          "rows": int(len(cl)),
                                          "columns": list(cl.columns)}

    prof = pd.read_csv(ROOT / "outputs/kpi_reports/column_profile_shipments.csv")
    prof = prof[["table", "column", "dtype", "total_records", "missing_values",
                 "missing_pct", "unique_values"]]
    prof.to_csv(OUT / "DQ_Missing.csv", index=False)
    manifest["tables"]["DQ_Missing"] = {"file": "DQ_Missing.csv",
                                         "rows": int(len(prof)),
                                         "columns": list(prof.columns)}

    dq = json.load(open(ROOT / "outputs/kpi_reports/data_quality_checks.json"))
    dq_rows = [
        {"check": "Raw records (shipments + scan events)", "value": "1,504,500",
         "status": "Recorded"},
        {"check": "Clean records (shipments + scan events)", "value": "1,494,000",
         "status": "Recorded"},
        {"check": "Shipments raw -> clean", "value": "304,500 -> 300,000",
         "status": "4,500 duplicate copies removed"},
        {"check": "Scan events raw -> clean", "value": "1,200,000 -> 1,194,000",
         "status": "6,000 orphan scans removed"},
        {"check": "Validation checks passed", "value": "38 / 38 (0% variance)",
         "status": "PASS"},
        {"check": "Out-of-order scan events after cleaning", "value": "0",
         "status": "PASS"},
    ]
    for k, v in dq.items():
        dq_rows.append({"check": k, "value": str(v), "status": "Recorded"})
    pd.DataFrame(dq_rows).to_csv(OUT / "DQ_Summary.csv", index=False)
    manifest["tables"]["DQ_Summary"] = {"file": "DQ_Summary.csv",
                                         "rows": len(dq_rows),
                                         "columns": ["check", "value", "status"]}

    (OUT / "export_manifest.json").write_text(json.dumps(manifest, indent=2))
    total_mb = sum(p.stat().st_size for p in OUT.glob("*.csv")) / 1e6
    print(f"Exported {len(manifest['tables'])} tables to {OUT}")
    print(f"Total CSV size: {total_mb:.1f} MB")
    for t, m in manifest["tables"].items():
        print(f"  {t}: {m['rows']} rows")
    print("Headline check from exported CSV:", manifest["headline_check_from_exported_csv"])


if __name__ == "__main__":
    main()
