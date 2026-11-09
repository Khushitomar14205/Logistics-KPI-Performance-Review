"""
data_transformation.py — star-schema modelling.
Rebuilds the analytical model from the CLEANED tables (idempotent):
  data/processed/shipments_clean.parquet
  data/processed/scan_events_clean.parquet
into:
  data/processed/trusted_analysis_table.parquet  (one row per shipment)
  data/reference/dim_date.csv

Star schema
  FACT_SHIPMENT (grain: one row per shipment)
      PK shipment_id; FKs origin_hub, destination_hub -> Dim_Hub;
      route_id -> Dim_Route; vehicle_id -> Dim_Vehicle;
      category -> Dim_Category; warehouse_id -> Dim_Warehouse;
      shipment_date -> Dim_Date.
  FACT_SCAN_EVENT (grain: one row per warehouse scan event)
      PK scan_event_id; FK shipment_id -> Fact_Shipment;
      hub_id -> Dim_Hub; warehouse_id -> Dim_Warehouse.
  DIM_HUB (12 hubs), DIM_ROUTE (60 routes + UNKNOWN), DIM_VEHICLE (180),
  DIM_CATEGORY (5 + Unknown), DIM_WAREHOUSE (24), DIM_DATE (2025 calendar).

Join safety: every join from Fact_Shipment to a dimension is many-to-one on a
unique dimension key, so joins cannot fan out shipment rows. This is verified
in outputs/validation_report.md (row counts before/after each join).
"""
import os
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC = os.path.join(ROOT, "data", "processed")
REF = os.path.join(ROOT, "data", "reference")

ship = pd.read_parquet(os.path.join(PROC, "shipments_clean.parquet"))
scans = pd.read_parquet(os.path.join(PROC, "scan_events_clean.parquet"))

# --- dwell pairs from the scan-event fact -------------------------------
# (shipments_clean already carries dwell columns from cleaning; drop them so
# this rebuild is a genuine recomputation from the scan-event fact)
ship = ship.drop(columns=[c for c in
                          ("dwell_origin_h", "dwell_dest_h",
                           "dwell_origin_valid", "dwell_dest_valid")
                          if c in ship.columns])
piv = scans.dropna(subset=["scan_ts_utc"]).pivot_table(
    index="shipment_id", columns="scan_type", values="scan_ts_utc", aggfunc="first")
dwell = pd.DataFrame(index=piv.index)
if "Loading" in piv and "Departure" in piv:
    dwell["dwell_origin_h"] = (piv["Departure"] - piv["Loading"]).dt.total_seconds() / 3600
if "Arrival" in piv and "Delivered" in piv:
    dwell["dwell_dest_h"] = (piv["Delivered"] - piv["Arrival"]).dt.total_seconds() / 3600
dwell["dwell_origin_valid"] = True
dwell["dwell_dest_valid"] = True
for col in ("dwell_origin_h", "dwell_dest_h"):
    if col in dwell:
        neg = dwell[col] < 0
        dwell.loc[neg, col.replace("_h", "_valid")] = False
dwell = dwell.reset_index()

# --- trusted analysis table (one row per shipment) ----------------------
ship = ship.merge(dwell, on="shipment_id", how="left", validate="one_to_one")
ship["shipment_date"] = pd.to_datetime(ship["shipment_date"])
ship["month"] = ship["shipment_date"].dt.to_period("M").astype(str)
ship["on_time"] = pd.Series(
    np.where(ship["promised_ts_utc"].notna() & ship["actual_ts_utc"].notna(),
             ship["actual_ts_utc"] <= ship["promised_ts_utc"], np.nan),
    dtype="boolean", index=ship.index)
ship["late_hours"] = np.where(
    ship["promised_ts_utc"].notna() & ship["actual_ts_utc"].notna(),
    (ship["actual_ts_utc"] - ship["promised_ts_utc"]).dt.total_seconds() / 3600, np.nan)
ship["dwell_avg_h"] = ship[["dwell_origin_h", "dwell_dest_h"]].mean(axis=1)

trusted = ship[["shipment_id", "shipment_date", "month", "origin_hub", "destination_hub",
                "route_id", "vehicle_id", "vehicle_type", "warehouse_id", "category",
                "status", "promised_ts_utc", "actual_ts_utc", "shipment_cost_inr",
                "cost_valid", "on_time", "late_hours",
                "dwell_origin_h", "dwell_dest_h", "dwell_avg_h",
                "dwell_origin_valid", "dwell_dest_valid"]].copy()
assert trusted["shipment_id"].is_unique, "trusted table must have one row per shipment"
assert len(trusted) == len(ship), "join must not change shipment row count"
trusted.to_parquet(os.path.join(PROC, "trusted_analysis_table.parquet"), index=False)
trusted.head(10_000).to_csv(os.path.join(PROC, "trusted_analysis_table_sample.csv"), index=False)

# --- dim_date ------------------------------------------------------------
dim_date = pd.DataFrame({"date": pd.date_range("2025-01-01", "2025-12-31")})
dim_date["date_key"] = dim_date["date"].dt.strftime("%Y%m%d").astype(int)
dim_date["month"] = dim_date["date"].dt.to_period("M").astype(str)
dim_date["month_name"] = dim_date["date"].dt.strftime("%B")
dim_date["quarter"] = "Q" + dim_date["date"].dt.quarter.astype(str)
dim_date["year"] = 2025
dim_date["day_of_week"] = dim_date["date"].dt.day_name()
dim_date.to_csv(os.path.join(REF, "dim_date.csv"), index=False)

print(f"trusted_analysis_table: {len(trusted):,} rows, {trusted['shipment_id'].nunique():,} unique shipment_ids")
print("DONE")
