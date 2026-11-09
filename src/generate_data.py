"""
generate_data.py — Logistics KPI Performance Review (SYNTHETIC DATA)
Generates messy raw logistics data for a fictional Kolkata-based cold-chain
distributor (12 hubs), 2025-01-01 .. 2025-12-31.

Scale target: ~300,000 shipments + ~1,200,000 warehouse scan events
(≈1.5M records total). All data is SYNTHETIC, generated with a fixed seed
for reproducibility. It is NOT real company data and must never be presented
as such.

Deliberate, documented data-quality defects injected (see README / cleaning log):
  * duplicate shipment_ids (~1.5%)
  * missing values (cost, timestamps, vehicle, category)
  * negative / zero shipment costs
  * invalid timestamp strings and mixed timezones (UTC 'Z', IST '+05:30', naive)
  * inconsistent hub identifiers (raw variants) — see hub crosswalk
  * invalid route / vehicle references (orphans)
  * out-of-order scan events within a shipment (~5% of shipments)
  * negative dwell (departure before arrival) in ~1% of scan pairs
  * orphan scan events referencing non-existent shipments (~0.5%)

Outputs:
  data/raw/shipments_raw.parquet
  data/raw/scan_events_raw.parquet
  data/reference/dim_hub.csv, dim_route.csv, dim_vehicle.csv,
                   dim_warehouse.csv, dim_category.csv, hub_crosswalk.csv
  data/raw/generation_manifest.json
"""
import json
import os
import numpy as np
import pandas as pd

SEED = 42
rng = np.random.default_rng(SEED)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "data", "raw")
REF = os.path.join(ROOT, "data", "reference")
os.makedirs(RAW, exist_ok=True)
os.makedirs(REF, exist_ok=True)

N_SHIPMENTS = 300_000
SCANS_PER_SHIPMENT = 4
START = pd.Timestamp("2025-01-01")
END = pd.Timestamp("2025-12-31")

# ---------------------------------------------------------------- hubs
HUBS = [
    ("KOL", "Kolkata", "West Bengal"),
    ("HWH", "Howrah", "West Bengal"),
    ("SIL", "Siliguri", "West Bengal"),
    ("DGP", "Durgapur", "West Bengal"),
    ("ASN", "Asansol", "West Bengal"),
    ("MLD", "Malda", "West Bengal"),
    ("GAU", "Guwahati", "Assam"),
    ("PAT", "Patna", "Bihar"),
    ("RNC", "Ranchi", "Jharkhand"),
    ("JSR", "Jamshedpur", "Jharkhand"),
    ("BBS", "Bhubaneswar", "Odisha"),
    ("CTC", "Cuttack", "Odisha"),
]
hub_codes = [h[0] for h in HUBS]
hub_names = {h[0]: h[1] for h in HUBS}

# Raw (messy) variants of each canonical hub id -> crosswalk
crosswalk_rows = []
for code, name, state in HUBS:
    variants = [
        f"{code}-01",
        f"{code.lower()}_{1}",
        f"{name} Hub 1",
        f"{name.upper()}",
        f"hub-{code.lower()}",
    ]
    for v in variants:
        crosswalk_rows.append({"raw_hub_id": v, "canonical_hub_id": code,
                               "hub_name": name, "state": state})
crosswalk = pd.DataFrame(crosswalk_rows)
crosswalk.to_csv(os.path.join(REF, "hub_crosswalk.csv"), index=False)

dim_hub = pd.DataFrame(
    [{"hub_id": c, "hub_name": n, "state": s, "region": "East India",
      "is_cold_chain": True} for c, n, s in HUBS])
dim_hub.to_csv(os.path.join(REF, "dim_hub.csv"), index=False)

# ---------------------------------------------------------------- routes
# 60 routes: origin != destination, weighted so Kolkata is the main origin
route_rows, route_keys = [], []
rid = 1
for o in hub_codes:
    for d in hub_codes:
        if o == d:
            continue
        route_rows.append({"route_id": f"RT-{rid:03d}", "origin_hub": o,
                           "destination_hub": d,
                           "distance_km": int(rng.integers(80, 900)),
                           "planned_transit_hours": int(rng.integers(6, 48))})
        route_keys.append((o, d))
        rid += 1
        if rid > 60:
            break
    if rid > 60:
        break
routes = pd.DataFrame(route_rows)
routes.to_csv(os.path.join(REF, "dim_route.csv"), index=False)
route_by_pair = {(r.origin_hub, r.destination_hub): r for r in routes.itertuples()}

# ---------------------------------------------------------------- vehicles
VEH_TYPES = ["Reefer Truck", "Reefer Van", "Insulated Truck",
             "Refrigerated Container", "Cold Storage Van"]
veh_rows = []
for i in range(1, 181):
    veh_rows.append({"vehicle_id": f"VEH-{i:03d}",
                     "vehicle_type": VEH_TYPES[i % len(VEH_TYPES)],
                     "capacity_kg": int(rng.integers(800, 12000)),
                     "home_hub": hub_codes[i % 12]})
vehicles = pd.DataFrame(veh_rows)
vehicles.to_csv(os.path.join(REF, "dim_vehicle.csv"), index=False)

# ---------------------------------------------------------------- warehouses
wh_rows = []
for i, (code, name, state) in enumerate(HUBS):
    for j in (1, 2):
        wh_rows.append({"warehouse_id": f"WH-{code}-{j}",
                        "hub_id": code, "warehouse_name": f"{name} Cold Store {j}",
                        "capacity_pallets": int(rng.integers(400, 2500))})
warehouses = pd.DataFrame(wh_rows)
warehouses.to_csv(os.path.join(REF, "dim_warehouse.csv"), index=False)
wh_by_hub = {c: [f"WH-{c}-1", f"WH-{c}-2"] for c in hub_codes}

# ---------------------------------------------------------------- categories
CATEGORIES = ["Frozen", "Chilled", "Ambient", "Pharma", "Dairy"]
cat_base_cost = {"Frozen": 4200, "Chilled": 3600, "Ambient": 2400,
                 "Pharma": 6800, "Dairy": 3100}
pd.DataFrame([{"category": c, "base_cost_inr": v}
              for c, v in cat_base_cost.items()]).to_csv(
    os.path.join(REF, "dim_category.csv"), index=False)

# ---------------------------------------------------------------- shipments
print("Generating shipments ...")
ship_dates = START + pd.to_timedelta(
    rng.integers(0, 365, size=N_SHIPMENTS), unit="D")
# origin hub: Kolkata weighted 35%
origin = rng.choice(hub_codes, size=N_SHIPMENTS,
                    p=[0.35, 0.12, 0.08, 0.07, 0.06, 0.05, 0.08, 0.07, 0.05, 0.04, 0.02, 0.01])
dest = np.empty(N_SHIPMENTS, dtype=object)
for i, o in enumerate(origin):
    choices = [c for c in hub_codes if c != o]
    dest[i] = choices[rng.integers(len(choices))]

route_id = np.empty(N_SHIPMENTS, dtype=object)
transit_h = np.empty(N_SHIPMENTS, dtype=float)
for i in range(N_SHIPMENTS):
    r = route_by_pair.get((origin[i], dest[i]))
    if r is None:  # pair not in the 60-route table -> nearest route id reused
        r = routes.iloc[rng.integers(len(routes))]
    route_id[i] = r.route_id
    transit_h[i] = r.planned_transit_hours

category = rng.choice(CATEGORIES, size=N_SHIPMENTS, p=[0.28, 0.26, 0.18, 0.12, 0.16])
vehicle = rng.choice(vehicles.vehicle_id.values, size=N_SHIPMENTS)
warehouse = np.array([rng.choice(wh_by_hub[o]) for o in origin])

# promised / actual delivery (UTC base)
dispatch_hour = rng.integers(0, 24, size=N_SHIPMENTS)
dispatch = ship_dates + pd.to_timedelta(dispatch_hour, unit="h")
promised = dispatch + pd.to_timedelta(transit_h, unit="h")
# delay: most deliveries arrive a little early (promised times carry buffer);
# ~10% are genuinely late with a long tail. Destination-hub and category
# effects create real segment variation for root-cause analysis.
delay_h = rng.normal(-6, 4, size=N_SHIPMENTS)
late = rng.random(N_SHIPMENTS) < 0.10
delay_h = np.where(late, np.abs(rng.normal(10, 8, size=N_SHIPMENTS)), delay_h)
hub_penalty = {"GAU": 5, "MLD": 4, "RNC": 3, "JSR": 2, "BBS": 2}
delay_h = delay_h + np.array([hub_penalty.get(d, 0) for d in dest])
delay_h = delay_h + np.where(category == "Pharma", -2, 0)
delay_h = delay_h + np.where(category == "Frozen", 1, 0)
actual = promised + pd.to_timedelta(delay_h, unit="h")

status = np.where(rng.random(N_SHIPMENTS) < 0.88, "Delivered", "Delivered")
r = rng.random(N_SHIPMENTS)
status = np.where((r >= 0.88) & (r < 0.93), "Delayed", status)
status = np.where((r >= 0.93) & (r < 0.96), "In Transit", status)
status = np.where((r >= 0.96) & (r < 0.985), "Cancelled", status)
status = np.where(r >= 0.985, "Returned", status)

cost = np.array([rng.lognormal(np.log(cat_base_cost[c]), 0.35) for c in category])
cost = np.round(cost, 2)

shipment_id = np.array([f"SHP-{i+1:06d}" for i in range(N_SHIPMENTS)], dtype=object)

ship = pd.DataFrame({
    "shipment_id": shipment_id,
    "shipment_date": ship_dates.strftime("%Y-%m-%d"),
    "origin_hub": origin,
    "destination_hub": dest,
    "route_id": route_id,
    "vehicle_id": vehicle,
    "vehicle_type": [VEH_TYPES[0]] * N_SHIPMENTS,
    "warehouse_id": warehouse,
    "category": category,
    "promised_delivery_ts": promised,
    "actual_delivery_ts": actual,
    "shipment_cost_inr": cost,
    "status": status,
})
# vehicle_type from vehicle table
vtype_map = dict(zip(vehicles.vehicle_id, vehicles.vehicle_type))
ship["vehicle_type"] = ship["vehicle_id"].map(vtype_map)

# --- messy hub identifiers: 30% of origin/destination use raw variants
raw_variant = dict(zip(crosswalk.raw_hub_id, crosswalk.canonical_hub_id))
variant_pool = {c: crosswalk[crosswalk.canonical_hub_id == c].raw_hub_id.tolist()
                for c in hub_codes}
for col in ("origin_hub", "destination_hub"):
    mask = rng.random(N_SHIPMENTS) < 0.30
    idx = np.where(mask)[0]
    ship.loc[idx, col] = [rng.choice(variant_pool[c]) for c in ship.loc[idx, col]]

# --- timestamp formatting: mixed timezones / invalid strings
def fmt_ts(series, frac_ist=0.35, frac_invalid=0.01, seed=SEED):
    r2 = np.random.default_rng(seed + 7)
    out = []
    for ts in series:
        u = r2.random()
        if u < frac_invalid:
            out.append(r2.choice(["2025-13-45 25:99:00", "not-a-date",
                                  "31/02/2025 10:00", ""]))
        elif u < frac_invalid + frac_ist:
            out.append((ts + pd.Timedelta(hours=5, minutes=30)).strftime(
                "%Y-%m-%dT%H:%M:%S+05:30"))
        elif u < frac_invalid + frac_ist + 0.45:
            out.append(ts.strftime("%Y-%m-%dT%H:%M:%SZ"))
        else:
            out.append(ts.strftime("%Y-%m-%d %H:%M:%S"))  # naive
    return out

ship["promised_delivery_ts"] = fmt_ts(ship["promised_delivery_ts"], seed=SEED)
ship["actual_delivery_ts"] = fmt_ts(ship["actual_delivery_ts"], seed=SEED + 1)

# --- missing values
for col, frac in [("shipment_cost_inr", 0.01), ("vehicle_id", 0.005),
                  ("category", 0.003), ("promised_delivery_ts", 0.008),
                  ("actual_delivery_ts", 0.012)]:
    mask = rng.random(N_SHIPMENTS) < frac
    ship.loc[mask, col] = np.nan if col != "promised_delivery_ts" else None
    if col in ("promised_delivery_ts", "actual_delivery_ts"):
        ship.loc[mask, col] = None

# --- negative / zero costs
neg = rng.random(N_SHIPMENTS) < 0.008
ship.loc[neg, "shipment_cost_inr"] = -np.round(rng.uniform(100, 5000, size=neg.sum()), 2)
zero = rng.random(N_SHIPMENTS) < 0.004
ship.loc[zero, "shipment_cost_inr"] = 0.0

# --- invalid route / vehicle references (orphans)
bad_route = rng.random(N_SHIPMENTS) < 0.005
ship.loc[bad_route, "route_id"] = "RT-999"
bad_veh = rng.random(N_SHIPMENTS) < 0.005
ship.loc[bad_veh, "vehicle_id"] = "VEH-999"

# --- duplicate shipment_ids (~1.5%): exact copies + copies with changed cost
n_dup = int(N_SHIPMENTS * 0.015)
dup_idx = rng.choice(N_SHIPMENTS, size=n_dup, replace=False)
dups = ship.iloc[dup_idx].copy()
dups.loc[dups.sample(frac=0.5, random_state=SEED).index, "shipment_cost_inr"] *= 1.1
ship = pd.concat([ship, dups], ignore_index=True)
ship = ship.sample(frac=1.0, random_state=SEED).reset_index(drop=True)

ship.to_parquet(os.path.join(RAW, "shipments_raw.parquet"), index=False)
print(f"shipments_raw: {len(ship):,} rows (incl. {n_dup:,} injected duplicates)")

# ---------------------------------------------------------------- scan events
print("Generating scan events ...")
n_base = N_SHIPMENTS
scan_rows = []
# vectorised per-shipment event generation in chunks
chunk = 50_000
all_scans = []
for start_i in range(0, n_base, chunk):
    end_i = min(start_i + chunk, n_base)
    sub = ship.iloc[start_i:end_i]
    # scan hub ids intentionally carry the same raw (messy) hub identifiers
    # as the shipment record; cleaning standardises them via the crosswalk
    n = len(sub)
    dest_canon = np.array([raw_variant.get(d, d) for d in sub["destination_hub"].values])
    dest_wh = np.array([rng.choice(wh_by_hub[d]) for d in dest_canon])
    base = pd.to_datetime(sub["shipment_date"]) + pd.to_timedelta(
        rng.integers(0, 24, size=n), unit="h")
    # event offsets (hours from dispatch): Loading 0, Depart origin 2, Arrive dest, Unload, Delivered
    t_load = base
    t_dep = base + pd.to_timedelta(rng.uniform(1, 4, size=n), unit="h")
    dwell = rng.uniform(1, 30, size=n)  # hub dwell hours
    t_arr = t_dep + pd.to_timedelta(rng.uniform(4, 40, size=n), unit="h")
    t_unl = t_arr + pd.to_timedelta(dwell, unit="h")
    # 1% negative dwell: unloading/delivery before arrival at destination hub
    neg_dwell = rng.random(n) < 0.01
    t_unl = np.where(neg_dwell, t_arr - pd.to_timedelta(rng.uniform(1, 5, size=n), unit="h"), t_unl)

    ev = pd.DataFrame({
        "shipment_id": np.repeat(sub["shipment_id"].values, SCANS_PER_SHIPMENT),
        "hub_id": np.stack([sub["origin_hub"].values, sub["origin_hub"].values,
                            sub["destination_hub"].values,
                            sub["destination_hub"].values], axis=1).ravel(),
        "warehouse_id": np.stack([sub["warehouse_id"].values, sub["warehouse_id"].values,
                                  dest_wh, dest_wh], axis=1).ravel(),
        "scan_type": np.tile(["Loading", "Departure", "Arrival", "Delivered"], n),
        "scan_ts_utc": np.stack([t_load.values, t_dep.values, t_arr.values, np.asarray(t_unl)], axis=1).ravel(),
    })
    all_scans.append(ev)
scans = pd.concat(all_scans, ignore_index=True)
scans["scan_event_id"] = [f"SCN-{i+1:07d}" for i in range(len(scans))]

# out-of-order: shuffle rows within ~5% of shipments (file order scrambled)
shuf_ids = rng.choice(ship["shipment_id"].unique(), size=int(n_base * 0.05), replace=False)
mask = scans["shipment_id"].isin(set(shuf_ids))
shuffled = scans[mask].sample(frac=1.0, random_state=SEED)
scans.loc[mask] = shuffled.values

# orphan scans (~0.5%): reference non-existent shipments
n_orph = int(len(scans) * 0.005)
orph_idx = rng.choice(len(scans), size=n_orph, replace=False)
scans.loc[orph_idx, "shipment_id"] = [
    f"SHP-9{900000 + i:06d}" for i in range(n_orph)]

# mixed timezone formatting for scan_ts
def fmt_scan(ts_series):
    r2 = np.random.default_rng(SEED + 11)
    out = []
    for ts in ts_series:
        u = r2.random()
        ts = pd.Timestamp(ts)
        if u < 0.30:
            out.append((ts + pd.Timedelta(hours=5, minutes=30)).strftime("%Y-%m-%dT%H:%M:%S+05:30"))
        elif u < 0.75:
            out.append(ts.strftime("%Y-%m-%dT%H:%M:%SZ"))
        else:
            out.append(ts.strftime("%Y-%m-%d %H:%M:%S"))
    return out

scans["scan_ts"] = fmt_scan(scans["scan_ts_utc"])
scans = scans[["scan_event_id", "shipment_id", "hub_id", "warehouse_id",
               "scan_type", "scan_ts"]]
scans.to_parquet(os.path.join(RAW, "scan_events_raw.parquet"), index=False)
print(f"scan_events_raw: {len(scans):,} rows (incl. {n_orph:,} orphan scans)")

manifest = {
    "project": "Logistics KPI Performance Review",
    "data_type": "SYNTHETIC — generated, not real company data",
    "seed": SEED,
    "date_range": ["2025-01-01", "2025-12-31"],
    "hubs": 12,
    "routes": len(routes),
    "vehicles": len(vehicles),
    "warehouses": len(warehouses),
    "shipments_raw_rows": int(len(ship)),
    "scan_events_raw_rows": int(len(scans)),
    "total_raw_records": int(len(ship) + len(scans)),
    "injected_defects": {
        "duplicate_shipment_rows": n_dup,
        "orphan_scan_rows": n_orph,
        "negative_dwell_shipments_pct": 1.0,
        "out_of_order_shipments_pct": 5.0,
        "messy_hub_identifier_pct": 30.0,
        "invalid_route_ref_pct": 0.5,
        "invalid_vehicle_ref_pct": 0.5,
    },
}
with open(os.path.join(RAW, "generation_manifest.json"), "w") as f:
    json.dump(manifest, f, indent=2)
print("Manifest:", json.dumps(manifest, indent=2))
print("DONE")
