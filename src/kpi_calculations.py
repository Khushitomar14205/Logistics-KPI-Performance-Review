"""
kpi_calculations.py — KPI calculations, segment comparisons, advanced
analysis (deciles, contribution, cohorts, root cause) and business impact.
All KPIs are computed from data/processed/trusted_analysis_table.parquet —
the single trusted analysis table. No other source is used for KPI numbers.

KPI definitions
  OTD %  = on-time eligible deliveries / eligible deliveries x 100.
           Eligible = status in (Delivered, Delayed) AND promised & actual
           timestamps both valid. Cancelled / Returned / In-Transit shipments
           are excluded from the denominator (they never completed delivery).
  Cost per shipment (INR) = sum(valid shipment costs) / count(valid costs).
           Costs flagged invalid (negative, zero, missing) are excluded.
  Hub dwell (hours) = departure ts - arrival ts at the same hub.
           Origin hub: Departure - Loading. Destination hub: Delivered - Arrival.
           Negative pairs are flagged invalid and excluded.
  YoY: the synthetic series covers one calendar year (2025) only, so no true
  prior-year period exists. "Prior period" comparisons use H1 2025 vs H2 2025
  and are labelled as such — never presented as year-over-year.
"""
import json
import os
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROC = os.path.join(ROOT, "data", "processed")
OUT = os.path.join(ROOT, "outputs", "kpi_reports")
os.makedirs(OUT, exist_ok=True)

t = pd.read_parquet(os.path.join(PROC, "trusted_analysis_table.parquet"))
t["shipment_date"] = pd.to_datetime(t["shipment_date"])
t["half"] = np.where(t["shipment_date"].dt.month <= 6, "H1 2025", "H2 2025")

# eligibility
t["otd_eligible"] = t["status"].isin(["Delivered", "Delayed"]) & t["on_time"].notna()
t["is_late"] = t["otd_eligible"] & (t["on_time"] == False)
t["is_ontime"] = t["otd_eligible"] & (t["on_time"] == True)
t["dwell_combined"] = t[["dwell_origin_h", "dwell_dest_h"]].mean(axis=1)


def kpis(df):
    el = df[df["otd_eligible"]]
    vc = df[df["cost_valid"]]
    dw = df[df["dwell_dest_valid"] & df["dwell_dest_h"].notna()]
    dwo = df[df["dwell_origin_valid"] & df["dwell_origin_h"].notna()]
    return {
        "shipments": int(len(df)),
        "otd_eligible": int(len(el)),
        "otd_pct": round(100 * el["on_time"].mean(), 2) if len(el) else None,
        "late_deliveries": int(el["is_late"].sum()),
        "cost_per_shipment_inr": round(vc["shipment_cost_inr"].mean(), 2) if len(vc) else None,
        "total_valid_cost_inr": round(vc["shipment_cost_inr"].sum(), 2),
        "valid_cost_shipments": int(len(vc)),
        "avg_dwell_dest_h": round(dw["dwell_dest_h"].mean(), 2) if len(dw) else None,
        "avg_dwell_origin_h": round(dwo["dwell_origin_h"].mean(), 2) if len(dwo) else None,
        "avg_dwell_combined_h": round(df["dwell_combined"].mean(), 2),
    }


def segment(df, col):
    rows = []
    for key, g in df.groupby(col, dropna=False):
        r = kpis(g)
        r[col] = key
        rows.append(r)
    out = pd.DataFrame(rows)
    return out


overall = kpis(t)
monthly = segment(t, "month")
hub_o = segment(t, "origin_hub")
hub_d = segment(t, "destination_hub")
route = segment(t, "route_id")
wh = segment(t, "warehouse_id")
cat = segment(t, "category")
veh = segment(t, "vehicle_type")
half = segment(t, "half")

# month-over-month changes
monthly = monthly.sort_values("month")
monthly["otd_mom_pp"] = monthly["otd_pct"].diff().round(2)
monthly["cost_mom_pct"] = monthly["cost_per_shipment_inr"].pct_change().mul(100).round(2)
monthly["dwell_mom_pct"] = monthly["avg_dwell_dest_h"].pct_change().mul(100).round(2)

for name, df in [("monthly_kpi", monthly), ("hub_origin_kpi", hub_o),
                 ("hub_destination_kpi", hub_d), ("route_kpi", route),
                 ("warehouse_kpi", wh), ("category_kpi", cat),
                 ("vehicle_type_kpi", veh), ("half_year_kpi", half)]:
    df.to_csv(os.path.join(OUT, f"{name}.csv"), index=False)

# ---- deciles ---------------------------------------------------------------
late_sorted = t[t["otd_eligible"]].sort_values("late_hours", ascending=False)
n_late = len(late_sorted)
bottom10_ship = late_sorted.head(max(1, n_late // 10))
top10_ship = late_sorted.tail(max(1, n_late // 10))
route_sorted = route.sort_values("otd_pct")
wh_sorted = wh.sort_values("avg_dwell_dest_h", ascending=False)

deciles = {
    "shipments_bottom10pc_by_lateness": {
        "count": int(len(bottom10_ship)),
        "mean_late_hours": round(float(bottom10_ship["late_hours"].mean()), 2),
        "mean_cost_inr": round(float(bottom10_ship["shipment_cost_inr"].mean()), 2),
    },
    "shipments_top10pc_by_earliness": {
        "count": int(len(top10_ship)),
        "mean_late_hours": round(float(top10_ship["late_hours"].mean()), 2),
        "mean_cost_inr": round(float(top10_ship["shipment_cost_inr"].mean()), 2),
    },
    "routes_bottom10pc_otd": route_sorted.head(max(1, len(route_sorted) // 10))[
        ["route_id", "shipments", "otd_pct", "cost_per_shipment_inr"]].to_dict("records"),
    "routes_top10pc_otd": route_sorted.tail(max(1, len(route_sorted) // 10))[
        ["route_id", "shipments", "otd_pct", "cost_per_shipment_inr"]].to_dict("records"),
    "warehouses_bottom10pc_dwell": wh_sorted.head(max(1, len(wh_sorted) // 10))[
        ["warehouse_id", "shipments", "avg_dwell_dest_h", "otd_pct"]].to_dict("records"),
    "warehouses_top10pc_dwell": wh_sorted.tail(max(1, len(wh_sorted) // 10))[
        ["warehouse_id", "shipments", "avg_dwell_dest_h", "otd_pct"]].to_dict("records"),
}
bottom10_ship.to_csv(os.path.join(OUT, "bottom10pc_shipments.csv"), index=False)
top10_ship.to_csv(os.path.join(OUT, "top10pc_shipments.csv"), index=False)

# ---- contribution analysis -------------------------------------------------
def contrib(df, col, metric_col, label):
    g = df.groupby(col, dropna=False)[metric_col].sum()
    tot = g.sum()
    out = (g / tot * 100).sort_values(ascending=False).round(2)
    return out

late_by_route = contrib(t[t["is_late"]], "route_id", "is_late", "late")
late_by_hub = contrib(t[t["is_late"]], "destination_hub", "is_late", "late")
late_by_cat = contrib(t[t["is_late"]], "category", "is_late", "late")
late_by_veh = contrib(t[t["is_late"]], "vehicle_type", "is_late", "late")
cost_by_route = t[t["cost_valid"]].groupby("route_id")["shipment_cost_inr"].sum().sort_values(ascending=False)
dwell_by_wh = t.groupby("warehouse_id")["dwell_dest_h"].mean().sort_values(ascending=False)

contribution = {
    "late_deliveries_total": int(t["is_late"].sum()),
    "late_share_by_route_top5_pct": {k: float(v) for k, v in late_by_route.head(5).items()},
    "late_share_by_dest_hub_pct": {k: float(v) for k, v in late_by_hub.items()},
    "late_share_by_category_pct": {k: float(v) for k, v in late_by_cat.items()},
    "late_share_by_vehicle_type_pct": {k: float(v) for k, v in late_by_veh.items()},
    "top5_routes_by_total_cost_inr": {k: round(float(v), 2) for k, v in cost_by_route.head(5).items()},
    "top5_warehouses_by_avg_dest_dwell_h": {k: round(float(v), 2) for k, v in dwell_by_wh.head(5).items()},
    # Pareto: how many routes account for 80% of late deliveries
    "pareto_routes_for_80pc_late": int((late_by_route.cumsum() <= 80).sum() + 1),
    "pareto_routes_total": int(len(late_by_route)),
}

# ---- cohorts ---------------------------------------------------------------
cohort_mc = t.groupby(["month", "category"]).agg(
    shipments=("shipment_id", "count"),
    otd_pct=("on_time", lambda s: round(100 * s.mean(), 2))).reset_index()
cohort_mc.to_csv(os.path.join(OUT, "cohort_month_category.csv"), index=False)
cohort_mh = t.groupby(["month", "destination_hub"]).agg(
    shipments=("shipment_id", "count"),
    otd_pct=("on_time", lambda s: round(100 * s.mean(), 2))).reset_index()
cohort_mh.to_csv(os.path.join(OUT, "cohort_month_hub.csv"), index=False)

# ---- business impact -------------------------------------------------------
# Routes: bottom quartile by OTD -> median OTD
r = route[route["route_id"] != "UNKNOWN"].copy()
r = r[r["otd_eligible"] >= 100]  # only routes with meaningful volume
q1_otd = r["otd_pct"].quantile(0.25)
median_otd = r["otd_pct"].median()
bottom_routes = r[r["otd_pct"] <= q1_otd].copy()
bottom_routes["gap_pp"] = median_otd - bottom_routes["otd_pct"]
bottom_routes["late_avoided_at_median"] = (
    bottom_routes["otd_eligible"] * bottom_routes["gap_pp"] / 100).round(0)
route_impact = {
    "routes_analysed": int(len(r)),
    "bottom_quartile_otd_threshold_pct": round(float(q1_otd), 2),
    "median_route_otd_pct": round(float(median_otd), 2),
    "bottom_quartile_routes": int(len(bottom_routes)),
    "bottom_quartile_eligible_shipments": int(bottom_routes["otd_eligible"].sum()),
    "bottom_quartile_current_otd_pct": round(
        float(100 - (bottom_routes["otd_eligible"] * (100 - bottom_routes["otd_pct"])).sum()
              / bottom_routes["otd_eligible"].sum()), 2) if bottom_routes["otd_eligible"].sum() else None,
    "late_deliveries_avoided_at_median": int(bottom_routes["late_avoided_at_median"].sum()),
}
# cost: bottom quartile by cost/shipment -> median cost/shipment
r2 = route[route["route_id"] != "UNKNOWN"].copy()
r2 = r2[r2["valid_cost_shipments"] >= 100]
q3_cost = r2["cost_per_shipment_inr"].quantile(0.75)
median_cost = r2["cost_per_shipment_inr"].median()
bottom_cost = r2[r2["cost_per_shipment_inr"] >= q3_cost].copy()
bottom_cost["saving_per_shipment"] = bottom_cost["cost_per_shipment_inr"] - median_cost
bottom_cost["saving_inr"] = bottom_cost["saving_per_shipment"] * bottom_cost["valid_cost_shipments"]
cost_impact = {
    "bottom_quartile_cost_threshold_inr": round(float(q3_cost), 2),
    "median_route_cost_per_shipment_inr": round(float(median_cost), 2),
    "bottom_quartile_routes": int(len(bottom_cost)),
    "bottom_quartile_valid_cost_shipments": int(bottom_cost["valid_cost_shipments"].sum()),
    "potential_saving_inr_at_median": round(float(bottom_cost["saving_inr"].sum()), 2),
}
# warehouses: bottom quartile by destination dwell -> median dwell
w = wh.copy()
w = w[w["shipments"] >= 100]
q3_dwell = w["avg_dwell_dest_h"].quantile(0.75)
median_dwell = w["avg_dwell_dest_h"].median()
bottom_wh = w[w["avg_dwell_dest_h"] >= q3_dwell].copy()
bottom_wh["excess_dwell_h_per_shipment"] = bottom_wh["avg_dwell_dest_h"] - median_dwell
bottom_wh["excess_dwell_h_total"] = bottom_wh["excess_dwell_h_per_shipment"] * bottom_wh["shipments"]
wh_impact = {
    "bottom_quartile_dwell_threshold_h": round(float(q3_dwell), 2),
    "median_warehouse_dest_dwell_h": round(float(median_dwell), 2),
    "bottom_quartile_warehouses": int(len(bottom_wh)),
    "bottom_quartile_shipments": int(bottom_wh["shipments"].sum()),
    "excess_dwell_hours_at_median": round(float(bottom_wh["excess_dwell_h_total"].sum()), 1),
}

scenarios = {}
for name, frac in [("conservative", 0.5), ("realistic", 1.0), ("optimistic", 1.0)]:
    late_avoided = route_impact["late_deliveries_avoided_at_median"] * frac
    saving = cost_impact["potential_saving_inr_at_median"] * frac
    if name == "optimistic":
        saving *= 1.10  # +10% efficiency gain on top of median convergence
    scenarios[name] = {
        "gap_closed_pct": int(frac * 100),
        "late_deliveries_avoided": int(round(late_avoided)),
        "potential_cost_saving_inr": round(float(saving), 2),
        "excess_dwell_hours_recovered": round(
            wh_impact["excess_dwell_hours_at_median"] * frac, 1),
    }

business_impact = {
    "routes_otd": route_impact,
    "routes_cost": cost_impact,
    "warehouses_dwell": wh_impact,
    "scenarios": scenarios,
    "assumptions": [
        "Modelled estimates on SYNTHETIC data; not a forecast of real savings.",
        "Improvement is modelled as bottom-quartile segments converging to the network median; no causal mechanism is assumed.",
        "Cost saving uses the route cost-per-shipment gap only; dwell-hour savings are reported in hours because no verified per-hour handling cost exists in the data.",
        "No behaviour change by customers, no capacity constraints and no implementation costs are included.",
        "Optimistic scenario adds a further 10% efficiency gain on top of median convergence; it is not guaranteed.",
    ],
}

summary = {
    "overall": overall,
    "half_year": half.to_dict("records"),
    "deciles": deciles,
    "contribution": contribution,
    "business_impact": business_impact,
    "notes": {
        "otd_definition": "Eligible = status Delivered/Delayed with valid promised and actual timestamps.",
        "yoy_note": "No true YoY: data covers 2025 only. H1 vs H2 2025 is the prior-period comparison.",
        "data": "SYNTHETIC data, seed 42. Not real company data.",
    },
}
with open(os.path.join(OUT, "kpi_summary.json"), "w") as f:
    json.dump(summary, f, indent=2, default=str)

print(json.dumps({"overall": overall,
                  "route_impact": route_impact,
                  "cost_impact": cost_impact,
                  "wh_impact": wh_impact,
                  "scenarios": scenarios}, indent=2, default=str))
print("DONE")
