"""
visualizations.py — 8 professional charts (matplotlib) from the KPI exports.
All numbers come from outputs/kpi_reports/*.csv (computed, not invented).
Style: consistent navy/teal palette, titled, labelled, units shown.
"""
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KPI = os.path.join(ROOT, "outputs", "kpi_reports")
VIZ = os.path.join(ROOT, "visualizations")
os.makedirs(VIZ, exist_ok=True)

NAVY = "#1F3864"; TEAL = "#2E75B6"; GREEN = "#548235"; AMBER = "#BF8F00"
RED = "#C00000"; GREY = "#7F7F7F"; LIGHT = "#D9E2F3"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10,
                     "axes.titlesize": 12, "axes.titleweight": "bold",
                     "axes.labelsize": 10, "figure.dpi": 150})

monthly = pd.read_csv(os.path.join(KPI, "monthly_kpi.csv"))
hub_d = pd.read_csv(os.path.join(KPI, "hub_destination_kpi.csv"))
hub_o = pd.read_csv(os.path.join(KPI, "hub_origin_kpi.csv"))
route = pd.read_csv(os.path.join(KPI, "route_kpi.csv"))
wh = pd.read_csv(os.path.join(KPI, "warehouse_kpi.csv"))
cat = pd.read_csv(os.path.join(KPI, "category_kpi.csv"))
import json
with open(os.path.join(KPI, "kpi_summary.json")) as f:
    summary = json.load(f)
OVERALL_OTD = summary["overall"]["otd_pct"]
OVERALL_COST = summary["overall"]["cost_per_shipment_inr"]
OVERALL_DWELL = summary["overall"]["avg_dwell_dest_h"]

def save(fig, name):
    fig.tight_layout()
    fig.savefig(os.path.join(VIZ, name), bbox_inches="tight")
    plt.close(fig)
    print("saved", name)

# 1. Monthly OTD trend
fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(monthly["month"], monthly["otd_pct"], marker="o", color=TEAL, lw=2)
ax.axhline(OVERALL_OTD, color=GREY, ls="--", lw=1, label=f"Network avg {OVERALL_OTD}%")
ax.axhline(90, color=GREEN, ls=":", lw=1, label="Target 90%")
ax.set_title("Monthly On-Time Delivery Trend — 2025 (synthetic data)")
ax.set_xlabel("Month"); ax.set_ylabel("OTD (%)")
ax.set_ylim(60, 100); ax.grid(axis="y", alpha=0.3); ax.legend()
plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
save(fig, "01_monthly_otd_trend.png")

# 2. Hub-wise OTD comparison
d = hub_d[hub_d["destination_hub"] != "UNKNOWN"].sort_values("otd_pct")
colors = [RED if v < 70 else AMBER if v < 80 else GREEN for v in d["otd_pct"]]
fig, ax = plt.subplots(figsize=(10, 5.5))
ax.barh(d["destination_hub"], d["otd_pct"], color=colors)
ax.axvline(OVERALL_OTD, color=NAVY, ls="--", lw=1.5, label=f"Network avg {OVERALL_OTD}%")
ax.set_title("Hub-wise On-Time Delivery % (by destination hub)")
ax.set_xlabel("OTD (%)"); ax.set_ylabel("Destination hub")
ax.set_xlim(0, 100); ax.legend()
for i, v in enumerate(d["otd_pct"]):
    ax.text(v + 1, i, f"{v:.1f}%", va="center", fontsize=9)
save(fig, "02_hub_otd_comparison.png")

# 3. Hub-wise dwell time
d = hub_d[hub_d["destination_hub"] != "UNKNOWN"].sort_values("avg_dwell_dest_h")
fig, ax = plt.subplots(figsize=(10, 5.5))
ax.barh(d["destination_hub"], d["avg_dwell_dest_h"], color=TEAL)
ax.axvline(OVERALL_DWELL, color=NAVY, ls="--", lw=1.5, label=f"Network avg {OVERALL_DWELL} h")
ax.set_title("Hub-wise Average Destination Dwell Time (hours)")
ax.set_xlabel("Average dwell time (hours)"); ax.set_ylabel("Destination hub")
ax.legend()
for i, v in enumerate(d["avg_dwell_dest_h"]):
    ax.text(v + 0.3, i, f"{v:.1f} h", va="center", fontsize=9)
save(fig, "03_hub_dwell_time.png")

# 4. Route performance ranking (top 15 + bottom 15 by OTD)
r = route[route["route_id"] != "UNKNOWN"].sort_values("otd_pct")
sel = pd.concat([r.head(15), r.tail(15)])
fig, ax = plt.subplots(figsize=(10, 7))
cols = [RED if v < 70 else AMBER if v < 80 else GREEN for v in sel["otd_pct"]]
ax.barh(sel["route_id"], sel["otd_pct"], color=cols)
ax.axvline(OVERALL_OTD, color=NAVY, ls="--", lw=1.5, label=f"Network avg {OVERALL_OTD}%")
ax.set_title("Route Performance Ranking — OTD % (bottom 15 and top 15 of 60 routes)")
ax.set_xlabel("OTD (%)"); ax.set_ylabel("Route")
ax.set_xlim(0, 100); ax.legend()
save(fig, "04_route_performance_ranking.png")

# 5. Category performance
c = cat[cat["category"] != "Unknown"].sort_values("otd_pct")
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11, 5))
a1.bar(c["category"], c["otd_pct"], color=TEAL)
a1.axhline(OVERALL_OTD, color=NAVY, ls="--", lw=1, label=f"Network avg {OVERALL_OTD}%")
a1.set_title("OTD % by Category"); a1.set_ylabel("OTD (%)"); a1.set_ylim(0, 100)
a1.legend(); plt.setp(a1.get_xticklabels(), rotation=30, ha="right")
a2.bar(c["category"], c["cost_per_shipment_inr"], color=AMBER)
a2.axhline(OVERALL_COST, color=NAVY, ls="--", lw=1, label=f"Network avg ₹{OVERALL_COST:,.0f}")
a2.set_title("Cost per Shipment (₹) by Category"); a2.set_ylabel("INR")
a2.legend(); plt.setp(a2.get_xticklabels(), rotation=30, ha="right")
fig.suptitle("Shipment Category Performance", fontweight="bold", fontsize=12)
save(fig, "05_category_performance.png")

# 6. Cost per shipment analysis (monthly + distribution by category)
fig, (a1, a2) = plt.subplots(1, 2, figsize=(11, 5))
a1.plot(monthly["month"], monthly["cost_per_shipment_inr"], marker="o", color=AMBER, lw=2)
a1.axhline(OVERALL_COST, color=NAVY, ls="--", lw=1, label=f"Network avg ₹{OVERALL_COST:,.0f}")
a1.set_title("Cost per Shipment by Month (₹)"); a1.set_ylabel("INR")
a1.grid(axis="y", alpha=0.3); a1.legend()
plt.setp(a1.get_xticklabels(), rotation=45, ha="right")
a2.bar(wh.sort_values("cost_per_shipment_inr", ascending=False).head(10)["warehouse_id"],
       wh.sort_values("cost_per_shipment_inr", ascending=False).head(10)["cost_per_shipment_inr"],
       color=TEAL)
a2.set_title("Top 10 Warehouses by Cost per Shipment (₹)")
a2.set_ylabel("INR"); plt.setp(a2.get_xticklabels(), rotation=45, ha="right")
fig.suptitle("Cost per Shipment Analysis", fontweight="bold", fontsize=12)
save(fig, "06_cost_per_shipment.png")

# 7. Warehouse performance comparison
w = wh.sort_values("avg_dwell_dest_h", ascending=False)
fig, ax = plt.subplots(figsize=(10, 6))
ax.barh(w["warehouse_id"], w["avg_dwell_dest_h"], color=TEAL)
ax.axvline(OVERALL_DWELL, color=NAVY, ls="--", lw=1.5, label=f"Network avg {OVERALL_DWELL} h")
ax.set_title("Warehouse Performance — Average Destination Dwell Time (hours)")
ax.set_xlabel("Average dwell time (hours)"); ax.set_ylabel("Warehouse")
ax.legend()
save(fig, "07_warehouse_performance.png")

# 8. Late delivery Pareto by route
late = pd.read_csv(os.path.join(KPI, "route_kpi.csv"))
late = late[late["route_id"] != "UNKNOWN"].copy()
late["late_n"] = (late["otd_eligible"] * (100 - late["otd_pct"]) / 100).round(0)
late = late.sort_values("late_n", ascending=False)
late["cum_pct"] = late["late_n"].cumsum() / late["late_n"].sum() * 100
fig, ax = plt.subplots(figsize=(11, 5.5))
ax.bar(late["route_id"], late["late_n"], color=TEAL, label="Late deliveries")
ax2 = ax.twinx()
ax2.plot(late["route_id"], late["cum_pct"], color=RED, marker="o", ms=3, label="Cumulative %")
ax2.axhline(80, color=RED, ls="--", lw=1)
ax2.set_ylabel("Cumulative % of late deliveries")
ax.set_title("Late Delivery Contribution — Pareto by Route")
ax.set_xlabel("Route (ranked by late deliveries)"); ax.set_ylabel("Late deliveries")
plt.setp(ax.get_xticklabels(), rotation=90, fontsize=7)
ax.legend(loc="upper left"); ax2.legend(loc="upper right")
save(fig, "08_late_delivery_pareto.png")
print("DONE")
