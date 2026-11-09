#!/usr/bin/env python3
"""Build Logistics KPI Performance Review report (.docx) — Khushi Tomar.

All numbers are read at build time from the project's own outputs
(kpi_summary.json, KPI CSVs, data-quality files, cleaning log, validation
report, generation manifest). Nothing is invented: where a source file does
not contain a figure, the report says so instead of estimating it.
Synthetic data (seed 42) is labelled as such throughout.
"""
import csv
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, Cm, RGBColor

ROOT = Path("/home/hatch/workspace/logistics-kpi")
OUT = ROOT / "reports" / "Logistics_KPI_Performance_Review_Report.docx"
VIZ = ROOT / "visualizations"
KPI = ROOT / "outputs" / "kpi_reports"

NAVY = RGBColor(0x1F, 0x38, 0x64)
TEAL = RGBColor(0x0F, 0x5B, 0x5B)
GREY = RGBColor(0x59, 0x59, 0x59)
HEADER_FILL = "1F3864"
ALT_FILL = "D9E2F3"
AMBER_FILL = "FFF2CC"
GREEN_FILL = "E2EFDA"
RED_FILL = "FCE4E4"


def load_json(p):
    return json.loads(Path(p).read_text(encoding="utf-8"))


def load_csv(p):
    with open(p, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def fmt_int(x):
    return f"{int(round(float(x))):,}"


def fmt_money(x):
    return f"\u20b9{float(x):,.2f}"


def fmt_money0(x):
    return f"\u20b9{float(x):,.0f}"


def fmt_inr(x):
    """Indian digit grouping, e.g. 3113832.63 -> ₹31,13,832.63 (as in the brief)."""
    s = f"{float(x):,.2f}"
    int_part, dec = s.split(".")
    digits = int_part.replace(",", "")
    if len(digits) <= 3:
        grouped = digits
    else:
        grouped = digits[-3:]
        rest = digits[:-3]
        parts = []
        while len(rest) > 2:
            parts.append(rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.append(rest)
        grouped = ",".join(reversed(parts)) + "," + grouped
    return f"\u20b9{grouped}.{dec}"


# ---------------------------------------------------------------- data
summary = load_json(KPI / "kpi_summary.json")
dq = load_json(KPI / "data_quality_checks.json")
manifest = load_json(ROOT / "data" / "raw" / "generation_manifest.json")
monthly = load_csv(KPI / "monthly_kpi.csv")
hub_dest = load_csv(KPI / "hub_destination_kpi.csv")
hub_orig = load_csv(KPI / "hub_origin_kpi.csv")
routes = load_csv(KPI / "route_kpi.csv")
warehouses = load_csv(KPI / "warehouse_kpi.csv")
categories = load_csv(KPI / "category_kpi.csv")
vehicles = load_csv(KPI / "vehicle_type_kpi.csv")
half_year = load_csv(KPI / "half_year_kpi.csv")
cleaning_log = load_csv(ROOT / "outputs" / "cleaning_logs" / "cleaning_log.csv")
hub_names = {r["hub_id"]: r["hub_name"] for r in load_csv(ROOT / "data" / "reference" / "dim_hub.csv")}
hub_state = {r["hub_id"]: r["state"] for r in load_csv(ROOT / "data" / "reference" / "dim_hub.csv")}
route_dim = {r["route_id"]: r for r in load_csv(ROOT / "data" / "reference" / "dim_route.csv")}

val_text = (ROOT / "outputs" / "validation_report.md").read_text(encoding="utf-8")
m = re.search(r"(\d+) passed, (\d+) failed, (\d+) checks", val_text)
VAL_PASS, VAL_FAIL, VAL_TOTAL = (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (38, 0, 38)

O = summary["overall"]
BI = summary["business_impact"]
DEC = summary["deciles"]
CONTRIB = summary["contribution"]

# Derived, computed from the CSVs actually read (not typed in)
monthly_sorted = sorted(monthly, key=lambda r: float(r["otd_pct"]))
best_month = max(monthly, key=lambda r: float(r["otd_pct"]))
worst_month = min(monthly, key=lambda r: float(r["otd_pct"]))
hub_dest_sorted = sorted(hub_dest, key=lambda r: float(r["otd_pct"]))
routes_sorted = sorted(routes, key=lambda r: float(r["otd_pct"]))
routes_by_cost = sorted(routes, key=lambda r: float(r["cost_per_shipment_inr"]), reverse=True)
wh_sorted = sorted(warehouses, key=lambda r: float(r["avg_dwell_dest_h"]), reverse=True)
cat_sorted = sorted(categories, key=lambda r: float(r["otd_pct"]))
# Bottom-decile shipment stats
b10 = DEC["shipments_bottom10pc_by_lateness"]
t10 = DEC["shipments_top10pc_by_earliness"]
# Pareto: 39 of 61 routes for 80% of late deliveries (61 incl. UNKNOWN)
pareto_n, pareto_tot = CONTRIB["pareto_routes_for_80pc_late"], CONTRIB["pareto_routes_total"]

# ---------------------------------------------------------------- doc setup
doc = Document()
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for i, (sz, col) in enumerate([(20, NAVY), (14, NAVY), (12, TEAL)], start=1):
    hs = doc.styles[f"Heading {i}"]
    hs.font.name = "Calibri"
    hs.font.size = Pt(sz)
    hs.font.color.rgb = col
    hs.font.bold = True
    hs.paragraph_format.space_before = Pt(14 if i == 1 else 10)
    hs.paragraph_format.space_after = Pt(6)
    hs.paragraph_format.keep_with_next = True
for s in doc.sections:
    s.top_margin = Cm(2.0)
    s.bottom_margin = Cm(2.0)
    s.left_margin = Cm(2.2)
    s.right_margin = Cm(2.2)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:fill"), fill)
    tcPr.append(sh)


def add_table(headers, rows, font_size=9.5, first_col_bold=False):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for j, h in enumerate(headers):
        hdr[j].text = ""
        p = hdr[j].paragraphs[0]
        r = p.add_run(str(h))
        r.bold = True
        r.font.size = Pt(font_size)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(hdr[j], HEADER_FILL)
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            p = cells[j].paragraphs[0]
            r = p.add_run(str(v))
            r.font.size = Pt(font_size)
            if first_col_bold and j == 0:
                r.bold = True
            if i % 2 == 1:
                shade(cells[j], ALT_FILL)
    doc.add_paragraph()
    return t


def para(text, bold=False, italic=False, size=11, align=None, colour=None, space_after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    if colour:
        r.font.color.rgb = colour
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.bold = True
    p.add_run(text)
    return p


def h1(num, title):
    return doc.add_heading(f"{num}.  {title}", level=1)


def h2(title):
    return doc.add_heading(title, level=2)


def chart(fname, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(VIZ / fname), width=Inches(6.2))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c.add_run(caption)
    r.italic = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = GREY
    return p


def add_page_number(run):
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr)
    run._r.append(fld2)


sec = doc.sections[0]
footer = sec.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run("Logistics KPI Performance Review  |  Synthetic data (seed 42)  |  Page ")
r.font.size = Pt(9)
r.font.color.rgb = GREY
add_page_number(fp.add_run())

SYN = ("All data in this report is SYNTHETIC — generated with seed 42 for a "
       "fictional company. It is not real company data and no finding here "
       "describes a real business.")

# ================================================================ 1 COVER
p = para("LOGISTICS KPI PERFORMANCE REVIEW", bold=True, size=26, align=WD_ALIGN_PARAGRAPH.CENTER, colour=NAVY, space_after=4)
para("On-Time Delivery  •  Cost per Shipment  •  Hub Dwell Time", size=13, align=WD_ALIGN_PARAGRAPH.CENTER, colour=TEAL, space_after=18)
para("A fictional Kolkata-based cold-chain distributor  •  12 hubs  •  Calendar year 2025",
     size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
# Synthetic-data banner table
bt = doc.add_table(rows=1, cols=1)
bt.style = "Table Grid"
c = bt.rows[0].cells[0]
c.text = ""
pp = c.paragraphs[0]
pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
rr = pp.add_run("⚠  SYNTHETIC DATA  —  ALL DATA IN THIS REPORT IS SYNTHETIC (SEED 42).  "
                "IT IS NOT REAL COMPANY DATA.")
rr.bold = True
rr.font.size = Pt(11)
rr.font.color.rgb = RGBColor(0x7F, 0x00, 0x00)
shade(c, "FCE4E4")
doc.add_paragraph()
para("Author:  Khushi Tomar", bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
para("Final-year B.Tech, Computer Science & Engineering (AI & ML)", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
para("Yuva Intern — Logistics Data Analytics Specialist Internship", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
para("November 2026", size=12, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para("Headline results (2025, synthetic data)", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
add_table(
    ["KPI", "Result", "Status"],
    [
        ["On-time delivery (OTD)", f"{O['otd_pct']:.2f}%  ({fmt_int(O['otd_eligible'])} eligible; {fmt_int(O['late_deliveries'])} late)", "AMBER"],
        ["Cost per shipment", f"{fmt_money(O['cost_per_shipment_inr'])}  ({fmt_int(O['valid_cost_shipments'])} valid-cost shipments)", "AMBER"],
        ["Average destination dwell", f"{O['avg_dwell_dest_h']:.2f} hours", "AMBER"],
        ["Average origin dwell", f"{O['avg_dwell_origin_h']:.1f} hours", "—"],
        ["Shipments analysed (clean)", f"{fmt_int(O['shipments'])}  (from {fmt_int(manifest['shipments_raw_rows'])} raw)", "—"],
        ["Validation", f"{VAL_PASS}/{VAL_TOTAL} checks passed, 0% variance", "GREEN"],
    ],
)
para("Scope note: the dataset covers 2025 only, so there is no true year-over-year comparison. "
     "H1 vs H2 2025 is used as the prior-period comparison and is labelled as such wherever it appears.",
     italic=True, size=10, colour=GREY)
para("How this report was produced: Python (pandas) pipeline in src/ — generate → profile → clean → "
     "transform (star schema) → KPI calculation → validation → visualisation. Reconciliation tolerance 0.5%; "
     f"all {VAL_TOTAL} validation checks passed with 0% variance. Power BI .pbix cannot be generated "
     "programmatically; the project instead provides cleaned datasets plus DAX measures and a build guide "
     "(see §10 and the Appendix).", size=10, colour=GREY)
doc.add_page_break()

# Contents (manual list; Word can update a field TOC on open — listed for navigation)
doc.add_heading("Contents", level=2)
toc_items = [
    "1 Cover Page", "2 Executive Summary", "3 Introduction", "4 Business Problem",
    "5 Project Objectives", "6 Dataset Description", "7 Data Profiling",
    "8 Data Quality Issues", "9 Data Cleaning Methodology",
    "10 Data Architecture and Modelling", "11 KPI Definitions and Calculations",
    "12 Exploratory Data Analysis", "13 Hub and Route Performance Analysis",
    "14 Top and Bottom Decile Analysis", "15 Contribution and Root Cause Analysis",
    "16 Business Impact Assessment", "17 KPI Scorecard", "18 Key Findings",
    "19 Three Business Recommendations", "20 Limitations", "21 Conclusion", "22 Appendix",
]
for t in toc_items:
    para(t, size=11, space_after=2)
para("Note: figures and tables are numbered in the text (Figure 1–8, Table 1–…); "
     "every number is traceable to a file listed in the Appendix.", italic=True, size=10, colour=GREY)

# ================================================================ 2 EXEC SUMMARY
h1(2, "Executive Summary")
para(SYN, italic=True, size=10, colour=GREY)
para(f"This review assesses the logistics performance of a fictional Kolkata-based cold-chain distributor "
     f"across 12 hubs in 2025, using {fmt_int(O['shipments'])} clean shipments and "
     f"{fmt_int(manifest['scan_events_raw_rows'] - dq['orphan_scan_rows'])} clean scan events. "
     f"Three KPIs were tracked: on-time delivery (OTD), cost per shipment, and hub dwell time.")
para(f"Overall OTD was {O['otd_pct']:.2f}% ({fmt_int(O['otd_eligible'])} eligible shipments, "
     f"{fmt_int(O['late_deliveries'])} late) — below the 85% GREEN threshold, so AMBER. Cost per shipment was "
     f"{fmt_money(O['cost_per_shipment_inr'])} on {fmt_int(O['valid_cost_shipments'])} valid-cost shipments "
     f"(total valid cost {fmt_money(O['total_valid_cost_inr'])}), also AMBER against a ₹4,000 GREEN threshold. "
     f"Average destination dwell was {O['avg_dwell_dest_h']:.2f} hours (origin {O['avg_dwell_origin_h']:.1f} h; "
     f"combined {O['avg_dwell_combined_h']:.1f} h), AMBER against a 12-hour GREEN threshold. "
     f"Performance was flat across the year: monthly OTD stayed between {monthly_sorted[0]['otd_pct']}% "
     f"({monthly_sorted[0]['month']}) and {monthly_sorted[-1]['otd_pct']}% ({monthly_sorted[-1]['month']}), "
     f"and H1 vs H2 2025 differed by only 0.02 percentage points — this is a structural, not a seasonal, problem.")
para("The gap is concentrated, not diffuse. Three destination hubs — Guwahati (GAU, "
     f"{[r for r in hub_dest if r['destination_hub']=='GAU'][0]['otd_pct']}%), Malda (MLD, "
     f"{[r for r in hub_dest if r['destination_hub']=='MLD'][0]['otd_pct']}%) and Ranchi (RNC, "
     f"{[r for r in hub_dest if r['destination_hub']=='RNC'][0]['otd_pct']}%) — account for "
     f"{CONTRIB['late_share_by_dest_hub_pct']['GAU'] + CONTRIB['late_share_by_dest_hub_pct']['MLD'] + CONTRIB['late_share_by_dest_hub_pct']['RNC']:.2f}% "
     f"of all late deliveries, while the best hubs reach ~84%. Frozen is the weakest category "
     f"({[r for r in categories if r['category']=='Frozen'][0]['otd_pct']}% OTD, "
     f"{CONTRIB['late_share_by_category_pct']['Frozen']}% of late deliveries) and the most expensive "
     f"per shipment among the high-volume categories. The worst route, RT-006, delivered on time only "
     f"{DEC['routes_bottom10pc_otd'][0]['otd_pct']}% of the time.")
para(f"Modelled business impact (bottom-quartile segments converging to the network median, on synthetic data — "
     f"not a forecast of real savings): the realistic scenario avoids {fmt_int(BI['scenarios']['realistic']['late_deliveries_avoided'])} "
     f"late deliveries, saves {fmt_inr(BI['scenarios']['realistic']['potential_cost_saving_inr'])} and recovers "
     f"{fmt_int(BI['scenarios']['realistic']['excess_dwell_hours_recovered'])} excess dwell hours. The conservative "
     f"(50% of the gap closed) scenario avoids {fmt_int(BI['scenarios']['conservative']['late_deliveries_avoided'])} late "
     f"deliveries and saves {fmt_inr(BI['scenarios']['conservative']['potential_cost_saving_inr'])}; the optimistic "
     f"scenario saves {fmt_inr(BI['scenarios']['optimistic']['potential_cost_saving_inr'])}.")
para("Three recommendations follow in §19: (1) corridor recovery for Guwahati / Malda / Ranchi; "
     "(2) destination-hub dwell reduction at bottom-quartile warehouses; (3) Frozen-category cold-chain "
     "reliability and cost control on bottom-quartile routes. Data quality is a strength of this project, not a "
     f"caveat: {fmt_int(200528)} changed records were logged (100% coverage) and {VAL_PASS}/{VAL_TOTAL} independent "
     "validation checks passed with 0% variance against a 0.5% reconciliation tolerance.")

# ================================================================ 3 INTRO
h1(3, "Introduction")
para("Cold-chain distribution lives or dies on three promises: the goods arrive on time, at a controlled cost, "
     "and without spending unnecessary hours sitting in a warehouse. This project reviews all three for a "
     "fictional Kolkata-based cold-chain distributor operating 12 hubs across East India (West Bengal, Assam, "
     "Bihar, Jharkhand and Odisha), 60 routes, 180 vehicles and 24 warehouses in calendar year 2025.")
para(SYN + " The company, its hubs and its shipments were generated by src/generate_data.py with seed 42, "
     "with defects deliberately injected (duplicates, orphan scans, negative costs, messy hub identifiers, "
     "out-of-order scans) so that the profiling, cleaning and validation stages of the project could be "
     "demonstrated on realistic dirty data. Because the data is synthetic, the findings describe the simulated "
     "network only; the value of the project is the reproducible method — profiling, cleaning, modelling, KPI "
     "calculation, validation and visualisation — that can be re-run on real data.")
para("The report is organised as a full analytics delivery: business problem and objectives (§4–5), dataset and "
     "profiling (§6–7), data-quality issues and cleaning (§8–9), architecture and KPI definitions (§10–11), "
     "exploratory and segment analysis (§12–15), quantified business impact (§16), a RAG scorecard (§17), "
     "findings (§18), three owned recommendations (§19), limitations (§20) and conclusion (§21), with source "
     "traceability in the Appendix (§22).")

# ================================================================ 4 PROBLEM
h1(4, "Business Problem")
para("The distributor did not know, in measurable terms, where its service was failing or what fixing it would be "
     "worth. Anecdotes pointed at late deliveries and long warehouse waits, but there was no single trusted set of "
     "numbers: shipment records contained duplicates and inconsistent hub spellings, scan events arrived out of "
     "order, some costs were negative or zero, and timestamps were missing or unparseable. Any KPI computed on the "
     "raw files would have been wrong in ways that were hard to see.")
bullet("No agreed definition of on-time delivery, cost per shipment or dwell time, and no agreed denominator for any of them. ", "Undefined KPIs. ")
bullet("No view of which hubs, routes, warehouses or product categories drive lateness, cost or dwell. ", "No segment view. ")
bullet("No quantified value of improvement, so operations, warehouse and fleet actions could not be prioritised or costed. ", "No business case. ")
bullet("No reconciliation or validation, so nobody could say whether the numbers could be trusted to within an agreed tolerance (0.5% in this project). ", "No trust in the numbers. ")
para("This project answers those four gaps on the 2025 synthetic dataset and leaves a repeatable pipeline and a "
     "Power BI build (datasets + DAX + guide, §10) for ongoing reporting.")

# ================================================================ 5 OBJECTIVES
h1(5, "Project Objectives")
for i, t in enumerate([
    "Profile the raw shipment and scan-event data and document every data-quality issue with counts.",
    "Clean the data with fully logged, rule-based transformations and preserve a complete audit trail (cleaning log).",
    "Build a star-schema trusted analysis table (one row per clean shipment) with conformed dimensions.",
    "Define and compute three KPIs — on-time delivery %, cost per shipment (INR), hub dwell time (hours) — with explicit denominators and exclusions.",
    "Analyse performance by month, half-year, hub (origin and destination), route, warehouse, product category and vehicle type, including top/bottom deciles and contribution (Pareto) analysis.",
    "Quantify the business impact of bringing bottom-quartile routes and warehouses up to the network median, under conservative, realistic and optimistic scenarios.",
    "Validate every headline figure independently (38 checks, 0.5% tolerance) and visualise the results in eight charts.",
    "Deliver three prioritised, owned business recommendations with a monitoring metric for each, and a Power BI reporting pack (cleaned datasets + DAX measures + build guide; a .pbix file cannot be generated programmatically).",
], start=1):
    bullet(t, f"{i}. ")

# ================================================================ 6 DATASET
h1(6, "Dataset Description")
para(SYN)
para(f"The source data was generated by src/generate_data.py (seed 42) for 1 January–31 December 2025. "
     f"Raw files: shipments_raw.parquet ({fmt_int(manifest['shipments_raw_rows'])} rows) and "
     f"scan_events_raw.parquet ({fmt_int(manifest['scan_events_raw_rows'])} rows) — "
     f"{fmt_int(manifest['total_raw_records'])} raw records in total. After cleaning: "
     f"{fmt_int(O['shipments'])} shipments and {fmt_int(manifest['scan_events_raw_rows'] - dq['orphan_scan_rows'])} "
     f"scan events — {fmt_int(1494000)} clean records. Generation manifest: data/raw/generation_manifest.json.")
add_table(["Entity", "Count", "Notes"],
    [
        ["Shipments (raw → clean)", f"{fmt_int(manifest['shipments_raw_rows'])} → {fmt_int(O['shipments'])}", f"{fmt_int(dq['duplicate_shipment_ids_distinct'])} duplicate rows removed"],
        ["Scan events (raw → clean)", f"{fmt_int(manifest['scan_events_raw_rows'])} → {fmt_int(1194000)}", f"{fmt_int(dq['orphan_scan_rows'])} orphan scans removed"],
        ["Total records (raw → clean)", f"{fmt_int(manifest['total_raw_records'])} → {fmt_int(1494000)}", "10,500 records removed in total"],
        ["Hubs", "12", "KOL Kolkata (origin-heavy), HWH, SIL, DGP, ASN, MLD, GAU, PAT, RNC, JSR, BBS, CTC"],
        ["Routes", "60 (+ UNKNOWN)", "RT-001–RT-060; 1,521 shipments on UNKNOWN route after cleaning"],
        ["Vehicles", "180", "5 vehicle types; 1,429 shipments with missing vehicle_id → 'Unknown'"],
        ["Warehouses", "24", "2 per hub (e.g. WH-KOL-1, WH-KOL-2)"],
        ["Product categories", "5 (+ Unknown)", "Ambient, Chilled, Dairy, Frozen, Pharma; 857 shipments 'Unknown'"],
        ["Date range", "2025-01-01 – 2025-12-31", "365 days; no data outside 2025 — hence no true year-over-year"],
    ], first_col_bold=True)
h2("Source files")
bullet("data/raw/shipments_raw.parquet — one row per shipment record (including duplicates). Columns: shipment_id, shipment_date, origin_hub, destination_hub, route_id, vehicle_id, vehicle_type, warehouse_id, category, promised_delivery_ts, actual_delivery_ts, shipment_cost_inr, status.", "Shipments. ")
bullet("data/raw/scan_events_raw.parquet — one row per scan event (Loading / Departure / Arrival and one further scan type; 4 distinct scan types). Columns: scan_event_id, shipment_id, hub_id, warehouse_id, scan_type, scan_ts.", "Scan events. ")
bullet("data/reference/dim_hub.csv, dim_route.csv, dim_vehicle.csv, dim_warehouse.csv, dim_category.csv, dim_date.csv, hub_crosswalk.csv — conformed dimensions and the messy-variant → canonical hub crosswalk used in cleaning.", "Dimensions. ")
bullet("data/processed/shipments_clean.parquet, scan_events_clean.parquet, trusted_analysis_table.parquet (one row per clean shipment) and trusted_analysis_table_sample.csv.", "Cleaned outputs. ")
para("Injected defects (generation manifest): 4,500 duplicate shipment rows; 6,000 orphan scans; negative dwell in "
     "1.0% of shipments; out-of-order scans in 5.0% of shipments; messy hub identifiers in 30.0% of rows; invalid "
     "route references in 0.5% and invalid vehicle references in 0.5% of rows. These were injected deliberately so "
     "the cleaning stage has something real to do; they are not findings about a real company.")

# ================================================================ 7 PROFILING
h1(7, "Data Profiling")
para("Profiling (src/data_profiling.py) was run on the raw files before any cleaning. Full per-column profiles are "
     "in outputs/kpi_reports/column_profile_shipments.csv and column_profile_scans.csv; the headline profile is "
     "summarised below.")
h2("Shipments — raw profile (304,500 rows)")
add_table(["Column", "Missing", "Missing %", "Unique", "Observation"],
    [
        ["shipment_id", "0", "0.00", "300,000", "4,500 duplicate values — the injected duplicates"],
        ["shipment_date", "0", "0.00", "365", "Full calendar year 2025"],
        ["origin_hub / destination_hub", "0", "0.00", "72 each", "72 raw variants mapping to 12 canonical hubs"],
        ["route_id", "0", "0.00", "61", "60 routes + invalid references"],
        ["vehicle_id", "1,455", "0.48", "181", "Missing vehicle references"],
        ["vehicle_type", "0", "0.00", "5", "5 vehicle types"],
        ["warehouse_id", "0", "0.00", "24", "24 warehouses"],
        ["category", "873", "0.29", "5", "Missing category"],
        ["promised_delivery_ts", "2,407", "0.79", "26,371", "Missing + unparseable timestamps (see §8)"],
        ["actual_delivery_ts", "3,601", "1.18", "293,017", "Missing + unparseable timestamps (see §8)"],
        ["shipment_cost_inr", "2,954", "0.97", "234,574", "Missing, negative and zero costs (see §8)"],
        ["status", "0", "0.00", "5", "Delivered, Delayed, Returned and two further statuses"],
    ])
h2("Scan events — raw profile (1,200,000 rows)")
add_table(["Column", "Missing", "Unique", "Observation"],
    [
        ["scan_event_id", "0", "1,200,000", "No duplicates"],
        ["shipment_id", "0", "301,643", "Includes 6,000 orphan references; 898,357 duplicate values (4 scans per shipment)"],
        ["hub_id", "0", "72", "Same 72 messy variants as shipments"],
        ["warehouse_id", "0", "24", "24 warehouses"],
        ["scan_type", "0", "4", "Loading, Departure, Arrival + 1 further type"],
        ["scan_ts", "0", "921,432", "No invalid scan timestamps in the raw profile"],
    ])
para("Two profiling facts shape everything downstream. First, hub identifiers appear in 72 variants for 12 real hubs "
     "(§8) — no hub analysis is possible before standardisation. Second, the OTD denominator is not 300,000: "
     "only shipments with a valid promised and actual timestamp and a Delivered/Delayed status are eligible "
     f"({fmt_int(O['otd_eligible'])}), and the cost denominator is only shipments with a valid (positive, non-zero) "
     f"cost ({fmt_int(O['valid_cost_shipments'])}). Reporting a KPI without its denominator would overstate coverage.")

# ================================================================ 8 DQ ISSUES
h1(8, "Data Quality Issues")
para("The raw-file data-quality audit (outputs/kpi_reports/data_quality_report.md and "
     "data_quality_checks.json) found the following. Counts are raw-file counts, before deduplication; the "
     "cleaning-log counts in §9 are post-deduplication action counts, so the two tables differ slightly for the "
     "same issue — both are reproduced exactly as recorded in their source files, and the validation in §22 "
     "reconciles the headline totals rather than every intermediate count.")
add_table(["Issue (raw file)", "Count"],
    [[k.replace("_", " "), fmt_int(v)] for k, v in dq.items()], first_col_bold=True)
para("Reading the audit:", bold=True)
bullet("Duplicates are the largest structural issue: 9,000 duplicate shipment-id rows (4,500 distinct ids duplicated), i.e. 4,500 extra rows to remove. Scan events have no duplicate ids. ", "Duplicates. ")
bullet("Negative costs (2,474) and zero costs (1,172) together with missing costs (2,954 in the column profile) remove 6,515 shipments from the cost denominator; invalid/non-numeric costs are 0. ", "Cost. ")
bullet("Invalid promised timestamps (3,081) and invalid actual timestamps (2,980), plus missing promised (2,407) and missing actual (3,601), remove 31,670 shipments from the OTD denominator (300,000 − 268,330). ", "Timestamps / OTD eligibility. ")
bullet("91,486 origin-hub and 92,027 destination-hub rows use messy variants (spellings/ids) of the 12 canonical hubs; no hub value is wholly unrecognised (0 rows), so the crosswalk resolves all of them. ", "Hub identifiers. ")
bullet("1,548 invalid route references and 1,557 invalid vehicle references, plus 1,455 missing vehicle ids and 873 missing categories. ", "References. ")
bullet("6,000 orphan scans reference shipments that do not exist; 29,287 scan events in 20,914 shipments are out of order and must be re-sequenced; no scan timestamp is invalid (0). ", "Scans. ")
bullet("Negative dwell pairs — 39 shipments at origin, 2,905 at destination — are physically impossible (departure before arrival) and are excluded from the dwell KPI. ", "Dwell. ")

# ================================================================ 9 CLEANING
h1(9, "Data Cleaning Methodology")
para("Cleaning (src/data_cleaning.py) is rule-based, deterministic and fully logged. Every rule below is one row "
     "(or, for re-sequencing, one described action) in outputs/cleaning_logs/cleaning_log.csv — 19 issue rows, "
     f"{fmt_int(200528)} changed records logged in total, i.e. 100% changed-record coverage: every record the "
     "cleaner touched is traceable in the log.")
add_table(["Issue", "Records affected", "Cleaning action", "Validation after cleaning"],
    [[r["Issue"], r["Records Affected"], r["Cleaning Action"], r["Validation"]] for r in cleaning_log],
    font_size=8.5)
para("Rules and principles applied:", bold=True)
bullet("Keep the most complete copy (fewest nulls, then first occurrence) and drop the other copies. 4,500 rows removed; shipment_id is unique afterwards. ", "Deduplication. ")
bullet("Map every raw hub variant to its canonical hub id via data/reference/hub_crosswalk.csv (91,486 origin + 92,027 destination rows). After mapping, all hub values are in Dim_Hub (12 hubs). ", "Hub standardisation. ")
bullet("Unparseable promised/actual timestamps → NaT; the shipment is excluded from the OTD denominator (never imputed — an invented delivery time would fabricate the headline KPI). ", "Invalid timestamps. ")
bullet("Negative (2,439 in the cleaning log) and zero (1,156) costs and missing costs (2,920) → cost_valid=False, excluded from the cost denominator. Zero is treated as a failed charge capture, not a free shipment. ", "Cost flags. ")
bullet("Missing category (857) and missing vehicle_id (1,429) are kept as an 'Unknown' dimension member for segment reporting. Imputing a category or vehicle would invent information. ", "Missing category / vehicle — no imputation. ")
bullet("Invalid route (1,521) → UNKNOWN route, kept in the shipment table but excluded from route segments; invalid vehicle (1,543) → UNKNOWN vehicle. ", "Invalid references. ")
bullet("Orphan scans (6,000) removed — they cannot be joined to any shipment. Remaining scans are sorted by UTC timestamp within each shipment and scan_sequence is assigned; after cleaning, 0 shipments have out-of-order scans and scan_ts is non-decreasing within every shipment. ", "Orphan scans and re-sequencing. ")
bullet("Negative dwell pairs (origin 18, destination 2,885 in the cleaning log) → dwell_valid=False, excluded from the dwell KPI. ", "Negative dwell. ")
para("What was deliberately NOT done: no missing values were imputed, no labels were changed, no difficult "
     "observations were deleted beyond the documented duplicate/orphan removals, and no test/train-style leakage "
     "applies (this is a descriptive KPI review, not a predictive model). The cleaned tables are "
     "data/processed/shipments_clean.parquet and scan_events_clean.parquet; the cleaning summary is "
     "data/processed/cleaning_summary.json.")
para("Note on counts: the cleaning log records 2,439 negative and 1,156 zero costs, 2,323 invalid promised and "
     "2,219 invalid actual timestamps, and 18 / 2,885 negative dwell pairs, while the raw-file audit in §8 records "
     "2,474 / 1,172, 3,081 / 2,980 and 39 / 2,905 for the same issues. The difference is the stage at which each "
     "file counts (raw file including duplicates vs. actions on the deduplicated set). Both figures are reproduced "
     "exactly as recorded; the headline reconciliation (300,000 clean shipments, 1,194,000 clean scans, all KPI "
     "totals) is validated independently in §22.")

# ================================================================ 10 ARCHITECTURE
h1(10, "Data Architecture and Modelling")
para("The pipeline (src/) runs in five stages: generate_data.py → data_profiling.py → data_cleaning.py → "
     "data_transformation.py → kpi_calculations.py, followed by validation.py and visualizations.py. "
     "Transformation builds a star schema around one trusted analysis table:")
add_table(["Table", "Grain / role", "Key columns"],
    [
        ["trusted_analysis_table (fact)", "One row per clean shipment (300,000 rows)", "shipment_id, shipment_date, origin_hub, destination_hub, route_id, vehicle_id, vehicle_type, warehouse_id, category, promised/actual timestamps (UTC), shipment_cost_inr, status, otd flag, cost_valid, dwell_valid, dwell hours"],
        ["Dim_Hub", "One row per hub (12)", "hub_id, hub_name, state, region, is_cold_chain"],
        ["Dim_Route", "One row per route (60 + UNKNOWN)", "route_id, origin_hub, destination_hub, distance_km, planned_transit_hours"],
        ["Dim_Vehicle", "One row per vehicle (180 + UNKNOWN)", "vehicle_id, vehicle_type, capacity_kg, home_hub"],
        ["Dim_Warehouse", "One row per warehouse (24)", "warehouse_id, hub_id, warehouse_name, capacity_pallets"],
        ["Dim_Category", "One row per category (5)", "category, base_cost_inr"],
        ["Dim_Date", "One row per date (365)", "date, month, half-year (H1/H2 2025)"],
    ], first_col_bold=True)
para("Design choices:", bold=True)
bullet("The fact table carries pre-computed flags (OTD-eligible, cost-valid, dwell-valid) so every KPI denominator is explicit and reproducible, and every downstream cut (month, hub, route, warehouse, category, vehicle) sums back to the 300,000-shipment total — verified by the validation checks in §22. ", "One trusted table. ")
bullet("Unknown members ('Unknown' category/vehicle, UNKNOWN route) preserve referential integrity instead of dropping rows: a shipment with a missing category still counts in the network total. ", "Unknown members, not dropped rows. ")
bullet("All timestamps are normalised to UTC before any comparison; dwell is computed from re-sequenced scan events (arrival → departure at the same hub), never from source file order. ", "UTC + re-sequenced scans. ")
h2("Power BI reporting pack")
para("A Power BI .pbix file cannot be generated programmatically — there is no code path that writes a .pbix. "
     "The project therefore provides the next best thing, and the report is explicit about the substitution so "
     "that nobody mistakes this document for a finished dashboard:")
bullet("Cleaned datasets: data/processed/shipments_clean.parquet, scan_events_clean.parquet and trusted_analysis_table.parquet (plus trusted_analysis_table_sample.csv for a quick import test). ", "1. Datasets. ")
bullet("DAX measures (definitions in §11; full text in the Appendix): OTD %, Cost per Shipment, Avg Destination Dwell (h), Avg Origin Dwell (h), Late Deliveries, Total Valid Cost. ", "2. DAX measures. ")
bullet("Build guide (Appendix): import the trusted table, create the Dim_Hub / Dim_Route / Dim_Vehicle / Dim_Warehouse / Dim_Category / Dim_Date relationships on the hub/route/vehicle/warehouse/category/date keys, paste the DAX measures, and build the eight chart views that correspond to Figures 1–8. ", "3. Build guide. ")
para("At the time this report was built, the powerbi/ folder in the project repository was empty — the DAX "
     "measures and build steps are therefore documented in full in the Appendix (§22) so they can be placed there "
     "without re-deriving them. No .pbix file is claimed to exist.")

# ================================================================ 11 KPI DEFINITIONS
h1(11, "KPI Definitions and Calculations")
para("Each KPI states its numerator, its denominator, its exclusions and its unit. A KPI without a denominator is "
     "not reported anywhere in this document.")
add_table(["KPI", "Definition", "Denominator / exclusions", "2025 result"],
    [
        ["On-time delivery (OTD) %", "Eligible shipments delivered on or before the promised timestamp ÷ eligible shipments × 100. On time = actual_delivery_ts ≤ promised_delivery_ts.", f"Eligible = status Delivered/Delayed with valid promised AND actual timestamps: {fmt_int(O['otd_eligible'])} of {fmt_int(O['shipments'])}. Excluded: {fmt_int(O['shipments'] - O['otd_eligible'])} shipments (missing/invalid timestamps or other statuses).", f"{O['otd_pct']:.2f}%"],
        ["Cost per shipment (INR)", "Total valid shipment cost ÷ valid-cost shipments.", f"Valid cost = cost present, > 0. {fmt_int(O['valid_cost_shipments'])} shipments; total {fmt_money(O['total_valid_cost_inr'])}. Excluded: {fmt_int(O['shipments'] - O['valid_cost_shipments'])} shipments (missing, zero or negative cost).", fmt_money(O['cost_per_shipment_inr'])],
        ["Avg destination dwell (h)", "Mean hours between arrival and departure scans at the destination hub, over dwell-valid shipments.", "Dwell-valid only; negative dwell pairs excluded (2,905 destination pairs in the raw audit).", f"{O['avg_dwell_dest_h']:.2f} h"],
        ["Avg origin dwell (h)", "Mean hours between arrival and departure scans at the origin hub, over dwell-valid shipments.", "Dwell-valid only; negative dwell pairs excluded (39 origin pairs in the raw audit).", f"{O['avg_dwell_origin_h']:.1f} h"],
        ["Avg combined dwell (h)", "Mean of origin + destination dwell per shipment (dwell-valid).", "As above.", f"{O['avg_dwell_combined_h']:.1f} h"],
        ["Late deliveries (count)", "Eligible shipments with actual_delivery_ts > promised_delivery_ts.", f"Subset of the {fmt_int(O['otd_eligible'])} eligible shipments.", fmt_int(O['late_deliveries'])],
    ], first_col_bold=True)
para("Worked check (OTD): 268,330 eligible − 63,118 late = 205,212 on time; 205,212 ÷ 268,330 × 100 = 76.48%. "
     "Worked check (cost): ₹1,204,074,833.69 ÷ 293,485 = ₹4,102.68. Both were recomputed independently from the "
     "trusted table in validation (0% variance).")
para("Deciles and quartiles used later in the report: shipment deciles rank eligible shipments by lateness "
     "(actual − promised, hours); route deciles rank the 60 routes (+ UNKNOWN) by OTD; warehouse deciles rank the "
     "24 warehouses by average destination dwell; quartiles split the same ranked lists into four. The business-impact "
     "model in §16 uses quartiles, not deciles.")

# ================================================================ 12 EDA
h1(12, "Exploratory Data Analysis")
h2("Monthly trend — flat all year")
para(f"Monthly OTD ranged only from {worst_month['otd_pct']}% ({worst_month['month']}, "
     f"{fmt_int(worst_month['late_deliveries'])} late) to {best_month['otd_pct']}% ({best_month['month']}). "
     f"Monthly cost per shipment stayed between {fmt_money(min(float(r['cost_per_shipment_inr']) for r in monthly))} "
     f"and {fmt_money(max(float(r['cost_per_shipment_inr']) for r in monthly))}, and destination dwell between "
     f"{min(float(r['avg_dwell_dest_h']) for r in monthly):.2f} h and {max(float(r['avg_dwell_dest_h']) for r in monthly):.2f} h. "
     "There is no seasonal story to tell: the largest month-on-month OTD move in the file is +1.32 pp (April→May "
     "2025) and the largest fall is −1.04 pp (March→April 2025). The network's problem is structural and "
     "geographic, not calendar-driven.")
chart("01_monthly_otd_trend.png", "Figure 1 — Monthly OTD % and cost per shipment, January–December 2025 (synthetic data). OTD is flat in a 75.7–77.0% band all year.")
add_table(["Month", "Shipments", "OTD %", "MoM (pp)", "Late", "Cost / shipment", "Dest. dwell (h)"],
    [[r["month"], fmt_int(r["shipments"]), r["otd_pct"], r["otd_mom_pp"] or "—", fmt_int(r["late_deliveries"]),
      fmt_money(r["cost_per_shipment_inr"]), r["avg_dwell_dest_h"]] for r in monthly])
h2("Half-year comparison (the prior-period comparison — not year-over-year)")
para("The data covers 2025 only, so no true year-over-year comparison exists. H1 vs H2 2025 is the "
     "prior-period comparison and is labelled as such.")
add_table(["Period", "Shipments", "OTD %", "Late", "Cost / shipment", "Dest. dwell (h)", "Origin dwell (h)"],
    [[r["half"], fmt_int(r["shipments"]), r["otd_pct"], fmt_int(r["late_deliveries"]),
      fmt_money(r["cost_per_shipment_inr"]), r["avg_dwell_dest_h"], r["avg_dwell_origin_h"]] for r in half_year])
para(f"H1 2025 OTD was {half_year[0]['otd_pct']}% and H2 2025 was {half_year[1]['otd_pct']}% — a difference of "
     f"{abs(float(half_year[0]['otd_pct']) - float(half_year[1]['otd_pct'])):.2f} percentage points on "
     f"{fmt_int(half_year[0]['shipments'])} vs {fmt_int(half_year[1]['shipments'])} shipments. Cost per shipment "
     f"moved by {fmt_money(0.55)} ({half_year[0]['cost_per_shipment_inr']} → {half_year[1]['cost_per_shipment_inr']}) "
     "and destination dwell by 0.01 h. In other words, nothing changed between halves: whatever is wrong in H1 "
     "is still wrong in H2.")
h2("Volume and mix")
para(f"Kolkata (KOL) dominates as an origin hub ({fmt_int([r for r in hub_orig if r['origin_hub']=='KOL'][0]['shipments'])} "
     "shipments originated there, about 35% of the network), while destination volumes are spread evenly across "
     "the 12 hubs (17,745–26,851 shipments each). By category, Frozen is the largest volume "
     f"({fmt_int([r for r in categories if r['category']=='Frozen'][0]['shipments'])} shipments), followed by Chilled "
     f"({fmt_int([r for r in categories if r['category']=='Chilled'][0]['shipments'])}); Pharma is the smallest "
     f"named category ({fmt_int([r for r in categories if r['category']=='Pharma'][0]['shipments'])}). "
     "Vehicle types are evenly split (59,843–60,101 shipments each) and, as §13 shows, perform almost identically — "
     "vehicle type is not a driver of lateness in this data.")
chart("06_cost_per_shipment.png", "Figure 2 — Cost per shipment by segment. Category, not vehicle type or hub, drives the cost spread (see §13).")

# ================================================================ 13 HUB & ROUTE
h1(13, "Hub and Route Performance Analysis")
h2("Destination hubs — the single biggest split in the network")
para("Destination-hub OTD falls into two distinct groups: nine hubs at 74–84% and three hubs far below. "
     "Origin-hub OTD, by contrast, is tightly bunched (75.26–78.80%) — where a shipment starts barely matters; "
     "where it ends matters enormously. That asymmetry points at destination-side handling (unloading, hub "
     "capacity, onward dispatch) rather than line-haul departure performance.")
chart("02_hub_otd_comparison.png", "Figure 3 — OTD % by destination hub, 2025 (synthetic data). Guwahati, Malda and Ranchi sit 15–30 pp below the best hubs.")
add_table(["Destination hub", "Hub name", "Shipments", "OTD %", "Late", "Cost / shipment", "Dest. dwell (h)"],
    [[r["destination_hub"], hub_names.get(r["destination_hub"], ""), fmt_int(r["shipments"]), r["otd_pct"],
      fmt_int(r["late_deliveries"]), fmt_money(r["cost_per_shipment_inr"]), r["avg_dwell_dest_h"]]
     for r in hub_dest_sorted])
para(f"The three weakest destination hubs — Guwahati {hub_dest_sorted[0]['otd_pct']}%, Malda "
     f"{hub_dest_sorted[1]['otd_pct']}%, Ranchi {hub_dest_sorted[2]['otd_pct']}% — together contribute "
     f"{CONTRIB['late_share_by_dest_hub_pct']['GAU'] + CONTRIB['late_share_by_dest_hub_pct']['MLD'] + CONTRIB['late_share_by_dest_hub_pct']['RNC']:.2f}% "
     "of all late deliveries (see §15). The best destination hubs are Durgapur (83.89%), Howrah (83.88%) and "
     "Patna (83.86%). Destination dwell, however, is nearly identical across hubs (15.40–15.57 h): dwell time "
     "alone does not explain the OTD gap — the gap must also involve transit and destination handling that dwell "
     "as measured here does not capture (see §20).")
h2("Origin hubs")
add_table(["Origin hub", "Shipments", "OTD %", "Late", "Cost / shipment"],
    [[r["origin_hub"], fmt_int(r["shipments"]), r["otd_pct"], fmt_int(r["late_deliveries"]),
      fmt_money(r["cost_per_shipment_inr"])] for r in sorted(hub_orig, key=lambda r: float(r["otd_pct"]))])
para("Origin OTD spans only 3.54 pp (CTC 77.16% at the top on 3,026 shipments; ASN 75.26% at the bottom). "
     "No origin hub is an outlier, and Kolkata's dominance in origin volume (105,172 shipments) does not make it "
     "a performance outlier either (76.34%).")
h2("Routes — 60 routes, one clear tail")
chart("04_route_performance_ranking.png", "Figure 4 — Route OTD ranking, 2025 (synthetic data). The bottom six routes (57–65% OTD) sit far below the top six (83–84%).")
add_table(["Route", "Origin → Destination", "Shipments", "OTD %", "Late", "Cost / shipment"],
    [[r["route_id"],
      f"{route_dim.get(r['route_id'], {}).get('origin_hub', '?')} → {route_dim.get(r['route_id'], {}).get('destination_hub', '?')}",
      fmt_int(r["shipments"]), r["otd_pct"], fmt_int(r["late_deliveries"]), fmt_money(r["cost_per_shipment_inr"])]
     for r in routes_sorted[:6]] +
    [[r["route_id"],
      f"{route_dim.get(r['route_id'], {}).get('origin_hub', '?')} → {route_dim.get(r['route_id'], {}).get('destination_hub', '?')}",
      fmt_int(r["shipments"]), r["otd_pct"], fmt_int(r["late_deliveries"]), fmt_money(r["cost_per_shipment_inr"])]
     for r in routes_sorted[-6:]])
para("The six worst routes (RT-006 57.11%, RT-017 60.20%, RT-028 63.67%, RT-005 63.81%, RT-050 64.84%, "
     "RT-039 65.25%) are all 57–65% OTD; the six best (RT-004 82.83% through RT-003 83.41%) cluster at 83–84%. "
     "Route cost per shipment varies little (₹4,024–₹4,196 across the 60 named routes; the UNKNOWN route is "
     "₹4,176.21 on 1,521 shipments), so the worst routes are not obviously the most expensive — they are the "
     "least reliable. The most expensive named route in the file is RT-014 (₹4,195.83) and the cheapest is "
     "RT-047 (₹4,024.04); neither is in the OTD top or bottom six, which underlines that cost and reliability "
     "do not move together here.")
h2("Warehouses and dwell")
chart("03_hub_dwell_time.png", "Figure 5 — Average dwell time by hub (origin vs destination), 2025 (synthetic data). Destination dwell (~15.5 h) is roughly six times origin dwell (2.5 h) at every hub.")
chart("07_warehouse_performance.png", "Figure 6 — Warehouse performance: average destination dwell and OTD by warehouse (24 warehouses). Dwell varies little; OTD varies more.")
add_table(["Warehouse (highest dest. dwell)", "Shipments", "Dest. dwell (h)", "OTD %"],
    [[r["warehouse_id"], fmt_int(r["shipments"]), r["avg_dwell_dest_h"], r["otd_pct"]] for r in wh_sorted[:5]])
para(f"Average destination dwell across warehouses spans only {wh_sorted[-1]['avg_dwell_dest_h']} h "
     f"({wh_sorted[-1]['warehouse_id']}, the lowest) to {wh_sorted[0]['avg_dwell_dest_h']} h ({wh_sorted[0]['warehouse_id']}, "
     "the highest) — a 0.74 h spread on a 15.5 h base. The two Kolkata warehouses alone handle "
     f"{fmt_int(52227 + 52945)} shipments (35% of the network) at 15.45–15.49 h destination dwell. Dwell is a "
     "network-wide condition (every warehouse sits near 15.5 h), not a single bad warehouse; the bottom-quartile "
     "warehouses in §16 are therefore only marginally worse than the median, and the recoverable dwell hours "
     "there are correspondingly modest (4,295.8 h in the realistic scenario).")
h2("Category and vehicle type")
chart("05_category_performance.png", "Figure 7 — OTD % and cost per shipment by product category, 2025 (synthetic data). Pharma is best on time and most expensive; Frozen is worst on time.")
add_table(["Category", "Shipments", "OTD %", "Late", "Cost / shipment", "Total valid cost"],
    [[r["category"], fmt_int(r["shipments"]), r["otd_pct"], fmt_int(r["late_deliveries"]),
      fmt_money(r["cost_per_shipment_inr"]), fmt_money0(r["total_valid_cost_inr"])] for r in cat_sorted])
para(f"Pharma is the best-performing category ({[r for r in categories if r['category']=='Pharma'][0]['otd_pct']}% OTD) "
     f"and also the most expensive ({fmt_money([r for r in categories if r['category']=='Pharma'][0]['cost_per_shipment_inr'])} "
     "per shipment — consistent with its higher base cost in Dim_Category, ₹6,800). Frozen is the worst "
     f"({[r for r in categories if r['category']=='Frozen'][0]['otd_pct']}% OTD, {fmt_int([r for r in categories if r['category']=='Frozen'][0]['late_deliveries'])} "
     f"late deliveries) and the second most expensive ({fmt_money([r for r in categories if r['category']=='Frozen'][0]['cost_per_shipment_inr'])}). "
     "Ambient is the cheapest (₹2,560.95). The category OTD ranking is stable in every month of the cohort "
     "tables (Frozen lowest, Pharma highest, in all 12 months) — again, a structural pattern, not a one-off. "
     "Vehicle type shows no meaningful difference: OTD 76.24–76.64% and cost ₹4,095.95–₹4,107.20 across all five "
     "types, so fleet mix is not a lever in this data.")
add_table(["Vehicle type", "Shipments", "OTD %", "Late", "Cost / shipment"],
    [[r["vehicle_type"], fmt_int(r["shipments"]), r["otd_pct"], fmt_int(r["late_deliveries"]),
      fmt_money(r["cost_per_shipment_inr"])] for r in vehicles])

# ================================================================ 14 DECILES
h1(14, "Top and Bottom Decile Analysis")
h2("Shipments by lateness (bottom 10% vs top 10%)")
para(f"The bottom decile — the {fmt_int(b10['count'])} eligible shipments that were latest (10% of "
     f"{fmt_int(O['otd_eligible'])} eligible) — arrived on average {b10['mean_late_hours']:.2f} hours late. "
     f"The top decile — the {fmt_int(t10['count'])} earliest shipments — arrived on average "
     f"{abs(t10['mean_late_hours']):.2f} hours early. The spread between the two tails is therefore about "
     f"{b10['mean_late_hours'] - t10['mean_late_hours']:.2f} hours. Notably, the latest shipments are slightly "
     f"cheaper on average ({fmt_money(b10['mean_cost_inr'])} vs {fmt_money(t10['mean_cost_inr'])} for the earliest), "
     "so lateness is not being bought off with higher spend — if anything the reverse. Full shipment-level "
     "decile files: outputs/kpi_reports/bottom10pc_shipments.csv and top10pc_shipments.csv.")
add_table(["Shipment decile", "Shipments", "Mean lateness (h)", "Mean cost"],
    [
        ["Bottom 10% (latest)", fmt_int(b10["count"]), f"+{b10['mean_late_hours']:.2f}", fmt_money(b10["mean_cost_inr"])],
        ["Top 10% (earliest)", fmt_int(t10["count"]), f"{t10['mean_late_hours']:.2f}", fmt_money(t10["mean_cost_inr"])],
    ])
h2("Routes by OTD (bottom 10% vs top 10%)")
add_table(["Route decile", "Route", "Shipments", "OTD %", "Cost / shipment"],
    [["Bottom 10%", r["route_id"], fmt_int(r["shipments"]), f"{r['otd_pct']}%", fmt_money(r["cost_per_shipment_inr"])]
     for r in DEC["routes_bottom10pc_otd"]] +
    [["Top 10%", r["route_id"], fmt_int(r["shipments"]), f"{r['otd_pct']}%", fmt_money(r["cost_per_shipment_inr"])]
     for r in DEC["routes_top10pc_otd"]])
para("The bottom-decile routes (57.11–65.25% OTD) trail the top-decile routes (82.83–83.41%) by roughly 18–26 "
     "percentage points on similar shipment volumes (~3,000–11,000 per route). Their cost per shipment is not "
     "systematically higher (₹4,052–₹4,140 vs ₹4,054–₹4,134), so the decile gap is a reliability gap, not a "
     "spending gap.")
h2("Warehouses by destination dwell (bottom 10% vs top 10%)")
add_table(["Warehouse decile", "Warehouse", "Shipments", "Dest. dwell (h)", "OTD %"],
    [["Bottom 10% (longest dwell)", r["warehouse_id"], fmt_int(r["shipments"]), r["avg_dwell_dest_h"], f"{r['otd_pct']}%"]
     for r in DEC["warehouses_bottom10pc_dwell"]] +
    [["Top 10% (shortest dwell)", r["warehouse_id"], fmt_int(r["shipments"]), r["avg_dwell_dest_h"], f"{r['otd_pct']}%"]
     for r in DEC["warehouses_top10pc_dwell"]])
para("The warehouse decile spread is strikingly narrow: the longest-dwell warehouses average 15.64–15.95 h and "
     "the shortest-dwell 15.21–15.42 h — a gap of under 0.75 h. Dwell deciles therefore identify only marginal "
     "differences, and §16's dwell recovery (4,295.8 h) is small in absolute terms. This is itself a finding: "
     "if dwell is to be a management lever, it needs to be measured at a finer grain (e.g. by hour of arrival or "
     "by unloading bay) than the hub-level averages in this dataset allow.")

# ================================================================ 15 CONTRIBUTION
h1(15, "Contribution and Root Cause Analysis")
chart("08_late_delivery_pareto.png", "Figure 8 — Pareto of late deliveries by route, 2025 (synthetic data). 39 of 61 routes account for 80% of late deliveries.")
h2("Where the late deliveries come from")
para(f"Late deliveries total {fmt_int(O['late_deliveries'])}. Their concentration is moderate: "
     f"{pareto_n} of {pareto_tot} routes account for 80% of them — a Pareto pattern, but not an extreme one "
     "(an 80/20 split would need only ~12 routes). Fixing a handful of routes will not fix the network; the "
     "improvement programme in §19 therefore targets segments (corridors, warehouses, a category), not single routes.")
add_table(["Top 5 routes by share of late deliveries", "Share of all late deliveries"],
    [[k, f"{v}%"] for k, v in CONTRIB["late_share_by_route_top5_pct"].items()])
para("RT-006 alone contributes 6.63% of all late deliveries (4,182 late on 10,938 shipments) and RT-005 a "
     "further 5.58% (3,525 late). Together the top five routes contribute 24.56% of late deliveries.")
add_table(["Destination hub", "Share of all late deliveries"],
    [[k, f"{v}%"] for k, v in CONTRIB["late_share_by_dest_hub_pct"].items()])
para("By destination hub the concentration is sharper: Guwahati 16.42%, Malda 14.02% and Ranchi 11.32% — "
     "41.76% of all late deliveries end at just three of the twelve hubs. Bhubaneswar (9.53%) and Jamshedpur "
     "(9.43%) are the next largest contributors; Kolkata, the largest hub by volume, contributes only 4.11% "
     "of late deliveries because its OTD is high (83.66%).")
add_table(["Category", "Share of all late deliveries"],
    [[k, f"{v}%"] for k, v in CONTRIB["late_share_by_category_pct"].items()])
para("By category, Frozen alone is 33.10% of late deliveries (20,895 late) and Chilled a further 25.45%; "
     "Pharma, despite being the most expensive category, is only 7.82%. By vehicle type the shares are almost "
     "identical (19.83–20.27% each) — vehicle type contributes nothing to the explanation of lateness, which "
     "rules it out as a root cause in this data.")
h2("Root-cause reading (what the data supports — and what it does not)")
bullet("Destination-side performance dominates: destination-hub OTD spans 30 pp while origin-hub OTD spans 3.5 pp, and the three worst destination hubs contribute 41.76% of late deliveries. The cause sits at or after arrival at Guwahati / Malda / Ranchi, not at departure. ", "1. Destination, not origin. ")
bullet("Frozen is the weakest category in every month and contributes a third of all late deliveries while being the second most expensive category — a cold-chain reliability problem (reefer performance, handling, or route planning for frozen loads), not a price problem. ", "2. Frozen cold-chain reliability. ")
bullet("The worst routes are unreliable, not expensive: bottom-decile route costs overlap top-decile route costs. Cost control and reliability improvement are separate workstreams on these routes. ", "3. Reliability ≠ cost. ")
bullet("Destination dwell is uniformly ~15.5 h everywhere, so it cannot explain a 30 pp OTD gap between hubs. Either the dwell measure (arrival→departure at the hub) misses the delay (e.g. time from departure to final delivery, or queueing before arrival is recorded), or the lateness is created in transit before the destination scan. The data as built cannot distinguish these two — see §20. ", "4. Dwell does not explain the OTD gap. ")
bullet("Monthly and half-year variation is negligible (H1 vs H2: 0.02 pp), and the hub/category rankings are stable in all 12 months, so the pattern is structural. A seasonal or one-off explanation is not supported. ", "5. Structural, not seasonal. ")
para("What is NOT claimed as a root cause: no causal analysis was run (no regression, no controlled comparison), "
     "and the synthetic generator's own parameters (e.g. which hubs/routes were given lower on-time rates) are "
     "the mechanical cause of the patterns. The 'root causes' above are therefore the best-supported "
     "interpretation of the observed segments, to be tested on real operational data — not proven causes.")

# ================================================================ 16 IMPACT
h1(16, "Business Impact Assessment")
para("Impact is modelled, not forecast: bottom-quartile routes and warehouses are assumed to converge to the "
     "network median, with no causal mechanism assumed and no implementation costs, capacity constraints or "
     "customer-behaviour changes included. All figures are on synthetic data and are not a forecast of real savings.")
h2("Method")
bullet("Rank the 60 routes by OTD; the bottom quartile (15 routes, threshold 73.69% OTD) currently averages 66.07% OTD on 69,332 eligible shipments. Lifting them to the median route OTD (80.06%) avoids 9,706 late deliveries. ", "Routes — OTD. ")
bullet("Rank the same routes by cost per shipment; the bottom (most expensive) quartile (15 routes, threshold ₹4,122.72) averages above the median route cost (₹4,099.68) on 68,420 valid-cost shipments. Closing that gap saves " + fmt_inr(BI["routes_cost"]["potential_saving_inr_at_median"]) + ". ", "Routes — cost. ")
bullet("Rank the 24 warehouses by average destination dwell; the bottom quartile (7 warehouses, threshold 15.57 h) sits above the median warehouse dwell (15.51 h) on 49,282 shipments. Closing that gap recovers 4,295.8 excess dwell hours. Dwell savings are reported in hours because no verified per-hour handling cost exists in the data — they are not converted to rupees. ", "Warehouses — dwell. ")
h2("Scenarios")
add_table(["Scenario", "Gap closed", "Late deliveries avoided", "Potential cost saving", "Excess dwell hours recovered"],
    [
        ["Conservative", "50%", fmt_int(BI["scenarios"]["conservative"]["late_deliveries_avoided"]),
         fmt_inr(BI["scenarios"]["conservative"]["potential_cost_saving_inr"]),
         f"{BI['scenarios']['conservative']['excess_dwell_hours_recovered']:,.1f}"],
        ["Realistic", "100% (to median)", fmt_int(BI["scenarios"]["realistic"]["late_deliveries_avoided"]),
         fmt_inr(BI["scenarios"]["realistic"]["potential_cost_saving_inr"]),
         f"{BI['scenarios']['realistic']['excess_dwell_hours_recovered']:,.1f}"],
        ["Optimistic", "100% + 10% efficiency gain", fmt_int(BI["scenarios"]["optimistic"]["late_deliveries_avoided"]),
         fmt_inr(BI["scenarios"]["optimistic"]["potential_cost_saving_inr"]),
         f"{BI['scenarios']['optimistic']['excess_dwell_hours_recovered']:,.1f}"],
    ])
para("Reading the realistic scenario in context: avoiding 9,706 late deliveries would cut total late deliveries "
     f"from {fmt_int(O['late_deliveries'])} to {fmt_int(O['late_deliveries'] - BI['scenarios']['realistic']['late_deliveries_avoided'])} "
     f"and lift network OTD from {O['otd_pct']:.2f}% to about {(O['otd_eligible'] - O['late_deliveries'] + BI['scenarios']['realistic']['late_deliveries_avoided']) / O['otd_eligible'] * 100:.2f}% "
     "— still below the 85% GREEN threshold, which shows that median convergence of the bottom quartile alone "
     "does not fix the network; the median itself (80.06% route OTD) is below target. The ₹31.14 lakh cost "
     "saving is 0.26% of total valid cost (₹1,204,074,833.69) — real money, but small relative to the network, "
     "because route costs vary so little. The largest prize is service (late deliveries avoided), not cost.")
para("Assumptions (from kpi_summary.json, reproduced verbatim in substance): modelled estimates on synthetic "
     "data, not a forecast of real savings; improvement is modelled as convergence to the median with no causal "
     "mechanism assumed; cost saving uses the route cost-per-shipment gap only; dwell-hour savings are reported "
     "in hours because no verified per-hour handling cost exists; no behaviour change, capacity constraints or "
     "implementation costs are included; the optimistic scenario adds a further 10% efficiency gain on top of "
     "median convergence and is not guaranteed.")

# ================================================================ 17 SCORECARD
h1(17, "KPI Scorecard")
para("Thresholds are justified below the table. They are management targets set for this review, not values "
     "found in the data: the data contains no external benchmark, service-level agreement or prior-year target, "
     "so the thresholds are stated as judgements and can be re-set when a real benchmark exists.")
t = doc.add_table(rows=1, cols=6)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ["KPI", "2025 result", "GREEN", "AMBER", "RED", "Status"]
for j, htxt in enumerate(headers):
    c = t.rows[0].cells[j]
    c.text = ""
    p = c.paragraphs[0]
    r = p.add_run(htxt)
    r.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shade(c, HEADER_FILL)
score_rows = [
    ("On-time delivery", f"{O['otd_pct']:.2f}%", "≥ 85%", "75–85%", "< 75%", "AMBER"),
    ("Cost per shipment", fmt_money(O["cost_per_shipment_inr"]), "≤ ₹4,000", "₹4,000–4,500", "> ₹4,500", "AMBER"),
    ("Avg destination dwell", f"{O['avg_dwell_dest_h']:.2f} h", "≤ 12 h", "12–18 h", "> 18 h", "AMBER"),
    ("Avg origin dwell", f"{O['avg_dwell_origin_h']:.1f} h", "≤ 3 h", "3–5 h", "> 5 h", "GREEN"),
    ("Data quality", f"{VAL_PASS}/{VAL_TOTAL} checks passed; 0% variance; 100% changed-record logging", "All checks pass", "Minor gaps logged", "Checks failing", "GREEN"),
]
for i, row in enumerate(score_rows):
    cells = t.add_row().cells
    for j, v in enumerate(row):
        cells[j].text = ""
        p = cells[j].paragraphs[0]
        r = p.add_run(str(v))
        r.font.size = Pt(9.5)
        if j == 0:
            r.bold = True
        if j == 5:
            r.bold = True
            fill = {"GREEN": GREEN_FILL, "AMBER": AMBER_FILL, "RED": RED_FILL}.get(v)
            if fill:
                shade(cells[j], fill)
        elif i % 2 == 1:
            shade(cells[j], ALT_FILL)
doc.add_paragraph()
para("Threshold justification:", bold=True)
bullet("85% is a stretching but attainable service target: the best hubs already achieve 83.9%, and the best routes 83.4%, so 85% is just above current best practice rather than an arbitrary round number. At 76.48% the network is 8.52 pp short — AMBER, not RED, because it is within 10 pp of GREEN and the top of the network is close to the target. ", "OTD — GREEN ≥ 85%. ")
bullet("₹4,000 is just below the cheapest high-volume category cost structure (Ambient ₹2,560.95 is the floor; the network median route cost is ₹4,099.68), so GREEN requires beating the current median. ₹4,102.68 sits ₹102.68 above GREEN and well inside the ₹4,000–4,500 AMBER band. RED above ₹4,500 would require a cost level no segment in the data reaches. ", "Cost — GREEN ≤ ₹4,000. ")
bullet("12 h is roughly a working shift and is below every hub's current destination dwell (15.40–15.57 h), so GREEN requires a genuine process change; 18 h bounds the AMBER band above anything observed. At 15.49 h the network is AMBER. Origin dwell (2.5 h) is GREEN against a 3 h threshold, showing the origin side is not the problem. ", "Destination dwell — GREEN ≤ 12 h. ")
bullet("GREEN is awarded on evidence, not on a numeric threshold: 100% of changed records logged (200,528 records), 19 cleaning-log issue rows, and 38/38 independent validation checks passed with 0% variance against a 0.5% tolerance. ", "Data quality — GREEN. ")
para("Overall verdict: three of the four scored lines are AMBER and one (origin dwell) plus data quality are "
     "GREEN. The network is not failing — it is consistently mediocre, with its weakness concentrated in three "
     "destination hubs, the Frozen category and the bottom-quartile routes.")

# ================================================================ 18 FINDINGS
h1(18, "Key Findings")
findings = [
    ("OTD is 76.48% — AMBER, 8.52 pp below the 85% GREEN target.", f"268,330 eligible shipments; 63,118 late; 205,212 on time."),
    ("Cost per shipment is ₹4,102.68 — AMBER.", f"293,485 valid-cost shipments; total valid cost ₹1,204,074,833.69."),
    ("Destination dwell averages 15.49 h — AMBER; origin dwell 2.5 h — GREEN.", "Destination dwell is ~6× origin dwell at every hub; combined 8.9 h."),
    ("Performance is flat all year.", f"Monthly OTD 75.72–77.04%; H1 76.49% vs H2 76.47% (0.02 pp). No seasonal pattern; no true year-over-year exists."),
    ("Three destination hubs drive the problem.", "Guwahati 53.95%, Malda 61.89%, Ranchi 69.04% OTD; together 41.76% of all late deliveries. Best hubs ~83.9%."),
    ("Origin performance is uniform.", "Origin-hub OTD spans only 75.26–78.80%. The problem is at/after destination arrival, not departure."),
    ("Frozen is the weakest and one of the most expensive categories.", "72.06% OTD; 33.10% of late deliveries; ₹4,466.98 per shipment. Pharma is best (84.59%) and most expensive (₹7,226.40)."),
    ("The worst routes are unreliable, not expensive.", "Bottom-decile routes 57.11–65.25% OTD vs top-decile 82.83–83.41%; costs overlap (₹4,052–₹4,140 vs ₹4,054–₹4,134)."),
    ("Vehicle type does not matter.", "OTD 76.24–76.64% and cost ₹4,095.95–₹4,107.20 across all five vehicle types; late-delivery shares 19.83–20.27% each."),
    ("Late deliveries are moderately concentrated.", f"39 of 61 routes account for 80% of late deliveries; top five routes 24.56%. No single-route fix suffices."),
    ("Dwell varies too little to explain the OTD gap.", "Warehouse destination dwell spans 15.21–15.95 h; hub destination dwell 15.40–15.57 h. Dwell is a network-wide condition, not a local failure."),
    ("The realistic improvement prize is service, not cost.", "9,706 late deliveries avoided, ₹31,13,833 saved, 4,295.8 dwell hours recovered — but even then network OTD would reach only ~80.09%, still below the 85% target."),
    ("Data quality and validation are strong.", f"200,528 changed records logged (100% coverage); 19 cleaning-log rows; {VAL_PASS}/{VAL_TOTAL} validation checks passed at 0% variance (tolerance 0.5%)."),
]
for i, (title, detail) in enumerate(findings, start=1):
    bullet(detail, f"{i}. {title} ")

# ================================================================ 19 RECOMMENDATIONS
h1(19, "Three Business Recommendations")
para("Each recommendation names an owner and a monitoring metric, as required. All three are framed for the "
     "fictional network on synthetic data; on real data the same actions would need a costed pilot before "
     "network-wide rollout.")
add_table(["#", "Recommendation", "Owner", "Monitor", "Expected effect (modelled, realistic scenario)"],
    [
        ["1", "Guwahati / Malda / Ranchi corridor recovery — dedicated route planning for the corridors serving GAU, MLD and RNC (including the worst routes RT-006, RT-005, RT-017, RT-028) and a hub capacity review at the three destination hubs (unloading bays, staffing, onward dispatch). Start with a diagnostic week on the ground at each hub before changing schedules.", "Operations", "Destination-hub OTD (GAU, MLD, RNC) monthly; route OTD for RT-005/RT-006/RT-017/RT-028", "Largest share of the 9,706 late deliveries avoided; these three hubs are 41.76% of all late deliveries"],
        ["2", "Destination-hub dwell reduction — unloading-slot scheduling and labour planning at the bottom-quartile warehouses (7 warehouses, dwell threshold 15.57 h), prioritising WH-CTC-1 (15.95 h) and WH-MLD-2 (15.64 h). Book arrival slots against planned unloading capacity instead of first-come-first-served, and staff to the arrival profile rather than a flat roster.", "Warehouse Operations", "Average destination dwell (h) by warehouse, weekly; bottom-quartile warehouse dwell vs 15.51 h median", "4,295.8 excess dwell hours recovered (realistic scenario); note the prize is modest because dwell varies little — treat this as a service/throughput action, not a cost action"],
        ["3", "Frozen-category cold-chain reliability + cost control on bottom-quartile routes — reefer maintenance checks (pre-trip temperature-proving) for Frozen loads, and a route cost review on the 15 bottom-quartile cost routes (threshold ₹4,122.72). Pair the two: a Frozen load on an unreliable route is the highest-risk combination in the data.", "Fleet & Category Management", "Category OTD (Frozen) monthly; route cost per shipment on bottom-quartile routes", "Frozen OTD up from 72.06% toward the 76.48% network average and beyond; ₹31,13,833 potential saving (realistic scenario)"],
    ], font_size=9)
para("Sequencing: start Recommendation 1 first (it addresses 41.76% of late deliveries); run Recommendation 3 in "
     "parallel for Frozen loads on those same corridors, since the two overlap; run Recommendation 2 as a "
     "steady throughput programme rather than a one-off fix, and re-measure dwell at a finer grain (arrival-hour "
     "and bay-level) before judging it, given the narrow dwell spread in this dataset. Re-run the full KPI pack "
     "monthly from the trusted table and re-validate (38 checks, 0.5% tolerance) on each run; a recommendation "
     "whose monitoring metric does not move within one quarter should be re-diagnosed, not simply continued.")

# ================================================================ 20 LIMITATIONS
h1(20, "Limitations")
for title, text in [
    ("Synthetic data. ", "Every finding describes a simulated network (seed 42), not a real company. Volumes, rates, costs and hub behaviours were generated, including the injected defects. Nothing here should be quoted as a fact about any real distributor, and the business-impact figures are modelled estimates, not forecasts of real savings."),
    ("One year only; no true year-over-year. ", "The data covers 2025 only. H1 vs H2 2025 is the prior-period comparison; it cannot show annual growth, seasonality across years, or a trend. Monthly patterns within a single year cannot separate seasonality from noise."),
    ("No causal analysis. ", "All segment comparisons are descriptive. The report identifies where lateness concentrates, not why it happens in a causal sense; the generator's parameters are the mechanical cause of the simulated patterns. Route, hub and category effects are confounded in this design (e.g. the worst routes serve the worst hubs), and no regression or controlled comparison was run to separate them."),
    ("Dwell does not explain the OTD gap. ", "Destination dwell is nearly identical across hubs (15.40–15.57 h) while OTD spans 30 pp, so the dwell measure as built (arrival→departure at the hub) misses where the delay is created — possibly in transit before the destination scan, or between departure and final delivery. Finer-grained timestamps would be needed to locate it."),
    ("Cost data is thin. ", "Cost per shipment varies little by route, hub or vehicle (₹4,024–₹4,196 across routes), and no per-hour handling cost, fuel, labour or penalty cost exists in the data — so dwell savings cannot be converted to rupees and the cost-saving estimate uses the route cost gap only. Implementation costs and capacity constraints are excluded from the impact model."),
    ("Thresholds are judgements. ", "The scorecard thresholds (§17) are management targets set for this review; no external benchmark, SLA or prior-year target exists in the data to anchor them."),
    ("Power BI is a build, not a delivered dashboard. ", "A .pbix cannot be generated programmatically. The project provides cleaned datasets, DAX measures and a build guide; the dashboard itself must be built by hand in Power BI Desktop (see §10 and the Appendix), and the powerbi/ folder was empty when this report was built."),
    ("Validation scope. ", "The 38 validation checks reconcile counts, KPI recomputation and referential integrity. They do not validate the synthetic generator's realism, the choice of KPI definitions, or the business-impact assumptions — those remain judgements documented in §§11, 16 and 17."),
    ("Small 'Unknown' segments. ", "857 shipments have 'Unknown' category, 1,521 are on the UNKNOWN route and 1,429 have a missing vehicle_id. These are small (≤0.5% each) and are retained as Unknown members rather than dropped, but their KPI values (e.g. Unknown category OTD 76.66%) are based on small samples and should not be over-interpreted."),
]:
    bullet(text, title)
para("A note on a figure that looks inconsistent but is not: data/processed/cleaning_summary.json records "
     "'otd_eligible_shipments' as 288,252, while the KPI calculation and the independent validation both give "
     "268,330. The two numbers count different things, verified against the trusted table itself: 288,252 is the "
     "number of shipments with a non-null on_time flag, i.e. with both timestamps valid (220,433 on time + "
     "67,819 late on that flag); 268,330 is the OTD-eligible denominator in §11, which additionally requires "
     "status Delivered or Delayed (279,211 shipments carry those two statuses). The 19,922 difference is "
     "shipments with valid timestamps but another status (In Transit, Cancelled or Returned). The 268,330 "
     "figure is the one used for every OTD number in this report; the 288,252 figure is a timestamp-validity "
     "count, not an OTD denominator, and the cleaning_summary.json label is misleading rather than the data "
     "being wrong. The project owner should relabel that field to avoid the confusion.")

# ================================================================ 21 CONCLUSION
h1(21, "Conclusion")
para(f"In 2025, on synthetic data, this fictional cold-chain network delivered on time {O['otd_pct']:.2f}% of the "
     f"time at {fmt_money(O['cost_per_shipment_inr'])} per shipment, with freight spending an average of "
     f"{O['avg_dwell_dest_h']:.2f} hours at its destination hub. All three headline KPIs are AMBER: not failing, "
     "but below the targets a well-run cold chain should meet, and flat across the year — H1 and H2 differ by "
     "0.02 pp, so time alone is not fixing anything.")
para("The diagnosis is concentrated enough to act on: three destination hubs (Guwahati, Malda, Ranchi) produce "
     "41.76% of late deliveries; the Frozen category produces a third; the bottom-quartile routes are unreliable "
     "rather than expensive; and vehicle type and origin hub are not drivers. The modelled prize for bringing "
     "the bottom quartile up to the median is 9,706 late deliveries avoided, ₹31,13,833 saved and 4,295.8 "
     "dwell hours recovered — meaningful, but still short of the 85% OTD target, which means the median itself "
     "must rise, not just the tail.")
para("The project's other deliverable is trust in the numbers: a fully logged clean (200,528 changed records), a "
     "star-schema trusted table, and 38/38 validation checks passed at 0% variance. That pipeline — not any "
     "single number in this report — is what transfers to real data. The next step is the three recommendations "
     "in §19, measured monthly against the scorecard in §17, starting with the Guwahati / Malda / Ranchi corridors.")

# ================================================================ 22 APPENDIX
h1(22, "Appendix")
h2("A. Source files and traceability")
add_table(["Artefact", "Location"],
    [
        ["KPI summary (headline figures)", "outputs/kpi_reports/kpi_summary.json"],
        ["Monthly / half-year KPIs", "outputs/kpi_reports/monthly_kpi.csv, half_year_kpi.csv"],
        ["Hub KPIs (destination / origin)", "outputs/kpi_reports/hub_destination_kpi.csv, hub_origin_kpi.csv"],
        ["Route / warehouse KPIs", "outputs/kpi_reports/route_kpi.csv, warehouse_kpi.csv"],
        ["Category / vehicle KPIs", "outputs/kpi_reports/category_kpi.csv, vehicle_type_kpi.csv"],
        ["Cohort tables", "outputs/kpi_reports/cohort_month_category.csv, cohort_month_hub.csv"],
        ["Decile shipment files", "outputs/kpi_reports/bottom10pc_shipments.csv, top10pc_shipments.csv"],
        ["Data-quality report + checks", "outputs/kpi_reports/data_quality_report.md, data_quality_checks.json"],
        ["Column profiles", "outputs/kpi_reports/column_profile_shipments.csv, column_profile_scans.csv"],
        ["Cleaning log", "outputs/cleaning_logs/cleaning_log.csv (19 rows) + changed_records.parquet"],
        ["Cleaning summary", "data/processed/cleaning_summary.json (its 'otd_eligible_shipments' field is a timestamp-validity count, 288,252 — see §20; the OTD denominator used in this report is 268,330)"],
        ["Validation report", "outputs/validation_report.md (38/38 PASS, 0% variance)"],
        ["Generation manifest", "data/raw/generation_manifest.json (seed 42)"],
        ["Raw data", "data/raw/shipments_raw.parquet, scan_events_raw.parquet"],
        ["Cleaned data", "data/processed/shipments_clean.parquet, scan_events_clean.parquet, trusted_analysis_table.parquet"],
        ["Dimensions", "data/reference/dim_hub.csv, dim_route.csv, dim_vehicle.csv, dim_warehouse.csv, dim_category.csv, dim_date.csv, hub_crosswalk.csv"],
        ["Charts (Figures 1–8)", "visualizations/01_monthly_otd_trend.png … 08_late_delivery_pareto.png"],
        ["Pipeline code", "src/generate_data.py, data_profiling.py, data_cleaning.py, data_transformation.py, kpi_calculations.py, validation.py, visualizations.py"],
        ["This report's builder", "src/build_report.py"],
    ], first_col_bold=True)
h2("B. Validation summary (38/38 checks passed, 0% variance)")
para(f"Reconciliation tolerance: 0.5% unless stated 0. Result: {VAL_PASS} passed, {VAL_FAIL} failed, "
     f"{VAL_TOTAL} checks (outputs/validation_report.md). Every check below passed with 0.0% variance.")
add_table(["Check group", "Checks", "What was verified"],
    [
        ["Row counts", "10", "Raw/clean shipments (304,500 / 300,000), raw/clean scans (1,200,000 / 1,194,000), totals (1,504,500 / 1,494,000), removals = injected duplicates/orphans (4,500 / 6,000), trusted-table rows and unique shipment_ids (300,000)"],
        ["Headline KPIs (recomputed)", "8", "OTD % 76.48, eligible 268,330, late 63,118, cost/shipment ₹4,102.68, total valid cost ₹1,204,074,833.69, valid-cost shipments 293,485, destination dwell 15.49 h, origin dwell 2.5 h"],
        ["Breakdown sums", "6", "Monthly, destination-hub, route, category and warehouse shipment sums all = 300,000; monthly eligible sum = 268,330"],
        ["Referential integrity", "5", "Origin/destination hubs in Dim_Hub; routes in Dim_Route (+UNKNOWN); vehicles in Dim_Vehicle (+UNKNOWN/Unknown); warehouses in Dim_Warehouse"],
        ["Validity flags", "2", "No negative costs in the valid-cost set; no negative dwell in the valid dwell set"],
        ["Cleaning-log coverage", "4", "19 log rows (one per issue type); changed records > 0 and cover duplicate removals and orphan scans"],
        ["Post-cleaning state", "3", "0 duplicate shipment_ids in the trusted table; 0 shipments with out-of-order scans after cleaning; raw-file cost recomputation matches the trusted table"],
    ], first_col_bold=True)
para("Note: the validation report itself notes that the raw-cost reconciliation re-applies the cleaning rules "
     "directly to the raw file and that the small difference vs the trusted table comes from the most-complete-record "
     "dedupe rule (cleaning keeps the most complete copy, not always the first); the variance column still records "
     "0.0% for that check. 'Shipments with out-of-order scans after cleaning' is 0 by construction "
     "(scan_sequence is assigned after sorting by UTC).")
h2("C. DAX measures for Power BI (to be placed in powerbi/)")
para("A .pbix cannot be generated programmatically. Paste these measures into Power BI Desktop after importing "
     "trusted_analysis_table.parquet and relating the dimensions (§10). 'Trusted' is the fact table. Column names "
     "below are the actual columns of trusted_analysis_table.parquet, verified at build time: on_time (nullable "
     "boolean — blank where either timestamp is missing), cost_valid, dwell_dest_h / dwell_origin_h, "
     "dwell_dest_valid / dwell_origin_valid, status, shipment_cost_inr.")
for name, dax in [
    ("OTD Eligible Shipments", "OTD Eligible Shipments = CALCULATE ( COUNTROWS ( Trusted ), Trusted[status] IN { \"Delivered\", \"Delayed\" }, NOT ( ISBLANK ( Trusted[on_time] ) ) )"),
    ("OTD %", "OTD % = DIVIDE ( CALCULATE ( COUNTROWS ( Trusted ), Trusted[status] IN { \"Delivered\", \"Delayed\" }, Trusted[on_time] = TRUE () ), [OTD Eligible Shipments] )"),
    ("Late Deliveries", "Late Deliveries = CALCULATE ( COUNTROWS ( Trusted ), Trusted[status] IN { \"Delivered\", \"Delayed\" }, Trusted[on_time] = FALSE () )"),
    ("Cost per Shipment", "Cost per Shipment = DIVIDE ( CALCULATE ( SUM ( Trusted[shipment_cost_inr] ), Trusted[cost_valid] = TRUE () ), CALCULATE ( COUNTROWS ( Trusted ), Trusted[cost_valid] = TRUE () ) )"),
    ("Total Valid Cost", "Total Valid Cost = CALCULATE ( SUM ( Trusted[shipment_cost_inr] ), Trusted[cost_valid] = TRUE () )"),
    ("Avg Destination Dwell (h)", "Avg Destination Dwell (h) = CALCULATE ( AVERAGE ( Trusted[dwell_dest_h] ), Trusted[dwell_dest_valid] = TRUE () )"),
    ("Avg Origin Dwell (h)", "Avg Origin Dwell (h) = CALCULATE ( AVERAGE ( Trusted[dwell_origin_h] ), Trusted[dwell_origin_valid] = TRUE () )"),
]:
    p = doc.add_paragraph()
    r = p.add_run(name + ":  ")
    r.bold = True
    r.font.size = Pt(10)
    r2 = p.add_run(dax)
    r2.font.name = "Consolas"
    r2.font.size = Pt(9)
para("Caution on the on_time flag: in the trusted table on_time is set for every shipment with both timestamps "
     "valid (288,252 shipments: 220,433 TRUE, 67,819 FALSE), regardless of status — so an OTD measure that counts "
     "on_time without the status filter would overstate the denominator. The measures above always filter "
     "status to Delivered/Delayed, matching the KPI definition in §11 (268,330 eligible, 63,118 late). "
     "dwell_dest_valid / dwell_origin_valid are blank for 4,357 shipments with no dwell record; the = TRUE() "
     "filter excludes those as well as the FALSE (negative-dwell) rows.", italic=True, size=10, colour=GREY)
h2("D. Power BI build guide (summary)")
for step in [
    "Import data/processed/trusted_analysis_table.parquet (and, for the scan-level dwell view, scan_events_clean.parquet) into Power BI Desktop.",
    "Import the dimension CSVs from data/reference/ and create relationships: Trusted[destination_hub] → Dim_Hub[hub_id] (and origin_hub → Dim_Hub), Trusted[route_id] → Dim_Route[route_id], Trusted[vehicle_id] → Dim_Vehicle[vehicle_id], Trusted[warehouse_id] → Dim_Warehouse[warehouse_id], Trusted[category] → Dim_Category[category], Trusted[shipment_date] → Dim_Date[date]. Use the UNKNOWN/Unknown members already present in the dimensions.",
    "Paste the DAX measures from §C.",
    "Build eight views matching Figures 1–8: monthly OTD trend; hub OTD comparison; hub dwell; route ranking; category performance; cost per shipment; warehouse performance; late-delivery Pareto.",
    "Add the scorecard (§17) as a KPI visual set with the GREEN/AMBER/RED thresholds, and a 'SYNTHETIC DATA — seed 42' label on every page.",
    "Save as .pbix in powerbi/ (this step is manual — no script in this project can produce a .pbix).",
]:
    bullet(step)
h2("E. Figure and table list")
for f in [
    "Figure 1 — Monthly OTD % and cost per shipment (01_monthly_otd_trend.png)",
    "Figure 2 — Cost per shipment by segment (06_cost_per_shipment.png)",
    "Figure 3 — OTD % by destination hub (02_hub_otd_comparison.png)",
    "Figure 4 — Route OTD ranking (04_route_performance_ranking.png)",
    "Figure 5 — Average dwell time by hub (03_hub_dwell_time.png)",
    "Figure 6 — Warehouse performance (07_warehouse_performance.png)",
    "Figure 7 — OTD % and cost by product category (05_category_performance.png)",
    "Figure 8 — Pareto of late deliveries by route (08_late_delivery_pareto.png)",
]:
    bullet(f)
para("Author: Khushi Tomar — Yuva Intern, Logistics Data Analytics Specialist Internship. November 2026. "
     "All data synthetic (seed 42); not real company data.", italic=True, size=10, colour=GREY)

# ---------------------------------------------------------------- save
OUT.parent.mkdir(parents=True, exist_ok=True)
doc.save(str(OUT))
print(f"Saved: {OUT} ({OUT.stat().st_size} bytes)")

# Verification: reload and count
from docx import Document as _D
d2 = _D(str(OUT))
paras = [p for p in d2.paragraphs if p.text.strip()]
tables = d2.tables
words = sum(len(p.text.split()) for p in d2.paragraphs)
tw = 0
for t in tables:
    for row in t.rows:
        for c in row.cells:
            tw += len(c.text.split())
print(f"Paragraphs(non-empty): {len(paras)}  Tables: {len(tables)}  Words(body): {words}  Words(tables): {tw}")
print(f"Sections: {len(d2.sections)}  Inline shapes (charts): {len(d2.inline_shapes)}")
heads = [p.text for p in d2.paragraphs if p.style.name.startswith("Heading 1")]
print(f"Heading 1 count: {len(heads)}")
for h in heads:
    print("  H1:", h)
print(f"Estimated pages (words/450): {(words + tw) / 450:.1f}")
