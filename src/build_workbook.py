#!/usr/bin/env python3
"""Build Logistics_KPI_Analysis_Workbook.xlsx (11 sheets) from verified outputs.

All numbers come from outputs/kpi_reports/*.json|*.csv, outputs/cleaning_logs/,
outputs/validation_report.md and data/raw/generation_manifest.json.
ALL data is SYNTHETIC (seed 42) — labelled as such throughout.
Run with: /home/hatch/workspace/logistics-kpi/.venv/bin/python src/build_workbook.py
"""
import csv
import json
import os
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path("/home/hatch/workspace/logistics-kpi")
OUT = ROOT / "excel" / "Logistics_KPI_Analysis_Workbook.xlsx"
KPI = ROOT / "outputs" / "kpi_reports"
CLEAN = ROOT / "outputs" / "cleaning_logs"
REF = ROOT / "data" / "reference"
PROC = ROOT / "data" / "processed"

# ---------------------------------------------------------------- styles
NAVY = "1F4E78"
NAVY_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="595959")
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
HEADER_FILL = NAVY_FILL
BODY_FONT = Font(name="Calibri", size=10)
BOLD = Font(name="Calibri", size=10, bold=True)
SMALL = Font(name="Calibri", size=9, color="595959")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
GREEN_FONT = Font(name="Calibri", size=10, bold=True, color="006100")
AMBER_FONT = Font(name="Calibri", size=10, bold=True, color="9C6500")
RED_FONT = Font(name="Calibri", size=10, bold=True, color="9C0006")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
WRAP_C = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(vertical="top", wrap_text=True)

FMT_INT = "#,##0"
FMT_PCT = "0.00%"          # value stored as fraction
FMT_PP = "0.00"
FMT_HRS = "0.00"
FMT_INR = "\u20b9#,##,##0.00"   # Indian grouping
FMT_INR0 = "\u20b9#,##,##0"


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def num(v):
    """Convert a CSV string to int/float where possible, else return as-is."""
    if v is None or v == "":
        return None
    try:
        if "." in str(v) or "e" in str(v).lower():
            return float(v)
        return int(v)
    except (TypeError, ValueError):
        return v


def title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = SUBTITLE_FONT
    c.alignment = WRAP
    ws.row_dimensions[2].height = 30
    return 4  # first free row after a blank row 3


def header_row(ws, row, headers, widths=None):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = WRAP_C
        c.border = BORDER
    ws.row_dimensions[row].height = 30
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w


def body_cell(ws, row, col, value, fmt=None, bold=False, fill=None, font=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font or (BOLD if bold else BODY_FONT)
    c.border = BORDER
    c.alignment = WRAP
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def finish_sheet(ws, header_row_idx, ncols, freeze=True):
    ws.auto_filter.ref = f"A{header_row_idx}:{get_column_letter(ncols)}{ws.max_row}"
    if freeze:
        ws.freeze_panes = ws.cell(row=header_row_idx + 1, column=1)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


# KPI column display spec: (csv key, header, format)
KPI_COLS = [
    ("shipments", "Shipments", FMT_INT),
    ("otd_eligible", "OTD-eligible shipments", FMT_INT),
    ("otd_pct", "OTD %", FMT_PCT),
    ("late_deliveries", "Late deliveries", FMT_INT),
    ("cost_per_shipment_inr", "Cost / shipment (\u20b9)", FMT_INR),
    ("total_valid_cost_inr", "Total valid cost (\u20b9)", FMT_INR),
    ("valid_cost_shipments", "Valid-cost shipments", FMT_INT),
    ("avg_dwell_dest_h", "Avg dwell \u2014 destination (h)", FMT_HRS),
    ("avg_dwell_origin_h", "Avg dwell \u2014 origin (h)", FMT_HRS),
    ("avg_dwell_combined_h", "Avg dwell \u2014 combined (h)", FMT_HRS),
]


def write_kpi_table(ws, start_row, rows, label_col, label_header, title=None):
    """Write a KPI table; returns next free row. otd_pct stored as fraction."""
    if title:
        ws.merge_cells(start_row=start_row, start_column=1,
                       end_row=start_row, end_column=len(KPI_COLS) + 1)
        c = ws.cell(row=start_row, column=1, value=title)
        c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        start_row += 1
    headers = [label_header] + [h for _, h, _ in KPI_COLS]
    widths = [18] + [16, 16, 10, 14, 16, 18, 16, 15, 14, 15]
    header_row(ws, start_row, headers, widths)
    r = start_row + 1
    for row in rows:
        body_cell(ws, r, 1, row[label_col], bold=True)
        for j, (key, _, fmt) in enumerate(KPI_COLS, start=2):
            v = num(row.get(key))
            if key == "otd_pct" and isinstance(v, (int, float)):
                v = v / 100.0
            body_cell(ws, r, j, v, fmt=fmt)
        r += 1
    ws.auto_filter.ref = (f"A{start_row}:{get_column_letter(len(headers))}{r - 1}")
    return r + 1


# ---------------------------------------------------------------- data
summary = load_json(KPI / "kpi_summary.json")
manifest = load_json(ROOT / "data" / "raw" / "generation_manifest.json")
clean_sum = load_json(PROC / "cleaning_summary.json")
dq = load_json(KPI / "data_quality_checks.json")
overall = summary["overall"]
bi = summary["business_impact"]
scen = bi["scenarios"]

monthly = read_csv(KPI / "monthly_kpi.csv")
hub_dest = read_csv(KPI / "hub_destination_kpi.csv")
hub_orig = read_csv(KPI / "hub_origin_kpi.csv")
route = read_csv(KPI / "route_kpi.csv")
warehouse = read_csv(KPI / "warehouse_kpi.csv")
category = read_csv(KPI / "category_kpi.csv")
vehicle = read_csv(KPI / "vehicle_type_kpi.csv")
half_year = read_csv(KPI / "half_year_kpi.csv")
prof_ship = read_csv(KPI / "column_profile_shipments.csv")
prof_scan = read_csv(KPI / "column_profile_scans.csv")
clean_log = read_csv(CLEAN / "cleaning_log.csv")
hubs = read_csv(REF / "dim_hub.csv")
crosswalk = read_csv(REF / "hub_crosswalk.csv")
sample_rows = read_csv(PROC / "trusted_analysis_table_sample.csv")[:20]

wb = Workbook()

# ============================================================ Sheet 1
ws = wb.active
ws.title = "Raw Data - Source Inventory"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
r = title_block(
    ws,
    "Logistics KPI Performance Review \u2014 Source Inventory (SYNTHETIC DATA)",
    "Fictional Kolkata-based cold-chain distributor, 12 hubs, 2025 only. ALL data is SYNTHETIC "
    "(seed 42), generated for the Yuva Intern Logistics Data Analytics Specialist Internship "
    "evaluation. It is NOT real company data and must never be presented as such.",
    5)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value=(
    "Excel cannot hold 1.5M rows: this sheet is an INVENTORY only. The full data lives in "
    "data/raw and data/processed as Parquet files (see table below). A 20-row sample of the "
    "trusted analysis table follows the inventory."))
c.font = Font(name="Calibri", size=10, bold=True, color="9C6500")
c.alignment = WRAP
ws.row_dimensions[r].height = 30
r += 2
inv_headers = ["File", "Records", "Rows (clean)", "Format / location", "Description"]
inv = [
    ("shipments_raw.parquet", 304500, 300000, "Parquet \u2014 data/raw",
     "Raw shipment records as generated, including 4,500 injected duplicate rows."),
    ("scan_events_raw.parquet", 1200000, 1194000, "Parquet \u2014 data/raw",
     "Raw warehouse scan events (4 per shipment), including 6,000 orphan scans."),
    ("shipments_clean.parquet", 300000, 300000, "Parquet \u2014 data/processed",
     "Shipments after cleaning (duplicates removed, most complete copy kept)."),
    ("scan_events_clean.parquet", 1194000, 1194000, "Parquet \u2014 data/processed",
     "Scan events after cleaning (orphan scans removed, re-sequenced by UTC)."),
    ("trusted_analysis_table.parquet", 300000, 300000, "Parquet \u2014 data/processed",
     "One row per clean shipment with derived KPI fields (OTD, cost validity, dwell). "
     "This is the single source for every KPI in this workbook."),
    ("trusted_analysis_table_sample.csv", 10000, 10000, "CSV \u2014 data/processed",
     "10,000-row sample of the trusted analysis table (first 20 rows reproduced below)."),
    ("dim_hub.csv", 12, 12, "CSV \u2014 data/reference", "12 hubs (KOL, HWH, SIL, DGP, ASN, MLD, GAU, PAT, RNC, JSR, BBS, CTC)."),
    ("dim_route.csv", 60, 60, "CSV \u2014 data/reference", "60 routes (+ UNKNOWN member for invalid refs)."),
    ("dim_vehicle.csv", 180, 180, "CSV \u2014 data/reference", "180 vehicles across 5 vehicle types."),
    ("dim_warehouse.csv", 24, 24, "CSV \u2014 data/reference", "24 warehouses, 2 per hub."),
    ("dim_category.csv", 5, 5, "CSV \u2014 data/reference", "5 product categories: Frozen, Chilled, Ambient, Pharma, Dairy."),
    ("dim_date.csv", 365, 365, "CSV \u2014 data/reference", "Calendar 2025-01-01 to 2025-12-31."),
    ("hub_crosswalk.csv", 60, 60, "CSV \u2014 data/reference",
     "60 messy raw hub spellings mapped to the 12 canonical hub ids."),
]
header_row(ws, r, inv_headers, [34, 12, 12, 24, 70])
inv_head = r
r += 1
for row in inv:
    for j, v in enumerate(row, start=1):
        fmt = FMT_INT if j in (2, 3) else None
        body_cell(ws, r, j, v, fmt=fmt, bold=(j == 1))
    r += 1
body_cell(ws, r, 1, "TOTAL (raw records)", bold=True)
body_cell(ws, r, 2, 1504500, fmt=FMT_INT, bold=True)
body_cell(ws, r, 3, 1494000, fmt=FMT_INT, bold=True)
body_cell(ws, r, 4, "1,504,500 raw \u2192 1,494,000 clean", bold=True)
body_cell(ws, r, 5, "10,500 records removed in cleaning: 4,500 duplicate shipments + 6,000 orphan scans.", bold=True)
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="Sample \u2014 first 20 rows of trusted_analysis_table_sample.csv (SYNTHETIC)")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
sample_cols = list(sample_rows[0].keys())
sample_widths = [13, 12, 8, 10, 12, 9, 9, 15, 13, 9, 9, 22, 22, 13, 9, 8, 9, 11, 11, 10, 12, 12]
# sample table is wider than the 5-col inventory; write from col 1 with its own widths
header_row(ws, r, sample_cols)
for j, w in enumerate(sample_widths, start=1):
    ws.column_dimensions[get_column_letter(j)].width = w
sample_head = r
r += 1
NUM_SAMPLE = {"shipment_cost_inr", "late_hours", "dwell_origin_h", "dwell_dest_h", "dwell_avg_h"}
for row in sample_rows:
    for j, col in enumerate(sample_cols, start=1):
        v = row[col]
        if col in NUM_SAMPLE:
            v = num(v)
            fmt = FMT_INR if col == "shipment_cost_inr" else FMT_HRS
        elif col in ("cost_valid", "on_time", "dwell_origin_valid", "dwell_dest_valid"):
            v = (v == "True")
            fmt = None
        else:
            fmt = None
        body_cell(ws, r, j, v, fmt=fmt)
    r += 1
ws.auto_filter.ref = f"A{sample_head}:{get_column_letter(len(sample_cols))}{r - 1}"
ws.freeze_panes = f"A{sample_head + 1}"
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ============================================================ Sheet 2
ws = wb.create_sheet("Data Dictionary")
r = title_block(
    ws,
    "Data Dictionary \u2014 Shipments, Scans, Derived Fields & Dimensions (SYNTHETIC DATA)",
    "Column, table, type and quality notes. Types and missing/unique counts are taken from "
    "column_profile_shipments.csv and column_profile_scans.csv (raw files). "
    "Hub crosswalk: hub_crosswalk.csv maps 60 messy raw hub spellings to the 12 canonical hub ids "
    "in dim_hub.csv (e.g. 'KOL-01', 'kol_1', 'Kolkata Hub 1', 'KOLKATA' \u2192 KOL). "
    "After mapping, every origin/destination hub is in Dim_Hub.",
    5)
headers = ["Column", "Table", "Type", "Description", "Quality notes (raw file)"]
header_row(ws, r, headers, [24, 26, 12, 62, 62])
dd_head = r
r += 1
# quality notes keyed by (table, column) from the profiles
prof = {(p["table"], p["column"]): p for p in prof_ship + prof_scan}


def qnote(table, column):
    p = prof.get((table, column))
    if not p:
        return ""
    miss = p["missing_values"]
    mp = p["missing_pct"]
    uniq = p["unique_values"]
    dup = p["duplicate_values"]
    bits = [f"missing {miss} ({mp}%)", f"unique {uniq}"]
    if dup not in ("", None):
        bits.append(f"duplicate values {dup}")
    if p.get("sample_values"):
        bits.append("e.g. " + p["sample_values"])
    return "; ".join(bits)


DD = [
    # shipments_raw
    ("shipment_id", "shipments_raw", "str", "Unique shipment identifier (SHP-000001 \u2026 SHP-300000).", qnote("shipments_raw", "shipment_id") + " 4,500 extra copies removed in cleaning; unique after cleaning."),
    ("shipment_date", "shipments_raw", "str", "Shipment (dispatch) date, 2025-01-01 to 2025-12-31.", qnote("shipments_raw", "shipment_date")),
    ("origin_hub", "shipments_raw", "str", "Origin hub, canonical id after crosswalk mapping (12 hubs).", qnote("shipments_raw", "origin_hub") + " 72 raw variants incl. messy spellings; 91,486 rows standardised."),
    ("destination_hub", "shipments_raw", "str", "Destination hub, canonical id after crosswalk mapping (12 hubs).", qnote("shipments_raw", "destination_hub") + " 72 raw variants; 92,027 rows standardised."),
    ("route_id", "shipments_raw", "str", "Route id (RT-001 \u2026 RT-060). Invalid refs set to UNKNOWN.", qnote("shipments_raw", "route_id") + " 1,548 invalid refs (0.5%) set to UNKNOWN."),
    ("vehicle_id", "shipments_raw", "str", "Vehicle id (VEH-001 \u2026 VEH-180). Invalid refs set to UNKNOWN.", qnote("shipments_raw", "vehicle_id") + " 1,557 invalid refs; 1,455 missing kept as 'Unknown'."),
    ("vehicle_type", "shipments_raw", "str", "Vehicle type: Reefer Van, Reefer Truck, Insulated Truck, Refrigerated Container, Cold Storage Van.", qnote("shipments_raw", "vehicle_type")),
    ("warehouse_id", "shipments_raw", "str", "Destination warehouse (WH-<hub>-1/2, 24 warehouses).", qnote("shipments_raw", "warehouse_id")),
    ("category", "shipments_raw", "str", "Product category: Frozen, Chilled, Ambient, Pharma, Dairy.", qnote("shipments_raw", "category") + " 873 missing kept as 'Unknown'."),
    ("promised_delivery_ts", "shipments_raw", "str", "Promised delivery timestamp (mixed formats/timezones in raw).", qnote("shipments_raw", "promised_delivery_ts") + " 2,407 missing; 3,081 invalid in data-quality checks; unparseable set to NaT."),
    ("actual_delivery_ts", "shipments_raw", "str", "Actual delivery timestamp (mixed formats/timezones in raw).", qnote("shipments_raw", "actual_delivery_ts") + " 3,601 missing; 2,980 invalid in data-quality checks; unparseable set to NaT."),
    ("shipment_cost_inr", "shipments_raw", "float64", "Shipment cost in INR.", qnote("shipments_raw", "shipment_cost_inr") + " 2,954 missing; 2,474 negative and 1,172 zero costs flagged cost_valid=False."),
    ("status", "shipments_raw", "str", "Shipment status: Delivered, Delayed, Returned, In Transit, Cancelled (5 values). OTD uses Delivered/Delayed only.", qnote("shipments_raw", "status")),
    # scans_raw
    ("scan_event_id", "scan_events_raw", "str", "Unique scan event identifier (SCN-0000001 \u2026).", qnote("scan_events_raw", "scan_event_id")),
    ("shipment_id", "scan_events_raw", "str", "Shipment the scan belongs to (foreign key to shipments).", qnote("scan_events_raw", "shipment_id") + " 6,000 orphan scans (no matching shipment) removed."),
    ("hub_id", "scan_events_raw", "str", "Hub where the scan occurred (72 raw variants; canonicalised).", qnote("scan_events_raw", "hub_id")),
    ("warehouse_id", "scan_events_raw", "str", "Warehouse where the scan occurred (24 warehouses).", qnote("scan_events_raw", "warehouse_id")),
    ("scan_type", "scan_events_raw", "str", "Scan type: Loading, Departure, Arrival, Delivered (4 per shipment).", qnote("scan_events_raw", "scan_type")),
    ("scan_ts", "scan_events_raw", "str", "Scan timestamp (mixed formats/timezones in raw); converted to UTC.", qnote("scan_events_raw", "scan_ts") + " 0 invalid timestamps; 29,287 events out of order in raw, re-sequenced."),
    # trusted derived
    ("month", "trusted_analysis_table (derived)", "str", "Shipment month, YYYY-MM (2025-01 \u2026 2025-12).", "Derived from shipment_date."),
    ("half", "trusted_analysis_table (derived)", "str", "Half-year: H1 2025 (Jan\u2013Jun) or H2 2025 (Jul\u2013Dec). Prior-period comparison; no true YoY exists.", "Derived from shipment_date."),
    ("promised_ts_utc", "trusted_analysis_table (derived)", "datetime (UTC)", "Promised delivery timestamp normalised to UTC; NaT if missing/unparseable.", "2,323 unparseable + 5,405 missing (cleaning log; counts differ from raw-profile checks because dedupe runs first)."),
    ("actual_ts_utc", "trusted_analysis_table (derived)", "datetime (UTC)", "Actual delivery timestamp normalised to UTC; NaT if missing/unparseable.", "2,219 unparseable + 6,479 missing (cleaning log)."),
    ("cost_valid", "trusted_analysis_table (derived)", "bool", "True when shipment_cost_inr is present and > 0.", "293,485 valid-cost shipments of 300,000."),
    ("on_time", "trusted_analysis_table (derived)", "bool", "True when actual_ts_utc <= promised_ts_utc (among OTD-eligible).", "OTD = on-time / OTD-eligible = 76.48%."),
    ("late_hours", "trusted_analysis_table (derived)", "float", "Actual minus promised delivery time, in hours (negative = early).", "Bottom 10% by lateness average +12.88 h late; top 10% by earliness average \u221212.09 h."),
    ("otd_eligible", "trusted_analysis_table (derived)", "bool", "True when status is Delivered/Delayed AND both timestamps are valid.", "268,330 eligible of 300,000. Note: cleaning_summary.json counts 288,252 shipments with valid timestamps before the status filter; the OTD denominator after the status filter is 268,330."),
    ("is_late / is_ontime", "trusted_analysis_table (derived)", "bool", "Late / on-time flags among OTD-eligible shipments.", "63,118 late deliveries."),
    ("dwell_origin_h", "trusted_analysis_table (derived)", "float", "Hours between arrival and departure scans at the origin hub.", "Mean 2.5 h; 39 shipments with negative origin dwell flagged invalid."),
    ("dwell_dest_h", "trusted_analysis_table (derived)", "float", "Hours between arrival and delivery scans at the destination hub.", "Mean 15.49 h; 2,905 shipments with negative destination dwell flagged invalid."),
    ("dwell_avg_h / dwell_combined", "trusted_analysis_table (derived)", "float", "Mean of origin and destination dwell (valid legs only).", "Mean 8.9 h."),
    ("dwell_origin_valid / dwell_dest_valid", "trusted_analysis_table (derived)", "bool", "True when the dwell leg is non-negative and computable.", "Dwell KPIs exclude invalid legs."),
    # dimensions
    ("hub_id / hub_name / state / region / is_cold_chain", "dim_hub", "str", "12 hubs: KOL Kolkata, HWH Howrah, SIL Siliguri, DGP Durgapur, ASN Asansol, MLD Malda (West Bengal); GAU Guwahati (Assam); PAT Patna (Bihar); RNC Ranchi, JSR Jamshedpur (Jharkhand); BBS Bhubaneswar, CTC Cuttack (Odisha). All East India, all cold-chain.", "12 rows; every origin/destination hub in the clean data is in this table."),
    ("route_id / origin_hub / destination_hub / distance_km / planned_transit_hours", "dim_route", "str / numeric", "60 routes RT-001 \u2026 RT-060 with origin, destination, distance and planned transit hours; plus an UNKNOWN member.", "61 rows incl. UNKNOWN."),
    ("vehicle_id / vehicle_type / capacity_kg / home_hub", "dim_vehicle", "str / numeric", "180 vehicles across the 5 vehicle types, with capacity and home hub.", "180 rows."),
    ("warehouse_id / hub_id / warehouse_name / capacity_pallets", "dim_warehouse", "str / numeric", "24 warehouses, 2 per hub (e.g. WH-KOL-1, WH-KOL-2).", "24 rows."),
    ("category / base_cost_inr", "dim_category", "str / numeric", "5 categories with base costs: Frozen \u20b94,200; Chilled \u20b93,600; Ambient \u20b92,400; Pharma \u20b96,800; Dairy \u20b93,100.", "5 rows."),
    ("date / month / quarter / year / day_of_week", "dim_date", "date", "Calendar for 2025 (365 days).", "365 rows."),
]
for row in DD:
    for j, v in enumerate(row, start=1):
        body_cell(ws, r, j, v, bold=(j == 1))
    ws.row_dimensions[r].height = 45
    r += 1
finish_sheet(ws, dd_head, 5)

# ============================================================ Sheet 3
ws = wb.create_sheet("Column Profile")
r = title_block(
    ws,
    "Column Profile \u2014 Raw Shipments & Raw Scan Events (SYNTHETIC DATA)",
    "Full profiles from column_profile_shipments.csv (13 columns, 304,500 rows) and "
    "column_profile_scans.csv (6 columns, 1,200,000 rows). Missing % and unique counts are "
    "measured on the RAW files, before cleaning.",
    9)
headers = ["Table", "Column", "Data type", "Total records", "Missing values",
           "Missing %", "Unique values", "Duplicate values", "Sample values"]
header_row(ws, r, headers, [18, 22, 11, 13, 13, 10, 13, 14, 55])
prof_head = r
r += 1
for p in prof_ship + prof_scan:
    body_cell(ws, r, 1, p["table"], bold=True)
    body_cell(ws, r, 2, p["column"], bold=True)
    body_cell(ws, r, 3, p["dtype"])
    body_cell(ws, r, 4, num(p["total_records"]), fmt=FMT_INT)
    body_cell(ws, r, 5, num(p["missing_values"]), fmt=FMT_INT)
    body_cell(ws, r, 6, num(p["missing_pct"]) / 100.0 if p["missing_pct"] not in ("", None) else None, fmt=FMT_PCT)
    body_cell(ws, r, 7, num(p["unique_values"]), fmt=FMT_INT)
    dv = p["duplicate_values"]
    body_cell(ws, r, 8, num(dv) if dv not in ("", None) else None, fmt=FMT_INT)
    body_cell(ws, r, 9, p["sample_values"])
    r += 1
finish_sheet(ws, prof_head, 9)

# ============================================================ Sheet 4
ws = wb.create_sheet("Cleaning Log")
r = title_block(
    ws,
    "Cleaning Log \u2014 Every Issue Found and How It Was Handled (SYNTHETIC DATA)",
    "Full cleaning log (19 issue types) from outputs/cleaning_logs/cleaning_log.csv. "
    "200,528 changed records were logged (100% of changed records). "
    "Removed: 4,500 duplicate shipment rows + 6,000 orphan scans = 10,500 records "
    "(1,504,500 raw \u2192 1,494,000 clean). All other issues were flagged or standardised, "
    "not deleted, so no shipment was dropped for a missing value alone.",
    5)
headers = ["Issue", "Records Affected", "Cleaning Action", "Reason", "Validation"]
header_row(ws, r, headers, [38, 16, 52, 48, 48])
cl_head = r
r += 1
for row in clean_log:
    body_cell(ws, r, 1, row["Issue"], bold=True)
    ra = row["Records Affected"]
    body_cell(ws, r, 2, num(ra) if ra not in ("see profiling (re-sequenced)",) else ra,
              fmt=FMT_INT if ra not in ("see profiling (re-sequenced)",) else None)
    body_cell(ws, r, 3, row["Cleaning Action"])
    body_cell(ws, r, 4, row["Reason"])
    body_cell(ws, r, 5, row["Validation"])
    ws.row_dimensions[r].height = 45
    r += 1
finish_sheet(ws, cl_head, 5)

# ============================================================ Sheet 5
ws = wb.create_sheet("Trusted Analysis Summary")
r = title_block(
    ws,
    "Trusted Analysis Summary \u2014 Headline KPIs (SYNTHETIC DATA, 2025 only)",
    "One row per clean shipment (300,000) in the trusted analysis table. "
    "All figures recomputed independently and validated: 38/38 checks PASS "
    "(outputs/validation_report.md). No true year-over-year exists \u2014 data covers 2025 only; "
    "H1 vs H2 2025 is the prior-period comparison.",
    4)
# --- overall KPI table
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="Overall KPIs \u2014 2025 (full year)")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
headers = ["KPI", "Value", "Denominator / basis", "Definition"]
header_row(ws, r, headers, [30, 20, 26, 62])
sum_head = r
r += 1
otd_frac = overall["otd_pct"] / 100.0
sum_rows = [
    ("On-time delivery (OTD)", otd_frac, "268,330 eligible of 300,000",
     "On-time / OTD-eligible. Eligible = status Delivered/Delayed with valid promised and actual timestamps.", FMT_PCT),
    ("Late deliveries", 63118, "268,330 eligible", "OTD-eligible shipments delivered after the promised time.", FMT_INT),
    ("Cost per shipment", 4102.68, "293,485 valid-cost shipments",
     "Total valid cost / valid-cost shipments. Cost valid = present and > 0.", FMT_INR),
    ("Total valid cost", 1204074833.69, "293,485 valid-cost shipments", "Sum of shipment_cost_inr over valid-cost shipments.", FMT_INR),
    ("Avg dwell \u2014 destination", 15.49, "valid destination-dwell legs",
     "Mean hours between arrival and delivery scans at the destination hub.", FMT_HRS),
    ("Avg dwell \u2014 origin", 2.5, "valid origin-dwell legs",
     "Mean hours between arrival and departure scans at the origin hub.", FMT_HRS),
    ("Avg dwell \u2014 combined", 8.9, "valid dwell legs",
     "Mean of origin and destination dwell (valid legs only).", FMT_HRS),
]
for kpi, val, basis, definition, fmt in sum_rows:
    body_cell(ws, r, 1, kpi, bold=True)
    body_cell(ws, r, 2, val, fmt=fmt, bold=True)
    body_cell(ws, r, 3, basis)
    body_cell(ws, r, 4, definition)
    r += 1
r += 1
# --- row counts
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="Row counts \u2014 raw vs clean")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
headers = ["Dataset", "Raw records", "Clean records", "Removed / note"]
header_row(ws, r, headers, [30, 20, 20, 62])
r += 1
for ds, raw, clean, note in [
    ("Shipments", 304500, 300000, "4,500 duplicate rows removed (most complete copy kept)."),
    ("Scan events", 1200000, 1194000, "6,000 orphan scans removed (no matching shipment)."),
    ("TOTAL", 1504500, 1494000, "10,500 records removed (0.70% of raw)."),
    ("Trusted analysis table", 300000, 300000, "One row per clean shipment; 300,000 unique shipment_ids; 0 duplicates remaining."),
]:
    body_cell(ws, r, 1, ds, bold=True)
    body_cell(ws, r, 2, raw, fmt=FMT_INT, bold=(ds == "TOTAL"))
    body_cell(ws, r, 3, clean, fmt=FMT_INT, bold=(ds == "TOTAL"))
    body_cell(ws, r, 4, note)
    r += 1
r += 1
# --- eligibility denominators
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="Eligibility denominators (why each KPI uses a different base)")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
headers = ["KPI", "Denominator", "Excluded", "Note"]
header_row(ws, r, headers, [30, 20, 26, 62])
r += 1
for kpi, den, excl, note in [
    ("OTD", "268,330", "31,670 shipments",
     "Excluded: non-Delivered/Delayed statuses and missing/unparseable promised or actual timestamps. "
     "cleaning_summary.json counts 288,252 shipments with valid timestamps before the status filter."),
    ("Cost per shipment", "293,485", "6,515 shipments",
     "Excluded: 2,954 missing + 2,474 negative + 1,172 zero costs (raw-profile counts; cleaning-log counts differ slightly because dedupe runs first)."),
    ("Dwell (destination)", "valid legs only", "2,905 negative-dwell shipments flagged",
     "Negative dwell legs excluded; 39 shipments also had negative origin dwell."),
]:
    body_cell(ws, r, 1, kpi, bold=True)
    body_cell(ws, r, 2, den)
    body_cell(ws, r, 3, excl)
    body_cell(ws, r, 4, note)
    ws.row_dimensions[r].height = 45
    r += 1
r += 1
# --- half year
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value="Prior-period comparison \u2014 H1 vs H2 2025 (NOT year-over-year; data covers 2025 only)")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
r = write_kpi_table(ws, r, half_year, "half", "Half-year")
# fix widths from write_kpi_table (11 cols) -> sheet only uses 4 cols here; widths already set, fine
r += 0
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
c = ws.cell(row=r, column=1, value=(
    "Reading: performance is essentially flat across halves \u2014 OTD 76.49% (H1) vs 76.47% (H2), "
    "cost \u20b94,102.40 vs \u20b94,102.95, destination dwell 15.49 h vs 15.50 h. "
    "Data quality: 38/38 validation checks PASS; 200,528 changed records logged (100%); "
    "0 duplicate shipment_ids and 0 out-of-order scans remain after cleaning."))
c.font = BODY_FONT
c.alignment = WRAP
ws.row_dimensions[r].height = 45
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1

# ============================================================ Sheet 6
ws = wb.create_sheet("Monthly KPI Analysis")
r = title_block(
    ws,
    "Monthly KPI Analysis \u2014 2025 (SYNTHETIC DATA)",
    "Month-on-month (MoM) change: OTD in percentage points; cost and destination dwell in %. "
    "January has no MoM value (no prior month in the data). OTD is flat all year (75.72%\u201377.04%); "
    "no month breaches 78% and none falls below 75%.",
    14)
headers = ["Month", "Shipments", "OTD-eligible", "OTD %", "Late deliveries",
           "Cost / shipment (\u20b9)", "Total valid cost (\u20b9)", "Valid-cost shipments",
           "Avg dwell \u2014 dest (h)", "Avg dwell \u2014 origin (h)", "Avg dwell \u2014 comb. (h)",
           "OTD MoM (pp)", "Cost MoM (%)", "Dwell MoM (%)"]
header_row(ws, r, headers, [10, 12, 12, 9, 13, 15, 17, 14, 13, 13, 13, 11, 11, 11])
mon_head = r
r += 1
for row in monthly:
    body_cell(ws, r, 1, row["month"], bold=True)
    body_cell(ws, r, 2, num(row["shipments"]), fmt=FMT_INT)
    body_cell(ws, r, 3, num(row["otd_eligible"]), fmt=FMT_INT)
    body_cell(ws, r, 4, num(row["otd_pct"]) / 100.0, fmt=FMT_PCT)
    body_cell(ws, r, 5, num(row["late_deliveries"]), fmt=FMT_INT)
    body_cell(ws, r, 6, num(row["cost_per_shipment_inr"]), fmt=FMT_INR)
    body_cell(ws, r, 7, num(row["total_valid_cost_inr"]), fmt=FMT_INR)
    body_cell(ws, r, 8, num(row["valid_cost_shipments"]), fmt=FMT_INT)
    body_cell(ws, r, 9, num(row["avg_dwell_dest_h"]), fmt=FMT_HRS)
    body_cell(ws, r, 10, num(row["avg_dwell_origin_h"]), fmt=FMT_HRS)
    body_cell(ws, r, 11, num(row["avg_dwell_combined_h"]), fmt=FMT_HRS)
    for j, key in enumerate(("otd_mom_pp", "cost_mom_pct", "dwell_mom_pct"), start=12):
        v = row[key]
        if key == "cost_mom_pct" or key == "dwell_mom_pct":
            body_cell(ws, r, j, num(v) / 100.0 if v not in ("", None) else None, fmt=FMT_PCT)
        else:
            body_cell(ws, r, j, num(v) if v not in ("", None) else None, fmt=FMT_PP)
    r += 1
mon_last = r - 1
finish_sheet(ws, mon_head, 14)
# --- OTD line chart
chart = LineChart()
chart.title = "On-time delivery % by month \u2014 2025 (synthetic)"
chart.style = 2
chart.y_axis.title = "OTD %"
chart.x_axis.title = "Month"
chart.y_axis.number_format = "0%"
chart.height = 7.5
chart.width = 15
data = Reference(ws, min_col=4, min_row=mon_head, max_row=mon_last)
cats = Reference(ws, min_col=1, min_row=mon_head + 1, max_row=mon_last)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
ws.add_chart(chart, f"A{mon_last + 2}")

# ============================================================ Sheet 7
ws = wb.create_sheet("Hub Performance")
r = title_block(
    ws,
    "Hub Performance \u2014 Destination, Origin & Warehouse (SYNTHETIC DATA, 2025)",
    "Destination hubs drive the story: 6 hubs sit at ~83\u201384% OTD while GAU (53.95%), MLD (61.89%) "
    "and RNC (69.04%) lag far behind. Origin-hub OTD is narrow (75.26%\u201378.80%) \u2014 the problem is "
    "at the destination end, not dispatch. Warehouse dwell is uniform (15.21\u201315.95 h); no warehouse "
    "is an outlier on dwell.",
    11)
r = write_kpi_table(ws, r, hub_dest, "destination_hub", "Destination hub",
                    title="Destination hub \u2014 where the shipment was delivered")
r = write_kpi_table(ws, r, hub_orig, "origin_hub", "Origin hub",
                    title="Origin hub \u2014 where the shipment was dispatched")
r = write_kpi_table(ws, r, warehouse, "warehouse_id", "Warehouse",
                    title="Warehouse \u2014 destination dwell by warehouse (24 warehouses)")
ws.freeze_panes = "A5"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1

# ============================================================ Sheet 8
ws = wb.create_sheet("Route Performance")
r = title_block(
    ws,
    "Route Performance \u2014 61 Routes incl. UNKNOWN (SYNTHETIC DATA, 2025)",
    "OTD ranges from 57.11% (RT-006) to 83.41% (RT-003). The 6 worst routes (bottom 10% by OTD) "
    "are RT-006, RT-017, RT-028, RT-005, RT-050 and RT-039. 39 of 61 routes account for 80% of all "
    "late deliveries (Pareto). Cost per shipment is nearly flat across routes (\u20b94,052\u2013\u20b94,140) \u2014 "
    "the route problem is lateness, not cost.",
    11)
r = write_kpi_table(ws, r, route, "route_id", "Route")
ws.freeze_panes = "A5"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1

# ============================================================ Sheet 9
ws = wb.create_sheet("Category Performance")
r = title_block(
    ws,
    "Category & Vehicle Performance (SYNTHETIC DATA, 2025)",
    "Frozen is the weakest category (72.06% OTD, 20,895 late \u2014 33.1% of all late deliveries) and "
    "the second-most expensive (\u20b94,466.98/shipment). Pharma is the best on time (84.59%) but the "
    "most expensive (\u20b97,226.40). Vehicle type barely matters: OTD 76.24%\u201376.64% across all 5 types.",
    11)
r = write_kpi_table(ws, r, category, "category", "Category",
                    title="Product category")
r = write_kpi_table(ws, r, vehicle, "vehicle_type", "Vehicle type",
                    title="Vehicle type")
ws.freeze_panes = "A5"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1

# ============================================================ Sheet 10
ws = wb.create_sheet("KPI Scorecard")
r = title_block(
    ws,
    "KPI Scorecard \u2014 2025 (SYNTHETIC DATA)",
    "One-page scorecard. Thresholds are the project scorecard definitions; status is computed from "
    "the 2025 full-year values in this workbook. Data quality is GREEN: 38/38 validation checks PASS.",
    5)
headers = ["KPI", "Value (2025)", "Target / threshold", "Status", "Justification"]
header_row(ws, r, headers, [28, 18, 34, 12, 72])
sc_head = r
r += 1


def status_cell(ws, r, value, fmt, status):
    body_cell(ws, r, 1, value[0], bold=True)
    body_cell(ws, r, 2, value[1], fmt=fmt, bold=True)
    body_cell(ws, r, 3, value[2])
    fill, font = {"GREEN": (GREEN_FILL, GREEN_FONT), "AMBER": (AMBER_FILL, AMBER_FONT),
                  "RED": (RED_FILL, RED_FONT)}[status]
    body_cell(ws, r, 4, status, fill=fill, font=font)
    body_cell(ws, r, 5, value[3])
    ws.row_dimensions[r].height = 45


status_cell(ws, r, ("On-time delivery (OTD)", otd_frac,
                    "GREEN \u226585% | AMBER 75\u201385% | RED <75%",
                    "76.48% is 8.52 pp below the 85% target and inside the AMBER band. "
                    "63,118 of 268,330 eligible deliveries were late; H1 (76.49%) and H2 (76.47%) are flat, "
                    "so this is a structural gap, not a seasonal dip."), FMT_PCT, "AMBER")
r += 1
status_cell(ws, r, ("Cost per shipment", 4102.68,
                    "GREEN \u2264\u20b94,000 | AMBER \u20b94,000\u2013\u20b94,500 | RED >\u20b94,500",
                    "\u20b94,102.68 sits \u20b9102.68 above the \u20b94,000 target, inside the AMBER band. "
                    "Total valid cost \u20b91,204,074,833.69 across 293,485 valid-cost shipments; "
                    "route-level cost is nearly flat, so savings must come from the worst routes, not a network-wide cut."), FMT_INR, "AMBER")
r += 1
status_cell(ws, r, ("Avg dwell \u2014 destination hub", 15.49,
                    "GREEN \u226412 h | AMBER 12\u201318 h | RED >18 h",
                    "15.49 h is 3.49 h above the 12 h target, inside the AMBER band. "
                    "Origin dwell is only 2.5 h, so the dwell problem is concentrated at destination hubs; "
                    "warehouse dwell is uniform (15.21\u201315.95 h), pointing to a network-wide process, not one bad site."), FMT_HRS, "AMBER")
r += 1
status_cell(ws, r, ("Data quality", "38/38 PASS",
                    "GREEN: 100% changed-record logging + all validation checks pass",
                    "38/38 validation checks PASS; 200,528 changed records logged (100%); "
                    "0 duplicate shipment_ids and 0 out-of-order scans remain after cleaning. "
                    "All KPIs in this workbook are traceable to the trusted analysis table."), None, "GREEN")
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value=(
    "Overall: 3 AMBER, 1 GREEN, 0 RED. The network is stable but below target on every operational KPI; "
    "the binding constraint is on-time delivery, concentrated in destination hubs GAU / MLD / RNC and routes RT-006 / RT-005."))
c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
c.alignment = WRAP
ws.row_dimensions[r].height = 30
finish_sheet(ws, sc_head, 5, freeze=False)
ws.freeze_panes = f"A{sc_head + 1}"

# ============================================================ Sheet 11
ws = wb.create_sheet("Business Impact Analysis")
r = title_block(
    ws,
    "Business Impact Analysis \u2014 Modelled Improvement to the Network Median (SYNTHETIC DATA)",
    "MODELLED ESTIMATES on synthetic data \u2014 not a forecast of real savings. Improvement is modelled "
    "as bottom-quartile segments converging to the network median; no causal mechanism is assumed.",
    6)
headers = ["Lever", "Current (bottom quartile)", "Network median", "Gap", "Eligible volume", "Modelled gain at median"]
header_row(ws, r, headers, [20, 22, 16, 14, 18, 30])
bi_head = r
r += 1
routes_otd = bi["routes_otd"]
routes_cost = bi["routes_cost"]
wh_dwell = bi["warehouses_dwell"]
bi_rows = [
    ("On-time delivery (routes)",
     f"{routes_otd['bottom_quartile_current_otd_pct']}% (15 routes, OTD \u2264 {routes_otd['bottom_quartile_otd_threshold_pct']}%)",
     f"{routes_otd['median_route_otd_pct']}%", "13.99 pp",
     f"{routes_otd['bottom_quartile_eligible_shipments']:,} eligible shipments",
     f"{routes_otd['late_deliveries_avoided_at_median']:,} late deliveries avoided"),
    ("Cost per shipment (routes)",
     f"\u20b9{routes_cost['bottom_quartile_cost_threshold_inr']:,.2f} (15 routes above threshold)",
     f"\u20b9{routes_cost['median_route_cost_per_shipment_inr']:,.2f}",
     f"\u20b9{routes_cost['bottom_quartile_cost_threshold_inr'] - routes_cost['median_route_cost_per_shipment_inr']:.2f}",
     f"{routes_cost['bottom_quartile_valid_cost_shipments']:,} valid-cost shipments",
     f"\u20b9{routes_cost['potential_saving_inr_at_median']:,.2f} saving"),
    ("Destination dwell (warehouses)",
     f"{wh_dwell['bottom_quartile_dwell_threshold_h']} h (7 warehouses above threshold)",
     f"{wh_dwell['median_warehouse_dest_dwell_h']} h", "0.06 h",
     f"{wh_dwell['bottom_quartile_shipments']:,} shipments",
     f"{wh_dwell['excess_dwell_hours_at_median']:,.1f} excess dwell hours recovered"),
]
for row in bi_rows:
    for j, v in enumerate(row, start=1):
        body_cell(ws, r, j, v, bold=(j in (1, 6)))
    ws.row_dimensions[r].height = 45
    r += 1
r += 1
# --- scenarios
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = ws.cell(row=r, column=1, value="Scenarios \u2014 gap closed toward the median")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
headers = ["Scenario", "Gap closed", "Late deliveries avoided", "Cost saving (\u20b9)",
           "Excess dwell hours recovered", "Note"]
header_row(ws, r, headers, [16, 12, 22, 18, 24, 46])
sc2_head = r
r += 1
for name, key, note in [
    ("Conservative", "conservative", "50% of the gap to the median is closed."),
    ("Realistic", "realistic", "Bottom-quartile segments fully converge to the median."),
    ("Optimistic", "optimistic", "Full median convergence plus a further 10% efficiency gain on cost; not guaranteed."),
]:
    s = scen[key]
    body_cell(ws, r, 1, name, bold=True)
    body_cell(ws, r, 2, s["gap_closed_pct"] / 100.0, fmt=FMT_PCT)
    body_cell(ws, r, 3, s["late_deliveries_avoided"], fmt=FMT_INT)
    body_cell(ws, r, 4, s["potential_cost_saving_inr"], fmt=FMT_INR)
    body_cell(ws, r, 5, s["excess_dwell_hours_recovered"], fmt=FMT_HRS)
    body_cell(ws, r, 6, note)
    r += 1
r += 1
# --- assumptions & limitations
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = ws.cell(row=r, column=1, value="Assumptions & limitations (from kpi_summary.json \u2014 read before quoting any figure)")
c.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
r += 1
for a in bi["assumptions"]:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value="\u2022  " + a)
    c.font = BODY_FONT
    c.alignment = WRAP
    ws.row_dimensions[r].height = 30
    r += 1
extra_limits = [
    "All figures are computed on SYNTHETIC data (seed 42) for a fictional distributor; they describe the synthetic network only.",
    "No true year-over-year comparison exists (2025 only); H1 vs H2 2025 is the prior-period comparison.",
    "Dwell-hour savings are reported in hours because no verified per-hour handling cost exists in the data \u2014 dwell is not monetised.",
    "No behaviour change by customers, no capacity constraints and no implementation costs are included.",
    "Bottom-quartile definitions: routes with OTD \u2264 73.69% (15 routes); routes with cost/shipment above \u20b94,122.72 (15 routes); warehouses with destination dwell above 15.57 h (7 warehouses).",
]
for a in extra_limits:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value="\u2022  " + a)
    c.font = BODY_FONT
    c.alignment = WRAP
    ws.row_dimensions[r].height = 30
    r += 1
ws.freeze_panes = f"A{bi_head + 1}"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1

# ---------------------------------------------------------------- save
OUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT)
print(f"Saved {OUT} ({OUT.stat().st_size:,} bytes)")
print("Sheets:", wb.sheetnames)
