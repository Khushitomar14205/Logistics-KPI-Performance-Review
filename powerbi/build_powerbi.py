"""Build the Power BI artifacts for the Logistics KPI Performance Review.

Outputs (under powerbi/):
  data/                 import-ready CSVs (see export_data.py)
  Logistics_KPI_Performance_Review.pbip/
      Logistics_KPI_Performance_Review.pbip
      Logistics_KPI_Performance_Review.SemanticModel/
          definition.pbism, model.bim, diagramLayout.json, .platform
      Logistics_KPI_Performance_Review.Report/
          definition.pbir, report.json, .platform
  Logistics_KPI_Performance_Review.pbit   (OPC template; see pbit notes)
  VALIDATION.md          (written by validate_powerbi.py)

The model and the report are authored here from the project spec
(powerbi/DAX_Measures.md + Power_BI_Dashboard_Guide.md). No LLM
generation is used: every DAX expression is copied verbatim from
DAX_Measures.md, every number comes from the executed project outputs.

Key format rules applied (see powerbi/FORMAT_NOTES.md for sources):
- model.bim is TMSL JSON, compatibilityLevel 1567, culture en-US,
  defaultPowerBIDataSourceVersion "powerBI_V3".
- Every CSV partition uses Table.TransformColumnTypes with an explicit
  "en-US" culture argument (dot-decimal CSVs; without it a comma-decimal
  display culture would read 1355.00 as 135500).
- month_name sorts by month (sortByColumn).
- A measure never shares a name with a column of its own table
  (case-insensitive) — checked by the validator.
- The .pbit's text parts (Version, DataModelSchema, Report/Layout,
  Metadata, Settings, DiagramLayout) are written UTF-16LE with BOM,
  as in real Desktop files; the .pbip text files are UTF-8, no BOM.
- The DataMashup part is a real MS-QDEFF stream (Package Parts OPC ZIP
  with /Config/Package.xml + /Formulas/Section1.m, minimal Metadata and
  Permissions); see build_data_mashup().
- No SecurityBindings part (Desktop re-signs on save).

Run:  .venv/bin/python powerbi/build_powerbi.py
Then: .venv/bin/python powerbi/validate_powerbi.py
"""
import json
import secrets
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PB = ROOT / "powerbi"
DATA = PB / "data"
NAME = "Logistics_KPI_Performance_Review"
PBIP_DIR = PB / f"{NAME}.pbip"
SM_DIR = PBIP_DIR / f"{NAME}.SemanticModel"
RP_DIR = PBIP_DIR / f"{NAME}.Report"

# ---------------------------------------------------------------------------
# Model spec
# Each column: (name, tmsl_type, m_type, extra dict)
# tmsl_type: string | int64 | double | dateTime | boolean
# m_type:    text | Int64.Type | number | date | datetime | logical
# ---------------------------------------------------------------------------
COL = {
    "Fact_Shipment": [
        ("shipment_id", "string", "text", {}),
        ("shipment_date", "dateTime", "date", {}),
        ("month", "string", "text", {"isHidden": True}),
        ("origin_hub", "string", "text", {}),
        ("destination_hub", "string", "text", {}),
        ("route_id", "string", "text", {}),
        ("vehicle_id", "string", "text", {}),
        ("vehicle_type", "string", "text", {}),
        ("warehouse_id", "string", "text", {}),
        ("category", "string", "text", {}),
        ("status", "string", "text", {}),
        ("promised_ts_utc", "dateTime", "datetime", {}),
        ("actual_ts_utc", "dateTime", "datetime", {}),
        ("shipment_cost_inr", "double", "number", {"formatString": "#,##0.00"}),
        ("cost_valid", "boolean", "logical", {}),
        ("on_time", "boolean", "logical", {}),
        ("late_hours", "double", "number", {"formatString": "0.00"}),
        ("dwell_origin_h", "double", "number", {"formatString": "0.00"}),
        ("dwell_dest_h", "double", "number", {"formatString": "0.00"}),
        ("dwell_avg_h", "double", "number", {"formatString": "0.00"}),
        ("dwell_origin_valid", "boolean", "logical", {}),
        ("dwell_dest_valid", "boolean", "logical", {}),
    ],
    "Dim_Hub": [
        ("hub_id", "string", "text", {}),
        ("hub_name", "string", "text", {}),
        ("state", "string", "text", {}),
        ("region", "string", "text", {}),
        ("is_cold_chain", "boolean", "logical", {}),
    ],
    "Dim_Route": [
        ("route_id", "string", "text", {}),
        ("origin_hub", "string", "text", {}),
        ("destination_hub", "string", "text", {}),
        ("distance_km", "int64", "Int64.Type", {}),
        ("planned_transit_hours", "int64", "Int64.Type", {}),
    ],
    "Dim_Vehicle": [
        ("vehicle_id", "string", "text", {}),
        ("vehicle_type", "string", "text", {}),
        ("capacity_kg", "int64", "Int64.Type", {}),
        ("home_hub", "string", "text", {}),
    ],
    "Dim_Warehouse": [
        ("warehouse_id", "string", "text", {}),
        ("hub_id", "string", "text", {}),
        ("warehouse_name", "string", "text", {}),
        ("capacity_pallets", "int64", "Int64.Type", {}),
    ],
    "Dim_Category": [
        ("category", "string", "text", {}),
        ("base_cost_inr", "double", "number", {"formatString": "#,##0"}),
    ],
    "Dim_Date": [
        ("date", "dateTime", "date", {}),
        ("date_key", "int64", "Int64.Type", {}),
        ("month", "string", "text", {}),
        ("month_name", "string", "text", {"sortByColumn": "month"}),
        ("quarter", "string", "text", {}),
        ("year", "int64", "Int64.Type", {}),
        ("day_of_week", "string", "text", {}),
    ],
}

KPI_COLS = [
    ("shipments", "int64", "Int64.Type", {}),
    ("otd_eligible", "int64", "Int64.Type", {}),
    ("otd_pct", "double", "number", {}),
    ("late_deliveries", "int64", "Int64.Type", {}),
    ("cost_per_shipment_inr", "double", "number", {}),
    ("total_valid_cost_inr", "double", "number", {}),
    ("valid_cost_shipments", "int64", "Int64.Type", {}),
    ("avg_dwell_dest_h", "double", "number", {}),
    ("avg_dwell_origin_h", "double", "number", {}),
    ("avg_dwell_combined_h", "double", "number", {}),
]
COL["Monthly_KPI"] = KPI_COLS + [
    ("month", "string", "text", {}),
    ("otd_mom_pp", "double", "number", {}),
    ("cost_mom_pct", "double", "number", {}),
    ("dwell_mom_pct", "double", "number", {}),
]
COL["Half_Year_KPI"] = KPI_COLS + [("half", "string", "text", {})]
COL["Hub_Destination_KPI"] = KPI_COLS + [("destination_hub", "string", "text", {})]
COL["Route_KPI"] = KPI_COLS + [("route_id", "string", "text", {})]
COL["Warehouse_KPI"] = KPI_COLS + [("warehouse_id", "string", "text", {})]
COL["Category_KPI"] = KPI_COLS + [("category", "string", "text", {})]
COL["Cleaning_Log"] = [
    # "Records Affected" is TEXT, not Int64: the source cleaning log
    # records "see profiling (re-sequenced)" for out-of-order scans.
    # (The measured count is 29,287 — outputs/kpi_reports/
    # data_quality_checks.json — but the log's own wording is kept
    # verbatim; typing it Int64 would break Desktop's refresh.)
    ("Issue", "string", "text", {}),
    ("Records Affected", "string", "text", {}),
    ("Cleaning Action", "string", "text", {}),
    ("Reason", "string", "text", {}),
    ("Validation", "string", "text", {}),
]
COL["DQ_Missing"] = [
    ("table", "string", "text", {}),
    ("column", "string", "text", {}),
    ("dtype", "string", "text", {}),
    ("total_records", "int64", "Int64.Type", {}),
    ("missing_values", "int64", "Int64.Type", {}),
    ("missing_pct", "double", "number", {}),
    ("unique_values", "int64", "Int64.Type", {}),
]
COL["DQ_Summary"] = [
    ("check", "string", "text", {}),
    ("value", "string", "text", {}),
    ("status", "string", "text", {}),
]

TABLES = list(COL.keys())

# Calculated column on Fact_Shipment (DAX_Measures.md §2 step 5).
CALC_COLUMNS = {
    "Fact_Shipment": [
        ("Half", "string",
         'IF ( MONTH ( Fact_Shipment[shipment_date] ) <= 6, "H1 2025", "H2 2025" )'),
    ]
}

# ---------------------------------------------------------------------------
# Measures — copied verbatim from powerbi/DAX_Measures.md §4-§5.
# (name, expression, formatString or None, description)
# ---------------------------------------------------------------------------
MEASURES = [
    ("Total Shipments", "COUNTROWS ( Fact_Shipment )", "#,##0",
     "Total rows in the trusted shipment table."),
    ("OTD Eligible Shipments",
     "CALCULATE (\n    COUNTROWS ( Fact_Shipment ),\n"
     '    Fact_Shipment[status] IN { "Delivered", "Delayed" },\n'
     "    NOT ( ISBLANK ( Fact_Shipment[on_time] ) )\n)",
     "#,##0", "Delivered/Delayed shipments with both timestamps valid."),
    ("On-Time Deliveries",
     "CALCULATE (\n    COUNTROWS ( Fact_Shipment ),\n"
     '    Fact_Shipment[status] IN { "Delivered", "Delayed" },\n'
     "    Fact_Shipment[on_time] = TRUE ()\n)",
     "#,##0", "Eligible shipments delivered on or before the promised time."),
    ("Late Deliveries",
     "CALCULATE (\n    COUNTROWS ( Fact_Shipment ),\n"
     '    Fact_Shipment[status] IN { "Delivered", "Delayed" },\n'
     "    Fact_Shipment[on_time] = FALSE ()\n)",
     "#,##0", "Eligible shipments delivered after the promised time."),
    ("OTD %",
     "DIVIDE ( [On-Time Deliveries], [OTD Eligible Shipments] )",
     "0.00%", "Full-table value: 76.48%."),
    ("OTD % (Origin Hub)",
     "CALCULATE (\n    [OTD %],\n"
     "    USERELATIONSHIP ( Fact_Shipment[origin_hub], Dim_Hub[hub_id] )\n)",
     "0.00%", "OTD sliced by ORIGIN hub through the inactive relationship."),
    ("OTD % MoM Change (pp)",
     "VAR CurrentOTD = [OTD %]\nVAR PriorOTD =\n"
     "    CALCULATE ( [OTD %], PREVIOUSMONTH ( Dim_Date[date] ) )\nRETURN\n"
     "    IF ( NOT ISBLANK ( PriorOTD ), ( CurrentOTD - PriorOTD ) * 100 )",
     "0.00", "Month-over-month change in percentage points. Jan 2025 is blank."),
    ("OTD % H1 2025",
     'CALCULATE ( [OTD %], Fact_Shipment[Half] = "H1 2025" )',
     "0.00%", "Full-table value: 76.49%."),
    ("OTD % H2 2025",
     'CALCULATE ( [OTD %], Fact_Shipment[Half] = "H2 2025" )',
     "0.00%", "Full-table value: 76.47%."),
    ("OTD % Half-Year Change (pp)",
     "( [OTD % H2 2025] - [OTD % H1 2025] ) * 100",
     "0.00", "Prior-period comparison (H1 vs H2 2025). NOT year-over-year. "
              "Full-table value: -0.02 pp."),
    ("Total Valid Cost (INR)",
     "CALCULATE (\n    SUM ( Fact_Shipment[shipment_cost_inr] ),\n"
     "    Fact_Shipment[cost_valid] = TRUE ()\n)",
     "\u20b9#,##0", "Full-table value: \u20b91,20,40,74,834."),
    ("Valid Cost Shipments",
     "CALCULATE (\n    COUNTROWS ( Fact_Shipment ),\n"
     "    Fact_Shipment[cost_valid] = TRUE ()\n)",
     "#,##0", "Full-table value: 293,485 of 300,000 shipments."),
    ("Cost per Shipment (INR)",
     "DIVIDE ( [Total Valid Cost (INR)], [Valid Cost Shipments] )",
     "\u20b9#,##0.00", "Full-table value: \u20b94,102.68."),
    ("Cost per Shipment MoM Change (%)",
     "VAR CurrentCost = [Cost per Shipment (INR)]\nVAR PriorCost =\n"
     "    CALCULATE ( [Cost per Shipment (INR)], PREVIOUSMONTH ( Dim_Date[date] ) )\n"
     "RETURN\n"
     "    IF ( NOT ISBLANK ( PriorCost ), DIVIDE ( CurrentCost - PriorCost, PriorCost ) )",
     "0.00%", "Month-over-month change in cost per shipment."),
    ("Avg Destination Dwell (h)",
     "CALCULATE (\n    AVERAGE ( Fact_Shipment[dwell_dest_h] ),\n"
     "    Fact_Shipment[dwell_dest_valid] = TRUE (),\n"
     "    NOT ( ISBLANK ( Fact_Shipment[dwell_dest_h] ) )\n)",
     "0.00", "Full-table value: 15.49 h."),
    ("Avg Origin Dwell (h)",
     "CALCULATE (\n    AVERAGE ( Fact_Shipment[dwell_origin_h] ),\n"
     "    Fact_Shipment[dwell_origin_valid] = TRUE (),\n"
     "    NOT ( ISBLANK ( Fact_Shipment[dwell_origin_h] ) )\n)",
     "0.00", "Full-table value: 2.50 h."),
    ("Avg Combined Dwell (h)",
     "AVERAGE ( Fact_Shipment[dwell_avg_h] )",
     "0.00", "Full-table value: 8.90 h."),
    ("Avg Destination Dwell MoM Change (%)",
     "VAR CurrentDwell = [Avg Destination Dwell (h)]\nVAR PriorDwell =\n"
     "    CALCULATE ( [Avg Destination Dwell (h)], PREVIOUSMONTH ( Dim_Date[date] ) )\n"
     "RETURN\n"
     "    IF ( NOT ISBLANK ( PriorDwell ), DIVIDE ( CurrentDwell - PriorDwell, PriorDwell ) )",
     "0.00%", "Month-over-month change in average destination dwell."),
    ("Avg Late Hours",
     "CALCULATE (\n    AVERAGE ( Fact_Shipment[late_hours] ),\n"
     "    NOT ( ISBLANK ( Fact_Shipment[late_hours] ) )\n)",
     "0.00", "Mean (actual - promised) in hours over eligible shipments."),
    ("Late Share of Eligible (%)",
     "DIVIDE ( [Late Deliveries], [OTD Eligible Shipments] )",
     "0.00%", "Full-table value: 23.52% (100% - 76.48%)."),
    ("OTD Gap to 85% Target (pp)",
     "( 0.85 - [OTD %] ) * 100",
     "0.00", "Positive = below the 85% Green threshold. Full-table value: +8.52 pp."),
    ("Shipments (Valid Cost %)",
     "DIVIDE ( [Valid Cost Shipments], [Total Shipments] )",
     "0.00%", "Full-table value: 97.83%."),
    ("OTD RAG Status",
     "VAR OTDValue = [OTD %]\nRETURN\n    SWITCH (\n        TRUE (),\n"
     '        ISBLANK ( OTDValue ), "\u2014 No data",\n'
     '        OTDValue >= 0.85, "\u25cf Green \u2014 on target (\u2265 85%)",\n'
     '        OTDValue >= 0.75, "\u25cf Amber \u2014 below target (75\u201385%)",\n'
     '        "\u25cf Red \u2014 needs action (< 75%)"\n    )',
     None, "Thresholds are analyst-defined reporting conventions, not measured targets. "
           "Full-table value: Amber (76.48%)."),
    ("Cost RAG Status",
     "VAR CostValue = [Cost per Shipment (INR)]\nRETURN\n    SWITCH (\n        TRUE (),\n"
     '        ISBLANK ( CostValue ), "\u2014 No data",\n'
     '        CostValue <= 4000, "\u25cf Green \u2014 at or below \u20b94,000",\n'
     '        CostValue <= 4500, "\u25cf Amber \u2014 \u20b94,000\u2013\u20b94,500",\n'
     '        "\u25cf Red \u2014 needs action (> \u20b94,500)"\n    )',
     None, "Full-table value: Amber (\u20b94,102.68)."),
    ("Dwell RAG Status",
     "VAR DwellValue = [Avg Destination Dwell (h)]\nRETURN\n    SWITCH (\n        TRUE (),\n"
     '        ISBLANK ( DwellValue ), "\u2014 No data",\n'
     '        DwellValue <= 12, "\u25cf Green \u2014 at or below 12 h",\n'
     '        DwellValue <= 18, "\u25cf Amber \u2014 12\u201318 h",\n'
     '        "\u25cf Red \u2014 needs action (> 18 h)"\n    )',
     None, "Full-table value: Amber (15.49 h)."),
    ("Network Scorecard Status",
     "VAR OTDStatus =\n"
     '    SWITCH ( TRUE (), [OTD %] >= 0.85, "Green", [OTD %] >= 0.75, "Amber", "Red" )\n'
     "VAR CostStatus =\n    SWITCH (\n        TRUE (),\n"
     "        [Cost per Shipment (INR)] <= 4000, \"Green\",\n"
     "        [Cost per Shipment (INR)] <= 4500, \"Amber\",\n        \"Red\"\n    )\n"
     "VAR DwellStatus =\n    SWITCH (\n        TRUE (),\n"
     "        [Avg Destination Dwell (h)] <= 12, \"Green\",\n"
     "        [Avg Destination Dwell (h)] <= 18, \"Amber\",\n        \"Red\"\n    )\n"
     "VAR Greens =\n    ( IF ( OTDStatus = \"Green\", 1, 0 ) + IF ( CostStatus = \"Green\", 1, 0 )"
     " + IF ( DwellStatus = \"Green\", 1, 0 ) )\n"
     "VAR Reds =\n    ( IF ( OTDStatus = \"Red\", 1, 0 ) + IF ( CostStatus = \"Red\", 1, 0 )"
     " + IF ( DwellStatus = \"Red\", 1, 0 ) )\n"
     "RETURN\n    SWITCH (\n        TRUE (),\n"
     '        Reds > 0, "\u25cf Red \u2014 " & Reds & " of 3 KPIs need action",\n'
     '        Greens = 3, "\u25cf Green \u2014 all 3 KPIs on target",\n'
     '        "\u25cf Amber \u2014 " & Greens & " of 3 KPIs on target, none Red"\n    )',
     None, "Full-table value: \u25cf Amber \u2014 0 of 3 KPIs on target, none Red."),
]

# ---------------------------------------------------------------------------
# Relationships: (fromTable, fromColumn, toTable, toColumn, isActive)
# NOTE: Dim_Warehouse[hub_id] -> Dim_Hub[hub_id] is DELIBERATELY NOT created.
# Fact_Shipment already reaches Dim_Hub directly (destination_hub, active),
# so adding the warehouse path would create two active filter paths between
# Fact_Shipment and Dim_Hub and Desktop would deactivate one unpredictably.
# Warehouse visuals reach hub names through the fact table instead.
# ---------------------------------------------------------------------------
RELATIONSHIPS = [
    ("Fact_Shipment", "destination_hub", "Dim_Hub", "hub_id", True),
    ("Fact_Shipment", "origin_hub", "Dim_Hub", "hub_id", False),
    ("Fact_Shipment", "route_id", "Dim_Route", "route_id", True),
    ("Fact_Shipment", "vehicle_id", "Dim_Vehicle", "vehicle_id", True),
    ("Fact_Shipment", "category", "Dim_Category", "category", True),
    ("Fact_Shipment", "warehouse_id", "Dim_Warehouse", "warehouse_id", True),
    ("Fact_Shipment", "shipment_date", "Dim_Date", "date", True),
]


def m_partition(table: str) -> list:
    """M query for one CSV partition, reading via the DataFolder parameter."""
    import csv
    with open(DATA / f"{table}.csv", newline="", encoding="utf-8") as fh:
        ncols = len(next(csv.reader(fh)))
    type_list = ", ".join(
        '{{"%s", %s}}' % (c[0], ("type " + c[2]) if c[2] != "Int64.Type" else "Int64.Type")
        for c in COL[table]
    )
    return [
        "let",
        f'    FilePath = Text.Combine({{DataFolder, "{table}.csv"}}, "\\"),',
        "    Source = Csv.Document(File.Contents(FilePath),",
        f"        [Delimiter=\",\", Columns={ncols}, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),",
        "    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),",
        f"    Typed = Table.TransformColumnTypes(Promoted, {{{type_list}}}, \"en-US\")",
        "in",
        "    Typed",
    ]


def build_model_bim() -> dict:
    tables = []
    for t in TABLES:
        cols = []
        for name, dtype, _mtype, extra in COL[t]:
            col = {"name": name, "dataType": dtype, "sourceColumn": name,
                   "summarizeBy": "none"}
            if "formatString" in extra:
                col["formatString"] = extra["formatString"]
            if extra.get("isHidden"):
                col["isHidden"] = True
            if "sortByColumn" in extra:
                col["sortByColumn"] = extra["sortByColumn"]
            cols.append(col)
        for name, dtype, expr in CALC_COLUMNS.get(t, []):
            cols.append({"name": name, "dataType": dtype, "expression": expr,
                         "summarizeBy": "none"})
        tab = {
            "name": t,
            "columns": cols,
            "partitions": [{
                "name": t,
                "mode": "import",
                "source": {"type": "m", "expression": m_partition(t)},
            }],
        }
        if t == "Fact_Shipment":
            tab["measures"] = [
                {"name": n, "expression": e, **({"formatString": f} if f else {}),
                 "description": d}
                for n, e, f, d in MEASURES
            ]
        tables.append(tab)
    rels = []
    for i, (ft, fc, tt, tc, active) in enumerate(RELATIONSHIPS):
        rel = {
            "name": f"rel_{i:02d}_{ft}_{fc}_to_{tt}_{tc}",
            "fromTable": ft, "fromColumn": fc,
            "toTable": tt, "toColumn": tc,
            "crossFilteringBehavior": "oneDirection",
            "isActive": active,
        }
        if tt == "Dim_Date":
            rel["joinOnDateBehavior"] = "dateAndTime"
        rels.append(rel)
    return {
        "name": NAME.replace("_", " "),
        "compatibilityLevel": 1567,
        "model": {
            "culture": "en-US",
            "sourceQueryCulture": "en-US",
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "defaultMode": "import",
            "dataAccessOptions": {"legacyRedirects": True,
                                  "returnErrorValuesAsNull": True},
            "tables": tables,
            "relationships": rels,
            "expressions": [
                {
                    "name": "DataFolder",
                    "kind": "m",
                    "expression": (
                        '"" meta [IsParameterQuery=true, Type="Text", '
                        'IsParameterQueryRequired=true]'
                    ),
                    "description": (
                        "Folder containing the powerbi/data CSVs (e.g. "
                        "C:\\Users\\Khushi\\Logistics-KPI-Performance-Review\\"
                        "powerbi\\data\\). Desktop prompts for it on first "
                        "open/refresh. Must end with a backslash."
                    ),
                }
            ],
        },
    }

# ---------------------------------------------------------------------------
# Report (PBIR-Legacy single report.json; the same JSON is used as
# Report/Layout inside the .pbit)
# ---------------------------------------------------------------------------
PAGE_W, PAGE_H = 1280, 720


def hexid() -> str:
    return secrets.token_hex(10)


def qref(table: str, field: str) -> str:
    return f"{table}.{field}"


def _col_expr(alias: str, prop: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Source": alias}},
                       "Property": prop}}


def _mea_expr(alias: str, prop: str) -> dict:
    return {"Measure": {"Expression": {"SourceRef": {"Source": alias}},
                        "Property": prop}}


class VisualSpec:
    """One data visual: (table, field, is_measure) bindings + roles."""

    def __init__(self, visual_type, x, y, w, h, roles, order_by=None,
                 objects=None):
        self.visual_type = visual_type
        self.x, self.y, self.w, self.h = x, y, w, h
        # roles: {"Category": [("Dim_Hub","hub_name",False)], "Y": [("Fact_Shipment","OTD %",True)]}
        self.roles = roles
        self.order_by = order_by  # ("Table","field",is_measure,"asc"|"desc")
        self.objects = objects or {}


def _build_visual(spec: VisualSpec) -> dict:
    # Assign one alias per table, in first-use order.
    aliases, tables = {}, []
    for role_fields in spec.roles.values():
        for (t, _f, _m) in role_fields:
            if t not in aliases:
                aliases[t] = chr(ord("a") + len(tables))
                tables.append(t)
    if spec.order_by:
        t = spec.order_by[0]
        if t not in aliases:
            aliases[t] = chr(ord("a") + len(tables))
            tables.append(t)
    projections = {}
    selects = []
    seen_selects = set()
    for role, fields in spec.roles.items():
        refs = []
        for (t, f, is_mea) in fields:
            refs.append({"queryRef": qref(t, f)})
            key = (t, f)
            if key not in seen_selects:
                seen_selects.add(key)
                entry = ((_mea_expr if is_mea else _col_expr)(aliases[t], f))
                entry["Name"] = qref(t, f)
                entry["NativeReferenceName"] = f
                selects.append(entry)
        projections[role] = refs
    proto = {
        "Version": 2,
        "From": [{"Name": aliases[t], "Entity": t, "Type": 0} for t in tables],
        "Select": selects,
    }
    if spec.order_by:
        t, f, is_mea, direction = spec.order_by
        expr = (_mea_expr if is_mea else _col_expr)(aliases[t], f)
        proto["OrderBy"] = [{"Direction": 2 if direction == "desc" else 1,
                             "Expression": expr}]
    single = {
        "visualType": spec.visual_type,
        "projections": projections,
        "prototypeQuery": proto,
        "drillFilterOtherVisuals": True,
    }
    if spec.objects:
        single["objects"] = spec.objects
    vid = hexid()
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": spec.x, "y": spec.y, "z": 1,
                                            "width": spec.w, "height": spec.h}}],
        "singleVisual": single,
    }
    return {"x": spec.x, "y": spec.y, "z": 1, "width": spec.w, "height": spec.h,
            "config": json.dumps(config, separators=(",", ":")),
            "filters": "[]"}


def textbox(x, y, w, h, paragraphs):
    """paragraphs: list of (text, size, bold, color)."""
    paras = []
    for text, size, bold, color in paragraphs:
        paras.append({
            "textRuns": [{"value": text,
                          "textStyle": {"fontSize": size, "bold": bold,
                                        "fontFamily": "Segoe UI",
                                        "color": color}}],
            "horizontalTextAlignment": "left",
        })
    vid = hexid()
    config = {
        "name": vid,
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": 1,
                                            "width": w, "height": h}}],
        "singleVisual": {
            "visualType": "textbox",
            "projections": {},
            "drillFilterOtherVisuals": True,
            "objects": {"general": {"paragraphs": paras}},
            "vcObjects": {"background": {"show": False},
                          "border": {"show": False}},
        },
    }
    return {"x": x, "y": y, "z": 1, "width": w, "height": h,
            "config": json.dumps(config, separators=(",", ":")),
            "filters": "[]"}


def slicer(x, y, w, h, table, column):
    return _build_visual(VisualSpec(
        "slicer", x, y, w, h, {"Values": [(table, column, False)]}))


def card(x, y, w, h, measure):
    return _build_visual(VisualSpec(
        "card", x, y, w, h, {"Values": [("Fact_Shipment", measure, True)]},
        objects={"categoryLabels": {"show": False},
                 "labels": {"show": True}}))


def rag_card(x, y, w, h, measure):
    return _build_visual(VisualSpec(
        "cardVisual", x, y, w, h, {"Data": [("Fact_Shipment", measure, True)]}))


def page_header(title):
    return [
        textbox(20, 8, 1240, 34,
                [(title, 20, True, "#1F3864")]),
        textbox(20, 44, 1240, 20,
                [("Synthetic data (seed 42) \u2014 fictional Kolkata cold-chain "
                  "distributor, 2025. Not real company data. Headline KPIs are "
                  "network-level; use the slicers to explore.",
                  11, False, "#595959")]),
    ]


def build_report() -> dict:
    F = "Fact_Shipment"
    sections = []

    # ---------------------------------------------------------- Page 1
    p1 = page_header("Logistics KPI Performance Review \u2014 Executive Overview")
    p1 += [slicer(20, 68, 300, 42, "Dim_Date", "month_name"),
           slicer(340, 68, 300, 42, "Dim_Category", "category"),
           slicer(660, 68, 300, 42, "Dim_Hub", "hub_name")]
    p1 += [card(20, 122, 300, 108, "Total Shipments"),
           card(340, 122, 300, 108, "OTD %"),
           card(660, 122, 300, 108, "Cost per Shipment (INR)"),
           card(980, 122, 280, 108, "Avg Destination Dwell (h)")]
    p1 += [rag_card(20, 232, 300, 34, "Network Scorecard Status"),
           rag_card(340, 232, 300, 34, "OTD RAG Status"),
           rag_card(660, 232, 300, 34, "Cost RAG Status"),
           rag_card(980, 232, 280, 34, "Dwell RAG Status")]
    p1.append(_build_visual(VisualSpec(
        "lineChart", 20, 278, 620, 200,
        {"Category": [("Dim_Date", "month_name", False)],
         "Y": [(F, "OTD %", True)]},
        order_by=("Dim_Date", "month", False, "asc"))))
    p1.append(_build_visual(VisualSpec(
        "lineClusteredColumnComboChart", 660, 278, 600, 200,
        {"Category": [("Dim_Date", "month_name", False)],
         "Y": [(F, "Cost per Shipment (INR)", True)],
         "Y2": [(F, "Avg Destination Dwell (h)", True)]},
        order_by=("Dim_Date", "month", False, "asc"))))
    p1.append(textbox(
        20, 492, 620, 92,
        [("Half-year comparison (prior-period, NOT year-over-year)", 13, True, "#1F3864"),
         ("H1 2025 OTD 76.49%  \u00b7  H2 2025 OTD 76.47%  \u00b7  Change \u22120.02 pp. "
          "The data covers 2025 only, so there is no true year-over-year comparison.",
          12, False, "#1F1F1F")]))
    p1.append(textbox(
        660, 492, 600, 92,
        [("Business impact \u2014 modelled estimate on synthetic data (not a forecast)", 13, True, "#1F3864"),
         ("Realistic case, bottom-quartile segments improving to the median: 9,706 late "
          "deliveries avoided \u00b7 \u20b931,13,833 potential saving \u00b7 4,295.8 excess "
          "dwell hours recovered. Conservative: 4,853 / \u20b915,56,916 / 2,147.9 h. "
          "Optimistic: 9,706 / \u20b934,25,216 / 4,295.8 h.",
          12, False, "#1F1F1F")]))
    p1.append(textbox(
        20, 598, 1240, 104,
        [("How to read this page", 13, True, "#1F3864"),
         ("OTD % 76.48% (Amber: Green \u226585%, Amber 75\u201385%, Red <75%)  \u00b7  "
          "Cost per shipment \u20b94,102.68 (Amber: Green \u2264\u20b94,000, Amber "
          "\u20b94,000\u2013\u20b94,500, Red >\u20b94,500)  \u00b7  Avg destination dwell 15.49 h "
          "(Amber: Green \u226412 h, Amber 12\u201318 h, Red >18 h). Thresholds are "
          "analyst-defined reporting conventions, not measured targets. Monthly OTD is "
          "flat (75.72%\u201377.04%) \u2014 the problem is structural and geographic, not seasonal.",
          12, False, "#1F1F1F")]))
    sections.append(("ExecutiveOverview", "Executive Overview", p1))

    # ---------------------------------------------------------- Page 2
    p2 = page_header("Hub & Warehouse Performance")
    p2 += [slicer(20, 68, 300, 42, "Dim_Hub", "state"),
           slicer(340, 68, 300, 42, "Dim_Category", "category"),
           slicer(660, 68, 300, 42, "Dim_Date", "quarter")]
    p2.append(_build_visual(VisualSpec(
        "clusteredBarChart", 20, 122, 620, 218,
        {"Category": [("Dim_Hub", "hub_name", False)],
         "Y": [(F, "OTD %", True)]},
        order_by=(F, "OTD %", True, "desc"))))
    p2.append(_build_visual(VisualSpec(
        "clusteredBarChart", 660, 122, 600, 218,
        {"Category": [("Dim_Hub", "hub_name", False)],
         "Y": [(F, "OTD % (Origin Hub)", True)]},
        order_by=(F, "OTD % (Origin Hub)", True, "desc"))))
    p2.append(_build_visual(VisualSpec(
        "tableEx", 20, 352, 620, 178,
        {"Values": [("Dim_Hub", "hub_name", False),
                    ("Dim_Hub", "state", False),
                    (F, "Total Shipments", True),
                    (F, "OTD Eligible Shipments", True),
                    (F, "Late Deliveries", True),
                    (F, "OTD %", True),
                    (F, "OTD RAG Status", True),
                    (F, "Cost per Shipment (INR)", True),
                    (F, "Avg Destination Dwell (h)", True)]},
        order_by=(F, "OTD %", True, "asc"))))
    p2.append(_build_visual(VisualSpec(
        "clusteredBarChart", 660, 352, 600, 178,
        {"Category": [("Dim_Warehouse", "warehouse_id", False)],
         "Y": [(F, "Avg Destination Dwell (h)", True)]},
        order_by=(F, "Avg Destination Dwell (h)", True, "desc"))))
    p2.append(_build_visual(VisualSpec(
        "tableEx", 20, 542, 620, 160,
        {"Values": [("Dim_Warehouse", "warehouse_id", False),
                    ("Dim_Warehouse", "warehouse_name", False),
                    ("Dim_Hub", "hub_name", False),
                    (F, "Total Shipments", True),
                    (F, "OTD %", True),
                    (F, "Avg Destination Dwell (h)", True),
                    (F, "Cost per Shipment (INR)", True)]},
        order_by=(F, "Avg Destination Dwell (h)", True, "desc"))))
    p2.append(_build_visual(VisualSpec(
        "scatterChart", 660, 542, 600, 160,
        {"X": [(F, "Avg Destination Dwell (h)", True)],
         "Y": [(F, "OTD %", True)],
         "Details": [("Dim_Warehouse", "warehouse_id", False)]})))
    sections.append(("HubWarehouse", "Hub & Warehouse Performance", p2))

    # ---------------------------------------------------------- Page 3
    p3 = page_header("Route & Shipment Analysis")
    p3 += [slicer(20, 68, 300, 42, "Dim_Hub", "hub_name"),
           slicer(340, 68, 300, 42, "Dim_Category", "category"),
           slicer(660, 68, 300, 42, "Dim_Date", "month_name")]
    p3.append(_build_visual(VisualSpec(
        "clusteredBarChart", 20, 122, 400, 200,
        {"Category": [("Dim_Route", "route_id", False)],
         "Y": [(F, "OTD %", True)]},
        order_by=(F, "OTD %", True, "asc"))))
    p3.append(_build_visual(VisualSpec(
        "clusteredBarChart", 440, 122, 400, 200,
        {"Category": [("Dim_Route", "route_id", False)],
         "Y": [(F, "OTD %", True)]},
        order_by=(F, "OTD %", True, "desc"))))
    p3.append(_build_visual(VisualSpec(
        "clusteredColumnChart", 860, 122, 400, 200,
        {"Category": [("Dim_Category", "category", False)],
         "Y": [(F, "OTD %", True)]},
        order_by=(F, "OTD %", True, "desc"))))
    p3.append(_build_visual(VisualSpec(
        "clusteredColumnChart", 20, 336, 400, 178,
        {"Category": [("Dim_Category", "category", False)],
         "Y": [(F, "Cost per Shipment (INR)", True)]},
        order_by=(F, "Cost per Shipment (INR)", True, "desc"))))
    p3.append(_build_visual(VisualSpec(
        "donutChart", 440, 336, 400, 178,
        {"Category": [("Dim_Category", "category", False)],
         "Y": [(F, "Late Deliveries", True)]})))
    p3.append(_build_visual(VisualSpec(
        "clusteredBarChart", 860, 336, 400, 178,
        {"Category": [("Dim_Vehicle", "vehicle_type", False)],
         "Y": [(F, "OTD %", True)]},
        order_by=(F, "OTD %", True, "desc"))))
    p3.append(_build_visual(VisualSpec(
        "tableEx", 20, 528, 620, 176,
        {"Values": [("Dim_Route", "route_id", False),
                    ("Dim_Route", "origin_hub", False),
                    ("Dim_Route", "destination_hub", False),
                    ("Dim_Route", "distance_km", False),
                    (F, "Total Shipments", True),
                    (F, "OTD %", True),
                    (F, "Late Deliveries", True),
                    (F, "Cost per Shipment (INR)", True),
                    (F, "Avg Destination Dwell (h)", True)]},
        order_by=(F, "OTD %", True, "asc"))))
    p3.append(_build_visual(VisualSpec(
        "clusteredBarChart", 660, 528, 600, 176,
        {"Category": [("Dim_Route", "route_id", False)],
         "Y": [(F, "Late Deliveries", True)]},
        order_by=(F, "Late Deliveries", True, "desc"))))
    sections.append(("RouteShipment", "Route & Shipment Analysis", p3))

    # ---------------------------------------------------------- Page 4
    p4 = page_header("Data Quality Monitoring")
    p4.append(textbox(20, 68, 300, 70,
                       [("Raw records", 12, True, "#595959"),
                        ("1,504,500", 18, True, "#1F3864"),
                        ("shipments + scan events", 11, False, "#595959")]))
    p4.append(textbox(340, 68, 300, 70,
                       [("Clean records", 12, True, "#595959"),
                        ("1,494,000", 18, True, "#1F3864"),
                        ("4,500 duplicate copies + 6,000 orphan scans removed",
                         11, False, "#595959")]))
    p4.append(textbox(660, 68, 300, 70,
                       [("Shipments 304,500 \u2192 300,000", 12, True, "#595959"),
                        ("Scans 1,200,000 \u2192 1,194,000", 12, True, "#1F3864"),
                        ("reconciliation within 0.5% tolerance", 11, False, "#595959")]))
    p4.append(textbox(980, 68, 280, 70,
                       [("Validation", 12, True, "#595959"),
                        ("38 / 38 checks PASS", 18, True, "#1F3864"),
                        ("0% variance", 11, False, "#595959")]))
    p4.append(_build_visual(VisualSpec(
        "tableEx", 20, 150, 620, 220,
        {"Values": [("Cleaning_Log", "Issue", False),
                    ("Cleaning_Log", "Records Affected", False),
                    ("Cleaning_Log", "Cleaning Action", False),
                    ("Cleaning_Log", "Reason", False),
                    ("Cleaning_Log", "Validation", False)]})))
    p4.append(_build_visual(VisualSpec(
        "clusteredColumnChart", 660, 150, 600, 220,
        {"Category": [("DQ_Missing", "column", False)],
         "Y": [("DQ_Missing", "missing_pct", False)]},
        order_by=("DQ_Missing", "missing_pct", False, "desc"))))
    p4.append(textbox(
        20, 386, 620, 92,
        [("Orphan scans & re-sequencing", 13, True, "#1F3864"),
         ("6,000 orphan scan events removed; 0 remain. 29,287 out-of-order scan events "
          "in 20,914 shipments re-sequenced by UTC timestamp; 0 out-of-order pairs "
          "remain after cleaning.", 12, False, "#1F1F1F")]))
    p4.append(textbox(
        660, 386, 600, 92,
        [("Negative dwell", 13, True, "#1F3864"),
         ("Negative dwell pairs flagged and excluded from the dwell KPI (not corrected): "
          "destination 2,885, origin 18 shipments (post-dedupe counts; raw profile: "
          "2,905 / 39).", 12, False, "#1F1F1F")]))
    p4.append(_build_visual(VisualSpec(
        "tableEx", 20, 492, 620, 200,
        {"Values": [("Monthly_KPI", "month", False),
                    ("Monthly_KPI", "shipments", False),
                    ("Monthly_KPI", "otd_pct", False),
                    ("Monthly_KPI", "cost_per_shipment_inr", False),
                    ("Monthly_KPI", "avg_dwell_dest_h", False)]})))
    p4.append(textbox(
        660, 492, 600, 200,
        [("What was NOT done", 13, True, "#1F3864"),
         ("No imputation: missing costs/timestamps stay excluded via validity flags; "
          "missing category/vehicle kept as Unknown/UNKNOWN members. No dwell-to-rupees "
          "conversion (no verified per-hour cost). Naive timestamps assumed UTC (rule C3). "
          "No true YoY \u2014 H1 vs H2 2025 only.", 12, False, "#1F1F1F"),
         ("Profile counts are pre-dedupe (raw) while cleaning-log counts are post-dedupe "
          "\u2014 both are correct at their stage.", 12, False, "#595959")]))
    sections.append(("DataQuality", "Data Quality Monitoring", p4))

    out_sections = []
    for i, (name, display, visuals) in enumerate(sections):
        out_sections.append({
            "name": name,
            "displayName": display,
            "displayOption": 1,
            "width": PAGE_W,
            "height": PAGE_H,
            "ordinal": i,
            "config": json.dumps({
                "name": name,
                "layouts": [{"id": 0, "position": {"x": 0, "y": 0,
                                                    "width": PAGE_W,
                                                    "height": PAGE_H}}],
                "singleVisualGroup": None,
            }, separators=(",", ":")),
            "filters": "[]",
            "visualContainers": visuals,
        })
    report_config = {
        "version": "5.70",
        "activeSectionIndex": 0,
        "defaultDrillFilterOtherVisuals": True,
        "settings": {
            "useNewFilterPaneExperience": True,
            "allowChangeFilterTypes": True,
            "useStylableVisualContainerHeader": True,
            "exportDataMode": 1,
        },
    }
    return {
        "config": json.dumps(report_config, separators=(",", ":")),
        "layoutOptimization": 0,
        "sections": out_sections,
    }

# ---------------------------------------------------------------------------
# DataMashup (MS-QDEFF stream) — real queries, not a stub
# ---------------------------------------------------------------------------
def section1_m() -> str:
    """Section1.m: the DataFolder parameter + one shared query per table.

    The M here is identical to the partition expressions in model.bim.
    """
    lines = ["section Section1;", ""]
    lines.append(
        'shared DataFolder = "" meta [IsParameterQuery=true, Type="Text", '
        'IsParameterQueryRequired=true];'
    )
    lines.append("")
    for t in TABLES:
        body = m_partition(t)
        # m_partition returns ["let", ...lines..., "in", "    Typed"]
        m_lines = [body[0]] + body[1:]
        lines.append(f"shared {t} = " + m_lines[0])
        lines.extend(m_lines[1:])
        lines.append("")
    return "\n".join(lines)


def build_data_mashup() -> bytes:
    """MS-QDEFF binary stream: 8-byte header + 5 length-prefixed parts.

    Parts in order: Package Parts (OPC ZIP), Metadata (XML),
    Metadata Content (empty), Permissions (XML), Permission Bindings
    (empty — Desktop then falls back to default permissions, see the
    MS-QDEFF notes in FORMAT_NOTES.md).
    """
    import io
    pkg = io.BytesIO()
    with zipfile.ZipFile(pkg, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="utf-8"?>\n'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
                   '  <Default Extension="xml" ContentType="application/xml" />\n'
                   '  <Override PartName="/Config/Package.xml" ContentType="application/xml" />\n'
                   '  <Override PartName="/Formulas/Section1.m" ContentType="application/octet-stream" />\n'
                   '</Types>\n')
        z.writestr("Config/Package.xml",
                   '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n'
                   '<Package xmlns="http://schemas.microsoft.com/DataMashup">\n'
                   '  <Version>2.0</Version>\n'
                   '  <MinVersion>2.0</MinVersion>\n'
                   '  <Culture>en-US</Culture>\n'
                   '</Package>\n')
        z.writestr("Formulas/Section1.m", section1_m().encode("utf-8"))
    package_parts = pkg.getvalue()
    metadata = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<Metadata xmlns="http://schemas.microsoft.com/DataMashup">\n'
        '  <AllFormulas />\n'
        '  <Formulas />\n'
        '</Metadata>\n'
    ).encode("utf-8")
    permissions = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<Permissions xmlns="http://schemas.microsoft.com/DataMashup">\n'
        '  <CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\n'
        '  <FirewallEnabled>false</FirewallEnabled>\n'
        '  <WorkbookGroupType xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xsi:nil="true" />\n'
        '</Permissions>\n'
    ).encode("utf-8")
    out = bytearray(b"\x01\x00\x00\x00\x01\x00\x00\x00")
    for part in (package_parts, metadata, b"", permissions, b""):
        out += len(part).to_bytes(4, "little")
        out += part
    return bytes(out)


# ---------------------------------------------------------------------------
# PBIP assembly
# ---------------------------------------------------------------------------
def _write_json(path: Path, payload, crlf=False):
    text = json.dumps(payload, indent=2, ensure_ascii=False)
    if crlf:
        text = text.replace("\n", "\r\n")
    path.write_text(text + ("\r\n" if crlf else "\n"), encoding="utf-8")


def _platform(display_name: str, item_type: str) -> dict:
    return {
        "version": "1.0",
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/"
                   "gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": item_type, "displayName": display_name},
        "config": {"logicalId": str(__import__("uuid").uuid4())},
    }


def _diagram_layout() -> dict:
    pos = {
        "Fact_Shipment": (560, 120), "Dim_Hub": (120, 40),
        "Dim_Route": (120, 200), "Dim_Vehicle": (120, 360),
        "Dim_Warehouse": (120, 520), "Dim_Category": (980, 40),
        "Dim_Date": (980, 200),
    }
    return {"version": "1.0", "diagrams": [{"nodes": {
        t: {"x": x, "y": y} for t, (x, y) in pos.items()}}]}


def build_pbip(model: dict, report: dict) -> None:
    import shutil
    if PBIP_DIR.exists():
        shutil.rmtree(PBIP_DIR)
    SM_DIR.mkdir(parents=True)
    RP_DIR.mkdir(parents=True)
    _write_json(PBIP_DIR / f"{NAME}.pbip", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/"
                   "pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })
    _write_json(SM_DIR / "definition.pbism", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2",
        "settings": {},
    })
    _write_json(SM_DIR / "model.bim", model)
    _write_json(SM_DIR / "diagramLayout.json", _diagram_layout())
    _write_json(SM_DIR / ".platform", _platform(NAME, "SemanticModel"))
    _write_json(RP_DIR / "definition.pbir", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "report/definitionProperties/2.0.0/schema.json",
        "version": "1.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}},
    })
    # PBIR-Legacy report.json: pretty JSON, CRLF, no BOM.
    _write_json(RP_DIR / "report.json", report, crlf=True)
    _write_json(RP_DIR / ".platform", _platform(NAME, "Report"))


# ---------------------------------------------------------------------------
# .pbit assembly (via nl2pbip's OPC builder — OPC/.pbit export path only)
# ---------------------------------------------------------------------------
def _utf16le(text: str) -> bytes:
    return b"\xff\xfe" + text.encode("utf-16-le")


def build_pbit(model: dict, report: dict) -> Path:
    import sys
    sys.path.insert(0, str(ROOT / ".venv/lib/python3.12/site-packages"))
    from nl2pbip.exporter.pbit_builder import PbitArchiveBuilder
    from nl2pbip.exporter.opc import (make_diagram_layout_json,
                                      make_metadata_json,
                                      make_security_bindings_json,
                                      make_settings_json)

    out = PB / f"{NAME}.pbit"
    b = PbitArchiveBuilder(output_path=out, template_name=NAME.replace("_", " "),
                           culture="en-US",
                           description="Logistics KPI Performance Review — "
                                       "synthetic cold-chain data (seed 42), 2025.")
    # Text parts are UTF-16LE with BOM, as in real Desktop files — EXCEPT
    # Version, which Desktop's ValidateVersion rejects if it carries a BOM
    # ("'﻿1.28' is not a valid .pbix file version number", Desktop Oct 2026).
    b.add_part("Version", "1.28".encode("utf-16-le"),
               content_type="application/json")
    b.add_part("DataModelSchema",
               _utf16le(json.dumps(model, indent=2, ensure_ascii=False)),
               content_type="application/json")
    b.add_part("Report/Layout",
               _utf16le(json.dumps(report, indent=2, ensure_ascii=False)),
               content_type="application/json")
    b.add_part("Metadata",
               _utf16le(json.dumps(make_metadata_json(
                   title=NAME.replace("_", " "),
                   description="Logistics KPI Performance Review — synthetic "
                               "cold-chain data (seed 42), 2025. Open in Power BI "
                               "Desktop, set the DataFolder parameter to the "
                               "powerbi/data folder, Refresh, then Save As .pbix.",
                   created="2026-11-09",
                   generator="logistics-kpi build_powerbi.py"), indent=2,
                   ensure_ascii=False)),
               content_type="application/json")
    b.add_part("Settings",
               _utf16le(json.dumps(make_settings_json(), indent=2)),
               content_type="application/json")
    b.add_part("DiagramLayout",
               _utf16le(json.dumps(_diagram_layout(), indent=2)),
               content_type="application/json")
    b.add_part("DataMashup", build_data_mashup(),
               content_type="application/octet-stream")
    b.finalize()
    return out


def main():
    model = build_model_bim()
    report = build_report()
    build_pbip(model, report)
    pbit = build_pbit(model, report)
    n_visuals = sum(len(s["visualContainers"]) for s in report["sections"])
    print(f"Model: {len(TABLES)} tables, {len(MEASURES)} measures, "
          f"{len(RELATIONSHIPS)} relationships, 1 parameter (DataFolder)")
    print(f"Report: {len(report['sections'])} pages, {n_visuals} visuals")
    print(f"PBIP folder: {PBIP_DIR}")
    print(f"PBIT: {pbit} ({pbit.stat().st_size / 1e6:.2f} MB)")
    print("Next: .venv/bin/python powerbi/validate_powerbi.py")


if __name__ == "__main__":
    main()
