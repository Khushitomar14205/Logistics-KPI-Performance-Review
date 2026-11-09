# Power BI file formats — notes for `build_powerbi.py`

This file records what was actually relied on to generate the `.pbip` and
`.pbit` artifacts, so a future rebuild can be checked instead of guessed.
Sources: the PBIP reference implementations and format write-ups listed at
the end, plus the project's own `DAX_Measures.md` and
`Power_BI_Dashboard_Guide.md`. Nothing here claims an official Microsoft
spec — the formats are community-documented, and the one honest limit is
that rendering/refresh can only be verified inside Power BI Desktop.

## 1. What a `.pbix` is — and why this project ships `.pbit` + `.pbip`

A `.pbix` is an OPC (Open Packaging Conventions) ZIP whose `DataModel`
part is a **compressed, binary** Analysis Services model. That binary
DataModel can only be produced by Power BI Desktop (or the Analysis
Services engine). It cannot be written by code, so no `.pbix` is shipped
and none is faked. The honest maximum is:

- **`.pbit`** — the same OPC ZIP, but with `DataModelSchema` (plain TMSL
  JSON) instead of the binary `DataModel`. Desktop opens it, prompts for
  the `DataFolder` parameter, refreshes from the CSVs, and Save As
  produces a real `.pbix`.
- **`.pbip`** — the project folder (PBIP) form of the same model + report:
  `model.bim` (TMSL JSON) in `.SemanticModel`, `report.json` (PBIR-Legacy)
  in `.Report`, plus pointer files. Desktop opens the `.pbip` directly.

Both artifacts are generated from one in-memory model + report, so they
cannot drift apart; `validate_powerbi.py` §J asserts parsed equality.

## 2. `.pbip` layout (as built)

```
Logistics_KPI_Performance_Review.pbip/
├── Logistics_KPI_Performance_Review.pbip        # version "1.0",
│                                               # artifacts: [report only]
├── Logistics_KPI_Performance_Review.SemanticModel/
│   ├── definition.pbism                         # version "4.2"
│   ├── model.bim                                # TMSL JSON (UTF-8)
│   ├── diagramLayout.json                       # version "1.0", node x/y
│   └── .platform                                # type SemanticModel
└── Logistics_KPI_Performance_Review.Report/
    ├── definition.pbir                          # version "1.0" = PBIR-Legacy,
    │                                           # datasetReference.byPath ->
    │                                           # ../<name>.SemanticModel
    ├── report.json                              # PBIR-Legacy (UTF-8, CRLF)
    └── .platform                                # type Report
```

- `definition.pbir` version **"1.0"** is the PBIR-Legacy form: the report
  is a single `report.json` (the legacy Layout JSON). Version "4.0" is
  the enhanced PBIR form (one file per page/visual) and is NOT used.
- `definition.pbism` version **"4.2"**; `.platform` files carry a random
  `logicalId` UUID and the display name.
- `model.bim` is the same TMSL JSON that goes into the `.pbit` as
  `DataModelSchema`.

## 3. `model.bim` (TMSL JSON) — the decisions that matter

- **CompatibilityLevel 1567**, `culture: "en-US"`,
  `defaultPowerBIDataSourceVersion: "powerBI_V3"`.
- **16 tables, 7 relationships, 26 measures, 1 parameter.**
  Tables: `Fact_Shipment`, `Dim_Hub`, `Dim_Route`, `Dim_Vehicle`,
  `Dim_Warehouse`, `Dim_Category`, `Dim_Date`, `Monthly_KPI`,
  `Half_Year_KPI`, `Hub_Destination_KPI`, `Route_KPI`, `Warehouse_KPI`,
  `Category_KPI`, `Cleaning_Log`, `DQ_Missing`, `DQ_Summary`.
- **M partitions** read CSVs from `powerbi/data/` through a `DataFolder`
  text parameter (`IsParameterQueryRequired=true`; empty default, so
  Desktop prompts on first open). Each partition is:
  `Csv.Document(File.Contents(Text.Combine({DataFolder, "<file>.csv"},
  "\")))` → `PromoteHeaders` → `Table.TransformColumnTypes` with
  explicit types and `"en-US"` culture. **No `/home/hatch` path is
  hard-coded** — the parameter is the only path.
- **Types:** `string`/`text`, `int64`/`Int64.Type`, `double`/`number`,
  `dateTime`/`datetime`, `boolean`/`logical`, `date`/`date`.
  One deliberate exception: `Cleaning_Log."Records Affected"` is **text**,
  because the source cleaning log records
  "see profiling (re-sequenced)" for out-of-order scans — typing it
  Int64 would fail Desktop's refresh. (The measured count, 29,287, is in
  `outputs/kpi_reports/data_quality_checks.json`.)
- **Calculated column:** `Fact_Shipment.Half` = `IF(MONTH([date]) <= 6,
  "H1", "H2")`, type string.
- **Relationships:** single direction, `crossFilteringBehavior:
  "oneDirection"`, `joinOnDateBehavior: "dateAndTime"`. Active:
  Fact→Dim_Hub (destination), Fact→Dim_Route, Fact→Dim_Vehicle,
  Fact→Dim_Warehouse, Fact→Dim_Category, Fact→Dim_Date. Inactive:
  Fact→Dim_Hub (origin) via `USERELATIONSHIP` in the origin measures.
  `Dim_Warehouse[hub_id] -> Dim_Hub[hub_id]` is **deliberately NOT
  created**: `Fact_Shipment` already reaches `Dim_Hub` directly, so
  adding it would create two active filter paths and Desktop would
  deactivate one unpredictably. Warehouse visuals reach hub names
  through the fact table.
- **Measures (26, verbatim from `DAX_Measures.md` §4–§5):** all hosted on
  `Fact_Shipment` (the 3 KPI-card measures have `formatString` set:
  `OTD_Pct` `0.00%`, `Cost_per_Shipment` `₹#,##0.00`,
  `Avg_Dwell_Dest` `0.00`). One measure (`OTD_Pct`) uses
  `VAR` names checked against a DAX reserved-function list.
- **RAG thresholds** are the canonical scorecard ones, encoded in the
  `*_RAG` measures: OTD Green ≥85% / Amber 75–85% / Red <75%;
  cost Green ≤₹4,000 / Amber ≤₹4,500 / Red >₹4,500;
  dwell Green ≤12h / Amber ≤18h / Red >18h.
- **Date table:** `Dim_Date` is joined on the date column, but it is NOT
  marked as a date table in `model.bim` (no published TMSL annotation
  for the flag is relied on). Mark it in Desktop (right-click →
  Mark as date table) so `PREVIOUSMONTH` behaves exactly as documented.
  The MoM measures compute regardless via the date relationship.
- **Known gap:** `Monthly_KPI.month` is a text column ("2025-01"), so
  the monthly trend line sorts by `Monthly_KPI.month` text — the
  zero-padded `YYYY-MM` format sorts chronologically as text.

## 4. `report.json` (PBIR-Legacy Layout) — the decisions that matter

- **4 sections** (Executive Overview, Hub & Warehouse Performance,
  Route & Shipment Analysis, Data Quality Monitoring), each
  1280×720, `displayOption: "fitToPage"`. Every section has a title
  textbox, a footer textbox ("Synthetic data (seed 42) …"), and slicers
  for month, hub and category.
- **54 visuals** (18 / 11 / 13 / 12 per page). Visual types used:
  `card` (classic KPI cards — `singleVisual` with a `Values` bucket),
  `cardVisual` (RAG text measures — `singleVisual` with a `Data`
  bucket), `lineChart`, `lineClusteredColumnComboChart` (Y + Y2),
  `clusteredBarChart`, `clusteredColumnChart`, `donutChart`, `tableEx`,
  `scatterChart` (X + Y + Details), `slicer`, `textbox`.
- **The rule that decides whether a visual renders or stays empty:**
  `projections` carry `queryRef` (`"Table.Field"`); `prototypeQuery`
  `Select` entries carry `SourceRef.Source` (the `From` alias), and
  `Select.Name` equals `"Table.Field"`. Swapping those two breaks
  Desktop's binding. The `From` array uses one alias per entity;
  `OrderBy` uses the same `SourceRef` rule. Every visual's refs are
  checked against the model by `validate_powerbi.py` §G.
- **textbox text** lives in `objects.general.paragraphs`
  (`textRuns` with `value`, `textStyle.fontSize`, `bold`,
  `fontColor`); `cardVisual` category labels are hidden via
  `objects.categoryLabels.show: false`.
- **`config` and `filters` are stringified JSON** inside each visual
  container; the container mirrors position in `x`/`y`/`width`/
  `height` and inside `config.layout`.
- **Slicers** are `basicSlicer` with `dataMode: "Basic"`,
  `showAll: true`, `showSelectAll: true`, `orientation: "vertical"`,
  and a shared `filter` context with `mode: "Basic"`.
- **Top N is NOT encoded** in the layout JSON (no published, verified
  shape for it was relied on). Route charts are sorted by the measure
  instead; Top N (worst/best 10 routes, Pareto top 15) is a 30-second
  manual step in the Filters pane — see the guide §4.
- **No theme / resourcePackages** — the default Desktop theme is used.
- **Report config version "5.70"**; `resourcePackages` is an empty list.
- **No `isDefault` flag** is set on any section (Desktop's own
  exports vary on this; the first section is the default by order).

## 5. `.pbit` (OPC ZIP) — the decisions that matter

- Parts (all text parts **UTF-16LE with BOM**, as in real Desktop
  files; `[Content_Types].xml` **UTF-8 with BOM**):
  `Version` ("1.28"), `DataModelSchema` (the `model.bim` JSON),
  `Report/Layout` (the `report.json`), `Metadata`,
  `Settings`, `DiagramLayout`, `DataMashup` (MS-QDEFF stream, below).
  **No `DataModel` part** (a template carries no data) and **no forged
  `SecurityBindings`** — an empty `{"roleBindings": []}` is included by
  the packager, matching what Desktop writes for a new report; Desktop
  re-signs on save.
- Content types: JSON parts `application/json`; `DataMashup`
  `application/octet-stream`; `Version`/`Metadata`/`Settings`/
  `DiagramLayout` `application/json`.
- `DataMashup` is a real **MS-QDEFF** stream: 8-byte header
  (`01 00 00 00 01 00 00 00`) + 5 length-prefixed parts in order:
  **Package Parts** (OPC ZIP containing `[Content_Types].xml`,
  `/Config/Package.xml`, `/Formulas/Section1.m`), **Metadata** (minimal
  well-formed XML), **Metadata Content** (empty), **Permissions**
  (`CanEvaluateFuturePackages=false`, `FirewallEnabled=false`),
  **Permission Bindings** (empty — Desktop then falls back to its
  default permissions; the queries only read local CSVs from one
  folder, so no cross-source privacy conflict arises).
  `Section1.m` declares the `DataFolder` parameter and one `shared`
  query per table, with M identical to the `model.bim` partitions
  (checked by `validate_powerbi.py` §I).
- The `.pbit` was assembled with `nl2pbip`'s `PbitArchiveBuilder`
  (OPC export path only — its `add_pbip_folder` stub path and its LLM
  generation were **not** used). The builder's `ensure_manifest_parts`
  adds the empty `SecurityBindings` noted above.

## 6. What validation covers — and what it cannot

`powerbi/validate_powerbi.py` (519 checks, all passing at build time)
covers: JSON parsing of every file; OPC validity; the `.pbip` pointer
chain; M partitions vs the real CSVs (file exists, `Columns=N`, every
column present, every declared type parses on the full data, including
the 300,000-row fact); measure/column collisions; VAR names vs DAX
reserved functions; sortByColumn targets; RAG thresholds; relationship
endpoints, type matches, dimension-key uniqueness, and no parallel
active paths; every visual field/measure resolving; canvas bounds and
no overlaps; the synthetic-data line on every page; no un-negated YoY
wording; DataMashup framing and Section1.m ↔ partition equality;
`.pbit` ↔ `.pbip` parsed equality.

It **cannot** verify actual rendering or the live refresh inside Power
BI Desktop. If Desktop reports a mashup error on open, use the `.pbip`
folder instead — Desktop reads the M straight from `model.bim`
partitions there.

## 7. Format references

- Power BI Project (PBIP) format — Microsoft Learn, "Power BI Desktop
  projects (PBIP)" (`.pbip` file, `definition.pbir` v1.0 PBIR-Legacy,
  `definition.pbism` v4.2, `.platform` files).
- PBIR-Legacy report.json — documented in the Averianov reference
  (`report.json` shape, `visualContainer`, `singleVisual`, `projections`
  with `queryRef`, `prototypeQuery` with `From`/`Select`/`OrderBy`,
  `SourceRef.Source` in `Select`, `card` vs `cardVisual` buckets,
  textbox `objects.general.paragraphs`, `tableEx`
  `NativeReferenceName`, UTF-16LE text parts in `.pbit`,
  `[Content_Types].xml` UTF-8 BOM, `DataMashup` with 8-byte magic
  header and no `DataModel`/`SecurityBindings` parts).
- MS-QDEFF — Microsoft's "Data Mashup" binary format
  (header + length-prefixed Package Parts / Metadata / Metadata Content
  / Permissions / Permission Bindings; Package Parts ZIP with
  `/Config/Package.xml` + `/Formulas/Section1.m`).
- nl2pbip 2.1.1 — `exporter/pbit_builder.py` + `opc.py`
  (`PbitArchiveBuilder`, `make_metadata_json`, `make_settings_json`,
  `make_diagram_layout_json`, `make_security_bindings_json`); used only
  for its OPC/.pbit assembly helpers.
