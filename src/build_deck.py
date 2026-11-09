#!/usr/bin/env python3
"""Build the 8-slide Logistics KPI Performance Review deck (python-pptx).

All figures are sourced from:
  outputs/kpi_reports/kpi_summary.json (+ monthly/hub/route/warehouse/category CSVs)
  outputs/validation_report.md (38/38 checks PASS, 0% variance)
  outputs/kpi_reports/data_quality_report.md / cleaning_log.csv
All data is SYNTHETIC (seed 42) for a FICTIONAL Kolkata-based cold-chain
distributor. No true YoY exists (2025 only); H1 vs H2 2025 is the
prior-period comparison.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE = "/home/hatch/workspace/logistics-kpi"
VIZ = os.path.join(BASE, "visualizations")
OUT = os.path.join(BASE, "reports", "Logistics_KPI_Insight_Deck.pptx")

NAVY = RGBColor(0x0F, 0x2A, 0x44)
TEAL = RGBColor(0x0E, 0x7C, 0x7B)
TEAL_DARK = RGBColor(0x0A, 0x5E, 0x5D)
LIGHT = RGBColor(0xED, 0xF2, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x2B, 0x3A)
GREY = RGBColor(0x5A, 0x6B, 0x7D)
GREEN = RGBColor(0x1E, 0x7D, 0x32)
GREEN_BG = RGBColor(0xE6, 0xF4, 0xEA)
AMBER = RGBColor(0xB2, 0x6A, 0x00)
AMBER_BG = RGBColor(0xFD, 0xF0, 0xD5)
RED = RGBColor(0xC6, 0x28, 0x28)
RULE = RGBColor(0xC9, 0xD6, 0xDF)

FONT = "Calibri"
SW, SH = 13.333, 7.5
ML = 0.55  # left margin

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]


def new_slide(title, kicker=None):
    s = prs.slides.add_slide(BLANK)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(0.95))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()
    bar.shadow.inherit = False
    tf = bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(ML)
    tf.margin_right = Inches(0.4)
    tf.margin_top = Inches(0.10)
    tf.margin_bottom = Inches(0.05)
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = FONT
    if kicker:
        p2 = tf.add_paragraph()
        p2.text = kicker
        p2.font.size = Pt(12)
        p2.font.color.rgb = RGBColor(0x9F, 0xD8, 0xD7)
        p2.font.name = FONT
    # teal rule under bar
    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.95),
                              Inches(SW), Pt(4))
    rule.fill.solid()
    rule.fill.fore_color.rgb = TEAL
    rule.line.fill.background()
    rule.shadow.inherit = False
    # footer
    n = len(prs.slides._sldIdLst)
    fb = s.shapes.add_textbox(Inches(ML), Inches(7.12), Inches(11.2), Inches(0.3))
    tfb = fb.text_frame
    tfb.word_wrap = True
    pb = tfb.paragraphs[0]
    pb.text = "Logistics KPI Performance Review  •  SYNTHETIC DATA (seed 42) — fictional company, not real company data"
    pb.font.size = Pt(9)
    pb.font.color.rgb = GREY
    pb.font.name = FONT
    pg = s.shapes.add_textbox(Inches(12.35), Inches(7.12), Inches(0.45), Inches(0.3))
    tpg = pg.text_frame
    tpg.word_wrap = True
    pp = tpg.paragraphs[0]
    pp.text = str(n)
    pp.alignment = PP_ALIGN.RIGHT
    pp.font.size = Pt(9)
    pp.font.color.rgb = GREY
    pp.font.name = FONT
    return s


def textbox(slide, x, y, w, h):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    return box, tf


def bullets(slide, x, y, w, h, items, size=14, space_after=8, color=DARK):
    """items: list of str, or (bold_lead, rest) tuples."""
    _, tf = textbox(slide, x, y, w, h)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 1.05
        if isinstance(item, tuple):
            r1 = p.add_run()
            r1.text = "•  " + item[0]
            r1.font.bold = True
            r1.font.size = Pt(size)
            r1.font.name = FONT
            r1.font.color.rgb = color
            r2 = p.add_run()
            r2.text = item[1]
            r2.font.size = Pt(size)
            r2.font.name = FONT
            r2.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = "•  " + item
            r.font.size = Pt(size)
            r.font.name = FONT
            r.font.color.rgb = color
    return tf


def para(tf, text, size=14, bold=False, color=DARK, space_after=6, first=False,
         align=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = FONT
    r.font.color.rgb = color
    return p


def kpi_card(slide, x, y, w, h, label, value, note, accent=TEAL):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT
    card.line.color.rgb = RULE
    card.line.width = Pt(1)
    card.shadow.inherit = False
    top = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                 Inches(w), Pt(5))
    top.fill.solid()
    top.fill.fore_color.rgb = accent
    top.line.fill.background()
    top.shadow.inherit = False
    tf = card.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.22)
    para(tf, label, size=12, bold=True, color=TEAL_DARK, first=True, space_after=2)
    para(tf, value, size=26, bold=True, color=NAVY, space_after=2)
    para(tf, note, size=11, color=GREY, space_after=0)


def style_table(table, col_widths, header=True, font_size=12.5,
                status_col=None):
    for j, w in enumerate(col_widths):
        table.columns[j].width = Inches(w)
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.12)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            tf = cell.text_frame
            tf.word_wrap = True
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.name = FONT
                    r.font.size = Pt(font_size)
            if i == 0 and header:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                for p in tf.paragraphs:
                    for r in p.runs:
                        r.font.color.rgb = WHITE
                        r.font.bold = True
            elif status_col is not None and j == status_col:
                val = cell.text.strip().upper()
                cell.fill.solid()
                if val == "AMBER":
                    cell.fill.fore_color.rgb = AMBER_BG
                elif val == "GREEN":
                    cell.fill.fore_color.rgb = GREEN_BG
                elif val == "RED":
                    cell.fill.fore_color.rgb = RGBColor(0xFD, 0xE8, 0xE8)
                for p in tf.paragraphs:
                    for r in p.runs:
                        r.font.bold = True
                        r.font.color.rgb = (AMBER if val == "AMBER" else
                                            GREEN if val == "GREEN" else RED)
            elif i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xF7, 0xF9, 0xFB)


def add_image_fit(slide, path, x, y, max_w, max_h):
    """Add picture scaled to fit inside (max_w, max_h) without distortion."""
    from PIL import Image
    iw, ih = Image.open(path).size
    scale = min(max_w / iw, max_h / ih)
    w, h = iw * scale, ih * scale
    cx = x + (max_w - w) / 2
    return slide.shapes.add_picture(path, Inches(cx), Inches(y),
                                    Inches(w), Inches(h))


# ---------------------------------------------------------------- slide 1
s = new_slide("Project Overview and Business Problem",
              "Logistics KPI Performance Review  •  Yuva Intern — Logistics Data Analytics Specialist Internship")
_, tf = textbox(s, ML, 1.35, 12.23, 1.5)
para(tf, "How reliable, how costly — and where is time lost?  A KPI review of a "
         "fictional Kolkata-based cold-chain distributor serving 12 hubs across "
         "East India in 2025.", size=17, color=NAVY, first=True, space_after=10)
para(tf, "Prepared by Khushi Tomar  •  Final-year B.Tech, CSE (AI & ML)",
     size=13, color=GREY, space_after=0)
bullets(s, ML, 3.05, 12.23, 2.6, [
    ("Business problem: ", "delivery reliability is below target — 63,118 of 268,330 "
     "eligible shipments arrived late (OTD 76.48%) — while cost (₹4,102.68 per "
     "shipment) and destination dwell (15.49 h) both sit above their green thresholds."),
    ("Objective: ", "measure on-time delivery, cost per shipment and hub dwell time; "
     "find where lateness, cost and dwell concentrate; and quantify what closing the "
     "gap to the network median would be worth."),
    ("Scope: ", "all shipments in calendar year 2025 (Jan–Dec), 12 destination hubs, "
     "60 routes, 24 warehouses, 6 product categories."),
], size=15, space_after=12)
_, tf2 = textbox(s, ML, 6.05, 12.23, 0.85)
para(tf2, "⚠  ALL DATA IS SYNTHETIC (seed 42), generated for this internship evaluation. "
          "It is NOT real company data and no finding here describes a real business.",
     size=13, bold=True, color=AMBER, first=True, space_after=0)

# ---------------------------------------------------------------- slide 2
s = new_slide("Dataset, Data Architecture and Methodology",
              "Scale  •  Star schema  •  Seven-step method")
bullets(s, ML, 1.25, 6.0, 2.2, [
    ("Scale: ", "1,504,500 raw records — 304,500 shipment rows and 1,200,000 scan "
     "events; 300,000 shipments and 1,194,000 scans after cleaning."),
    ("Coverage: ", "Jan–Dec 2025 only. No true year-over-year exists; H1 vs H2 2025 "
     "is the prior-period comparison used in this review."),
    ("Sources: ", "synthetic shipment + scan-event files (seed 42) and reference "
     "dimensions built alongside them."),
], size=14, space_after=10)
_, tf = textbox(s, ML, 3.55, 6.0, 0.4)
para(tf, "Star schema — trusted analysis table", size=15, bold=True, color=NAVY,
     first=True, space_after=0)
bullets(s, ML, 4.0, 6.0, 2.7, [
    ("Facts: ", "Fact_Shipment — 1 row per shipment (300,000); Fact_Scan_Event — "
     "1 row per scan (1,194,000)."),
    ("Dimensions: ", "Dim_Hub (12), Dim_Route (60 + UNKNOWN), Dim_Vehicle (180), "
     "Dim_Category (5 + Unknown), Dim_Warehouse (24), Dim_Date (2025 calendar)."),
], size=14, space_after=10)
_, tf = textbox(s, 6.95, 1.25, 5.83, 0.4)
para(tf, "Method — 7 steps", size=15, bold=True, color=NAVY, first=True,
     space_after=0)
bullets(s, 6.95, 1.70, 5.83, 4.9, [
    ("1. Generate ", "— synthetic data, seed 42, with known defects injected."),
    ("2. Profile ", "— column profiles and integrity checks on raw data."),
    ("3. Clean ", "— deduplicate, standardise, flag invalid values (slide 3)."),
    ("4. Transform ", "— build the star schema and trusted analysis table."),
    ("5. Calculate KPIs ", "— OTD, cost per shipment, dwell, by month / hub / route / warehouse / category."),
    ("6. Validate ", "— 38 independent recomputation checks: 38/38 PASS, 0% variance."),
    ("7. Visualise ", "— 8 charts; this deck summarises the headline views."),
], size=14, space_after=9)

# ---------------------------------------------------------------- slide 3
s = new_slide("Data Quality Issues and Cleaning Process",
              "What was wrong with the raw data — and what cleaning did about it")
bullets(s, ML, 1.25, 12.23, 3.1, [
    ("Duplicates: ", "9,000 duplicate shipment rows (4,500 extra copies) — removed, "
     "keeping the most complete record; 0 duplicate shipment IDs remain."),
    ("Orphan scans: ", "6,000 scan events with no matching shipment — removed; "
     "referential integrity holds after cleaning."),
    ("Messy hub identifiers: ", "91,486 origin and 92,027 destination hub values in "
     "non-standard spellings — standardised via hub_crosswalk.csv to the 12 canonical hubs."),
    ("Bad timestamps & costs: ", "invalid / missing promised and actual timestamps "
     "excluded from the OTD denominator; negative (2,439), zero (1,156) and missing "
     "(2,920) costs excluded from the cost KPI — 293,485 valid-cost shipments remain."),
    ("Scan order & dwell: ", "29,287 out-of-order scan events in 20,914 shipments "
     "re-sequenced by UTC; negative dwell pairs flagged and excluded (2,885 destination, 18 origin)."),
], size=14, space_after=9)
_, tf = textbox(s, ML, 4.75, 12.23, 0.4)
para(tf, "Cleaning discipline", size=15, bold=True, color=NAVY, first=True,
     space_after=0)
bullets(s, ML, 5.2, 12.23, 1.6, [
    ("19 issue types logged, ", "200,528 changed records logged — every change is traceable; "
     "nothing was silently dropped or imputed (missing category / vehicle kept as “Unknown”)."),
    ("Result: ", "300,000 clean shipments, 1,194,000 clean scans; validation 38/38 PASS, 0% variance."),
], size=14, space_after=8)

# ---------------------------------------------------------------- slide 4
s = new_slide("Executive KPI Scorecard",
              "2025 full year  •  All three headline KPIs are AMBER — below green, above red")
kpi_card(s, ML, 1.25, 2.92, 1.55, "ON-TIME DELIVERY", "76.48%",
         "268,330 eligible · 63,118 late", accent=AMBER)
kpi_card(s, 3.72, 1.25, 2.92, 1.55, "COST PER SHIPMENT", "₹4,102.68",
         "₹1,204,074,834 total valid cost", accent=AMBER)
kpi_card(s, 6.89, 1.25, 2.92, 1.55, "AVG DESTINATION DWELL", "15.49 h",
         "Origin dwell 2.5 h · combined 8.9 h", accent=AMBER)
kpi_card(s, 10.06, 1.25, 2.72, 1.55, "DATA QUALITY", "GREEN",
         "38/38 validation checks passed", accent=GREEN)
rows, cols = 5, 4
tbl_shape = s.shapes.add_table(rows, cols, Inches(ML), Inches(3.05),
                               Inches(12.23), Inches(1.95))
table = tbl_shape.table
headers = ["KPI", "2025 result", "Thresholds", "Status"]
data = [
    ("On-time delivery (OTD)", "76.48%",
     "GREEN ≥ 85%  •  AMBER 75–85%  •  RED < 75%", "AMBER"),
    ("Cost per shipment", "₹4,102.68",
     "GREEN ≤ ₹4,000  •  AMBER ₹4,000–4,500  •  RED > ₹4,500", "AMBER"),
    ("Avg destination dwell", "15.49 h",
     "GREEN ≤ 12 h  •  AMBER 12–18 h  •  RED > 18 h", "AMBER"),
    ("Data quality", "38/38 checks passed, 0% variance",
     "GREEN: 100% changed-record logging, all checks pass", "GREEN"),
]
for j, h in enumerate(headers):
    table.cell(0, j).text = h
for i, row in enumerate(data, start=1):
    for j, val in enumerate(row):
        table.cell(i, j).text = val
style_table(table, [3.1, 3.0, 4.63, 1.5], font_size=13, status_col=3)
bullets(s, ML, 5.25, 12.23, 1.7, [
    ("OTD 76.48% — AMBER: ", "85% is the service level a reliable cold chain should promise; the network is 8.5 pp short of it."),
    ("Cost ₹4,102.68 — AMBER: ", "₹4,000 is the cost anchor; the network runs ₹102.68 above it per shipment."),
    ("Dwell 15.49 h — AMBER: ", "12 h is the same-day unload target; destination dwell runs 3.5 h over it, while origin dwell (2.5 h) is not the problem."),
], size=13.5, space_after=6)

# ---------------------------------------------------------------- slide 5
s = new_slide("Monthly Delivery Performance",
              "OTD by month, Jan–Dec 2025  •  No true YoY — 2025 only")
add_image_fit(s, os.path.join(VIZ, "01_monthly_otd_trend.png"),
              ML, 1.30, 8.1, 4.55)
_, tf = textbox(s, ML, 5.95, 8.1, 0.9)
para(tf, "Chart 01 — Monthly OTD trend (synthetic data, seed 42). Monthly OTD "
         "stays in a narrow 75.7–77.0% band all year.",
     size=11, color=GREY, first=True, space_after=0)
bullets(s, 8.95, 1.30, 3.83, 4.6, [
    ("Flat all year: ", "monthly OTD ranges only 75.72% (Apr) to 77.04% (May) — no seasonal peak or trough."),
    ("No improvement: ", "H1 2025 OTD 76.49% vs H2 2025 76.47% — the prior-period comparison is essentially unchanged."),
    ("Cost and dwell flat too: ", "monthly cost ₹4,077–₹4,127 and destination dwell 15.42–15.59 h — the network is stable, not deteriorating."),
], size=14, space_after=12)

# ---------------------------------------------------------------- slide 6
s = new_slide("Hub, Warehouse and Route Analysis",
              "Lateness is concentrated — in three hubs, a few routes, and one category")
add_image_fit(s, os.path.join(VIZ, "02_hub_otd_comparison.png"),
              ML, 1.25, 6.35, 3.55)
_, tf = textbox(s, ML, 4.90, 6.35, 0.4)
para(tf, "Chart 02 — OTD by destination hub", size=11, color=GREY, first=True,
     space_after=0)
add_image_fit(s, os.path.join(VIZ, "04_route_performance_ranking.png"),
              7.15, 1.25, 5.63, 3.55)
_, tf = textbox(s, 7.15, 4.90, 5.63, 0.4)
para(tf, "Chart 04 — Route performance ranking (60 routes + UNKNOWN)",
     size=11, color=GREY, first=True, space_after=0)
bullets(s, ML, 5.35, 12.23, 1.6, [
    ("Worst hubs: ", "Guwahati 53.95% (10,364 late — 16.4% of all late deliveries), Malda 61.89%, Ranchi 69.04%; best is Durgapur 83.89%. Nine hubs sit at 83–84%."),
    ("Worst routes: ", "RT-006 57.11% (4,182 late), RT-017 60.20%, RT-028 63.67% — vs best route RT-003 83.41%. Bottom-quartile routes (OTD ≤ 73.69%) hold 69,332 eligible shipments at 66.07% OTD."),
    ("Frozen is the weak category: ", "72.06% OTD (lowest) and 33.1% of all late deliveries; Pharma is best at 84.59%. Warehouse dwell is narrow (15.21–15.95 h) — dwell is network-wide, not one bad warehouse."),
], size=13.5, space_after=8)

# ---------------------------------------------------------------- slide 7
s = new_slide("Business Impact and Key Findings",
              "What closing the gap to the network median would be worth — modelled estimates on synthetic data")
rows, cols = 4, 5
tbl_shape = s.shapes.add_table(rows, cols, Inches(ML), Inches(1.30),
                               Inches(12.23), Inches(1.85))
table = tbl_shape.table
headers = ["Scenario", "Gap closed", "Late deliveries avoided",
           "Cost saving", "Excess dwell recovered"]
data = [
    ("Conservative", "50%", "4,853", "₹15,56,916", "2,147.9 h"),
    ("Realistic", "100% to median", "9,706", "₹31,13,833", "4,295.8 h"),
    ("Optimistic", "100% + 10% gain", "9,706", "₹34,25,216", "4,295.8 h"),
]
for j, h in enumerate(headers):
    table.cell(0, j).text = h
for i, row in enumerate(data, start=1):
    for j, val in enumerate(row):
        table.cell(i, j).text = val
style_table(table, [2.3, 2.2, 2.73, 2.5, 2.5], font_size=13)
_, tf = textbox(s, ML, 3.35, 12.23, 0.4)
para(tf, "Key findings", size=15, bold=True, color=NAVY, first=True, space_after=0)
bullets(s, ML, 3.80, 12.23, 2.2, [
    ("Concentrated, not diffuse: ", "39 of 61 routes account for 80% of late deliveries; Guwahati, Malda and Ranchi together account for 41.8% of all late deliveries."),
    ("Realistic case is the planning case: ", "9,706 late deliveries avoided, ₹31,13,833 saved and 4,295.8 excess dwell hours recovered if bottom-quartile routes, costs and warehouses converge to the network median."),
    ("Read the numbers honestly: ", "these are modelled estimates on synthetic data, not a forecast of real savings — no causal mechanism, capacity constraint, customer-behaviour change or implementation cost is included."),
], size=14, space_after=9)
_, tf = textbox(s, ML, 6.35, 12.23, 0.6)
para(tf, "Basis: bottom-quartile routes (15 of 60, median route OTD 80.06%), bottom-quartile route costs (median ₹4,099.68) and bottom-quartile warehouses (7, median dwell 15.51 h).",
     size=11, color=GREY, first=True, space_after=0)

# ---------------------------------------------------------------- slide 8
s = new_slide("Three Prioritised Recommendations and Conclusion",
              "Fix the concentrated gap first — then hold the network at the median")
recs = [
    ("1", "Guwahati / Malda / Ranchi corridor recovery",
     "Dedicated route planning + hub capacity review for the three worst destination hubs (OTD 53.95% / 61.89% / 69.04%).",
     "Owner: Operations", "Monitor: destination-hub OTD"),
    ("2", "Destination-hub dwell reduction",
     "Unloading-slot scheduling and labour planning at bottom-quartile warehouses — dwell is 15.49 h vs a 12 h green target.",
     "Owner: Warehouse Operations", "Monitor: avg destination dwell"),
    ("3", "Frozen-category cold-chain reliability + cost control on bottom-quartile routes",
     "Reefer maintenance checks and route cost review — Frozen is the weakest category (72.06% OTD, 33.1% of late deliveries).",
     "Owner: Fleet & Category Management", "Monitor: category OTD and route cost / shipment"),
]
y = 1.30
for num, title, action, owner, monitor in recs:
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(ML),
                              Inches(y), Inches(12.23), Inches(1.35))
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT
    card.line.color.rgb = RULE
    card.line.width = Pt(1)
    card.shadow.inherit = False
    numbox = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ML + 0.18),
                                Inches(y + 0.18), Inches(0.62), Inches(0.62))
    numbox.fill.solid()
    numbox.fill.fore_color.rgb = TEAL
    numbox.line.fill.background()
    numbox.shadow.inherit = False
    ntf = numbox.text_frame
    ntf.word_wrap = True
    np_ = ntf.paragraphs[0]
    np_.alignment = PP_ALIGN.CENTER
    nr = np_.add_run()
    nr.text = num
    nr.font.size = Pt(22)
    nr.font.bold = True
    nr.font.color.rgb = WHITE
    nr.font.name = FONT
    _, tf = textbox(s, ML + 1.05, y + 0.12, 11.0, 1.15)
    para(tf, title, size=15, bold=True, color=NAVY, first=True, space_after=2)
    para(tf, action, size=13, color=DARK, space_after=2)
    para(tf, owner + "   •   " + monitor, size=12, bold=True, color=TEAL_DARK,
         space_after=0)
    y += 1.55
_, tf = textbox(s, ML, 6.10, 12.23, 0.8)
para(tf, "Conclusion: the network is stable, not deteriorating — the gap is concentrated "
         "in three hubs, fifteen routes and one category, and it is measurable and "
         "closable to the median. Start with the Guwahati / Malda / Ranchi corridor.",
     size=14, bold=True, color=NAVY, first=True, space_after=0)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
prs.save(OUT)
print("Saved:", OUT)
print("Slides:", len(prs.slides.__iter__.__self__._sldIdLst))
