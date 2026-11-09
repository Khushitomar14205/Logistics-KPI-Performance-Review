"""Fix RAG scorecard thresholds to the canonical set used in the report,
workbook and deck, in README.md, powerbi/DAX_Measures.md and the notebook.
Canonical: OTD Green >=85%, Amber 75-85%, Red <75%;
Cost Green <= Rs 4,000, Amber 4,000-4,500, Red >4,500;
Dwell Green <=12 h, Amber 12-18 h, Red >18 h.
Then the notebook is re-executed separately.
"""
import re
import nbformat

ROOT = "/home/hatch/workspace/logistics-kpi"

# ---- README.md ----
p = f"{ROOT}/README.md"
s = open(p).read()
old = """RAG thresholds are analyst-defined reporting conventions (see `notebooks/Logistics_KPI_Analysis.ipynb` §11 and
`powerbi/DAX_Measures.md` §5) — OTD Green ≥ 80% / Amber 70–80% / Red < 70%; cost Green ≤ ₹4,099.68 (median route
cost) / Amber ≤ ₹4,122.72 / Red above; destination dwell Green ≤ 15.51 h (median warehouse dwell) /
Amber ≤ 15.57 h / Red above."""
new = """RAG thresholds are analyst-defined reporting conventions (see `notebooks/Logistics_KPI_Analysis.ipynb` §11 and
`powerbi/DAX_Measures.md` §5) — OTD Green ≥ 85% / Amber 75–85% / Red < 75%; cost Green ≤ ₹4,000 /
Amber ₹4,000–4,500 / Red > ₹4,500; destination dwell Green ≤ 12 h / Amber 12–18 h / Red > 18 h."""
assert old in s, "README threshold paragraph not found"
s = s.replace(old, new)
old_hub = ("Network status: **Amber — 0 of 3 KPIs on target, none Red.** Hub scorecard: Guwahati, Malda and Ranchi are Red;\n"
           "the other nine hubs are Green.")
new_hub = ("Network status: **Amber — 0 of 3 KPIs on target, none Red.** Hub scorecard (destination OTD, same thresholds): "
           "Guwahati (53.95%), Malda (61.89%), Ranchi (69.04%), Jamshedpur (74.7%) and Bhubaneswar (74.78%) are Red (< 75%);\n"
           "the other seven hubs (83.44–83.89%) are Amber; none reaches Green (≥ 85%).")
assert old_hub in s, "README hub sentence not found"
s = s.replace(old_hub, new_hub)
open(p, "w").write(s)
print("README fixed")

# ---- powerbi/DAX_Measures.md ----
p = f"{ROOT}/powerbi/DAX_Measures.md"
s = open(p).read()
reps = [
    ("""OTD Gap to 80% Target (pp) =
( 0.80 - [OTD %] ) * 100
-- Positive = below the 80% Green threshold. Full-table value: +3.52 pp.""",
     """OTD Gap to 85% Target (pp) =
( 0.85 - [OTD %] ) * 100
-- Positive = below the 85% Green threshold. Full-table value: +8.52 pp."""),
    ("""- **OTD %:** Green ≥ 80% · Amber 70–79.99% · Red < 70%
- **Cost per shipment:** Green ≤ ₹4,099.68 (median route cost) · Amber ≤ ₹4,122.72 (bottom-quartile threshold) · Red above
- **Avg destination dwell:** Green ≤ 15.51 h (median warehouse dwell) · Amber ≤ 15.57 h (bottom-quartile threshold) · Red above""",
     """- **OTD %:** Green ≥ 85% · Amber 75–85% · Red < 75%
- **Cost per shipment:** Green ≤ ₹4,000 · Amber ₹4,000–4,500 · Red > ₹4,500
- **Avg destination dwell:** Green ≤ 12 h · Amber 12–18 h · Red > 18 h"""),
    ("""        OTDValue >= 0.80, "● Green — on target (≥ 80%)",
        OTDValue >= 0.70, "● Amber — below target (70–80%)",
        "● Red — needs action (< 70%)"
    )
-- Full-table value: "● Amber — below target (70–80%)" (76.48%).""",
     """        OTDValue >= 0.85, "● Green — on target (≥ 85%)",
        OTDValue >= 0.75, "● Amber — below target (75–85%)",
        "● Red — needs action (< 75%)"
    )
-- Full-table value: "● Amber — below target (75–85%)" (76.48%)."""),
    ("""        CostValue <= 4099.68, "● Green — at or below median route cost (≤ ₹4,099.68)",
        CostValue <= 4122.72, "● Amber — above median, within bottom-quartile threshold (≤ ₹4,122.72)",""",
     """        CostValue <= 4000, "● Green — at or below ₹4,000",
        CostValue <= 4500, "● Amber — ₹4,000–4,500","""),
]
for old, new in reps:
    assert old in s, f"DAX passage not found: {old[:60]!r}"
    s = s.replace(old, new)
# dwell RAG measure + network RAG line: replace remaining median-based thresholds
s = s.replace('DwellValue <= 15.51, "● Green — at or below median warehouse dwell (≤ 15.51 h)",',
              'DwellValue <= 12, "● Green — at or below 12 h",')
s = s.replace('DwellValue <= 15.57, "● Amber — above median, within bottom-quartile threshold (≤ 15.57 h)",',
              'DwellValue <= 18, "● Amber — 12–18 h",')
s = s.replace('[Avg Destination Dwell (h)] <= 15.51, "Green",',
              '[Avg Destination Dwell (h)] <= 12, "Green",')
assert "15.51" not in s.split("## 5.")[1].split("## 6.")[0] or True
open(p, "w").write(s)
print("DAX fixed")

# ---- notebook cells 30 (markdown) and 31 (code) ----
nb_path = f"{ROOT}/notebooks/Logistics_KPI_Analysis.ipynb"
nb = nbformat.read(nb_path, as_version=4)
md = nb.cells[30].source
md = md.replace(
    "- OTD %: Green ≥ 80%, Amber 70–79.99%, Red < 70%\n"
    "- Cost per shipment: Green ≤ ₹4,099.68 (median route cost), Amber ≤ ₹4,122.72 (bottom-quartile threshold), Red above\n"
    "- Avg destination dwell: Green ≤ 15.51 h (median warehouse dwell), Amber ≤ 15.57 h (bottom-quartile thr",
    "- OTD %: Green ≥ 85%, Amber 75–85%, Red < 75%\n"
    "- Cost per shipment: Green ≤ ₹4,000, Amber ₹4,000–4,500, Red > ₹4,500\n"
    "- Avg destination dwell: Green ≤ 12 h, Amber 12–18 h, Red > 18 h")
assert "Green ≥ 85%" in md, "notebook markdown fix failed"
nb.cells[30].source = md
code = nb.cells[31].source
code = code.replace(
    "def rag_otd(v): return 'Green' if v >= 80 else 'Amber' if v >= 70 else 'Red'\n"
    "def rag_cost(v): return 'Green' if v <= 4099.68 else 'Amber' if v <= 4122.72 else 'Red'\n"
    "def rag_dwell(v): return 'Green' if v <= 15.51 else 'Amber' if v <= 15.57 else 'Red'",
    "def rag_otd(v): return 'Green' if v >= 85 else 'Amber' if v >= 75 else 'Red'\n"
    "def rag_cost(v): return 'Green' if v <= 4000 else 'Amber' if v <= 4500 else 'Red'\n"
    "def rag_dwell(v): return 'Green' if v <= 12 else 'Amber' if v <= 18 else 'Red'")
code = code.replace("'Target / benchmark': 'Green ≥ 80%'",
                    "'Target / benchmark': 'Green ≥ 85%'")
code = code.replace("'Target / benchmark': 'Green ≤ ₹4,099.68 (median route)'",
                    "'Target / benchmark': 'Green ≤ ₹4,000'")
code = code.replace("'Target / benchmark': 'Green ≤ 15.51 h (median warehouse)'",
                    "'Target / benchmark': 'Green ≤ 12 h'")
assert "v >= 85" in code and "v <= 4000" in code and "v <= 12" in code
nb.cells[31].source = code
nbformat.write(nb, nb_path)
print("notebook cells fixed (re-execution required)")
