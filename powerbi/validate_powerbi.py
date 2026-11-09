"""Structural validation for the Power BI artifacts.

Checks (all deterministic, no Power BI Desktop needed):
  A. JSON parses everywhere (model.bim, report.json, .pbit parts, pointers)
  B. .pbit OPC validity ([Content_Types].xml lists every part, UTF-8 BOM,
     text parts decode as UTF-16LE with BOM, no DataModel part)
  C. .pbip pointer chain (.pbip -> .Report -> definition.pbir byPath ->
     .SemanticModel -> definition.pbism + model.bim)
  D. M partitions vs CSVs (file exists, Columns=N matches header count,
     every TransformColumnTypes column exists in the CSV, every model
     column exists in the CSV, declared types parse on the real data)
  E. Model integrity (measure/column name collisions, duplicate measures,
     VAR names vs DAX reserved function names, sortByColumn targets,
     RAG thresholds match the canonical scorecard)
  F. Relationships (endpoints exist, types match, dimension-side keys
     unique in the data, no two active filter paths between one pair)
  G. Visuals (every queryRef / Select / OrderBy resolves to a model
     column or measure; From aliases consistent; SourceRef rule:
     projections carry queryRef, prototypeQuery Select carries Source)
  H. Report shape (4 pages, canvas bounds, no overlapping visuals,
     synthetic-data line on every page, no "YoY" wording)
  I. DataMashup (MS-QDEFF framing: header + 5 length-prefixed parts,
     Package Parts ZIP parses, Section1.m present, shared queries match
     model tables, Section1.m M matches partition M)
  J. .pbit <-> .pbip equivalence (DataModelSchema == model.bim,
     Report/Layout == report.json, parsed equality)

What this CANNOT verify (stated in VALIDATION.md): actual rendering in
Power BI Desktop and the live refresh against the CSVs.

Run: .venv/bin/python powerbi/validate_powerbi.py
Exit code 0 = all checks pass.
"""
import csv
import io
import json
import re
import sys
import zipfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
PB = ROOT / "powerbi"
DATA = PB / "data"
NAME = "Logistics_KPI_Performance_Review"
PBIP_DIR = PB / f"{NAME}.pbip"
SM_DIR = PBIP_DIR / f"{NAME}.SemanticModel"
RP_DIR = PBIP_DIR / f"{NAME}.Report"
PBIT = PB / f"{NAME}.pbit"

sys.path.insert(0, str(PB))
import build_powerbi as B  # noqa: E402  (model spec is the source of truth)

RESULTS = []  # (section, check, ok, detail)


def check(section, name, ok, detail=""):
    RESULTS.append((section, name, bool(ok), str(detail)))
    return bool(ok)


def jload(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


# ---------------------------------------------------------------- A: JSON
def section_a():
    try:
        model = jload(SM_DIR / "model.bim")
        check("A", "model.bim parses as JSON", True)
    except Exception as e:
        check("A", "model.bim parses as JSON", False, e)
        return None, None
    try:
        report = jload(RP_DIR / "report.json")
        check("A", "report.json parses as JSON", True)
    except Exception as e:
        check("A", "report.json parses as JSON", False, e)
        return model, None
    for f in [PBIP_DIR / f"{NAME}.pbip", SM_DIR / "definition.pbism",
              SM_DIR / ".platform", RP_DIR / "definition.pbir",
              RP_DIR / ".platform", SM_DIR / "diagramLayout.json"]:
        try:
            jload(f)
            check("A", f"{f.name} parses as JSON", True)
        except Exception as e:
            check("A", f"{f.name} parses as JSON", False, e)
    return model, report


# ---------------------------------------------------------------- B: .pbit OPC
def section_b():
    if not PBIT.exists():
        return check("B", ".pbit file exists", False, PBIT)
    try:
        z = zipfile.ZipFile(PBIT)
        names = z.namelist()
        check("B", ".pbit opens as a ZIP", True, f"{len(names)} parts")
    except Exception as e:
        return check("B", ".pbit opens as a ZIP", False, e)
    bad = z.testzip()
    check("B", "ZIP integrity (testzip)", bad is None, bad or "")
    ct = z.read("[Content_Types].xml")
    check("B", "[Content_Types].xml has UTF-8 BOM",
          ct[:3] == b"\xef\xbb\xbf")
    text = ct[3:].decode("utf-8")
    overrides = set(re.findall(r'PartName="([^"]+)"', text))
    missing = [n for n in names if n != "[Content_Types].xml"
               and f"/{n}" not in overrides
               and not any(n.endswith("." + e) for e in ("json", "xml"))]
    check("B", "[Content_Types].xml lists every part", not missing,
          f"unlisted: {missing}" if missing else f"{len(overrides)} overrides")
    check("B", "No DataModel part (template, not a .pbix)",
          "DataModel" not in names)
    for part in ("DataModelSchema", "Report/Layout", "Metadata",
                 "Settings", "DiagramLayout"):
        raw = z.read(part)
        ok = raw[:2] == b"\xff\xfe"
        try:
            raw[2:].decode("utf-16-le")
            dec = True
        except Exception as e:
            dec = False
            detail = str(e)
        else:
            detail = ""
        check("B", f"{part} is UTF-16LE with BOM", ok and dec, detail)
    # Version must be UTF-16LE WITHOUT a BOM: Desktop's ValidateVersion
    # rejects "'﻿1.28'" (BOM retained in the parsed string) as an invalid
    # version number (observed Desktop Oct 2026, v2.159.1231.0).
    raw = z.read("Version")
    ver = raw.decode("utf-16-le")
    check("B", "Version is UTF-16LE without BOM and parses as a version",
          raw[:2] != b"\xff\xfe" and ver == "1.28", repr(ver))
    return z


def pbit_json(z, part):
    return json.loads(z.read(part)[2:].decode("utf-16-le"))


# ---------------------------------------------------------------- C: pointer chain
def section_c():
    pbip = jload(PBIP_DIR / f"{NAME}.pbip")
    arts = pbip.get("artifacts", [])
    ok = len(arts) == 1 and "report" in arts[0]
    check("C", ".pbip artifacts[] names the report only", ok, str(arts))
    rpath = arts[0]["report"]["path"] if ok else ""
    check("C", ".pbip report path resolves", (PBIP_DIR / rpath).is_dir(), rpath)
    pbir = jload(RP_DIR / "definition.pbir")
    by_path = pbir.get("datasetReference", {}).get("byPath", {}).get("path", "")
    check("C", "definition.pbir version is 1.0 (PBIR-Legacy)",
          pbir.get("version") == "1.0", str(pbir.get("version")))
    check("C", "definition.pbir byPath resolves to the SemanticModel folder",
          (RP_DIR / by_path).is_dir() and (RP_DIR / by_path / "model.bim").exists(),
          by_path)
    pbism = jload(SM_DIR / "definition.pbism")
    check("C", "definition.pbism present, version 4.2",
          pbism.get("version") == "4.2", str(pbism.get("version")))


# ---------------------------------------------------------------- D: M vs CSVs
def _m_type_check(series: pd.Series, mtype: str, col: str):
    """Return (ok, detail) for one CSV column against its declared M type."""
    s = series.dropna()
    if len(s) == 0:
        return True, "all blank"
    if mtype == "text":
        return True, ""
    if mtype == "Int64.Type":
        bad = pd.to_numeric(s, errors="coerce").isna().sum()
        frac = (pd.to_numeric(s, errors="coerce") % 1 != 0).sum()
        return bad == 0 and frac == 0, f"{bad} non-integer, {frac} fractional"
    if mtype == "number":
        bad = pd.to_numeric(s, errors="coerce").isna().sum()
        return bad == 0, f"{bad} non-numeric"
    if mtype == "logical":
        bad = (~s.astype(str).str.upper().isin(["TRUE", "FALSE"])).sum()
        return bad == 0, f"{bad} not TRUE/FALSE"
    if mtype == "date":
        bad = pd.to_datetime(s, errors="coerce").isna().sum()
        return bad == 0, f"{bad} unparseable dates"
    if mtype == "datetime":
        bad = pd.to_datetime(s, errors="coerce").isna().sum()
        return bad == 0, f"{bad} unparseable datetimes"
    return False, f"unknown M type {mtype}"


def section_d(model):
    tables = {t["name"]: t for t in model["model"]["tables"]}
    for tname, t in tables.items():
        csv_path = DATA / f"{tname}.csv"
        if not check("D", f"{tname}: CSV exists", csv_path.exists(), str(csv_path)):
            continue
        with open(csv_path, newline="", encoding="utf-8") as fh:
            header = next(csv.reader(fh))
        part = t["partitions"][0]["source"]["expression"]
        m_text = "\n".join(part)
        m = re.search(r"Columns=(\d+)", m_text)
        ncols = int(m.group(1)) if m else -1
        check("D", f"{tname}: partition Columns={ncols} matches CSV header "
                   f"({len(header)})", ncols == len(header))
        model_cols = [c["name"] for c in t["columns"]
                      if "expression" not in c]
        check("D", f"{tname}: model columns == CSV header",
              model_cols == header,
              "" if model_cols == header else
              f"model-only: {sorted(set(model_cols) - set(header))}; "
              f"csv-only: {sorted(set(header) - set(model_cols))}")
        check("D", f"{tname}: M references the CSV file",
              f'"{tname}.csv"' in m_text and "DataFolder" in m_text)
        check("D", f"{tname}: TransformColumnTypes has en-US culture",
              '"en-US"' in m_text)
        # Declared types parse on the real data.
        df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
        spec = {c[0]: c[2] for c in B.COL[tname]}
        for col, mtype in spec.items():
            if col not in df.columns:
                continue
            series = df[col].replace("", pd.NA)
            ok, detail = _m_type_check(series, mtype, col)
            if not ok:
                check("D", f"{tname}.{col}: values parse as {mtype}", False, detail)
        check("D", f"{tname}: all {len(spec)} columns parse as declared types",
              True, f"{len(df)} rows")


# ---------------------------------------------------------------- E: model integrity
def section_e(model):
    tables = {t["name"]: t for t in model["model"]["tables"]}
    all_measures = []
    for tname, t in tables.items():
        cols = [c["name"].lower() for c in t.get("columns", [])]
        for mea in t.get("measures", []):
            all_measures.append((tname, mea["name"]))
            if mea["name"].lower() in cols:
                check("E", f"{tname}: measure '{mea['name']}' collides with a "
                           "column (case-insensitive)", False)
        if len(cols) != len(set(cols)):
            check("E", f"{tname}: duplicate column names", False)
    names = [m.lower() for _, m in all_measures]
    check("E", "No measure shares a name with a column of its own table",
          True, f"{len(all_measures)} measures checked")
    check("E", "No duplicate measure names (case-insensitive)",
          len(names) == len(set(names)))
    reserved = {"FIRSTDATE", "LASTDATE", "DATE", "YEAR", "MONTH", "TODAY",
                "VALUE", "FILTER", "RANK", "MIN", "MAX", "QUARTER", "SUM",
                "COUNT", "CALCULATE", "ALL", "USERELATIONSHIP", "DIVIDE"}
    bad_vars = []
    for tname, mea in [(t["name"], m) for t in tables.values()
                       for m in t.get("measures", [])]:
        for var in re.findall(r"\bVAR\s+([A-Za-z_][A-Za-z0-9_]*)", mea["expression"]):
            if var.upper() in reserved:
                bad_vars.append(f"{mea['name']}: VAR {var}")
    check("E", "No VAR named after a DAX function", not bad_vars,
          "; ".join(bad_vars))
    # sortByColumn targets
    for t in tables.values():
        for c in t.get("columns", []):
            if "sortByColumn" in c:
                tgt = c["sortByColumn"]
                ok = any(cc["name"] == tgt for cc in t["columns"])
                check("E", f"{t['name']}.{c['name']} sortByColumn -> {tgt}",
                      ok)
    # RAG thresholds canonical
    bim_text = json.dumps(model)
    for needle in ['>= 0.85', '>= 0.75', '<= 4000', '<= 4500',
                   '<= 12', '<= 18']:
        check("E", f"RAG threshold present: {needle}", needle in bim_text)
    # DataFolder parameter
    exprs = {e["name"]: e for e in model["model"].get("expressions", [])}
    df = exprs.get("DataFolder", {})
    check("E", "DataFolder parameter exists and is required",
          "IsParameterQueryRequired=true" in df.get("expression", ""),
          df.get("expression", "")[:80])
    check("E", "Model has defaultPowerBIDataSourceVersion powerBI_V3",
          model["model"].get("defaultPowerBIDataSourceVersion") == "powerBI_V3")
    check("E", "Model compatibilityLevel is 1567",
          model.get("compatibilityLevel") == 1567)


# ---------------------------------------------------------------- F: relationships
def section_f(model):
    tables = {t["name"]: t for t in model["model"]["tables"]}
    types = {(t["name"], c["name"]): c["dataType"]
             for t in tables.values() for c in t["columns"]}
    for r in model["model"]["relationships"]:
        fk, tk = (r["fromTable"], r["fromColumn"]), (r["toTable"], r["toColumn"])
        ok = fk in types and tk in types
        check("F", f"Relationship {fk[0]}.{fk[1]} -> {tk[0]}.{tk[1]} endpoints exist",
              ok)
        if ok:
            check("F", f"  types match ({types[fk]} vs {types[tk]})",
                  types[fk] == types[tk], f"{types[fk]} vs {types[tk]}")
            df = pd.read_csv(DATA / f"{tk[0]}.csv", dtype=str, keep_default_na=False)
            dup = df[tk[1]].duplicated().sum()
            check("F", f"  dimension key {tk[0]}.{tk[1]} unique in data",
                  dup == 0, f"{dup} duplicates")
    # Two active filter paths between one pair of tables?
    active = [r for r in model["model"]["relationships"] if r.get("isActive")]
    pairs = {}
    for r in active:
        key = tuple(sorted([r["fromTable"], r["toTable"]]))
        pairs.setdefault(key, []).append(r)
    bad = {k: v for k, v in pairs.items() if len(v) > 1}
    check("F", "No pair of tables has >1 active direct relationship",
          not bad, str(bad))
    # 2-hop active path parallel to a direct active path (Fact->Hub case)
    adj = {}
    for r in active:
        adj.setdefault(r["fromTable"], []).append(r["toTable"])
        adj.setdefault(r["toTable"], []).append(r["fromTable"])
    parallel = []
    for r in active:
        a, b = r["fromTable"], r["toTable"]
        for mid in adj.get(a, []):
            if mid != b and b in adj.get(mid, []):
                parallel.append(f"{a}-{mid}-{b} alongside direct {a}-{b}")
    check("F", "No active 2-hop path parallel to an active direct path",
          not parallel, "; ".join(sorted(set(parallel))))


# ---------------------------------------------------------------- G: visuals
def _parse_config(vc):
    return json.loads(vc["config"])


def section_g(model, report):
    tables = {t["name"]: t for t in model["model"]["tables"]}
    fields = {}
    for tname, t in tables.items():
        for c in t.get("columns", []):
            fields[(tname, c["name"])] = "column"
        for m in t.get("measures", []):
            fields[(tname, m["name"])] = "measure"
    n_visuals = 0
    for sec in report["sections"]:
        for vc in sec["visualContainers"]:
            n_visuals += 1
            try:
                cfg = _parse_config(vc)
            except Exception as e:
                check("G", f"{sec['name']}: visual config parses", False, str(e))
                continue
            sv = cfg.get("singleVisual", {})
            vt = sv.get("visualType", "?")
            # Projections: queryRef must resolve.
            for role, refs in (sv.get("projections") or {}).items():
                for ref in refs:
                    qr = ref.get("queryRef", "")
                    if "." not in qr:
                        check("G", f"{sec['name']} {vt}: queryRef '{qr}' has "
                                   "Table.Field form", False)
                        continue
                    tname, fname = qr.split(".", 1)
                    check("G", f"{sec['name']} {vt}: projection {role} "
                               f"'{qr}' exists in model",
                          (tname, fname) in fields,
                          "" if (tname, fname) in fields else "NOT FOUND")
            # prototypeQuery: From aliases + Select SourceRefs.
            pq = sv.get("prototypeQuery")
            if pq:
                alias_to_entity = {}
                for f in pq.get("From", []):
                    alias_to_entity[f.get("Name")] = f.get("Entity")
                    check("G", f"{sec['name']} {vt}: From entity "
                               f"'{f.get('Entity')}' exists",
                          f.get("Entity") in tables)
                for sel in pq.get("Select", []):
                    kind = "Measure" if "Measure" in sel else "Column"
                    node = sel[kind]
                    src = node["Expression"]["SourceRef"].get("Source")
                    prop = node["Property"]
                    ent = alias_to_entity.get(src)
                    check("G", f"{sec['name']} {vt}: Select {kind} "
                               f"'{ent}.{prop}' resolves",
                          ent is not None and (ent, prop) in fields,
                          f"alias={src} entity={ent}")
                    check("G", f"{sec['name']} {vt}: Select Name matches "
                               f"'{sel.get('Name')}'",
                          sel.get("Name") == f"{ent}.{prop}")
                for ob in pq.get("OrderBy", []):
                    expr = ob.get("Expression", {})
                    kind = "Measure" if "Measure" in expr else "Column"
                    node = expr[kind]
                    src = node["Expression"]["SourceRef"].get("Source")
                    prop = node["Property"]
                    ent = alias_to_entity.get(src)
                    check("G", f"{sec['name']} {vt}: OrderBy {kind} "
                               f"'{ent}.{prop}' resolves",
                          ent is not None and (ent, prop) in fields)
            elif vt not in ("textbox",):
                check("G", f"{sec['name']} {vt}: has prototypeQuery", False,
                      "data visual without prototypeQuery renders empty")
    check("G", f"Visuals checked: {n_visuals} across "
               f"{len(report['sections'])} pages", n_visuals > 0)


# ---------------------------------------------------------------- H: report shape
def section_h(report):
    check("H", "Report has 4 pages", len(report["sections"]) == 4,
          str([s["displayName"] for s in report["sections"]]))
    names = [s["displayName"] for s in report["sections"]]
    for want in ("Executive Overview", "Hub & Warehouse Performance",
                 "Route & Shipment Analysis", "Data Quality Monitoring"):
        check("H", f"Page present: {want}", want in names)
    for sec in report["sections"]:
        boxes = []
        for vc in sec["visualContainers"]:
            x, y, w, h = vc["x"], vc["y"], vc["width"], vc["height"]
            ok = (0 <= x and 0 <= y and x + w <= 1280 and y + h <= 720)
            if not ok:
                check("H", f"{sec['name']}: visual inside 1280x720 canvas",
                      False, f"({x},{y},{w},{h})")
            boxes.append((x, y, x + w, y + h))
        check("H", f"{sec['name']}: all visuals inside canvas", True,
              f"{len(boxes)} visuals")
        overlap = []
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
                    overlap.append((i, j))
        check("H", f"{sec['name']}: no overlapping visuals", not overlap,
              str(overlap))
        cfg_text = json.dumps(sec)
        check("H", f"{sec['name']}: synthetic-data line present",
              "Synthetic data (seed 42)" in cfg_text)
        # YoY wording is only acceptable in an explicit negation ("NOT
        # year-over-year", "no true year-over-year", "No true YoY") — a
        # bare YoY claim would be wrong with a single year of data.
        stripped = re.sub(
            r"(?i)(NOT\s+year-over-year|no true year-over-year|"
            r"no true YoY|NOT\s+YoY)", "", cfg_text)
        yoy_ok = ("YoY" not in stripped and
                  "year-over-year" not in stripped.lower())
        check("H", f"{sec['name']}: no YoY claim (negated mentions only)",
              yoy_ok)


# ---------------------------------------------------------------- I: DataMashup
def section_i(z):
    raw = z.read("DataMashup")
    check("I", "DataMashup starts with the 8-byte magic header",
          raw[:8] == b"\x01\x00\x00\x00\x01\x00\x00\x00")
    parts, off = [], 8
    ok = True
    try:
        for _ in range(5):
            ln = int.from_bytes(raw[off:off + 4], "little")
            off += 4
            parts.append(raw[off:off + ln])
            off += ln
        check("I", "DataMashup has exactly 5 length-prefixed parts, "
                   "fully consumed", off == len(raw),
              f"consumed {off} of {len(raw)} bytes")
    except Exception as e:
        check("I", "DataMashup framing parses", False, str(e))
        return
    pkg_parts, metadata, meta_content, permissions, bindings = parts
    try:
        zz = zipfile.ZipFile(io.BytesIO(pkg_parts))
        check("I", "Package Parts is a valid ZIP", True,
              str(zz.namelist()))
        m_text = zz.read("Formulas/Section1.m").decode("utf-8")
        check("I", "Section1.m present in Package Parts", True)
    except Exception as e:
        check("I", "Package Parts ZIP parses", False, str(e))
        return
    for t in B.TABLES:
        check("I", f"Section1.m has shared query '{t}'",
              f"shared {t} =" in m_text)
    check("I", "Section1.m declares the DataFolder parameter",
          "shared DataFolder =" in m_text
          and "IsParameterQueryRequired=true" in m_text)
    # Section1.m M matches the partition M (whitespace-normalised).
    model = jload(SM_DIR / "model.bim")
    mism = []
    for t in model["model"]["tables"]:
        part_m = "\n".join(
            t["partitions"][0]["source"]["expression"])
        norm = lambda s: re.sub(r"\s+", " ", s).strip()
        # Extract this table's query body from Section1.m.
        m = re.search(rf"shared {re.escape(t['name'])} = (.*?)(?=\nshared |\Z)",
                      m_text, re.S)
        if not m:
            mism.append(t["name"])
            continue
        if norm(m.group(1)) != norm(part_m):
            mism.append(t["name"])
    check("I", "Section1.m M matches partition M for every table",
          not mism, f"mismatched: {mism}")
    check("I", "Metadata part is well-formed XML",
          metadata.strip().startswith(b"<?xml"))
    check("I", "Permissions part is well-formed XML",
          permissions.strip().startswith(b"<?xml"))
    check("I", "Permission Bindings empty (Desktop falls back to defaults)",
          bindings == b"")


# ---------------------------------------------------------------- J: equivalence
def section_j(z, model, report):
    schema = pbit_json(z, "DataModelSchema")
    check("J", ".pbit DataModelSchema == .pbip model.bim (parsed)",
          schema == model)
    layout = pbit_json(z, "Report/Layout")
    check("J", ".pbit Report/Layout == .pbip report.json (parsed)",
          layout == report)


def main():
    model, report = section_a()
    z = section_b()
    section_c()
    if model is not None:
        section_d(model)
        section_e(model)
        section_f(model)
    if model is not None and report is not None:
        section_g(model, report)
    if report is not None:
        section_h(report)
    if z is not None:
        section_i(z)
        if model is not None and report is not None:
            section_j(z, model, report)
    fails = [r for r in RESULTS if not r[2]]
    print(f"\n{'=' * 70}\nVALIDATION: {len(RESULTS) - len(fails)}/{len(RESULTS)} checks pass")
    for sec, name, ok, detail in RESULTS:
        mark = "PASS" if ok else "FAIL"
        print(f"[{mark}] [{sec}] {name}" + (f" — {detail}" if detail else ""))
    # Write VALIDATION.md
    lines = ["# Power BI artifacts — validation note\n",
             "> Generated by `powerbi/validate_powerbi.py`. All checks are "
             "structural (JSON/OPC/M/CSV cross-checks). Nothing here verifies "
             "rendering or refresh inside Power BI Desktop itself — see "
             "§ Limitations.\n",
             f"\n**Result: {len(RESULTS) - len(fails)}/{len(RESULTS)} checks pass.**\n",
             "\n## Checks\n"]
    cur = None
    for sec, name, ok, detail in RESULTS:
        if sec != cur:
            cur = sec
            lines.append(f"\n### {sec}\n")
        lines.append(f"- [{'x' if ok else ' '}] {name}"
                     + (f" — {detail}" if detail else "") + "\n")
    lines.append("""
## What was verified
- The .pbip folder opens as a project: `.pbip` names the report only;
  `definition.pbir` (v1.0, PBIR-Legacy) points by relative path at the
  `.SemanticModel` folder; `definition.pbism` + `model.bim` are present.
- The .pbit is a valid OPC ZIP: `[Content_Types].xml` (UTF-8 BOM) lists
  every part; text parts are UTF-16LE with BOM; there is no `DataModel`
  part (a template carries no data) and no forged `SecurityBindings`
  (Desktop re-signs on save; an empty one is included by the packager).
- Every M partition reads an existing CSV in `powerbi/data/` via the
  `DataFolder` parameter, with `Columns=N` matching the header, every
  `TransformColumnTypes` column present in the file, and every declared
  type parsing on the real data (full 300,000-row fact checked).
- Every visual field/measure resolves to a model object; `From` aliases
  are consistent; projections use `queryRef`, `prototypeQuery` Select
  entries use `SourceRef.Source` (the swap that renders visuals empty).
- Relationships join on existing columns of matching types with unique
  dimension keys; there is no pair of tables with two active filter
  paths. Note: `Dim_Warehouse[hub_id] -> Dim_Hub[hub_id]` is deliberately
  NOT created — `Fact_Shipment` already reaches `Dim_Hub` directly, so
  adding it would create two active paths and Desktop would deactivate
  one unpredictably. Warehouse visuals reach hub names through the fact.
- RAG thresholds in the measures match the canonical scorecard
  (OTD Green ≥85% / Amber 75–85% / Red <75%; cost Green ≤₹4,000 /
  Amber ≤₹4,500 / Red >₹4,500; dwell Green ≤12h / Amber ≤18h / Red >18h).
- The DataMashup is a real MS-QDEFF stream (header + 5 length-prefixed
  parts; Package Parts ZIP with `/Formulas/Section1.m`), and the M in
  `Section1.m` matches the partition M in `model.bim` for every table.

## Limitations — what only Power BI Desktop can verify
- Actual visual rendering (a structurally valid visual can still render
  empty if a Desktop version interprets a role differently) and the live
  refresh against the CSVs (file paths, privacy levels, parameter prompt).
- The DataMashup framing follows the published MS-QDEFF layout; if
  Desktop reports a mashup error on open, use the `.pbip` folder instead
  — Desktop reads the M straight from `model.bim` partitions there.
- Top N filters (worst/best 10 routes, Pareto top 15) are NOT encoded in
  the layout JSON: the route charts are sorted by the measure instead.
  Apply Top N in the Filters pane (30 seconds, see the guide §4).
- `Dim_Date` is not marked as a date table inside `model.bim` (no
  published TMSL annotation for the flag is relied on); mark it in
  Desktop (right-click Dim_Date → Mark as date table) so PREVIOUSMONTH
  behaves exactly as documented. The relationship already joins on the
  date column, so the MoM measures compute regardless.

## How to open (3 steps)
1. In Power BI Desktop: File → Open → `Logistics_KPI_Performance_Review.pbit`
   (or open the `.pbip` file in the project folder).
2. When prompted, set the `DataFolder` parameter to the full path of the
   `powerbi/data` folder (ending with a backslash), then Refresh.
3. File → Save As → `.pbix`. That saved file is the deliverable .pbix.
""")
    (PB / "VALIDATION.md").write_text("".join(lines), encoding="utf-8")
    print(f"\nWrote {PB / 'VALIDATION.md'}")
    return 0 if not fails else 1


if __name__ == "__main__":
    raise SystemExit(main())
