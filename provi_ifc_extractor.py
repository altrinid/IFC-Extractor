#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ProVI IFC Extractor
Extracts civil / road-engineering data from ProVI IFC4X3 exports:

  - Alignments (horizontal, vertical and cant segments)
  - Roads and road parts
  - Pavements and pavement courses
  - Bridges and bridge parts
  - All remaining physical elements (kerbs, earthworks, signs, …)

Each category is written to its own sheet of an Excel workbook, or to a
separate CSV file.

Usage:
  python provi_ifc_extractor.py model.ifc
  python provi_ifc_extractor.py model.ifc -o output.xlsx
  python provi_ifc_extractor.py model.ifc --csv-dir ./output
"""

import argparse
import csv
import os
from typing import Any, Dict, List, Optional, Set, Tuple

try:
    import ifcopenshell
    import ifcopenshell.util.element
except ImportError:
    raise SystemExit("Please install ifcopenshell:  pip install ifcopenshell")

try:
    import pandas as pd
    import openpyxl  # noqa: F401  (engine for pandas .xlsx output)
    PANDAS = True
except ImportError:
    PANDAS = False


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def normalize(val) -> str:
    """Render any IFC attribute value as a plain string."""
    if val is None:
        return ""
    if hasattr(val, "wrappedValue"):
        return str(val.wrappedValue)
    return str(val)


def get_name(entity) -> str:
    for attr in ("Name", "GlobalId", "Tag"):
        v = getattr(entity, attr, None)
        if v:
            return str(v)
    return f"{entity.is_a()}_{entity.id()}"


def get_psets(entity) -> Dict[str, str]:
    """
    Flat dict of every property/quantity, keyed as 'PsetName:PropName'.
    Uses ifcopenshell's util with a hand-rolled fallback.
    """
    out: Dict[str, str] = {}
    try:
        psets = ifcopenshell.util.element.get_psets(entity, include_inherited=True)
        for grp, vals in psets.items():
            if isinstance(vals, dict):
                for k, v in vals.items():
                    if k == "id":          # util adds the STEP id of the pset – skip it
                        continue
                    out[f"{grp}:{k}"] = normalize(v)
            else:
                out[grp] = normalize(vals)
        return out
    except Exception:
        pass

    # Fallback: walk IsDefinedBy → IfcRelDefinesByProperties
    try:
        for rel in getattr(entity, "IsDefinedBy", []) or []:
            if not rel.is_a("IfcRelDefinesByProperties"):
                continue
            props = getattr(rel, "RelatingPropertyDefinition", None)
            if not props:
                continue
            if props.is_a("IfcPropertySet"):
                for p in props.HasProperties or []:
                    out[f"{props.Name}:{p.Name}"] = normalize(getattr(p, "NominalValue", None))
            elif props.is_a("IfcElementQuantity"):
                for q in props.Quantities or []:
                    for field in ("LengthValue", "AreaValue", "VolumeValue",
                                  "CountValue", "WeightValue", "TimeValue"):
                        if getattr(q, field, None) is not None:
                            out[f"{props.Name}:{q.Name}"] = normalize(getattr(q, field))
                            break
    except Exception:
        pass
    return out


def decomposed_children(entity) -> List[Any]:
    """
    Direct children via *either* IfcRelNests or IfcRelAggregates.

    IFC4X3 relates an IfcAlignment to its horizontal/vertical/cant layouts
    (and each layout to its segments) with IfcRelNests, but some exporters
    use aggregation, so both inverses are honoured.
    """
    children: List[Any] = []
    for rel in getattr(entity, "IsNestedBy", []) or []:
        children.extend(rel.RelatedObjects or [])
    for rel in getattr(entity, "IsDecomposedBy", []) or []:
        children.extend(rel.RelatedObjects or [])
    return children


def _material_name(mat) -> str:
    if mat is None:
        return ""
    if mat.is_a("IfcMaterial"):
        return mat.Name or ""
    if mat.is_a("IfcMaterialLayerSetUsage"):
        ls = mat.ForLayerSet
        return (ls.LayerSetName or "") if ls else ""
    if mat.is_a("IfcMaterialLayerSet"):
        return mat.LayerSetName or ""
    if mat.is_a("IfcMaterialLayer"):
        return (mat.Material.Name or "") if mat.Material else ""
    if mat.is_a("IfcMaterialProfileSetUsage"):
        ps = mat.ForProfileSet
        return (ps.Name or "") if ps else ""
    if mat.is_a("IfcMaterialProfileSet"):
        return mat.Name or ""
    if mat.is_a("IfcMaterialConstituentSet"):
        return mat.Name or ""
    if mat.is_a("IfcMaterialList"):
        return ", ".join((m.Name or "") for m in (mat.Materials or []))
    return getattr(mat, "Name", "") or ""


def get_material(entity) -> str:
    names = []
    for rel in getattr(entity, "HasAssociations", []) or []:
        if rel.is_a("IfcRelAssociatesMaterial"):
            n = _material_name(rel.RelatingMaterial)
            if n:
                names.append(n)
    return ", ".join(names)


def _point_xy(point) -> Tuple[str, str]:
    if point is None:
        return "", ""
    coords = getattr(point, "Coordinates", None)
    if coords and len(coords) >= 2:
        return normalize(coords[0]), normalize(coords[1])
    return "", ""


def base_row(entity, extra: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    row = {
        "GlobalId":       normalize(getattr(entity, "GlobalId", "")),
        "Entity":         entity.is_a(),
        "Name":           get_name(entity),
        "Description":    normalize(getattr(entity, "Description", "")),
        "ObjectType":     normalize(getattr(entity, "ObjectType", "")),
        "PredefinedType": normalize(getattr(entity, "PredefinedType", "")),
    }
    if extra:
        row.update(extra)
    row.update(get_psets(entity))
    return row


# ---------------------------------------------------------------------------
# Alignment extraction
# ---------------------------------------------------------------------------

def _segments(layout) -> List[Any]:
    return [c for c in decomposed_children(layout) if c.is_a("IfcAlignmentSegment")]


def extract_horizontal(alignment_name: str, horiz) -> List[Dict]:
    rows, dist = [], 0.0
    for seg in _segments(horiz):
        d = seg.DesignParameters
        if not d:
            continue
        x, y = _point_xy(getattr(d, "StartPoint", None))
        length = getattr(d, "SegmentLength", None)
        rows.append({
            "AlignmentName":          alignment_name,
            "SegmentId":              normalize(seg.GlobalId),
            "SegmentName":            normalize(seg.Name),
            "PredefinedType":         normalize(getattr(d, "PredefinedType", "")),
            "StartDistAlong":         f"{dist:.6f}",
            "SegmentLength":          normalize(length),
            "StartPointX":            x,
            "StartPointY":            y,
            "StartDirection":         normalize(getattr(d, "StartDirection", "")),
            "StartRadiusOfCurvature": normalize(getattr(d, "StartRadiusOfCurvature", "")),
            "EndRadiusOfCurvature":   normalize(getattr(d, "EndRadiusOfCurvature", "")),
            "GravityCenterLineHeight": normalize(getattr(d, "GravityCenterLineHeight", "")),
        })
        if length is not None:
            try:
                dist += float(length)
            except (TypeError, ValueError):
                pass
    return rows


def extract_vertical(alignment_name: str, vert) -> List[Dict]:
    rows = []
    for seg in _segments(vert):
        d = seg.DesignParameters
        if not d:
            continue
        rows.append({
            "AlignmentName":     alignment_name,
            "SegmentId":         normalize(seg.GlobalId),
            "SegmentName":       normalize(seg.Name),
            "PredefinedType":    normalize(getattr(d, "PredefinedType", "")),
            "StartDistAlong":    normalize(getattr(d, "StartDistAlong", "")),
            "HorizontalLength":  normalize(getattr(d, "HorizontalLength", "")),
            "StartHeight":       normalize(getattr(d, "StartHeight", "")),
            "StartGradient":     normalize(getattr(d, "StartGradient", "")),
            "EndGradient":       normalize(getattr(d, "EndGradient", "")),
            "RadiusOfCurvature": normalize(getattr(d, "RadiusOfCurvature", "")),
        })
    return rows


def extract_cant(alignment_name: str, cant) -> List[Dict]:
    rows = []
    rail_head = normalize(getattr(cant, "RailHeadDistance", ""))
    for seg in _segments(cant):
        d = seg.DesignParameters
        if not d:
            continue
        rows.append({
            "AlignmentName":    alignment_name,
            "SegmentId":        normalize(seg.GlobalId),
            "SegmentName":      normalize(seg.Name),
            "PredefinedType":   normalize(getattr(d, "PredefinedType", "")),
            "StartDistAlong":   normalize(getattr(d, "StartDistAlong", "")),
            "HorizontalLength": normalize(getattr(d, "HorizontalLength", "")),
            "StartCantLeft":    normalize(getattr(d, "StartCantLeft", "")),
            "EndCantLeft":      normalize(getattr(d, "EndCantLeft", "")),
            "StartCantRight":   normalize(getattr(d, "StartCantRight", "")),
            "EndCantRight":     normalize(getattr(d, "EndCantRight", "")),
            "RailHeadDistance": rail_head,
        })
    return rows


def extract_alignments(model):
    """Returns (summary, horizontal, vertical, cant) row lists."""
    summary, horiz_rows, vert_rows, cant_rows = [], [], [], []
    try:
        alignments = model.by_type("IfcAlignment")
    except Exception:
        return summary, horiz_rows, vert_rows, cant_rows

    for alignment in alignments:
        name = get_name(alignment)
        summary.append(base_row(alignment))
        for child in decomposed_children(alignment):
            if child.is_a("IfcAlignmentHorizontal"):
                horiz_rows.extend(extract_horizontal(name, child))
            elif child.is_a("IfcAlignmentVertical"):
                vert_rows.extend(extract_vertical(name, child))
            elif child.is_a("IfcAlignmentCant"):
                cant_rows.extend(extract_cant(name, child))
    return summary, horiz_rows, vert_rows, cant_rows


# ---------------------------------------------------------------------------
# Spatial / physical element extraction
# ---------------------------------------------------------------------------

def _collect(model, classes: List[str], seen: Set[str],
             with_material: bool = False) -> List[Dict]:
    rows = []
    for ifc_class in classes:
        try:
            entities = model.by_type(ifc_class)
        except Exception:
            continue  # class not present in this schema
        for ent in entities:
            gid = getattr(ent, "GlobalId", None)
            if gid in seen:
                continue
            seen.add(gid)
            extra = {"Material": get_material(ent)} if with_material else None
            rows.append(base_row(ent, extra))
    return rows


def extract_roads(model, seen):
    return _collect(model, ["IfcRoad", "IfcRoadPart"], seen)


def extract_bridges(model, seen):
    return _collect(model, ["IfcBridge", "IfcBridgePart"], seen)


def extract_pavements(model, seen):
    return _collect(model, ["IfcPavement", "IfcCourse"], seen, with_material=True)


def extract_remaining(model, seen):
    """All remaining physical elements not already captured elsewhere."""
    return _collect(model, ["IfcElement"], seen, with_material=True)


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def _ordered_keys(rows: List[Dict]) -> List[str]:
    keys: List[str] = []
    for r in rows:
        for k in r:
            if k not in keys:
                keys.append(k)
    return keys


def _write_csv(rows: List[Dict], path: str, label: str):
    if not rows:
        print(f"[SKIP] {label}: 0 rows")
        return
    keys = _ordered_keys(rows)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in keys})
    print(f"[OK] {label}: {len(rows)} rows → {path}")


def write_csv_dir(sheets: Dict[str, List[Dict]], directory: str):
    os.makedirs(directory, exist_ok=True)
    for name, rows in sheets.items():
        _write_csv(rows, os.path.join(directory, f"{name}.csv"), name)
    print(f"\nSaved CSVs → {directory}")


def write_xlsx(sheets: Dict[str, List[Dict]], path: str):
    if not PANDAS:
        raise SystemExit("pandas + openpyxl required for Excel output.  "
                         "pip install pandas openpyxl  (or use --csv-dir)")
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    wrote_any = False
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, rows in sheets.items():
            if not rows:
                continue
            df = pd.DataFrame(rows, columns=_ordered_keys(rows))
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
            print(f"[OK] {sheet_name}: {len(rows)} rows")
            wrote_any = True
        if not wrote_any:
            # ExcelWriter needs at least one sheet to produce a valid file
            pd.DataFrame({"info": ["No extractable data found"]}).to_excel(
                writer, sheet_name="Empty", index=False)
    print(f"\nSaved → {path}")


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

def build_sheets(model) -> Dict[str, List[Dict]]:
    seen: Set[str] = set()
    align_summary, horiz, vert, cant = extract_alignments(model)
    # order matters: specific categories claim their elements before the catch-all
    roads     = extract_roads(model, seen)
    bridges   = extract_bridges(model, seen)
    pavements = extract_pavements(model, seen)
    remaining = extract_remaining(model, seen)
    return {
        "Alignments":           align_summary,
        "Alignment_Horizontal": horiz,
        "Alignment_Vertical":   vert,
        "Alignment_Cant":       cant,
        "Roads":                roads,
        "Bridges":              bridges,
        "Pavements":            pavements,
        "Elements":             remaining,
    }


def run(ifc_path: str, out_xlsx: Optional[str], csv_dir: Optional[str]):
    print(f"Opening {ifc_path} …")
    model = ifcopenshell.open(ifc_path)
    print(f"Schema: {model.schema}")

    sheets = build_sheets(model)

    total = sum(len(v) for v in sheets.values())
    print("\nSummary:")
    for k, v in sheets.items():
        print(f"  {k:<22}: {len(v)}")
    print(f"  {'TOTAL':<22}: {total}")

    if out_xlsx:
        write_xlsx(sheets, out_xlsx)
    elif csv_dir:
        write_csv_dir(sheets, csv_dir)
    else:
        default = os.path.splitext(ifc_path)[0] + "_provi_extract"
        if PANDAS:
            write_xlsx(sheets, default + ".xlsx")
        else:
            write_csv_dir(sheets, default)


def main():
    ap = argparse.ArgumentParser(
        description="ProVI IFC Extractor – civil/road data from ProVI IFC4X3 exports")
    ap.add_argument("ifc", help="Path to ProVI IFC file")
    ap.add_argument("-o", "--out", default=None,
                    help="Output Excel file (.xlsx). "
                         "Default: <ifc_name>_provi_extract.xlsx")
    ap.add_argument("--csv-dir", default=None, metavar="DIR",
                    help="Write one CSV per sheet into DIR (instead of Excel)")
    args = ap.parse_args()

    if args.out and args.csv_dir:
        ap.error("Use either -o (Excel) or --csv-dir, not both.")

    run(args.ifc, args.out, args.csv_dir)


if __name__ == "__main__":
    main()
