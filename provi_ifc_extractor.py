#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ProVI IFC Extractor
Extracts civil/road engineering data from ProVI IFC4X3 exports:
  - Alignments (horizontal, vertical, cant segments)
  - Roads, pavements, and pavement courses
  - Bridges and bridge parts
  - General elements + Psets

Usage:
  python provi_ifc_extractor.py model.ifc
  python provi_ifc_extractor.py model.ifc -o output.xlsx
  python provi_ifc_extractor.py model.ifc --csv-dir ./output
"""

import argparse
import csv
import os
from typing import Any, Dict, List, Optional, Tuple

try:
    import ifcopenshell
    import ifcopenshell.util.element
except ImportError:
    raise SystemExit("Please install ifcopenshell:  pip install ifcopenshell")

try:
    import pandas as pd
    import openpyxl  # noqa: F401
    PANDAS = True
except ImportError:
    PANDAS = False


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalize(val) -> str:
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


def get_psets(entity) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    try:
        psets = ifcopenshell.util.element.get_psets(entity, include_inherited=True)
        for grp, vals in psets.items():
            if isinstance(vals, dict):
                for k, v in vals.items():
                    out[f"{grp}:{k}"] = normalize(v)
            else:
                out[grp] = normalize(vals)
        return out
    except Exception:
        pass

    try:
        for rel in getattr(entity, "IsDefinedBy", []) or []:
            props = getattr(rel, "RelatingPropertyDefinition", None)
            if not props:
                continue
            if props.is_a("IfcPropertySet"):
                for p in props.HasProperties or []:
                    out[f"{props.Name}:{p.Name}"] = normalize(
                        getattr(p, "NominalValue", None)
                    )
            elif props.is_a("IfcElementQuantity"):
                for q in props.Quantities or []:
                    for field in ("LengthValue", "AreaValue", "VolumeValue", "CountValue", "WeightValue"):
                        if hasattr(q, field) and getattr(q, field) is not None:
                            out[f"{props.Name}:{q.Name}"] = normalize(getattr(q, field))
                            break
    except Exception:
        pass
    return out


def _nested_objects(model, parent) -> List[Any]:
    """Return objects nested under parent via IfcRelNests."""
    result = []
    for rel in model.by_type("IfcRelNests"):
        if rel.RelatingObject == parent:
            result.extend(rel.RelatedObjects or [])
    return result


def _aggregated_objects(model, parent) -> List[Any]:
    """Return objects aggregated under parent via IfcRelAggregates."""
    result = []
    for rel in model.by_type("IfcRelAggregates"):
        if rel.RelatingObject == parent:
            result.extend(rel.RelatedObjects or [])
    return result


def _point_coords(point) -> Tuple[str, str]:
    if point is None:
        return "", ""
    coords = getattr(point, "Coordinates", None)
    if coords and len(coords) >= 2:
        return normalize(coords[0]), normalize(coords[1])
    return "", ""


# ---------------------------------------------------------------------------
# Alignment extraction
# ---------------------------------------------------------------------------

def extract_horizontal_segments(model, alignment_name: str, horiz) -> List[Dict]:
    rows = []
    for seg in _nested_objects(model, horiz):
        if not seg.is_a("IfcAlignmentSegment"):
            continue
        d = getattr(seg, "DesignParameters", None)
        if not d:
            continue
        x, y = _point_coords(getattr(d, "StartPoint", None))
        rows.append({
            "AlignmentName":          alignment_name,
            "SegmentId":              normalize(getattr(seg, "GlobalId", "")),
            "SegmentName":            normalize(getattr(seg, "Name", "")),
            "PredefinedType":         normalize(getattr(d, "PredefinedType", "")),
            "StartDistAlong":         normalize(getattr(d, "StartDistAlong", "")),
            "HorizontalLength":       normalize(getattr(d, "HorizontalLength", "")),
            "StartPointX":            x,
            "StartPointY":            y,
            "StartDirection":         normalize(getattr(d, "StartDirection", "")),
            "StartRadiusOfCurvature": normalize(getattr(d, "StartRadiusOfCurvature", "")),
            "EndRadiusOfCurvature":   normalize(getattr(d, "EndRadiusOfCurvature", "")),
            "IsEntry":                normalize(getattr(d, "IsEntry", "")),
            "GravityCenterLineHeight": normalize(getattr(d, "GravityCenterLineHeight", "")),
        })
    return rows


def extract_vertical_segments(model, alignment_name: str, vert) -> List[Dict]:
    rows = []
    for seg in _nested_objects(model, vert):
        if not seg.is_a("IfcAlignmentSegment"):
            continue
        d = getattr(seg, "DesignParameters", None)
        if not d:
            continue
        rows.append({
            "AlignmentName":      alignment_name,
            "SegmentId":          normalize(getattr(seg, "GlobalId", "")),
            "SegmentName":        normalize(getattr(seg, "Name", "")),
            "PredefinedType":     normalize(getattr(d, "PredefinedType", "")),
            "StartDistAlong":     normalize(getattr(d, "StartDistAlong", "")),
            "HorizontalLength":   normalize(getattr(d, "HorizontalLength", "")),
            "StartHeight":        normalize(getattr(d, "StartHeight", "")),
            "StartGradient":      normalize(getattr(d, "StartGradient", "")),
            "EndGradient":        normalize(getattr(d, "EndGradient", "")),
            "RadiusOfCurvature":  normalize(getattr(d, "RadiusOfCurvature", "")),
            "IsConvex":           normalize(getattr(d, "IsConvex", "")),
        })
    return rows


def extract_cant_segments(model, alignment_name: str, cant) -> List[Dict]:
    rows = []
    rail_head = normalize(getattr(cant, "RailHeadDistance", ""))
    for seg in _nested_objects(model, cant):
        if not seg.is_a("IfcAlignmentSegment"):
            continue
        d = getattr(seg, "DesignParameters", None)
        if not d:
            continue
        rows.append({
            "AlignmentName":    alignment_name,
            "SegmentId":        normalize(getattr(seg, "GlobalId", "")),
            "SegmentName":      normalize(getattr(seg, "Name", "")),
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


def extract_alignments(model) -> Tuple[List[Dict], List[Dict], List[Dict], List[Dict]]:
    """
    Returns (alignment_summary, horiz_rows, vert_rows, cant_rows).
    """
    summaries: List[Dict] = []
    horiz_rows: List[Dict] = []
    vert_rows: List[Dict] = []
    cant_rows: List[Dict] = []

    try:
        alignments = model.by_type("IfcAlignment")
    except Exception:
        return summaries, horiz_rows, vert_rows, cant_rows

    for alignment in alignments:
        name = get_name(alignment)
        psets = get_psets(alignment)
        summary: Dict[str, Any] = {
            "GlobalId":      normalize(getattr(alignment, "GlobalId", "")),
            "Name":          normalize(getattr(alignment, "Name", "")),
            "Description":   normalize(getattr(alignment, "Description", "")),
            "ObjectType":    normalize(getattr(alignment, "ObjectType", "")),
            "PredefinedType": normalize(getattr(alignment, "PredefinedType", "")),
        }
        summary.update(psets)
        summaries.append(summary)

        for child in _aggregated_objects(model, alignment):
            if child.is_a("IfcAlignmentHorizontal"):
                horiz_rows.extend(extract_horizontal_segments(model, name, child))
            elif child.is_a("IfcAlignmentVertical"):
                vert_rows.extend(extract_vertical_segments(model, name, child))
            elif child.is_a("IfcAlignmentCant"):
                cant_rows.extend(extract_cant_segments(model, name, child))

    return summaries, horiz_rows, vert_rows, cant_rows


# ---------------------------------------------------------------------------
# Road / Pavement extraction
# ---------------------------------------------------------------------------

def extract_roads(model) -> List[Dict]:
    rows = []
    try:
        road_types = ["IfcRoad", "IfcFacility"]
    except Exception:
        return rows

    for ifc_type in road_types:
        try:
            for road in model.by_type(ifc_type):
                row: Dict[str, Any] = {
                    "GlobalId":      normalize(getattr(road, "GlobalId", "")),
                    "Entity":        road.is_a(),
                    "Name":          get_name(road),
                    "Description":   normalize(getattr(road, "Description", "")),
                    "ObjectType":    normalize(getattr(road, "ObjectType", "")),
                    "PredefinedType": normalize(getattr(road, "PredefinedType", "")),
                }
                row.update(get_psets(road))
                rows.append(row)
        except Exception:
            continue
    return rows


def extract_pavements(model) -> List[Dict]:
    rows = []
    pavement_types = ["IfcPavement", "IfcCourse"]
    for ifc_type in pavement_types:
        try:
            for elem in model.by_type(ifc_type):
                row: Dict[str, Any] = {
                    "GlobalId":      normalize(getattr(elem, "GlobalId", "")),
                    "Entity":        elem.is_a(),
                    "Name":          get_name(elem),
                    "Description":   normalize(getattr(elem, "Description", "")),
                    "ObjectType":    normalize(getattr(elem, "ObjectType", "")),
                    "PredefinedType": normalize(getattr(elem, "PredefinedType", "")),
                    "Material":      _get_material(elem),
                }
                row.update(get_psets(elem))
                rows.append(row)
        except Exception:
            continue
    return rows


def _get_material(entity) -> str:
    try:
        for rel in model_global.by_type("IfcRelAssociatesMaterial"):  # type: ignore[name-defined]
            if entity in (rel.RelatedObjects or []):
                mat = rel.RelatingMaterial
                if hasattr(mat, "Name") and mat.Name:
                    return str(mat.Name)
                if hasattr(mat, "ForLayerSet") and mat.ForLayerSet:
                    return str(getattr(mat.ForLayerSet, "LayerSetName", ""))
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
# Bridge extraction
# ---------------------------------------------------------------------------

def extract_bridges(model) -> List[Dict]:
    rows = []
    bridge_types = ["IfcBridge", "IfcBridgePart"]
    for ifc_type in bridge_types:
        try:
            for elem in model.by_type(ifc_type):
                row: Dict[str, Any] = {
                    "GlobalId":      normalize(getattr(elem, "GlobalId", "")),
                    "Entity":        elem.is_a(),
                    "Name":          get_name(elem),
                    "Description":   normalize(getattr(elem, "Description", "")),
                    "ObjectType":    normalize(getattr(elem, "ObjectType", "")),
                    "PredefinedType": normalize(getattr(elem, "PredefinedType", "")),
                }
                row.update(get_psets(elem))
                rows.append(row)
        except Exception:
            continue
    return rows


# ---------------------------------------------------------------------------
# General elements extraction (fallback / overview)
# ---------------------------------------------------------------------------

GENERAL_CLASSES = [
    "IfcWall", "IfcSlab", "IfcBeam", "IfcColumn", "IfcPile",
    "IfcEarthworksCut", "IfcEarthworksFill",
    "IfcKerb", "IfcRailing", "IfcSignal", "IfcSign",
    "IfcLinearElement", "IfcBuiltElement",
]


def extract_general(model) -> List[Dict]:
    seen = set()
    rows = []
    for ifc_type in GENERAL_CLASSES:
        try:
            for elem in model.by_type(ifc_type):
                gid = getattr(elem, "GlobalId", None)
                if gid in seen:
                    continue
                seen.add(gid)
                row: Dict[str, Any] = {
                    "GlobalId":      normalize(gid),
                    "Entity":        elem.is_a(),
                    "Name":          get_name(elem),
                    "Description":   normalize(getattr(elem, "Description", "")),
                    "ObjectType":    normalize(getattr(elem, "ObjectType", "")),
                    "PredefinedType": normalize(getattr(elem, "PredefinedType", "")),
                }
                row.update(get_psets(elem))
                rows.append(row)
        except Exception:
            continue
    return rows


# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------

def _write_csv(rows: List[Dict], path: str, label: str):
    if not rows:
        print(f"[SKIP] No {label} found.")
        return
    keys: List[str] = []
    for r in rows:
        for k in r:
            if k not in keys:
                keys.append(k)
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
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


def write_xlsx(sheets: Dict[str, List[Dict]], path: str):
    if not PANDAS:
        raise SystemExit("pandas + openpyxl required for Excel output.  pip install pandas openpyxl")
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, rows in sheets.items():
            if not rows:
                continue
            df = pd.DataFrame(rows)
            safe_name = sheet_name[:31]  # Excel sheet name limit
            df.to_excel(writer, sheet_name=safe_name, index=False)
            print(f"[OK] Sheet '{safe_name}': {len(rows)} rows")
    print(f"\nSaved → {path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

model_global = None  # module-level reference for _get_material


def run(ifc_path: str, out_xlsx: Optional[str], csv_dir: Optional[str]):
    global model_global
    print(f"Opening {ifc_path} …")
    model = ifcopenshell.open(ifc_path)
    model_global = model

    schema = model.schema
    print(f"Schema: {schema}")

    print("Extracting alignments …")
    align_summary, horiz, vert, cant = extract_alignments(model)

    print("Extracting roads …")
    roads = extract_roads(model)

    print("Extracting pavements …")
    pavements = extract_pavements(model)

    print("Extracting bridges …")
    bridges = extract_bridges(model)

    print("Extracting general elements …")
    general = extract_general(model)

    sheets: Dict[str, List[Dict]] = {
        "Alignments":            align_summary,
        "Alignment_Horizontal":  horiz,
        "Alignment_Vertical":    vert,
        "Alignment_Cant":        cant,
        "Roads":                 roads,
        "Pavements":             pavements,
        "Bridges":               bridges,
        "Elements":              general,
    }

    total = sum(len(v) for v in sheets.values())
    print(f"\nTotal records: {total}")
    for k, v in sheets.items():
        print(f"  {k}: {len(v)}")

    if out_xlsx:
        write_xlsx(sheets, out_xlsx)
    elif csv_dir:
        write_csv_dir(sheets, csv_dir)
    else:
        # Default: Excel next to the IFC file
        default_out = os.path.splitext(ifc_path)[0] + "_provi_extract.xlsx"
        if PANDAS:
            write_xlsx(sheets, default_out)
        else:
            default_dir = os.path.splitext(ifc_path)[0] + "_provi_extract"
            write_csv_dir(sheets, default_dir)


def main():
    ap = argparse.ArgumentParser(
        description="ProVI IFC Extractor – extracts civil/road data from ProVI IFC4X3 exports"
    )
    ap.add_argument("ifc", help="Path to ProVI IFC file")
    ap.add_argument(
        "-o", "--out",
        default=None,
        help="Output Excel file (.xlsx). Default: <ifc_name>_provi_extract.xlsx"
    )
    ap.add_argument(
        "--csv-dir",
        default=None,
        metavar="DIR",
        help="Output directory for CSV files (one per sheet). Mutually exclusive with -o."
    )
    args = ap.parse_args()

    if args.out and args.csv_dir:
        ap.error("Use either -o (Excel) or --csv-dir, not both.")

    run(args.ifc, args.out, args.csv_dir)


if __name__ == "__main__":
    main()
