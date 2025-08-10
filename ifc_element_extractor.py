#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import csv
import os
from typing import List, Dict, Any

try:
    import ifcopenshell
except ImportError:
    raise SystemExit("Please install ifcopenshell:  pip install ifcopenshell")

try:
    import pandas as pd   # optional
    PANDAS = True
except Exception:
    PANDAS = False


# -------- Helpers --------
def get_name(entity) -> str:
    for attr in ("Name", "GlobalId", "Tag"):
        v = getattr(entity, attr, None)
        if v:
            return str(v)
    return f"{entity.is_a()}_{entity.id()}"


def get_level(entity) -> str:
    """
    Attempts to retrieve the level/floor name via IfcRelContainedInSpatialStructure.
    Returns '' if not found.
    """
    try:
        if hasattr(entity, "ContainedInStructure"):
            # IFC4
            rels = entity.ContainedInStructure or []
        else:
            # Generic approach
            rels = [r for r in (entity.ContainedInStructure or [])]

    except Exception:
        rels = []

    # Alternative: check inverse relationships
    if not rels:
        try:
            # Use inverted relationships
            for inv in entity.IsContainedIn:
                if inv.is_a("IfcRelContainedInSpatialStructure") and inv.RelatingStructure:
                    return get_name(inv.RelatingStructure)
        except Exception:
            pass

    for rel in rels:
        try:
            if rel.is_a("IfcRelContainedInSpatialStructure") and rel.RelatingStructure:
                return get_name(rel.RelatingStructure)
        except Exception:
            pass
    return ""


def get_psets(entity) -> Dict[str, Any]:
    """
    Returns a flat dict of all properties (Psets + Qto):
    { 'Pset_WallCommon:FireRating': 'REI60', 'Qto_WallBaseQuantities:Length': 12.3, ... }
    """
    out = {}
    try:
        psets = ifcopenshell.util.element.get_psets(entity, include_inherited=True)  # type: ignore
    except Exception:
        # Fallback without util
        psets = {}
        # Through IsDefinedBy → IfcRelDefinesByProperties → IfcPropertySet/IfcElementQuantity
        try:
            for rel in getattr(entity, "IsDefinedBy", []) or []:
                props = getattr(rel, "RelatingPropertyDefinition", None)
                if not props:
                    continue
                if props.is_a("IfcPropertySet"):
                    for p in props.HasProperties or []:
                        key = f"{props.Name}:{p.Name}"
                        out[key] = getattr(p, "NominalValue", getattr(p, "Description", None))
                elif props.is_a("IfcElementQuantity"):
                    for q in props.Quantities or []:
                        val = None
                        for f in ("LengthValue", "AreaValue", "VolumeValue", "CountValue", "WeightValue", "TimeValue"):
                            if hasattr(q, f) and getattr(q, f) is not None:
                                val = getattr(q, f)
                                break
                        key = f"{props.Name}:{q.Name}"
                        out[key] = val
        else:
            pass

    if psets:
        # util returns already flat structure, but grouped by Pset
        flat = {}
        for grp, vals in psets.items():
            if isinstance(vals, dict):
                for k, v in vals.items():
                    flat[f"{grp}:{k}"] = v
            else:
                flat[grp] = vals
        out.update(flat)

    return out


def gather_elements(model, classes: List[str]) -> List[Any]:
    elems = []
    if not classes or classes == ["*"]:
        # Select all elements that have a GlobalId (filter out irrelevant objects)
        for ent in model:
            try:
                if hasattr(ent, "GlobalId") and ent.GlobalId:
                    elems.append(ent)
            except Exception:
                pass
        return elems

    for c in classes:
        try:
            elems.extend(model.by_type(c))
        except Exception:
            pass
    return elems


def normalize(val):
    if val is None:
        return ""
    if hasattr(val, "wrappedValue"):
        return val.wrappedValue
    return str(val)


# -------- Main extraction --------
def extract(ifc_path: str, out_csv: str, classes: List[str], top_props: List[str], limit: int = 0):
    model = ifcopenshell.open(ifc_path)

    elements = gather_elements(model, classes)
    if limit > 0:
        elements = elements[:limit]

    # Base columns
    base_cols = ["GlobalId", "Entity", "Name", "Level"]
    rows = []
    all_dyn_keys = set()

    # Process elements and collect base attributes + psets
    for e in elements:
        try:
            base = {
                "GlobalId": getattr(e, "GlobalId", ""),
                "Entity": e.is_a(),
                "Name": get_name(e),
                "Level": get_level(e)
            }
            # Try to get top_props as direct attributes
            for p in top_props:
                base[p] = normalize(getattr(e, p, ""))
            # psets
            pset_dict = get_psets(e)
            for k, v in pset_dict.items():
                if v is None:
                    continue
                key = str(k)
                base[key] = normalize(v)
                all_dyn_keys.add(key)
            rows.append(base)
        except Exception:
            continue

    # Final column order
    dyn_cols = sorted(list(all_dyn_keys))
    # Include top_props if not already present
    extra_top = [p for p in top_props if p not in base_cols and p not in dyn_cols]
    header = base_cols + extra_top + dyn_cols

    # Write CSV
    os.makedirs(os.path.dirname(os.path.abspath(out_csv)) or ".", exist_ok=True)
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=header)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in header})

    print(f"[OK] Extracted {len(rows)} elements → {out_csv}")
    print(f"Classes: {', '.join(classes) if classes else 'ALL'}")
    if PANDAS:
        try:
            df = pd.read_csv(out_csv)
            print(df.head(10).to_string(index=False))
        except Exception:
            pass


# -------- CLI --------
def main():
    ap = argparse.ArgumentParser(
        description="IFC Element Extractor: export elements + properties (Psets/Qto) to CSV"
    )
    ap.add_argument("ifc", help="Path to IFC file, e.g. model.ifc")
    ap.add_argument("-o", "--out", default="ifc_elements.csv", help="Output CSV path (default: ifc_elements.csv)")
    ap.add_argument(
        "-c", "--classes",
        default="IfcWall,IfcDoor,IfcWindow",
        help="Comma-separated IFC classes to extract (use * for all). Default: IfcWall,IfcDoor,IfcWindow"
    )
    ap.add_argument(
        "-p", "--props",
        default="PredefinedType,Tag",
        help="Comma-separated top-level attributes to try to read (in addition to Psets), e.g. Name,Tag,PredefinedType"
    )
    ap.add_argument("--limit", type=int, default=0, help="Limit number of elements (debug)")

    args = ap.parse_args()

    classes = [c.strip() for c in args.classes.split(",")] if args.classes else []
    props = [p.strip() for p in args.props.split(",")] if args.props else []

    extract(args.ifc, args.out, classes, props, args.limit)


if __name__ == "__main__":
    main()
