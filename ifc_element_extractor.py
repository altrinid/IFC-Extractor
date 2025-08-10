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

# Для Excel-выгрузки используем pandas (опционально)
try:
    import pandas as pd
    PANDAS = True
except Exception:
    PANDAS = False


def get_name(entity) -> str:
    for attr in ("Name", "GlobalId", "Tag"):
        v = getattr(entity, attr, None)
        if v:
            return str(v)
    return f"{entity.is_a()}_{entity.id()}"


def get_level(entity) -> str:
    # Пытаемся достать этаж через связи
    try:
        for inv in getattr(entity, "IsContainedIn", []) or []:
            if inv.is_a("IfcRelContainedInSpatialStructure") and inv.RelatingStructure:
                return get_name(inv.RelatingStructure)
    except Exception:
        pass
    return ""


def get_psets(entity) -> Dict[str, Any]:
    out = {}
    # Попытка через util
    try:
        from ifcopenshell.util.element import get_psets  # type: ignore
        psets = get_psets(entity, include_inherited=True)
        flat = {}
        for grp, vals in psets.items():
            if isinstance(vals, dict):
                for k, v in vals.items():
                    flat[f"{grp}:{k}"] = v
            else:
                flat[grp] = vals
        out.update(flat)
        return out
    except Exception:
        pass

    # Фолбэк без util: пройти связь IsDefinedBy
    try:
        for rel in getattr(entity, "IsDefinedBy", []) or []:
            ps = getattr(rel, "RelatingPropertyDefinition", None)
            if not ps:
                continue
            if ps.is_a("IfcPropertySet"):
                for p in ps.HasProperties or []:
                    key = f"{ps.Name}:{p.Name}"
                    val = getattr(p, "NominalValue", getattr(p, "Description", None))
                    out[key] = getattr(val, "wrappedValue", val)
            elif ps.is_a("IfcElementQuantity"):
                for q in ps.Quantities or []:
                    val = None
                    for f in ("LengthValue", "AreaValue", "VolumeValue", "CountValue",
                              "WeightValue", "TimeValue"):
                        if hasattr(q, f) and getattr(q, f) is not None:
                            val = getattr(q, f)
                            break
                    out[f"{ps.Name}:{q.Name}"] = val
    except Exception:
        pass
    return out


def gather_elements(model, classes: List[str]) -> List[Any]:
    if not classes or classes == ["*"]:
        return [e for e in model if hasattr(e, "GlobalId") and getattr(e, "GlobalId", None)]
    elems = []
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


def extract(ifc_path: str, classes: List[str], top_props: List[str], limit: int = 0):
    model = ifcopenshell.open(ifc_path)
    elements = gather_elements(model, classes)
    if limit > 0:
        elements = elements[:limit]

    base_cols = ["GlobalId", "Entity", "Name", "Level"]
    rows, all_dyn_keys = [], set()

    for e in elements:
        try:
            base = {
                "GlobalId": getattr(e, "GlobalId", ""),
                "Entity": e.is_a(),
                "Name": get_name(e),
                "Level": get_level(e)
            }
            for p in top_props:
                base[p] = normalize(getattr(e, p, ""))
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

    dyn_cols = sorted(list(all_dyn_keys))
    extra_top = [p for p in top_props if p not in base_cols and p not in dyn_cols]
    header = base_cols + extra_top + dyn_cols
    return header, rows


def write_csv(header, rows, out_csv: str):
    os.makedirs(os.path.dirname(os.path.abspath(out_csv)) or ".", exist_ok=True)
    import csv
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=header)
        writer.writeheader()
        for r in rows:
            writer.writerow({k: r.get(k, "") for k in header})


def write_xlsx(header, rows, out_xlsx: str, sheet: str = "Elements"):
    if not PANDAS:
        raise SystemExit("Excel export requires pandas:  pip install pandas openpyxl")
    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx)) or ".", exist_ok=True)
    df = pd.DataFrame([{k: r.get(k, "") for k in header} for r in rows], columns=header)
    df.to_excel(out_xlsx, index=False, sheet_name=sheet)


def main():
    ap = argparse.ArgumentParser(description="IFC Element Extractor → CSV/XLSX")
    ap.add_argument("ifc", help="Path to IFC file, e.g. model.ifc")
    ap.add_argument("-c", "--classes", default="IfcWall,IfcDoor,IfcWindow",
                    help="Comma-separated IFC classes (use * for all).")
    ap.add_argument("-p", "--props", default="PredefinedType,Tag",
                    help="Comma-separated top-level attributes to include.")
    ap.add_argument("--limit", type=int, default=0, help="Limit number of elements (debug).")
    ap.add_argument("--csv", help="Output CSV path, e.g. out.csv")
    ap.add_argument("--xlsx", help="Output Excel path, e.g. out.xlsx")
    ap.add_argument("--sheet", default="Elements", help="Excel sheet name (default: Elements)")

    args = ap.parse_args()

    classes = [c.strip() for c in args.classes.split(",")] if args.classes else []
    props = [p.strip() for p in args.props.split(",")] if args.props else []

    header, rows = extract(args.ifc, classes, props, args.limit)

    wrote = False
    if args.csv:
        write_csv(header, rows, args.csv)
        print(f"[OK] CSV → {args.csv}  (rows: {len(rows)})")
        wrote = True
    if args.xlsx:
        write_xlsx(header, rows, args.xlsx, args.sheet)
        print(f"[OK] XLSX → {args.xlsx}  (rows: {len(rows)})")
        wrote = True

    if not wrote:
        default = "ifc_elements.csv"
        write_csv(header, rows, default)
        print(f"[OK] CSV → {default}  (rows: {len(rows)})")


if __name__ == "__main__":
    main()
