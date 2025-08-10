# IFC Element Extractor

Export IFC elements (classes + properties) to **CSV** or **Excel**.  
Tested with IFC2x3/IFC4.

## Features
- Select IFC classes (e.g., `IfcWall,IfcDoor,IfcWindow` or `*` for all).
- Export base fields (`GlobalId, Entity, Name, Level`) + **all** Psets/Qto as flat columns.
- Output to CSV or XLSX.

## Install
```bash
python -m venv .venv && . .venv/bin/activate  # (Windows: .venv\Scripts\activate)
pip install -r requirements.txt
```

## Usage
```bash
# Default → CSV (ifc_elements.csv)
python ifc_element_extractor.py model.ifc

# Specific classes + extra attributes → CSV
python ifc_element_extractor.py model.ifc -c "IfcWall,IfcDoor" -p "Name,Tag,PredefinedType" --csv out.csv

# All classes → Excel
python ifc_element_extractor.py model.ifc -c "*" --xlsx out.xlsx

# Limit rows (debug)
python ifc_element_extractor.py model.ifc --limit 200 --csv sample.csv
```

## Output
- **CSV/XLSX columns**:  
  `GlobalId, Entity, Name, Level, [your top-level attributes], Pset_*:* / Qto_*:*`

## Notes
- Excel export requires `pandas` + `openpyxl`.
- Works standalone (no Revit/Dynamo needed).

## License
MIT
