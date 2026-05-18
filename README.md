IFC Element Extractor – Python Scripts
This repository contains two Python scripts for extracting data from IFC files.

---

## 1. General IFC Extractor (`ifc_element_extractor.py`)

Extracts elements and their properties from any IFC2x3 or IFC4 file and exports them to CSV.
Supports filtering by IFC class and collects all available Pset/Qto data.

### Usage

```bash
pip install -r requirements.txt
python ifc_element_extractor.py model.ifc -o output.csv
```

**Options**

| Flag | Default | Description |
|------|---------|-------------|
| `-o` | `ifc_elements.csv` | Output CSV path |
| `-c` | `IfcWall,IfcDoor,IfcWindow` | Comma-separated IFC classes (use `*` for all) |
| `-p` | `PredefinedType,Tag` | Extra top-level attributes to include |
| `--limit` | 0 (off) | Cap element count (debug) |

### Example Output

| GlobalId | Entity | Name | Level | Pset_WallCommon:FireRating |
|----------|--------|------|-------|---------------------------|
| 3kd9... | IfcWall | ExtWall01 | Level1 | REI60 |

---

## 2. ProVI IFC Extractor (`provi_ifc_extractor.py`)

Specialised extractor for **ProVI** road/civil design exports (IFC4X3).
Extracts alignment geometry, road layers, pavements, and bridge structures into
separate sheets of an Excel workbook (or individual CSV files).

### Output sheets

| Sheet | Contents |
|-------|----------|
| `Alignments` | Alignment summary + property sets |
| `Alignment_Horizontal` | Horizontal segments (line, arc, clothoid, …) |
| `Alignment_Vertical` | Vertical segments (constant grade, parabolic arc, …) |
| `Alignment_Cant` | Cant/super-elevation segments |
| `Roads` | `IfcRoad` / `IfcFacility` elements |
| `Pavements` | `IfcPavement` and `IfcCourse` layers |
| `Bridges` | `IfcBridge` and `IfcBridgePart` elements |
| `Elements` | General civil elements (kerbs, earthworks, signs, …) |

### Usage

```bash
pip install -r requirements.txt

# Default: produces <model>_provi_extract.xlsx next to the IFC file
python provi_ifc_extractor.py model.ifc

# Explicit Excel output
python provi_ifc_extractor.py model.ifc -o output.xlsx

# CSV output (one file per sheet)
python provi_ifc_extractor.py model.ifc --csv-dir ./output
```

### Horizontal segment fields

`AlignmentName`, `SegmentId`, `PredefinedType`, `StartDistAlong`,
`HorizontalLength`, `StartPointX`, `StartPointY`, `StartDirection`,
`StartRadiusOfCurvature`, `EndRadiusOfCurvature`, `IsEntry`

Supported `PredefinedType` values: `LINE`, `CIRCULARARC`, `CLOTHOID`,
`CUBIC`, `HELMERTCURVE`, `BLOSSCURVE`, `COSINECURVE`, `SINECURVE`, `VIENNESEBEND`

### Vertical segment fields

`AlignmentName`, `SegmentId`, `PredefinedType`, `StartDistAlong`,
`HorizontalLength`, `StartHeight`, `StartGradient`, `EndGradient`,
`RadiusOfCurvature`, `IsConvex`

Supported `PredefinedType` values: `CONSTANTGRADIENT`, `PARABOLICARC`, `CIRCULARARC`

---

## Prerequisites

- Python 3.11+
- `ifcopenshell >= 0.7`
- `pandas >= 2.0` + `openpyxl >= 3.1` (Excel output, optional for the general extractor)

```bash
pip install -r requirements.txt
```

---

Credits: Rodion Dykhanov – for learning and demonstration purposes.
