IFC Element Extractor – Python Script
This Python script extracts elements and their properties from IFC files and exports them to CSV or Excel format.
It supports filtering by IFC class, retrieving base attributes (e.g., Name, Level, PredefinedType), and collecting all available property set (Pset) and quantity (Qto) data.

How it Works
Opens an IFC2x3 or IFC4 model using ifcopenshell.

Selects elements from specified classes (e.g., IfcWall, IfcDoor, IfcWindow) or all classes.

Retrieves base attributes and all Pset/Qto data.

Writes the extracted data to CSV or Excel.

How to Use
Install Python 3.11+ and dependencies:

bash
Kopieren
Bearbeiten
pip install -r requirements.txt
Run the script:

bash
Kopieren
Bearbeiten
python ifc_element_extractor.py model.ifc --csv output.csv
Optional: export to Excel:

bash
Kopieren
Bearbeiten
python ifc_element_extractor.py model.ifc --xlsx output.xlsx
Example Output
GlobalId	Entity	Name	Level	Pset_WallCommon:FireRating
3kd9...	IfcWall	ExtWall01	Level1	REI60

Prerequisites
Python 3.11 or newer

ifcopenshell

pandas + openpyxl for Excel export (optional)

Credits
Author: Rodion Dykhanov
For learning and demonstration purposes.

