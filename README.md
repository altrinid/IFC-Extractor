# Room Data Exporter

This conceptual Python script demonstrates how to extract room data from a Revit model using the Revit API and export it to a CSV file. The script is intended to run in an environment that supports RevitPythonShell or pyRevit.

## Features
- Collects all rooms in the active document
- Extracts room name, number, level, area, and ceiling height
- Calculates a simple occupancy load
- Writes the data to a CSV file in the system temporary directory

## Usage
1. Install **RevitPythonShell** or **pyRevit**.
2. Load `RoomDataExporter.py` in the script environment.
3. Run the script from within Revit. The output CSV will be placed in your temporary directory (e.g., `C:\\Temp\\room_report.csv`).

The occupancy load uses an example factor of 100 square feet per person. Adjust the value in the script to match your project requirements.
