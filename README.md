# Revit Parameters Export – Dynamo Script

This Dynamo script allows you to export key parameters (such as Name, Level, Category, and dimensions) from selected Revit elements directly to an Excel file.

## How it Works

- Selects all elements of a given category (e.g., Walls, Doors, Windows) in your Revit model.
- Extracts user-specified parameters (e.g., “Type Name”, “Level”, “Area”, etc.).
- Writes the parameter values to an Excel spreadsheet for easy review and further analysis.

## How to Use

1. Open your Revit project and start Dynamo.
2. Open the provided `.dyn` script file.
3. Select the category you wish to export (can duplicate for multiple categories).
4. Enter parameter names as string inputs.
5. Specify your Excel file path.
6. Run the script. The data will be written to the Excel file.

## Example Output

| ElementId | Type Name | Level    | Area   |
|-----------|-----------|----------|--------|
| 123456    | Basic 200 | Level 1  | 24.5   |
| 789012    | Basic 150 | Level 2  | 10.2   |

## Prerequisites

- Autodesk Revit (tested with 2023)
- Dynamo for Revit (2.x or higher)
- Excel installed (for best results)

## Credits

Author: Rodion Dykhanov  
For learning and demonstration purposes.
