import clr
import csv
import math
from System.IO import Path

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, BuiltInParameter,
                               UnitUtils, UnitTypeId)

doc = DocumentManager.Instance.CurrentDBDocument

rooms = (FilteredElementCollector(doc)
         .OfCategory(BuiltInCategory.OST_Rooms)
         .WhereElementIsNotElementType()
         .ToElements())

output_path = Path.Combine(Path.GetTempPath(), 'room_report.csv')

with open(output_path, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Room Name', 'Number', 'Level', 'Area (sf)', 'Ceiling Height (ft)', 'Occupancy Load'])

    factor = 100.0  # occupant load factor

    for room in rooms:
        name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
        number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
        level = room.Level.Name if room.Level else ''
        area = UnitUtils.ConvertFromInternalUnits(room.Area, UnitTypeId.SquareFeet)
        height = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble()
        height_ft = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Feet)
        occupancy = math.ceil(area / factor)
        writer.writerow([name, number, level, round(area, 2), round(height_ft, 2), occupancy])

print('Room data exported to', output_path)
