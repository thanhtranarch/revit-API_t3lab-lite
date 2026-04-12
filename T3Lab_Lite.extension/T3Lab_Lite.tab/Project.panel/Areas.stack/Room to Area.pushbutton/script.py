# -*- coding: utf-8 -*-
"""
Room to Area

This script will loop through existing areas to asign roomsr.

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""
__author__ ="Tran Tien Thanh"
__title__ = "Room to Area"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==================================================
from rpw import revit, DB, db
from Autodesk.Revit.DB import (
    Transaction,
    View,
    FilteredElementCollector,
    BuiltInCategory,
    ViewType)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
active_view_id=doc.ActiveView.Id

# CLASS/FUNCTIONS
# ==================================================
def compare_box(room, area):
    room_box = room.get_BoundingBox(doc.ActiveView)
    area_box = area.get_BoundingBox(doc.ActiveView)
    if area_box is not None and room_box is not None:
        compare = [
            abs(room_box.Min.X - area_box.Min.X),
            abs(room_box.Min.Y - area_box.Min.Y),
            abs(room_box.Max.X - area_box.Max.X),
            abs(room_box.Max.Y - area_box.Max.Y),
            abs(room_box.Max.Z - area_box.Max.Z)]
        result = all(i <= 4 for i in compare)
        return result
def assign_room_to_area(room,area):
    try:
        occupancy = room.LookupParameter("Occupancy").AsString()
        room_type = room.LookupParameter("Room Type").AsString()
        name = room.LookupParameter("Name").AsString()
        level= room.LookupParameter("Level").AsValueString()
        level_splited=level.split(" ",1)
        # bal_area=area.LookupParameter("BALCONY AREA").AsDouble()
        area.LookupParameter("Occupancy").Set(occupancy)
        area.LookupParameter("Room Type").Set(room_type)
        area.LookupParameter("Name").Set(name)
        area.LookupParameter("Balcony Area").Set(0)
        area.LookupParameter("LEVEL").Set(level_splited[-1])
    except:
        pass

# MAIN SCRIPT
# ==================================================
# Get elements
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
areas = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).ToElements()
with Transaction(doc, "Assign Rooms to Areas") as t:
    t.Start()
    count=0
    for area in areas:
        # if area.AreaScheme.Name.strip()=='GS - DWELLING UNIT':.
        try:
            for room in rooms:
                result = compare_box(room, area)
                if result:
                    count+=1
                    assign_room_to_area(room,area)
        except Exception as e:
            print(str(e))
    t.Commit()
    
if count > 0:
    TaskDialog.Show("Rooms to Areas", "Done!")
else:
    TaskDialog.Show("Rooms to Areas", "There is not assigned value")