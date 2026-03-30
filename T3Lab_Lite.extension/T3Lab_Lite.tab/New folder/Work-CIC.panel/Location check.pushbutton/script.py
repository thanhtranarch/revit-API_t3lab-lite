# -*- coding: utf-8 -*-
"""

Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Check"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from collections import namedtuple

"""--------------------------------------------------"""
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""
floors  = FilteredElementCollector(doc).OfClass(Floor).WhereElementIsNotElementType().ToElements()
walls   = FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType().ToElements()
doors   = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()
windows  = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()
railings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing).WhereElementIsNotElementType().ToElements()
levels  = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
ceilings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()


Levels= namedtuple("Levels",["obj","elev"])
level_info=[]
for level in levels:
    level_elev=level.LookupParameter("Elevation")
    level_info.append(Levels(level,level_elev.AsDouble()))


# Start a transaction
t = Transaction(doc, "Set Location")
t.Start()


for f in floors:
    try:
        level_param_f = f.LookupParameter("Level")

        if level_param_f:
            level_f = level_param_f.AsValueString()[0:2]  # Extract the first two characters of the level string
            # Set a custom parameter instead of "Location"
            f.LookupParameter("Element Name").Set("Sàn")

            if "FL" == level_param_f.AsValueString()[2:4]:
                f.LookupParameter("Location").Set("{}FL/ Tầng {}".format(level_f, level_f))
            elif "BS" == level_param_f.AsValueString()[2:4]:
                f.LookupParameter("Location").Set("{}BS/ Hầm {}".format(level_f, level_f))
            if "GRFL" == level_param_f.AsValueString()[0:4]:
                f.LookupParameter("Location").Set("GRFL/ Tầng trệt")
            if "1RFL" == level_param_f.AsValueString()[0:4]:
                f.LookupParameter("Location").Set("{}FL/ Tầng mái thấp".format(level_f))
            elif "2RFL" == level_param_f.AsValueString()[0:4]:
                f.LookupParameter("Location").Set("{}FL/ Tầng mái cao".format(level_f))
    except: 
        pass
 
for c in ceilings:
    try:
        level_param_c = c.LookupParameter("Level")
        if level_param_c:
            level_c = level_param_c.AsValueString()[0:2]  # Extract the first two characters of the level string
            # Set a custom parameter instead of "Location"
            c.LookupParameter("Element Name").Set("Trần")
            
            if "FL" == level_param_c.AsValueString()[2:4]:
                c.LookupParameter("Location").Set("{}FL/ Tầng {}".format(level_c, level_c))
            elif "BS" == level_param_c.AsValueString()[2:4]:
                c.LookupParameter("Location").Set("{}BS/ Hầm {}".format(level_c, level_c))
            if "GRFL" == level_param_c.AsValueString()[0:4]:
                c.LookupParameter("Location").Set("GRFL/ Tầng trệt")
            if "1RFL" == level_param_c.AsValueString()[0:4]:
                c.LookupParameter("Location").Set("{}FL/ Tầng mái thấp".format(level_c))
            elif "2RFL" == level_param_c.AsValueString()[0:4]:
                c.LookupParameter("Location").Set("{}FL/ Tầng mái cao".format(level_c))
    except:
        pass
                
for w in walls:
    try:
        level_param_w = w.LookupParameter("Base Constraint")
        family_w       = w.LookupParameter("Family")
        if level_param_w:
            level_w = level_param_w.AsValueString()[0:2]  # Extract the first two characters of the level string
            # Set a custom parameter instead of "Location"

            if "FL" == level_param_w.AsValueString()[2:4]:
                w.LookupParameter("Location").Set("{}FL/ Tầng {}".format(level_w, level_w))
            elif "BS" == level_param_w.AsValueString()[2:4]:
                w.LookupParameter("Location").Set("{}BS/ Hầm {}".format(level_w, level_w))
            if "GRFL" == level_param_w.AsValueString()[0:4]:
                w.LookupParameter("Location").Set("GRFL/ Tầng trệt")
            if "1RFL" == level_param_w.AsValueString()[0:4]:
                w.LookupParameter("Location").Set("{}FL/ Tầng mái thấp".format(level_w))
            elif "2RFL" == level_param_w.AsValueString()[0:4]:
                w.LookupParameter("Location").Set("{}FL/ Tầng mái cao".format(level_w))
        if family_w:
            if family_w.AsValueString()=="Basic Wall":
                w.LookupParameter("Element Name").Set("Tường")
            elif family_w.AsValueString()=="Curtain Wall":
                if "LOUVER" in w.LookupParameter("Type").AsValueString().upper():
                    w.LookupParameter("Element Name").Set("Lưới ngăn kỹ thuật")
                else:
                    w.LookupParameter("Element Name").Set("Tường kính")
    except:
        pass
for d in doors:
    try:
        if d.LookupParameter("Family"):
            if "OPENING" in d.LookupParameter("Family").AsValueString().upper():
                d.LookupParameter("Element Name").Set("")
                d.LookupParameter("Location").Set("")
                d.LookupParameter("Mark").Set("")
            else:
                level_param_d = d.LookupParameter("Level")
                if level_param_d:
                    level_d = level_param_d.AsValueString()[0:2]  # Extract the first two characters of the level string
                    # Set a custom parameter instead of "Location"
                    d.LookupParameter("Element Name").Set("Cửa")

                    if "FL" == level_param_d.AsValueString()[2:4]:
                        d.LookupParameter("Location").Set("{}FL/ Tầng {}".format(level_d, level_d))
                    elif "BS" == level_param_d.AsValueString()[2:4]:
                        d.LookupParameter("Location").Set("{}BS/ Hầm {}".format(level_d, level_d))
                    if "GRFL" == level_param_d.AsValueString()[0:4]:
                        d.LookupParameter("Location").Set("GRFL/ Tầng trệt")
                    if "1RFL" == level_param_d.AsValueString()[0:4]:
                        d.LookupParameter("Location").Set("{}FL/ Tầng mái thấp".format(level_d))
                    elif "2RFL" == level_param_d.AsValueString()[0:4]:
                        d.LookupParameter("Location").Set("{}FL/ Tầng mái cao".format(level_d))
    except:
        pass


for wd in windows:
    try:
        level_param_wd = wd.LookupParameter("Level")
        family_wd   = wd.LookupParameter("Family")
        if level_param_wd:
            level_wd = level_param_wd.AsValueString()[0:2]  # Extract the first two characters of the level string
            # Set a custom parameter instead of "Location"

            if "FL" == level_param_wd.AsValueString()[2:4]:
                wd.LookupParameter("Location").Set("{}FL/ Tầng {}".format(level_wd, level_wd))
            elif "BS" == level_param_wd.AsValueString()[2:4]:
                wd.LookupParameter("Location").Set("{}BS/ Hầm {}".format(level_wd, level_wd))
            if "GRFL" == level_param_wd.AsValueString()[0:4]:
                wd.LookupParameter("Location").Set("GRFL/ Tầng trệt")
            if "1RFL" == level_param_wd.AsValueString()[0:4]:
                wd.LookupParameter("Location").Set("{}FL/ Tầng mái thấp".format(level_wd))
            elif "2RFL" == level_param_wd.AsValueString()[0:4]:
                wd.LookupParameter("Location").Set("{}FL/ Tầng mái cao".format(level_wd))
            if "LOUVER" in family_wd.AsValueString().upper():
                wd.LookupParameter("Element Name").Set("Cửa thông gió")
            else:
                wd.LookupParameter("Element Name").Set("Cửa sổ")
    except:
        pass
for r in railings:
    try:
    # Get the bounding box of the railing
        bounding_box = r.get_BoundingBox(None)  # Use None for the coordinate system
        if bounding_box is not None:
            min_point = bounding_box.Min.Z
            for level in level_info:
                if abs(level.elev - min_point) < 1:
                    level_r=level.obj.Name[0:2]

                    if "FL" == level.obj.Name[2:4]:
                        r.LookupParameter("Location").Set("{}FL/ Tầng {}".format(level_r, level_r))
                    elif "BS" == level.obj.Name[2:4]:
                        r.LookupParameter("Location").Set("{}BS/ Hầm {}".format(level_r, level_r))
                    if "GRFL" == level.obj.Name[0:4]:
                        r.LookupParameter("Location").Set("GRFL/ Tầng trệt")
                    if "1RFL" == level.obj.Name[0:4]:
                        wd.LookupParameter("Location").Set("{}FL/ Tầng mái thấp".format(level_r))
                    elif "2RFL" == level.obj.Name[0:4]:
                        wd.LookupParameter("Location").Set("{}FL/ Tầng mái cao".format(level_r))
                    r.LookupParameter("Element Name").Set("Lan can")

    except:
        pass


    
t.Commit()

