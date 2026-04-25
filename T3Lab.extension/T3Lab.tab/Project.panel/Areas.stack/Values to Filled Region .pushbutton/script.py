
"""
Opening Assign Values
Assign Values to Filled Region

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Opening Assign Values"
__version__ = "1.0.0"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from rpw import *
from pyrevit import *
import math
"""--------------------------------------------------"""
uidoc= __revit__.ActiveUIDocument
doc=__revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""
active_view_id=doc.ActiveView.Id
detailcomponents_in_view = FilteredElementCollector(doc,active_view_id).\
    OfCategory(BuiltInCategory.OST_DetailComponents).ToElements()
filledregion=[detail for detail in detailcomponents_in_view if detail.Name == "Detail Filled Region"]
filledregion_filter_opening=[fg for fg in filledregion if fg.LookupParameter("Family and Type").AsValueString() == "Filled region: _Area of Opening"]
filledregion_filter_wall=[fg for fg in filledregion if fg.LookupParameter("Family and Type").AsValueString() == "Filled region: _Area of Wall"]
levels = FilteredElementCollector(doc, active_view_id).OfCategory(BuiltInCategory.OST_Levels).ToElements()
level_dict = {level.Name: level.LookupParameter("Elevation").AsDouble() if level.LookupParameter("Elevation") else None for level in levels}
"""--------------------------------------------------"""
def assign_fg_opening(fg):
    number=math.ceil(round(fg.LookupParameter("Area").AsDouble(),5))
    fg.LookupParameter("GMG_AREA ROUNDUP").Set(number)
def assign_fg_wall(fg):
    number=math.floor(fg.LookupParameter("Area").AsDouble())
    fg.LookupParameter("GMG_AREA OF WALL").Set(number)
def assign_percent(fg):
    area_wall   = fg.LookupParameter("GMG_AREA OF WALL").AsDouble()
    area_opening= fg.LookupParameter("GMG_AREA OF OPENING").AsDouble()
    period=0
    if area_wall is not None and area_wall!=0:
        period=area_opening/area_wall
    fg.LookupParameter("GMG_PERCENT OF OPENING").Set(period)
def get_BoundingBox(element):
    options = Options()
    options.IncludeNonVisibleObjects = True
    bbox = element.get_BoundingBox(None)
    return bbox
# def check_overlap(curve_wall,curve_opening):
#     #boundaries of wall
#     wallMinX = curve_wall.Min.X
#     wallMinY = curve_wall.Min.Y
#     wallMinZ = curve_wall.Min.Z
#     wallMaxX = curve_wall.Max.X
#     wallMaxY = curve_wall.Max.Y
#     wallMaxZ = curve_wall.Max.Z
#     #boundaries of opening
#     openingMinX = curve_opening.Min.X
#     openingMinY = curve_opening.Min.Y
#     openingMinZ = curve_opening.Min.Z
#     openingMaxX = curve_opening.Max.X
#     openingMaxY = curve_opening.Max.Y
#     openingMaxZ = curve_opening.Max.Z
#     #Compare
#     compare1 = False
#     compare2 = False
#     compare3 = False
#     compare4 = False
#     if openingMinX==openingMaxX:
#         compare1 = openingMinY >= wallMinY
#         compare2 = openingMinZ >= wallMinZ
#         compare3 = openingMaxY <= wallMaxY
#         compare4 = openingMaxZ <= wallMaxZ
#     if openingMinY==openingMaxY:
#         compare1 = openingMinX >= wallMinX
#         compare2 = openingMinZ >= wallMinZ
#         compare3 = openingMaxX <= wallMaxX
#         compare4 = openingMaxZ <= wallMaxZ
#     if openingMinZ==openingMaxZ:
#         compare1 = openingMinX >= wallMinX
#         compare2 = openingMinY >= wallMinY
#         compare3 = openingMaxX <= wallMaxX
#         compare4 = openingMaxY <= wallMaxY
# 
#     if compare1 and compare2 and compare3 and compare4:
#         return True
#     else:
#         return False
def check_overlap(curve_wall,curve_opening):
    #boundaries of wall
    wallMinX = round(curve_wall.Min.X,10)
    wallMinY = round(curve_wall.Min.Y,10)
    wallMinZ = round(curve_wall.Min.Z,10)
    wallMaxX = round(curve_wall.Max.X,10)
    wallMaxY = round(curve_wall.Max.Y,10)
    wallMaxZ = round(curve_wall.Max.Z,10)
    #boundaries of opening
    openingMinX = round(curve_opening.Min.X,10)
    openingMinY = round(curve_opening.Min.Y,10)
    openingMinZ = round(curve_opening.Min.Z,10)
    openingMaxX = round(curve_opening.Max.X,10)
    openingMaxY = round(curve_opening.Max.Y,10)
    openingMaxZ = round(curve_opening.Max.Z,10)
    if openingMinX!=openingMaxX and openingMinY!=openingMaxY and openingMinZ!=openingMaxZ:
        compare_1 = openingMinZ >= wallMinZ
        compare_2 = openingMaxZ <= wallMaxZ
        # print(openingMinZ)
        # print(openingMaxZ)
        # print(wallMinZ)
        # print(wallMaxZ)
        # print(compare_1)
        # print(compare_2)
        # print("__________________")
        if compare_1 and compare_2:
            # Vector of wall
            minmax_wall_XY = (wallMaxX - wallMinX, wallMaxY - wallMinY)

            # Vector of opening-wall
            minmin_open_XY = (openingMinX - wallMinX, openingMinY - wallMinY)
            minmax_open_XY = (openingMaxX - wallMinX, openingMaxY - wallMinY)
            divide_1_1 = round(minmin_open_XY[0] / minmax_wall_XY[0],10)
            divide_1_2 = round(minmin_open_XY[1] / minmax_wall_XY[1],10)
            divide_2_1 = round(minmax_open_XY[0] / minmax_wall_XY[0],10)
            divide_2_2 = round(minmax_open_XY[1] / minmax_wall_XY[1],10)
            compare_1=divide_1_1 == divide_1_2
            compare_2=divide_2_1 == divide_2_2
            if compare_1 and compare_2 and 0 <= divide_1_1 <= 1 and 0 <= divide_2_1 <= 1:
                return True
            return False
    else:
        compare1 = False
        compare2 = False
        compare3 = False
        compare4 = False
        if openingMinX==openingMaxX:
            compare1 = openingMinY >= wallMinY
            compare2 = openingMinZ >= wallMinZ
            compare3 = openingMaxY <= wallMaxY
            compare4 = openingMaxZ <= wallMaxZ
        if openingMinY==openingMaxY:
            compare1 = openingMinX >= wallMinX
            compare2 = openingMinZ >= wallMinZ
            compare3 = openingMaxX <= wallMaxX
            compare4 = openingMaxZ <= wallMaxZ
        if openingMinZ==openingMaxZ:
            compare1 = openingMinX >= wallMinX
            compare2 = openingMinY >= wallMinY
            compare3 = openingMaxX <= wallMaxX
            compare4 = openingMaxY <= wallMaxY

        if compare1 and compare2 and compare3 and compare4:
            return True
        else:
            return False
def checklevel(fg_wall ):
    # Assuming get_BoundingBox is defined elsewhere
    curve_wall = get_BoundingBox(fg_wall)
    wallMinZ = curve_wall.Min.Z

    # Assuming level_dict is defined and contains levels with their elevations
    for level_name, level_elevation in level_dict.items():
        # Comparing with tolerance (e.g., 0.001) due to floating-point precision issues
        if level_elevation < 0:
            fg_wall.LookupParameter("GMG_LEVEL OF ALLOWABLE OPENING").Set("1ST FLOOR")
        elif abs(wallMinZ - level_elevation) < 0.001:
            # Splitting level name assuming it's a string
            name_split = level_name.split(" ",1)
            if name_split:
                level_suffix = name_split[-1]
                # Assuming "GMG_LEVEL OF ALLOWABLE OPENING" parameter exists and can be set
                fg_wall.LookupParameter("GMG_LEVEL OF ALLOWABLE OPENING").Set(level_suffix)
                break


t=Transaction(doc,"Assign Values to Filled Region")
t.Start()
for fg_opening in filledregion_filter_opening:
    assign_fg_opening(fg_opening)
for fg_wall in filledregion_filter_wall:
    try:
        assign_fg_wall(fg_wall)
        curve_wall = get_BoundingBox(fg_wall)
        area_opening=0
        area_wall=fg_wall.LookupParameter("GMG_AREA OF WALL")
        separa=fg_wall.LookupParameter("GMG_SEPARATION DISTANCE")
        if separa is not None:
            if separa.AsString().endswith("10'"):
                fg_wall.LookupParameter("GMG_ALLOWABLE OPENING").Set("25%")
            elif separa.AsString().endswith("15'"):
                fg_wall.LookupParameter("GMG_ALLOWABLE OPENING").Set("45%")
            elif separa.AsString().endswith("20'"):
                fg_wall.LookupParameter("GMG_ALLOWABLE OPENING").Set("75%")
            elif separa.AsString().endswith("25'"):
                fg_wall.LookupParameter("GMG_ALLOWABLE OPENING").Set("NO LIMIT")
            elif separa.AsString().endswith("30'"):
                fg_wall.LookupParameter("GMG_ALLOWABLE OPENING").Set("NO LIMIT")
        for fg_opening in filledregion_filter_opening:
            curve_opening = get_BoundingBox(fg_opening)
            check_in=check_overlap(curve_wall,curve_opening)
            if check_in:
                area_opening+=fg_opening.LookupParameter("GMG_AREA ROUNDUP").AsDouble()
        fg_wall.LookupParameter("GMG_AREA OF OPENING").Set(area_opening)
        assign_percent(fg_wall)
        checklevel(fg_wall)
    except:
        continue
t.Commit()