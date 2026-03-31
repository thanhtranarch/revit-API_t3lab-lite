# -*- coding: UTF-8 -*-
"""
Intersection2D

Check Intersection of Filled Region

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__  = "Intersection2D"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from rpw import *
from pyrevit import *

# DEFINE VARIABLES
# ==================================================
uidoc= __revit__.ActiveUIDocument
doc=__revit__.ActiveUIDocument.Document
active_view_id=doc.ActiveView.Id

# CLASS/FUNCTIONS
# ==================================================
def calculate_normal_vector(profile):
    points = []
    for curve in profile:
        for line in curve:
            points.append(line.Origin)

    # Use the first three unique points to calculate the normal
    unique_points = list(set(points))[:3]
    p1, p2, p3 = unique_points
    v1 = p2 - p1
    v2 = p3 - p1
    normal = v1.CrossProduct(v2).Normalize()
    return normal
def check_overlap(curve_wall, curve_opening, normal_vector):
    try:
        # Convert curves to solids
        solid_wall = GeometryCreationUtilities.CreateExtrusionGeometry(curve_wall, normal_vector, 1)
        solid_opening = GeometryCreationUtilities.CreateExtrusionGeometry(curve_opening, normal_vector, 1)
        # Check for overlap using ExecuteBooleanOperation
        solid_inter = BooleanOperationsUtils.ExecuteBooleanOperation(solid_opening, solid_wall, BooleanOperationsType.Intersect)
        print(solid_opening.Volume)
        print(solid_wall.Volume)
        print(solid_inter.Volume)
        if abs(solid_inter.Volume - solid_wall.Volume) < 1:
            print(1)
            return True
        else:
            print(2)
            return False
    except Exception as e:
        print("Error in check_overlap: {}".format(e))
        print(3)
        return False
    
# MAIN SCRIPT
# ==================================================
#get elements
detailcomponents_in_view = FilteredElementCollector(doc,active_view_id).OfCategory(BuiltInCategory.OST_DetailComponents).ToElements()
filledregion=[detail for detail in detailcomponents_in_view if detail.Name == "Detail Filled Region"]
filledregion_filter_opening=[fg for fg in filledregion if fg.LookupParameter("Type").AsValueString() == "Diagonal_cross-hatch"]
filledregion_filter_wall=[fg for fg in filledregion if fg.LookupParameter("Type").AsValueString() == "Solid_Black"]
levels = FilteredElementCollector(doc, active_view_id).OfCategory(BuiltInCategory.OST_Levels).ToElements()
level_dict = {level.Name: level.LookupParameter("Elevation").AsDouble() if level.LookupParameter("Elevation") else None for level in levels}
"""--------------------------------------------------"""
print(len(filledregion_filter_opening))
print(len(filledregion_filter_wall))
# Iterate over filled regions for walls and openings
count=0
normal_vector=calculate_normal_vector(filledregion_filter_wall[0].GetBoundaries())
for fg_wall in filledregion_filter_wall:
    curve_wall = fg_wall.GetBoundaries()  # Assuming you have a method to get boundaries

    print("Normal vector:", normal_vector)
    for fg_opening in filledregion_filter_opening:
        curve_opening = fg_opening.GetBoundaries()  # Assuming you have a method to get boundaries
        print("Checking overlap between:", fg_wall.Name, "and", fg_opening.Name)
        # Check for overlap
        overlap = check_overlap(curve_wall, curve_opening, normal_vector)
        if overlap:
            count+=1
        print("Overlap result:", overlap)
        print("------------------------------------------")
print(count)


