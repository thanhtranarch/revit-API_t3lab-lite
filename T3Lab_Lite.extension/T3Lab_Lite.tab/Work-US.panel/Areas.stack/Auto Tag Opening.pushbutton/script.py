
"""
Tag Area Opening
Auto Tag Opening

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Tag Area Opening"

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
filledregion_filter=[fg for fg in filledregion if fg.LookupParameter("Family and Type").AsValueString() == "Filled region: _Area of Opening"]
tag_types=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DetailComponentTags).WhereElementIsElementType().ToElements()
tagtotag = next((t for t in tag_types if t.LookupParameter("Type Name").AsString() == "Area Tag for Fill Region"), None)
"""--------------------------------------------------"""
def get_filled_region_center(filled_region):
    # Get the boundary curves of the filled region
    boundary_curves = filled_region.GetBoundaries()

    # Initialize variables to store total X, Y, and Z coordinates
    total_x = 0
    total_y = 0
    total_z = 0
    total_points = 0

    # Loop through each curve loop in the boundary curves
    for curve_loop in boundary_curves:
        # Loop through each curve in the curve loop
        for curve in curve_loop:
            # Get the start and end points of the curve
            start_point = curve.GetEndPoint(0)
            end_point = curve.GetEndPoint(1)

            # Calculate the midpoint of the curve
            midpoint = XYZ((start_point.X + end_point.X) / 2,
                           (start_point.Y + end_point.Y) / 2,
                           (start_point.Z + end_point.Z) / 2)

            # Add the midpoint coordinates to the total
            total_x += midpoint.X
            total_y += midpoint.Y
            total_z += midpoint.Z
            total_points += 1

    # Calculate the average X, Y, and Z coordinates
    if total_points > 0:
        avg_x = total_x / total_points
        avg_y = total_y / total_points
        avg_z = total_z / total_points
        return XYZ(avg_x, avg_y, avg_z)
    else:
        return None
def tag_filled_region(filled_region):
    curves = filled_region.GetBoundaries()
    centerpoint = get_filled_region_center(filled_region)
    if centerpoint:
        tag = IndependentTag.Create(
            doc,
            tagtotag.Id,
            active_view_id,
            Reference(filled_region),
            False,
            TagOrientation.Horizontal,
            centerpoint)  # Passing the center point
        tag.ChangeTypeId(tagtotag.Id)
"""--------------------------------------------------"""
if tagtotag is not None:
    t = Transaction(doc, "Tag Opening-FR")
    t.Start()
    for filled_region in filledregion_filter:
        tag_filled_region(filled_region)
    t.Commit()
else:
    TaskDialog.Show("Tag", "Tag type not found or defined.")