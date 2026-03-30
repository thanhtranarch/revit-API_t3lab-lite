"""
Legend (Types List)
Create Types List of Element
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Types List"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

"""--------------------------------------------------"""
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
"""--------------------------------------------------"""
active_view=doc.ActiveView
active_level=doc.ActiveView.GenLevel
#Get Types
all_walls_types = FilteredElementCollector(doc).OfClass(WallType).ToElements()
# all_floors_types= FilteredElementCollector(doc).OfCategory(BuiltInCategory.OTS_floors).OfClass(FloorType).ToElements()


#Function
def create_wall(origin,wall_type):
    pt_start=origin
    pt_end=XYZ(origin.X+2,origin.Y,origin.Z)
    curve=Line.CreateBound(pt_start,pt_end)
    H=10
    O=0
    flip=False
    struc=False
    wall=Wall.Create(doc,curve,wall_type.Id, active_level.Id, H, O, flip, struc)
    return wall
#Origin

X=0
Y=0
Z=0

t=Transaction(doc,"Create Wall's List")
t.Start()
#Create WallType:
for wall_type in all_walls_types:
    origin = XYZ(X, Y, Z)
    create_wall(origin, wall_type)
    Y-=2
t.Commit()
