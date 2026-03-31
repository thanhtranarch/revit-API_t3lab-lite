"""Assign Demolition to Walls"""
__author__='Tran Tien Thanh - trantienthanh.arch@gmail.com'

from Autodesk.Revit.DB import *

uidoc=__revit__.ActiveUIDocument
doc = uidoc.Document

walls= FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType().ToElements()

t=Transaction(doc,"Assign Demolition to Walls")
t.Start()

for wall in walls:
    demo_wall_id=wall.DemolishedPhaseId
    demo_wall=wall.LookupParameter('Demolished_Wall')

    if demo_wall_id==0:
        demo_wall.Set('Yes')
    else:
        demo_wall.Set('No')
t.Commit()

