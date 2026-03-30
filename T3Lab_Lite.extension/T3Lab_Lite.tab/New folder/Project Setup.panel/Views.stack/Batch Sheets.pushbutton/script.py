# -*- coding: utf-8 -*-
"""
Batch Sheets
Create Multi Sheets

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""
__author__  = "Tran Tien Thanh"
__title__   = "Batch Sheets"


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

