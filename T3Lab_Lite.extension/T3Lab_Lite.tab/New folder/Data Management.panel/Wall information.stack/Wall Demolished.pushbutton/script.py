# -*- coding: utf-8 -*-
"""
Demolished Wall
Check Demolished Walls

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Demolished Wall"

from Autodesk.Revit.DB import *

"""--------------------------------------------------"""
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""
walls   = FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType().ToElements()
phase_id= map(lambda x: x.Id , FilteredElementCollector(doc).OfClass(Phase).ToElements())


# t=Transaction(doc,"Assign Demolition to Walls")
# t.Start()
if walls:
    count = 0
    for wall in walls:
        demo_wall_id= wall.DemolishedPhaseId
        if demo_wall_id in phase_id:
            count+=1
    if count != 0:
        print("There are {} Demolished Wall".format(count))
    else:
        print("There is not Demolished Wall")

else:
    print("There is not Wall in project")
# t.Commit()

