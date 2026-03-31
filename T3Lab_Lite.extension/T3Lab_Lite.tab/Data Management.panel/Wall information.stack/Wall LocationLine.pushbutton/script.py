# -*- coding: utf-8 -*-
"""
Wall LocationLine
Set Wall LocationLine

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Wall LocationLine"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

"""--------------------------------------------------"""
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""
walls   = FilteredElementCollector(doc).OfClass(Wall).WhereElementIsNotElementType().ToElements()
selection=uidoc.Selection.GetElementIds()
wall_selected=[doc.GetElement(i) for i in selection if isinstance(doc.GetElement(i), Wall)]
#0
#1
#2
#3
#4
t = Transaction(doc, "Set Wall LocationLine")
t.Start()
for wall in wall_selected:
    for para in wall.GetOrderedParameters():
        if para.Definition.Name == "Location Line":
            # print(para.AsValueString())
            para.Set(4)
            # print(WallLocationLine.CoreExterior)

t.Commit()

