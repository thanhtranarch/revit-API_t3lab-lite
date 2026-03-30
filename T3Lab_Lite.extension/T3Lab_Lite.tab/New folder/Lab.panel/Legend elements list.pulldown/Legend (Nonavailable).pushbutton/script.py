# -*- coding: utf-8 -*-
"""
Upper Text
Upper Text of NameView & TitleBlock
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Upper(ViewName&TitleBlock)"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, View, ViewType
from Autodesk.Revit.UI import TaskDialog

"""--------------------------------------------------"""
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""
views   = FilteredElementCollector(doc).OfClass(View).ToElements()
sheets  = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()


# Collect views and sheets
view_type=[ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation, ViewType.AreaPlan, ViewType.DraftingView, ViewType.Legend, ViewType.EngineeringPlan, ViewType.Section, ViewType.Detail]
views = FilteredElementCollector(doc).OfClass(View).ToElements()
sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
# Start a transaction
# t = Transaction(doc, "Upper(NameView&TitleBlock)")
# t.Start()



from collections import namedtuple
Viewname = namedtuple ("Viewname",["viewname","titleonsheet"])
sectionlist=[]
# Process views
for view in views:
    if view.ViewType == ViewType.Section:
        sheetnum = view.LookupParameter("Sheet Number")
        title_on_sheet = view.LookupParameter("Title on Sheet")
        if sheetnum:
            if sheetnum.AsString().startswith("A4"):
                viewnamesheet=Viewname(view.Name, title_on_sheet.AsString() )
                sectionlist.append(viewnamesheet)


sortviews=sorted(map(lambda x: x.titleonsheet,sectionlist))
for sortview in sortviews:
    print(sortview)
    print("\n")
# t.Commit()
