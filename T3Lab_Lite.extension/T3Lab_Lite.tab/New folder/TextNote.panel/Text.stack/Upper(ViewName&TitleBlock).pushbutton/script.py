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
t = Transaction(doc, "Upper(NameView&TitleBlock)")
t.Start()


# Function to convert text to uppercase if it contains lowercase letters
def convert_to_upper(text):
    if any(c.islower() for c in text):
        return text.upper()
    return text


# Process views
for view in views:
    if view.ViewType in view_type:
        try:
            view.Name = convert_to_upper(view.Name)
            title_on_sheet_param = view.LookupParameter("Title on Sheet")
            if title_on_sheet_param:
                title_on_sheet_param.Set(convert_to_upper(title_on_sheet_param.AsString()))
        except Exception as e:
            print("Error processing view:", e)
            print(view.Name)
# Process sheets
for sheet in sheets:
    try:
        sheet_name_param = sheet.LookupParameter("Sheet Name")
        if sheet_name_param:
            sheet_name_param.Set(convert_to_upper(sheet_name_param.AsString()))
    except Exception as e:
        print("Error processing sheet:", e)
t.Commit()
