# -*- coding: utf-8 -*-
"""
Remove Legend & Schedule not on sheet

This script will loop through sheets to check if legend and schedule are present on the sheet or not.

Author: Tran Tien Thanh
Email: trantienthanh909@gmail.com
LinkedIn: linkedin.com/in/sunarch7899/
"""

__author__ ="Tran Tien Thanh"
__title__ = "Remove View"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# DEFINE VARIABLES
# ==================================================
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
active_view_id = doc.ActiveView.Id

# MAIN SCRIPT
# ==================================================
#get elements
views = FilteredElementCollector(doc).OfClass(View).ToElements()
legend = [i for i in views if i.ViewType == ViewType.Legend]
schedule = [i for i in views if i.ViewType == ViewType.Schedule]
sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

viewonsheets_lst = []
# Iterate through each sheet
for sheet in sheets:
    # Print the sheet name
    # print("Sheet Name: ", sheet.Name)
    
    # Get the views on the sheet
    view_ids = sheet.GetAllPlacedViews()
    
    # Iterate through each view on the sheet
    
    for view_id in view_ids:
        view = doc.GetElement(view_id)
        viewonsheets_lst.append(view.Name)
with Transaction(doc, "Delete Legend&Schedule not on Sheet") as t:
    t.Start()
    for i in legend:
        if i.Name not in viewonsheets_lst:
            doc.Delete(i.Id)
    for i in schedule:
        if i.Name not in viewonsheets_lst:
            doc.Delete(i.Id)
    t.Commit()