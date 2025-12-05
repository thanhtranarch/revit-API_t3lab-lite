# -*- coding: utf-8 -*-
"""
Delete Dimensions by Name
Search Dimensions by name, select from list, and delete selected ones

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Del Dim"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.0'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, script
from pyrevit.forms import ask_for_string, SelectFromList

# DEFINE VARIABLES
# ==================================================
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# CLASS/FUNCTIONS
# ==================================================
class DimensionItem(object):
    def __init__(self, dim):
        self.dim = dim
        self.name = dim.Name if dim.Name else "<Unnamed>"
        self.view = doc.GetElement(dim.OwnerViewId)
        self.view_name = self.view.Name if self.view else "Unknown View"
        self.label = "{}  [View: {}]".format(self.name, self.view_name)
        self.dim_id = dim.Id

    def __str__(self):
        return self.label

def collect_dimensions_by_name(keyword):
    dims = FilteredElementCollector(doc)\
        .OfClass(Dimension)\
        .WhereElementIsNotElementType()\
        .ToElements()

    results = []
    for dim in dims:
        if keyword.lower() in (dim.Name or "").lower():
            results.append(DimensionItem(dim))
    return results

# MAIN SCRIPT
# ==================================================
keyword = ask_for_string(prompt="Enter keyword to find dimensions:", default="DIM")

if not keyword:
    output.print_md("*No keyword entered.*")
else:
    dim_list = collect_dimensions_by_name(keyword)

    if not dim_list:
        output.print_md("**No matching dimensions found.**")
    else:
        selected_dims = SelectFromList.show(dim_list, multiselect=True, title="Select dimensions to delete")
        if selected_dims:
            t = Transaction(doc, "Delete Selected Dimensions")
            t.Start()
            for item in selected_dims:
                doc.Delete(item.dim_id)
            t.Commit()
            output.print_md("**Deleted {} dimension(s).**".format(len(selected_dims)))
        else:
            output.print_md("*No dimensions selected for deletion.*")
