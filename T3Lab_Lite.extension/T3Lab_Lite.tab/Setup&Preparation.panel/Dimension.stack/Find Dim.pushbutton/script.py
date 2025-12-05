# -*- coding: utf-8 -*-
"""
Find View of Dimension
Search Dimensions placed in model by name and jump to their View

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Find Dim"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.4'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, script
from pyrevit.forms import ask_for_string, SelectFromList

# DEFINE VARIABLES
# ==================================================
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# CLASS/FUNCTIONS
# ==================================================
class DimensionResult(object):
    def __init__(self, dim_element, view_element):
        self.dim = dim_element
        self.view = view_element
        self.dim_id = dim_element.Id
        self.view_id = view_element.Id if view_element else None
        self.name = dim_element.Name if dim_element.Name else "<Unnamed>"
        self.view_name = view_element.Name if view_element else "Unknown View"

    def __str__(self):
        return "{}  [View: {}]".format(self.name, self.view_name)

def find_dimensions_by_name(name_filter):
    dims = FilteredElementCollector(doc)\
        .OfClass(Dimension)\
        .WhereElementIsNotElementType()\
        .ToElements()

    result = []
    for dim in dims:
        if name_filter.lower() in (dim.Name or "").lower():
            view = doc.GetElement(dim.OwnerViewId)
            if view:  # only dimensions with valid view
                result.append(DimensionResult(dim, view))
    return result

# MAIN SCRIPT
# ==================================================
search_key = ask_for_string(prompt='Enter dimension name or keyword:', default='DIM')

if not search_key:
    output.print_md('*No input provided.*')
else:
    dim_results = find_dimensions_by_name(search_key)
    if not dim_results:
        output.print_md('**No matching dimensions found for `{}`.**'.format(search_key))
    else:
        selected = SelectFromList.show(dim_results, multiselect=False, title='Select Dimension to jump to its View')
        if selected and selected.view:
            uidoc.ActiveView = selected.view
            uidoc.ShowElements(selected.dim_id)
            output.print_md('**Jumped to View `{}` containing Dimension `{}`.**'.format(selected.view_name, selected.name))
