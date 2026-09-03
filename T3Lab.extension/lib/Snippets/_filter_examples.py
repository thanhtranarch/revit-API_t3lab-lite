# -*- coding: utf-8 -*-
"""
Filter Examples Snippets

Example code snippets for Revit element filtering.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Filter Examples Snippets"

#>>>>>>>>>> IMPORTS
import clr, os
from Autodesk.Revit.DB import *

#>>>>>>>>>> VARIABLES
# `__revit__` members are unavailable when no UIDocument is active, and at
# module scope that kills the import outright. Resolve defensively; the entry
# point reports the real problem (see Snippets._host.resolve_doc()).
try:
    doc = __revit__.ActiveUIDocument.Document
except Exception:
    doc = None
try:
    uidoc = __revit__.ActiveUIDocument
except Exception:
    uidoc = None
try:
    app = __revit__.Application
except Exception:
    app = None

#>>>>>>>>>> STRING FILTER
def create_string_filter(key_parameter, element_value, caseSensitive = True):
    """Function to create a RevitAPI filter."""
    f_parameter         = ParameterValueProvider(ElementId(key_parameter))  #sheet.SheetNumber
    f_parameter_value   = element_value #e.g. element.Category.Id           #element.GetPara
    caseSensitive       = True
    f_rule              = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value, caseSensitive)
    return ElementParameterFilter(f_rule)

#>>>>>>>>>> MAIN
if __name__ == '__main__':

    #>>>>>>>>>> ACTIVE VIEW
    view = doc.ActiveView
    #>>>>>>>>>> CREATE FILTER FROM VIEW's VIEWER SHEET NUMBER
    my_filter = create_string_filter(key_parameter = BuiltInParameter.SHEET_NUMBER,
                                     element_value= view.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER).AsString())
    #>>>>>>>>>> GET SHEET
    sheet = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().WherePasses(my_filter).FirstElement()

    #>>>>>>>>>> PRINT RESULTS
    if sheet:   print('Sheet Found: {} - {}'.format(sheet.SheetNumber, sheet.Name))
    else:       print('No sheet associated with the given view: {}'.format(view.Name))