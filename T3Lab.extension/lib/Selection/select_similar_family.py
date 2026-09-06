# -*- coding: utf-8 -*-
"""
Select Similar Family

Select all elements of the same family type as the selection.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Select Similar Family"
__doc__ = """Version = 1.0
Date    = 22.08.2022
_____________________________________________________________________
Description:
Select all instances in the project of the same Family.
_____________________________________________________________________
How-to:
- Select a single element
- Get All instances of the same family in Model
_____________________________________________________________________
Last update:
- [22.08.2022] - 1.0 RELEASE
_____________________________________________________________________
"""
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import clr
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
try:
    from Snippets._host import resolve_doc, resolve_uidoc, get_revit_version
except ImportError:
    try:
        import importlib
        import Snippets._host
        importlib.reload(Snippets._host)
        from Snippets._host import resolve_doc, resolve_uidoc, get_revit_version
    except Exception:
        from Snippets._host import resolve_doc, get_revit_version
        def resolve_uidoc(candidate=None):
            if candidate is not None and hasattr(candidate, 'Document') and candidate.Document is not None:
                return candidate
            try:
                from pyrevit import revit
                return revit.uidoc
            except Exception:
                return None
from Snippets._compat import make_eid

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝ FUNCTION
# ==================================================
def select_similar_by_family(uidoc=None, mode='model'):
    uidoc = resolve_uidoc(uidoc)
    if not uidoc:
        return
    doc = uidoc.Document
    selected_elements = list(uidoc.Selection.GetElementIds())

    try:
        # Check that Only single item is selected!
        if len(selected_elements) != 1:
            from pyrevit import forms
            forms.alert('You need to select only 1 element.', title=__title__, exitscript=True)

        selected_element = doc.GetElement(selected_elements[0])

        # CREATE FILTER RULE
        elem_type_id = selected_element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsElementId()
        elem_type = doc.GetElement(elem_type_id)
        elem_family_name = elem_type.FamilyName
        f_parameter = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME))
        f_parameter_value = elem_family_name

        rvt_year = get_revit_version(doc)
        if rvt_year and rvt_year < 2022:
            try:
                f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value, True)
            except Exception:
                f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value)
        else:
            f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value)

        # CREATE FILTER
        filter_family_name = ElementParameterFilter(f_rule)

        # GET ELEMENTS
        elements_by_f_name = []
        if mode   == 'model':
            elements_by_f_name = FilteredElementCollector(doc)\
                    .WherePasses(filter_family_name).WhereElementIsNotElementType().ToElementIds()
        elif mode == 'view':
            elements_by_f_name = FilteredElementCollector(doc, doc.ActiveView.Id)\
                    .WherePasses(filter_family_name).WhereElementIsNotElementType().ToElementIds()

        # SET SELECTION
        if elements_by_f_name:
            uidoc.Selection.SetElementIds(List[ElementId](elements_by_f_name))
    except:
        print('{} is not supported with this tool.'.format(type(selected_element)))