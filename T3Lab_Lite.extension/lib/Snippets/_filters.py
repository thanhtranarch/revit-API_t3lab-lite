# -*- coding: utf-8 -*-
"""
Filters Snippets

Code snippets for Revit element filter classes.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Filters Snippets"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================

from Autodesk.Revit.DB import *
from pyrevit.forms import alert
# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
doc   = __revit__.ActiveUIDocument.Document     # Document   class from RevitAPI that represents project. Used to Create, Delete, Modify and Query elements from the project.
uidoc = __revit__.ActiveUIDocument              # UIDocument class from RevitAPI that represents Revit project opened in the Revit UI.
app   = __revit__.Application                   # Represents the Autodesk Revit Application, providing access to documents, options and other application wide data and settings.

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝ FUNCTIONS
# ==================================================
def create_filter(key_parameter, element_value):
    """Function to create a RevitAPI filter."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_parameter_value = element_value  # e.g. element.Category.Id
    f_rule = FilterElementIdRule(f_parameter, FilterNumericEquals(), f_parameter_value)
    filter = ElementParameterFilter(f_rule)
    return filter

# EXAMPLE GET GROUP INSTANCE
# filter = create_filter(BuiltInParameter.ELEM_TYPE_PARAM, group_type_id)
# group = FilteredElementCollector(doc).WherePasses(filter).FirstElement()


def get_family_types(family_name):
    """Function to get FamilyTypes of a given FamilyName. It has to be written exactly the same."""
    pvp         = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME))
    condition   = FilterStringEquals()
    ruleValue   = family_name
    fRule       = FilterStringRule(pvp, condition, ruleValue, True)
    my_filter   = ElementParameterFilter(fRule)

    family_types = FilteredElementCollector(doc).WherePasses(my_filter).WhereElementIsElementType().ToElements()

    if not family_types:
        alert("Could not find a Family with a name: " + ruleValue, title = 'Family Not Found.', exitscript=True)

    return family_types