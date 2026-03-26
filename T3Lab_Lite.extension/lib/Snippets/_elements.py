# -*- coding: utf-8 -*-
"""
Elements Snippets

Code snippets for common Revit element operations.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Elements Snippets"

from Autodesk.Revit.DB import *



all_floor_types = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_Floors).WhereElementIsElementType().ToElements()

def dict_name_element(given_elements, dotNet=False):
    dict_output = {Element.Name.GetValue(fr): fr for fr in given_elements}


    return dict_output
