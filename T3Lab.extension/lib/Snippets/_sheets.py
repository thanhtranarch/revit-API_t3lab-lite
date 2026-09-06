# -*- coding: utf-8 -*-
"""
Sheets Snippets

Code snippets for working with Revit sheets.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Sheets Snippets"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import UIDocument
try:
    from Snippets._host import resolve_doc, host_uiapp
    default_doc, _ = resolve_doc()
    _uiapp = host_uiapp()
    default_uidoc = _uiapp.ActiveUIDocument if _uiapp else None
except Exception:
    default_doc = None
    default_uidoc = None




def get_views_on_sheet(sheet, uidoc=None):
    """Function to return all views found on the given sheet."""
    doc = (sheet.Document if sheet and hasattr(sheet, 'Document') else None) or (uidoc.Document if uidoc else default_doc)
    if not doc:
        return []
    viewports_ids   = sheet.GetAllViewports()
    viewports       = [doc.GetElement(viewport_id)  for viewport_id in viewports_ids]
    views_ids       = [viewport.ViewId              for viewport    in viewports]
    views           = [doc.GetElement(view_id)      for view_id     in views_ids]
    return views


def get_titleblock_on_sheet(sheet, uidoc=None):
    """Function to get TitleBlock from given ViewSheet.
    It will not return any TitleBlocks if there are more than 1 on ViewSheet.
    :returns TitleBlock"""
    #TODO THIS FUNCTION IS OBSOLETE
    doc = (sheet.Document if sheet and hasattr(sheet, 'Document') else None) or (uidoc.Document if uidoc else default_doc)
    if not doc:
        return None

    all_TitleBlocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
    title_blocks_on_sheet = []

    for title_block in all_TitleBlocks:
        if title_block.OwnerViewId == sheet.Id:
            title_blocks_on_sheet.append(title_block)

    if not title_blocks_on_sheet:
        print("***No TitleBlocks were found on given ViewSheet ({}***".format(sheet.SheetNumber))

    elif len(title_blocks_on_sheet) > 1:
        print("***There are more than 1 TitleBlock on given ViewSheet ({})****".format(sheet.SheetNumber))

    else:
        return title_blocks_on_sheet[0]


def get_titleblocks_from_sheet(sheet, uidoc=None):
    #type:(ViewSheet, UIDocument) -> list
    """Function to get TitleBlocks from the given ViewSheet.
    :param sheet: ViewSheet that has TitleBlock/
    :param uidoc: UIDocument of the Project
    :return:      list of TitleBlocks that are placed on the given Sheet."""
    doc = (sheet.Document if sheet and hasattr(sheet, 'Document') else None) or (uidoc.Document if uidoc else default_doc)
    if not doc:
        return []

    # CREATE A RULE
    rule_value = sheet.SheetNumber
    param_sheet_number = ElementId(BuiltInParameter.SHEET_NUMBER)
    f_pvp = ParameterValueProvider(param_sheet_number)
    evaluator = FilterStringEquals()
    try:
        # Revit 2022+
        f_rule = FilterStringRule(f_pvp, evaluator, rule_value)
    except Exception:
        # Revit 2021 and earlier
        f_rule = FilterStringRule(f_pvp, evaluator, rule_value, True)

    # CREATE A FILTER
    tb_filter = ElementParameterFilter(f_rule)

    tb = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks) \
        .WhereElementIsNotElementType().WherePasses(tb_filter).ToElements()

    return list(tb)

