# -*- coding: utf-8 -*-
"""
Upper Text
Convert View Names and TitleBlock text to uppercase.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__ = "Upper(ViewName&TitleBlock)"

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, View, ViewType
)
from pyrevit import revit, script

logger = script.get_logger()
doc = revit.doc

SUPPORTED_VIEW_TYPES = [
    ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation,
    ViewType.AreaPlan, ViewType.DraftingView, ViewType.Legend,
    ViewType.EngineeringPlan, ViewType.Section, ViewType.Detail
]


def convert_to_upper(text):
    if text and any(c.islower() for c in text):
        return text.upper()
    return text


views = FilteredElementCollector(doc).OfClass(View).ToElements()
sheets = (FilteredElementCollector(doc)
          .OfCategory(BuiltInCategory.OST_TitleBlocks)
          .WhereElementIsNotElementType()
          .ToElements())

with revit.Transaction("Upper(NameView&TitleBlock)"):
    for view in views:
        if view.ViewType not in SUPPORTED_VIEW_TYPES:
            continue
        try:
            view.Name = convert_to_upper(view.Name)
            title_param = view.LookupParameter("Title on Sheet")
            if title_param and title_param.AsString():
                title_param.Set(convert_to_upper(title_param.AsString()))
        except Exception as e:
            logger.warning("Error processing view '{}': {}".format(view.Name, e))

    for sheet in sheets:
        try:
            sheet_name_param = sheet.LookupParameter("Sheet Name")
            if sheet_name_param and sheet_name_param.AsString():
                sheet_name_param.Set(convert_to_upper(sheet_name_param.AsString()))
        except Exception as e:
            logger.warning("Error processing title block {}: {}".format(sheet.Id, e))
