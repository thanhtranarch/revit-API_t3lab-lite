# -*- coding: utf-8 -*-
"""
Upper All Text

Convert View Names, TitleBlock parameters, and Dimension text overrides
to uppercase in the current project/view.

Author: Tran Tien Thanh
"""

__author__ = "Tran Tien Thanh"
__title__ = "Upper All Text\n(Views, Sheets & Dims)"
__version__ = "1.0.0"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, View, ViewType, Dimension
from Autodesk.Revit.UI import TaskDialog

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
active_view_id = uidoc.ActiveView.Id

def convert_to_upper(text):
    if text and any(c.islower() for c in text):
        return text.upper()
    return text

def convert_text_to_uppercase(text):
    if text is None or text == "":
        return ""
    return str(text).upper()

def update_dimension_text_to_uppercase(dim):
    try:
        if dim.HasOneSegment():
            dim.Above = convert_text_to_uppercase(dim.Above if dim.Above else "")
            dim.Below = convert_text_to_uppercase(dim.Below if dim.Below else "")
            dim.Prefix = convert_text_to_uppercase(dim.Prefix if dim.Prefix else "")
            dim.Suffix = convert_text_to_uppercase(dim.Suffix if dim.Suffix else "")
            if dim.ValueOverride:
                dim.ValueOverride = convert_text_to_uppercase(dim.ValueOverride)
        else:
            for segment in dim.Segments:
                segment.Above = convert_text_to_uppercase(segment.Above if segment.Above else "")
                segment.Below = convert_text_to_uppercase(segment.Below if segment.Below else "")
                segment.Prefix = convert_text_to_uppercase(segment.Prefix if segment.Prefix else "")
                segment.Suffix = convert_text_to_uppercase(segment.Suffix if segment.Suffix else "")
                if segment.ValueOverride:
                    segment.ValueOverride = convert_text_to_uppercase(segment.ValueOverride)
        return True
    except Exception:
        return False

def main():
    view_type = [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation, 
                 ViewType.AreaPlan, ViewType.DraftingView, ViewType.Legend, 
                 ViewType.EngineeringPlan, ViewType.Section, ViewType.Detail]
    views = FilteredElementCollector(doc).OfClass(View).ToElements()
    sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
    dimensions = FilteredElementCollector(doc, active_view_id).OfClass(Dimension).ToElements()
    
    with Transaction(doc, "Upper All Text") as t:
        t.Start()
        
        # 1. Process views
        for view in views:
            if view.ViewType in view_type:
                try:
                    view.Name = convert_to_upper(view.Name)
                    title_on_sheet_param = view.LookupParameter("Title on Sheet")
                    if title_on_sheet_param and title_on_sheet_param.AsString():
                        title_on_sheet_param.Set(convert_to_upper(title_on_sheet_param.AsString()))
                except Exception:
                    pass
                    
        # 2. Process sheets
        for sheet in sheets:
            try:
                sheet_name_param = sheet.LookupParameter("Sheet Name")
                if sheet_name_param and sheet_name_param.AsString():
                    sheet_name_param.Set(convert_to_upper(sheet_name_param.AsString()))
            except Exception:
                pass
                
        # 3. Process dimensions
        success_dim_count = 0
        for dim in dimensions:
            if update_dimension_text_to_uppercase(dim):
                success_dim_count += 1
                
        t.Commit()
        
    TaskDialog.Show("Upper All Text", "Finished converting Views, Sheets, and {} Dimensions in current view to UPPERCASE.".format(success_dim_count))

if __name__ == "__main__":
    main()
