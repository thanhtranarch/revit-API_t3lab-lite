# -*- coding: utf-8 -*-
"""
Legend to Excel

Export legend data to Excel file

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ ="Tran Tien Thanh"
__title__ = "From Legend"

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
clr.AddReference("System")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ViewType, ElementId
from RevitServices.Persistence import DocumentManager
from System import Enum
import openpyxl

# Get the active Revit document
doc = DocumentManager.Instance.CurrentDBDocument

# Function to export legend data to Excel
def export_legend_to_excel(legend_view_name, excel_path):
    # Find the legend view by name
    legends = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.View)
    legend_view = None
    
    for view in legends:
        if view.ViewType == ViewType.Legend and view.Name == legend_view_name:
            legend_view = view
            break
    
    if legend_view is None:
        print("Legend view '{legend_view_name}' not found.".format(legend_view_name))
        return

    # Collect all text notes and symbols in the legend
    elements = FilteredElementCollector(doc, legend_view.Id).WhereElementIsNotElementType().ToElements()

    # Create a new Excel workbook and sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = legend_view_name

    # Write headers
    sheet.append(["Element ID", "Element Type", "Text"])

    # Iterate over elements in the legend
    for element in elements:
        element_id = element.Id.IntegerValue
        element_type = element.GetType().Name

        if hasattr(element, 'Text'):
            text = element.Text
        else:
            text = "N/A"

        # Write element data to Excel
        sheet.append([element_id, element_type, text])

    # Save the Excel file
    workbook.save(excel_path)
    print("Legend '{legend_view_name}' exported to '{excel_path}'.".format(legend_view_name,excel_path))

# Define the legend view name and Excel file path
legend_view_name = select # Replace with your actual legend view name
excel_path = select_paths(r"C:\path\to\your\output.xlsx")  # Replace with your desired Excel file path

# Export the legend to Excel
export_legend_to_excel(legend_view_name, excel_path)