__author__ ="Tran Tien Thanh"
__title__ = "CAD to model"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import xlrd
from pyrevit import forms
"""--------------------------------------------------"""
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""

from pyrevit import revit, DB

# Function to read data from Excel
def read_excel_data(file_path):
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    data = []
    for row_num in range(1, sheet.nrows):  # Skip header row
        row_data = [sheet.cell_value(row_num, col_num) for col_num in range(sheet.ncols)]
        data.append(row_data)
    return data

# Function to display data in terminal
def display_data_in_terminal(data):
    for row in data:
        print('\t'.join(str(item) for item in row))

# Main function
def main():
    # Path to Excel file
    excel_file_path = r"C:\Users\trant\OneDrive\Desktop\Book1.xlsx"

    # Read data from Excel
    data = read_excel_data(excel_file_path)

    # Display data in terminal
    display_data_in_terminal(data)



# t = Transaction(doc, "Create Wall Type")
# t.Start()
#
# wall_type_name = "Test 2"
#
# # Find an existing basic wall type to use as a template
# basic_wall_type = FilteredElementCollector(doc).OfClass(WallType) \
#     .WhereElementIsElementType().FirstElement()
# print(basic_wall_type)
# if basic_wall_type:
#     # Duplicate the existing basic wall type to create a new wall type
#     new_wall_type = basic_wall_type.Duplicate(wall_type_name)
#     t.Commit()
#     print("New wall type created successfully.")