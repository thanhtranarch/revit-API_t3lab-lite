# -*- coding: utf-8 -*-
"""
Central File
Create Central File

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Central File"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import (
    FilteredWorksetCollector,
    Workset,
    WorksetKind,
    Transaction,
    SaveAsOptions,
    WorksharingSaveAsOptions,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms

# DEFINE VARIABLES
# ==================================================    
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# CLASS/FUNCTIONS
# ==================================================
def create_central():
    file_name = forms.save_file()
    if file_name:
        file_name_without_extension = file_name.rsplit(".", 1)[0]
        central_file_path = file_name_without_extension + ".rvt"
        try:
            options = WorksharingSaveAsOptions()
            options.SaveAsCentral = True
            saveas_option = SaveAsOptions()
            saveas_option.MaximumBackups = 10
            saveas_option.OverwriteExistingFile = True
            saveas_option.SetWorksharingOptions(options)
            doc.SaveAs(central_file_path, saveas_option)
            TaskDialog.Show("Central Created", "Central File was created")
        except Exception as e:
            print("Failed to create Central File:", e)
    else:
        TaskDialog.Show("Central Created", "Please set path")

# MAIN SCRIPT
# ==================================================
if __name__ == '__main__':
    try:
        create_central()
    except Exception as e:
        print("Failed to create Central File:", e)
