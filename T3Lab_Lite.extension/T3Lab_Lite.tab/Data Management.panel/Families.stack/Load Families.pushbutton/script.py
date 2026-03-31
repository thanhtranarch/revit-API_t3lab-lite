"""
Load Families
To load multiple families into the project at once

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Load Families"

import os
from Autodesk.Revit.DB import Document, Family, Transaction
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
from pyrevit import script

logger = script.get_logger()
output = script.get_output()

doc = __revit__.ActiveUIDocument.Document


def load_family(family_path):
    t = Transaction(doc, "Load Family")
    t.Start()

    try:
        doc.LoadFamily(family_path)
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        print("Error loading family:", str(e))
        return False

directory = r"A:\4 KTA Project\00. Library\01. Family Revit"

if not directory or not os.path.exists(directory):
    logger.error("Invalid or empty directory.")
    script.exit()

family_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(".rfa")]
family_name=[f for f in os.listdir(directory) if f.endswith(".rfa")]
selected_families = forms.SelectFromList.show(family_name,
                                              title="Select Families",
                                              width=500,
                                              button_name="Load Families",
                                              multiselect=True)
count=0
if selected_families:

    for selected_family in selected_families:
        for family_file in family_files:
            if selected_family == os.path.basename(family_file):
                count+=1
                success = load_family(family_file)
if count != 0:
    TaskDialog.Show("Load Family","Family(s) has been loaded into the project")
else:
    forms.alert("Please select at least a family", title="Load Family")