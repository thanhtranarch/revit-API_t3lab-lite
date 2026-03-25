# -*- coding: utf-8 -*-
"""
Delete TextNoteType by Name
Search TextNoteTypes by name, select from list, and delete selected ones

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Del TextType"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.0'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, script
from pyrevit.forms import ask_for_string, SelectFromList

# DEFINE VARIABLES
# ==================================================
doc = revit.doc
output = script.get_output()

# CLASS/FUNCTIONS
# ==================================================
class TextNoteTypeItem(object):
    def __init__(self, txttype):
        self.txttype = txttype
        self.name = txttype.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        self.label = "{} [ID: {}]".format(self.name, txttype.Id)

    def __str__(self):
        return self.label

def collect_textnotetypes_by_name(keyword):
    types = FilteredElementCollector(doc).OfClass(TextNoteType).WhereElementIsElementType().ToElements()
    results = []
    for t in types:
        name = t.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        if keyword.lower() in (name or "").lower():
            results.append(TextNoteTypeItem(t))
    return results

# MAIN SCRIPT
# ==================================================
keyword = ask_for_string(prompt="Enter keyword to find TextNoteTypes:", default="LB")

if not keyword:
    output.print_md("*No keyword entered.*")
else:
    texttype_list = collect_textnotetypes_by_name(keyword)

    if not texttype_list:
        output.print_md("**No matching TextNoteTypes found.**")
    else:
        selected_types = SelectFromList.show(texttype_list, multiselect=True, title="Select TextNoteTypes to delete")
        if selected_types:
            t = Transaction(doc, "Delete Selected TextNoteTypes")
            t.Start()
            for item in selected_types:
                try:
                    doc.Delete(item.txttype.Id)
                except Exception as e:
                    output.print_md("Cannot delete '{}': {}".format(item.name, str(e)))
            t.Commit()
            output.print_md("**Deleted {} TextNoteType(s).**".format(len(selected_types)))
        else:
            output.print_md("*No TextNoteTypes selected for deletion.*")