# -*- coding: utf-8 -*-
"""
Upper Text
Upper TextNote
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "UPPER"

from Autodesk.Revit.DB import FilteredElementCollector, TextNote, Transaction

"""--------------------------------------------------"""
doc             = __revit__.ActiveUIDocument.Document
active_view_id  = doc.ActiveView.Id
"""--------------------------------------------------"""
# Collect text_notes in view
text_notes=FilteredElementCollector(doc, active_view_id).OfClass(TextNote).ToElements()

def uppercheck(text):
    if any(c.islower() for c in text):
        return False
    return True

# Start a transaction
t = Transaction(doc, "Upper TextNote")
t.Start()
for text in text_notes:
    if not uppercheck(text.Text) :
        text.Text=text.Text.upper()
t.Commit()
