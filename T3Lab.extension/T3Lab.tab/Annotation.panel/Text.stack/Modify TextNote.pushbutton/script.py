# -*- coding: utf-8 -*-
"""
Modify TextNote
Change TextNote type for matching text in the active view.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__ = "Adjust TextNote (py)"
__version__ = "1.0.0"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, TextNoteType
from pyrevit import revit, script

logger = script.get_logger()
doc = revit.doc

active_view_id = doc.ActiveView.Id

textnotes = (FilteredElementCollector(doc, active_view_id)
             .OfCategory(BuiltInCategory.OST_TextNotes)
             .ToElements())
textnotes_type = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()

text_com = r'TYP.'
texttype = r'1/16" Swis721 Lt Bt'
text_id = next(
    (i.Id for i in textnotes_type
     if i.LookupParameter("Type Name").AsString().strip() == texttype),
    None
)

if text_id is None:
    logger.warning("TextNoteType '{}' not found in document.".format(texttype))
else:
    with revit.Transaction("Change TextType"):
        for textnote in textnotes:
            if textnote.Text.strip() == text_com:
                textnote.ChangeTypeId(text_id)
