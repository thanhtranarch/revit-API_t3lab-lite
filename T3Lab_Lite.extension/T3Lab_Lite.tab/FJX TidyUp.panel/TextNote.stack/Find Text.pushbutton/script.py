# -*- coding: utf-8 -*-
"""
Find View of TextNote
Search TextNotes placed in model by name and jump to their View

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Find Text"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.0'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, script
from pyrevit.forms import ask_for_string, SelectFromList

# DEFINE VARIABLES
# ==================================================
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# CLASS/FUNCTIONS
# ==================================================
class TextNoteResult(object):
    def __init__(self, textnote, view_element):
        self.textnote = textnote
        self.view = view_element
        self.text_id = textnote.Id
        self.view_id = view_element.Id if view_element else None
        self.name = textnote.Text if textnote.Text else "<Empty Text>"
        self.view_name = view_element.Name if view_element else "Unknown View"

    def __str__(self):
        return "{}  [View: {}]".format(self.name, self.view_name)

def find_textnotes_by_content(keyword):
    textnotes = FilteredElementCollector(doc)\
        .OfClass(TextNote)\
        .WhereElementIsNotElementType()\
        .ToElements()

    results = []
    for tn in textnotes:
        if keyword.lower() in (tn.Text or "").lower():
            view = doc.GetElement(tn.ViewId)
            if view:
                results.append(TextNoteResult(tn, view))
    return results

# MAIN SCRIPT
# ==================================================
search_key = ask_for_string(prompt='Enter TextNote content or keyword:', default='NOTE')

if not search_key:
    output.print_md('*No input provided.*')
else:
    tn_results = find_textnotes_by_content(search_key)
    if not tn_results:
        output.print_md('**No matching TextNotes found for `{}`.**'.format(search_key))
    else:
        selected = SelectFromList.show(tn_results, multiselect=False, title='Select TextNote to jump to its View')
        if selected and selected.view:
            uidoc.ActiveView = selected.view
            uidoc.ShowElements(selected.text_id)
            output.print_md('**Jumped to View `{}` containing TextNote: `{}`.**'.format(selected.view_name, selected.name))