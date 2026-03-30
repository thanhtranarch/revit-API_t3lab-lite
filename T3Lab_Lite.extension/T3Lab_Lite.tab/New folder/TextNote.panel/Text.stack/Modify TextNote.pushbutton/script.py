"""
Modify TextNote (py)
Modify Types, Format, Content,...
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Adjust TextNote (py)"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from rpw import *
from pyrevit import *
"""--------------------------------------------------"""
uidoc= __revit__.ActiveUIDocument
doc=__revit__.ActiveUIDocument.Document
activeview=doc.ActiveView.Id
"""--------------------------------------------------"""
textnotes=FilteredElementCollector(doc,activeview).OfCategory(BuiltInCategory.OST_TextNotes).ToElements()
textnotes_type=FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
"""--------------------------------------------------"""
# t=Transaction(doc,"Uper Typ. Text")
# t.Start()
# text_com=r'TYP.'
# for textnote in textnotes:
#     text=textnote.Text.strip()
#     uppertext=text.upper().strip()
#     if text!= text_com and uppertext==text_com:
#         textnote.Text=uppertext
# t.Commit()
"""--------------------------------------------------"""
text_com=r'TYP.'
texttype=r'1/16" Swis721 Lt Bt'
# texttype=r'3/32" Arial'
text_id=next((i.Id for i in textnotes_type if i.LookupParameter("Type Name").AsString().strip()==texttype),None)
t=Transaction(doc,"Change TextType")
t.Start()
for textnote in textnotes:
    if textnote.Text.strip()==text_com:
        textnote.ChangeTypeId(text_id)
t.Commit()

# text=textnote.Text.strip()
