"""
Cycle Type
Quickly Cycle Type
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Types in Family"



from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import TaskDialog

"""--------------------------------------------------"""
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
"""--------------------------------------------------"""
if not doc.IsFamilyDocument:
    TaskDialog.Show('pyRevitPlus', 'Must be in Family Document.')
else:
    family_types = [x for x in doc.FamilyManager.Types]
    sorted_type_names = sorted([x.Name for x in family_types])
    current_type = doc.FamilyManager.CurrentType
    for n, type_name in enumerate(sorted_type_names):
        if type_name == current_type.Name:
            try:
                next_family_type_name = sorted_type_names[n + 1]
            except IndexError:
                next_family_type_name = sorted_type_names[0]
    for family_type in family_types:
        if family_type.Name == next_family_type_name:
            t = Transaction(doc, 'Cycle Type')
            t.Start()
            doc.FamilyManager.CurrentType = family_type
            t.Commit()
