from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms

uidoc= __revit__.ActiveUIDocument
doc=__revit__.ActiveUIDocument.Document
active_view_id= doc.ActiveView.Id


views = FilteredElementCollector(doc).OfClass(View).ToElements()

t = Transaction(doc, "Replace View Name")
t.Start()

# for view in views:
#     sheet_number_param = view.LookupParameter("Sheet Number")
#     detail_number_param = view.LookupParameter("Detail Number")
#     try:
#         if sheet_number_param and (sheet_number_param.AsString() in ("A562", "A563")) and detail_number_param:
#             new_name =sheet_number_param.AsString()+ " CORRIDOR - 2ND-6TH ELEVATION " + detail_number_param.AsString()
#             view.Name = new_name
#             new_title = "CORRIDOR - 2ND-6TH ELEVATION " + detail_number_param.AsString()
#             a=view.LookupParameter("Title on Sheet").Set(new_title)
#     except Exception as e:

t.Commit()
# for view in views:
#     word="Post_"
#     if view.Name.startswith(word):
#         replace_word = "POST - "
#         namesort=view.Name[len(word):]
#         view.Name=replace_word + namesort
# for view in views:
#     word="OPENINGS"
#     primary_group=view.LookupParameter("Primary View Group")
#     if primary_group:
#         if view.Name.endswith(word) and primary_group.AsString() == "03_elevations":
#             replace_word = "ALLOWABLE OPENING - "
#             namesort=view.Name[:-len(word)]
#             view.Name=replace_word + namesort
# for view in views:
#     word="INTERIOR ELEVATIONS"
#     primary_group=view.LookupParameter("Primary View Group")
#     if primary_group:
#         if view.Name.startswith(word) and primary_group.AsString() == "03_elevations":
#             replace_word = "PARKING OPENING - "
#             namesort=view.Name
#             view.Name=replace_word + namesort
# t = Transaction(doc, "Replace Ter_View Group")
# t.Start()
# for view in views:
#     primary_group = view.LookupParameter("Primary View Group")
#     second_group  = view.LookupParameter("Secondary View Group")
#     ter_group     = view.LookupParameter("Tertiary View Group")
#     sheetnumber   = view.LookupParameter("Sheet Number")
#     if primary_group and sheetnumber:
#
#         if primary_group.AsString()=="09_details" and sheetnumber.AsString().startswith("G3"):
#             ter_group.Set("Gs - 300 signage")
# for view in views:
#     primary_group = view.LookupParameter("Primary View Group")
#     second_group  = view.LookupParameter("Secondary View Group")
#     ter_group     = view.LookupParameter("Tertiary View Group")
#     sheetnumber   = view.LookupParameter("Sheet Number")
#     if primary_group and sheetnumber:
#         if primary_group.AsString()=="09_details" and sheetnumber.AsString().startswith("A99"):
#             ter_group.Set("090 mailbox")
# for view in views:
#     primary_group = view.LookupParameter("Primary View Group")
#     second_group  = view.LookupParameter("Secondary View Group")
#     ter_group     = view.LookupParameter("Tertiary View Group")
#     sheetnumber   = view.LookupParameter("Sheet Number")
#     if primary_group and sheetnumber:
#         if not sheetnumber.AsString().startswith("-") :
#             second_group.Set("sheet views")
# t.Commit()
# t.Commit()
