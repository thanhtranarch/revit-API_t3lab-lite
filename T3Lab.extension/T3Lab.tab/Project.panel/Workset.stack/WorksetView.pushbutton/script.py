# -*- coding: utf-8 -*-
"""
Workset Views
Create View with Workset

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Workset Views"
__version__ = "1.0.0"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
"""--------------------------------------------------"""
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""

if doc.IsWorkshared:
    ws_colects = set(ws for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
else:
    forms.alert("No user worksets found in the workshared document.")

views=[view.Name for view in FilteredElementCollector(doc).OfClass(View).ToElements()]
viewtypes = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
threeD_viewid=next((viewtype.Id for viewtype in viewtypes if viewtype.ViewFamily==ViewFamily.ThreeDimensional),None)


if threeD_viewid is None:
    forms.alert("No 3D view family type found.")
    raise ValueError("No 3D view family type found")

def create_3d_views(worksets):
    t = Transaction(doc, "Create 3D Workset")
    t.Start()
    count=0
    try:
        for ws in worksets:
            if not (ws.Name in views):
                view3d = View3D.CreateIsometric(doc, threeD_viewid)
                count+=1
                if view3d:
                    view3d.Name = ws.Name
                    for ws in worksets:
                        if view3d.Name == ws.Name:
                            view3d.SetWorksetVisibility(ws.Id, WorksetVisibility.Visible)
                        else:
                            view3d.SetWorksetVisibility(ws.Id, WorksetVisibility.Hidden)
        if count==0:
            forms.alert("3D Workset(s) was existed")
        else:
            TaskDialog.Show("Create 3D Workset(s)", "3D Workset(s) was create, place at ???")
    except Exception as e:
        t.RollBack()
        forms.alert("An error occurred: {}".format(e))
    else:
        t.Commit()

create_3d_views(ws_colects)
