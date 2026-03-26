# -*- coding: utf-8 -*-
"""
Make 3D View

Create a 3D view from the current view or selection.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Make 3D View"

from pyrevit import revit, HOST_APP

for model in __models__:
    doc = HOST_APP.app.OpenDocumentFile(model)
    with revit.Transaction("Create ThreeDView", doc=doc):
        revit.create.create_3d_view("ThreeDView", doc=doc)
    doc.Save()
