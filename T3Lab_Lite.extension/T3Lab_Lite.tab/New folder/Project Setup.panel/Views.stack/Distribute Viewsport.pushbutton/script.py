# -*- coding: utf-8 -*-
"""
Distribute Viewsport

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""
__author__  = "Tran Tien Thanh"
__title__   = "Distribute Viewsport"


import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
import math

from rpw.ui.forms import *
from collections import *
"""--------------------------------------------------"""
uidoc= __revit__.ActiveUIDocument
doc=__revit__.ActiveUIDocument.Document
activeview=doc.ActiveView.Id
"""--------------------------------------------------"""
tb_active=FilteredElementCollector(doc,activeview).OfCategory(BuiltInCategory.OST_TitleBlocks).FirstElement()
viewports=FilteredElementCollector(doc,activeview).OfClass(Viewport).ToElements()
"""--------------------------------------------------"""

class DistributeViewPort:
    def __init__(self,tb_active):
        self.element=tb_active
        self.name=tb_active.Name
        self.width=tb_active.LookupParameter("Sheet Width").AsDouble()
        self.height=tb_active.LookupParameter("Sheet Height").AsDouble()
    def checkFamily(self):
        family = self.element.LookupParameter("Family").AsValueString()
        if "Title Block - D" in family:
            grider_l, grider_r, grider_t = 0.183333333333342, 0.250000000000000, 0.0937500000000107
        elif "Title Block - E " in family:
            grider_l, grider_r, grider_t = 0, 0, 0
        elif "Title Block - E1" in family:
            grider_l, grider_r, grider_t = 0.0416666666666767, 0.250000000000000, 0.0416666666666767
        else:
            # Define default values if family not recognized
            grider_l, grider_r, grider_t = 0, 0, 0
        grider_b = grider_t  # Assuming bottom margin same as top
        return grider_l, grider_r, grider_t, grider_b,family
    def checkType(self):
        grider_l, grider_r, grider_t, grider_b, family = self.checkFamily()
        if "(1)Legend" in self.name:
            column_size = 0.333333333333336
        elif "(2)Legend" in self.name:
            column_size = 0.66666666666667
        else:
            column_size = 0
        layout_region_w = self.width - (grider_l + grider_r) - column_size
        layout_region_h = self.height - (grider_t + grider_b)
        return layout_region_w, layout_region_h, column_size
    def distribute(self,number_row,number_column,viewports):
        grider_l, grider_r, grider_t, grider_b,family = self.checkFamily()
        layout_region_w, layout_region_h,column_size=self.checkType()
        strip = 0.0416666666666663  # 1/2"
        for vp in viewports:
            try:
                para = vp.LookupParameter("Detail Number")
                if  para:
                    key=para.AsString()
                    box_w = layout_region_w/number_column
                    box_h = layout_region_h/number_row
                    ratio = number_row**(-1)
                    number= math.ceil(float(key)*ratio)
                    key_location=int(key) % number_row
                    if "Title Block - D" in family:
                        new_X = box_w * (number_column - number+0.5) + grider_l #need check line code
                    elif "Title Block - E1" in family:
                        new_X = box_w * (number_column - number+0.5) + grider_l - 0.5  #need check line code
                    if key_location!=0:
                        new_Y = grider_b+(box_h-strip)*0.5+strip+(box_h*(key_location-1))
                    else:
                        new_Y = grider_b+(box_h-strip)*0.5+strip+(box_h*(number_row-1))

                    vp_z=vp.GetBoxCenter().Z
                    vp.SetBoxCenter(XYZ(new_X,new_Y,vp_z))
            except:
                pass

"""--------------------------------------------------"""
components=[Label("Row"),TextBox("number_row",Text=""), Label("Column"),TextBox("number_column", Text=""),Separator(),Button("Run")]
"""--------------------------------------------------"""
form=FlexForm("Grid View",components)
form.show()
value=form.values
"""--------------------------------------------------"""
# General Information for layout
if value:
    number_row      = int(value['number_row'])
    number_column   = int(value['number_column'])
# General Information for 24x36(D Archs)
    if tb_active:
        tb=DistributeViewPort(tb_active)
        a=tb.checkFamily()
        b=tb.checkType()
        t = Transaction(doc, "Change Location of Viewport")
        t.Start()
        tb.distribute(number_row, number_column, viewports)
        t.Commit()
    else:
        pass
        # TaskDialog.Show("Distribute - Viewsport", "None Title Block in Sheet")
else:
    # TaskDialog.Show("Distribute - Viewsport", "Please enter information about rows and columns")
    print("Please enter information about rows and columns")