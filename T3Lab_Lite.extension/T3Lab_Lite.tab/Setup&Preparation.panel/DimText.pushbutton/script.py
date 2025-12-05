# -*- coding: utf-8 -*-
"""
Text in Dimension

Adjust Dimension Text

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Text in Dimension"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# DEFINE VARIABLES
# ==================================================
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
active_view_id = uidoc.ActiveView.Id


# CLASS/FUNCTIONS
# ==================================================
def turnoff_leader(dim):
    ord_para_list = dim.GetOrderedParameters()
    for para in ord_para_list:
        if para.Definition.Name == "Leader":
            para.Set(0)

def setvalue(dim,prefix, sufix, above, below, sub_z, overide):
    if dim.HasOneSegment():
        dim.Above = above
        dim.Below = below
        dim.Prefix = prefix
        dim.Suffix = sufix
        dim.ValueOverride = overide
        if sub_z !=0:
                d.TextPosition = d.TextPosition.Add(XYZ(0, 0, sub_z))
    
    else:
        for d in dim.Segments:
            d.Above     = above 
            d.Below     = below
            d.Prefix    = prefix
            d.Suffix   = sufix 
            d.ValueOverride = overide
            if sub_z !=0:
                d.TextPosition = d.TextPosition.Add(XYZ(0, 0, sub_z))
                
# MAIN SCRIPT
# ==================================================
with Transaction(doc, "Adjust DimentionsText") as t:
    t.Start()
    count = 0
    #setdim
    # prefix="(E) ±"
    # below="V.I.F"
    
    #provide value for below and prefix
    prefix = ""
    sufix = ""
    above = ""
    # below = "CANOPY TO AC LEDGE SLAB EDGE"
    # below = "(E)"
    # below = "M&E SPACE"
    # below = "CANOPY TO SLAB EDGE"
    below = "CLEAR"
    # below = "CANOPY TO FACADE LINE"
    overide = ""
    sub_z=0
    
    selection = uidoc.Selection.GetElementIds()
    dims = [doc.GetElement(i) for i in selection if isinstance(doc.GetElement(i), Dimension)]

    for dim in dims:
        # turnoff_leader(dim)
        setvalue(dim,prefix, sufix, above, below, sub_z, overide)
        count += 1
    if count > 0:
        TaskDialog.Show("Text in Dimension", "Done!")
    t.Commit()