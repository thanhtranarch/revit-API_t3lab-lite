# -*- coding: utf-8 -*-
"""
Vertical Bottom
Align Bottom TextNote

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__  = "Tran Tien Thanh"
__title__   = "Bottom"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.UI import *
from collections import namedtuple

import unittest
import logging
import sys
import os

# logging.basicConfig(filename="aligntext.log", level=logging.NOTSET, format='[%(asctime)s] - %(name)s - %(levelname)s - %(message)s - %(thread)d')
# logger=logging.getLogger(__name__)
# logger.info("Check bug rất mê mõi!!!")
# log_directory = os.path.dirname(os.path.abspath("aligntext.log"))
# print("Log file directory:", log_directory)

doc     = __revit__.ActiveUIDocument.Document
uidoc   = __revit__.ActiveUIDocument
active_view_id = uidoc.ActiveView.Id


import s_align



LeaderLocation = namedtuple("LeaderLocation", ["obj","obj_leader","anchor","elbow", "end"])

textnotes=FilteredElementCollector(doc,active_view_id).OfClass(TextNote).WhereElementIsNotElementType().ToElementIds()
selection=uidoc.Selection.GetElementIds()
textnote_selected = [doc.GetElement(i) for i in selection if isinstance(doc.GetElement(i), TextNote)]
center=0
middle=0

t = Transaction(doc, "Align TextNote")
t.Start()

textnote_class = s_align.SetTextNote(*textnote_selected)
textnote_class.leaderset(LeaderAtachement.TopLine)
leader_location_list=s_align.get_leader(textnote_selected)
textnote_class.align_text("bottom", center, middle)
s_align.set_leaderdefault(leader_location_list,textnote_selected)

t.Commit()