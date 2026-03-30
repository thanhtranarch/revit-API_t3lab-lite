import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms
import math
"""--------------------------------------------------"""
uidoc= __revit__.ActiveUIDocument
doc=__revit__.ActiveUIDocument.Document
activeview=doc.ActiveView.Id
"""--------------------------------------------------"""
viewports=FilteredElementCollector(doc,activeview).OfClass(Viewport).ToElements()
