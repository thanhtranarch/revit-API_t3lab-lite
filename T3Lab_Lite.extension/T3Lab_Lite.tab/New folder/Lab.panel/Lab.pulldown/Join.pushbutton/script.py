"""Join Selected Objects
"""

from pyrevit import framework
from pyrevit import revit, DB
from pyrevit import forms

from Autodesk.Revit.DB import *

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

# Get the selected elements
selected_elements = FilteredElementCollector(doc).OfClass(JoinableElement).ToElements()

# Check if at least two elements are selected
if len(selected_elements) >= 2:
    # Create a new JoinGeometryUtils object
    join_utils = JoinGeometryUtils(doc)

    # Attempt to join the selected elements
    try:
        join_utils.JoinElements(selected_elements)
        print("Elements joined successfully!")
    except Exception as e:
        print("Failed to join elements:", e)
else:
    print("Please select at least two joinable elements.")
