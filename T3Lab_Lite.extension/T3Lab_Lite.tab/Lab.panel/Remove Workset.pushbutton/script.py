# -*- coding: UTF-8 -*-
"""
Remove Workset
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# FUNCTION TO REMOVE WORKSET
# ==================================================
workset_name = "Unit"

# Function to delete a workset
def delete_workset(doc, workset_name,workset_moveto):
    worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
    
    # Find the workset by name
    workset_to_delete = next((ws for ws in worksets if ws.Name == workset_name), None)
    workset_to_move = next((ws for ws in worksets if ws.Name == workset_moveto), None)
    
    if not workset_to_delete:
        forms.alert("Workset '{}' not found.".format(workset_name), exitscript=True)
        return
    
    # Attempt to delete the workset
    with Transaction(doc, "Delete Workset {}".format(workset_name)) as t:
        t.Start()
        try:
            # Create DeleteWorksetSettings to transfer elements instead of deleting them
            delete_settings = DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, workset_to_move.Id)
            
            # Delete the workset but transfer the elements
            WorksetTable.DeleteWorkset(doc, workset_to_delete.Id, delete_settings)
            t.Commit()
            forms.alert("Workset '{}' deleted successfully.".format(workset_name))
        except Exception as e:
            t.RollBack()
            forms.alert("Failed to delete workset: {}".format(str(e)), exitscript=True)

# MAIN EXECUTION
# ==================================================
delete_workset(doc, workset_name)