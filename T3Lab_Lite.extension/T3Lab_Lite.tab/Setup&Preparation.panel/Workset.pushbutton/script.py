# -*- coding: utf-8 -*-
"""
Workset
    
Assign Workset to Project
and Remove Unused Workset

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""
__author__ = "Tran Tien Thanh"
__title__ = "Workset"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import Workset, WorksetTable, Transaction, FilteredWorksetCollector, WorksetKind, DeleteWorksetSettings, DeleteWorksetOption, WorksharingSaveAsOptions, SaveAsOptions, WorksharingUtils
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
import os

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# Base workset list
base_workset_list = [
    "01_Shared Levels and Grids_CORE_OFF",
    "01_Shared Levels and Grids_PH_OFF",
    "01_Shared Levels and Grids_RA_OFF",
    "01_Shared Levels and Grids_SA_OFF",
    "01_Shared Levels and Grids_ROOF_OFF",
    "01_Shared Levels and Grids_for Coordination",
    "02_Link Architecture Models_OFF",
    "02_Link Architecture Models_Attachment",
    "03_Link Structural Models_OFF",
    "04_Link Interior Models_OFF",
    "05_Link Facade Models_OFF",
    "06_Link Site Models_OFF",
    "07_Link Landscape Models_OFF",
    "08_Link Other 3D Data_OFF",
    "09_Link MEP Models_OFF",
    "10_Do not use_OFF",
    "11_Link Cad Consultant_OFF",
    "11_Link Cad Internal_OFF",
    "11_Link Cad Subcon_OFF",
    "12_Link PBU Models",
    "ARC_3DLine-3DText",
    "ARC_3DRoomTag",
    "ARC_Ancillary",
    "ARC_AreaRoomSpace",
    "ARC_BMU",
    "ARC_Ceiling",
    "ARC_DoorAndWindow",
    "ARC_ExteriallWallAndFacade",
    "ARC_ExteriorRoofAndCanopy",
    "ARC_FireProvision",
    "ARC_FloorFinish",
    "ARC_FloorStructural_OFF",
    "ARC_Floor",
    "ARC_Furniture",
    "ARC_Matchline",
    "ARC_Misc",
    "ARC_NonPBU",
    "ARC_NonStructureWall",
    "ARC_ParkingLots",
    "ARC_PlantingSoil",
    "ARC_Railing",
    "ARC_Ramp",
    "ARC_RoadAndPavement",
    "ARC_SanitaryAndDrainage",
    "ARC_Signage",
    "ARC_StructuralCore_OFF",
    "ARC_StructuralColumn_OFF",
    "ARC_StructuralSlabElement_OFF",
    "ARC_StructureWall_OFF",
    "ARC_Temporary_OFF",
    "ARC_Tile Line (Model)",
    "ARC_Toilets",
    "ARC_WallExterior",
    "ARC_WallFinish",
    "ARC_WallInterior",
    "Workset1"
]

# Function to get dynamic workset list based on filename
def get_workset_list():
    """
    Get workset list based on file name.
    If filename contains 'CORE', add 'ARC_TRELLIS' to the list.
    """
    workset_list = base_workset_list[:]  # Create a copy of base list
    
    try:
        # Get the document file path
        file_path = doc.PathName
        if file_path:
            # Extract filename without extension and convert to uppercase
            filename = os.path.splitext(os.path.basename(file_path))[0].upper()
            print("Current filename (uppercase): {}".format(filename))
            
            # Check if filename contains 'CORE' (both in uppercase)
            if '_MD_CR_ZZ_' in filename:
                workset_list.append("ARC_TRELLIS")
                workset_list.append("ARC_Staircase")
                print("Filename contains 'CORE' - Added 'ARC_TRELLIS' to workset list")
            else:
                print("Filename doesn't contain 'CORE' - Using standard workset list")
        else:
            print("Document not saved yet - Using standard workset list")
            
    except Exception as e:
        print("Error checking filename: {}".format(e))
        print("Using standard workset list")
    
    return workset_list

# CLASS/FUNCTIONS
# ==================================================

def lcs(str1, str2):
    """Find longest common subsequence between two strings"""
    m, n = len(str1), len(str2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if str1[i - 1] == str2[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    lcs_str = ""
    i, j = m, n
    while i > 0 and j > 0:
        if str1[i - 1] == str2[j - 1]:
            lcs_str = str1[i - 1] + lcs_str
            i -= 1
            j -= 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1

    return lcs_str

def find_best_match(target_str, string_list):
    """Find best matching string from list using LCS algorithm"""
    max_lcs = ""
    best_match = None
    for s in string_list:
        current_lcs = lcs(target_str, s)
        if len(current_lcs) > len(max_lcs):
            max_lcs = current_lcs
            best_match = s
    return best_match

def enable_worksharing():
    """Enable worksharing with default worksets"""
    try:
        t = Transaction(doc, "Enable Worksharing")
        t.Start()
        doc.EnableWorksharing("_SHARED LEVELS & GRIDS", "_ARCHITECT")
        t.Commit()
        print("Worksharing enabled successfully")
        return True
    except Exception as e:
        print("Failed to enable worksharing: {}".format(e))
        if t.HasStarted():
            t.RollBack()
        return False
        
def create_worksets(workset_names, existing_names):
    """Create worksets that don't already exist"""
    created = []
    for name in workset_names:
        if name not in existing_names:
            t = Transaction(doc, "Create Workset: {}".format(name))
            t.Start()
            try:
                Workset.Create(doc, name)
                t.Commit()
                created.append(name)
                print("Created workset: '{}'".format(name))
            except Exception as e:
                t.RollBack()
                print("Failed to create workset '{}': {}".format(name, e))
    return created

def check_editable_worksets(worksets):
    """Check if worksets are editable and report status"""
    non_editable = []
    for ws in worksets:
        try:
            if not ws.IsEditable:
                non_editable.append(ws.Name)
                print("Workset '{}' is not editable.".format(ws.Name))
        except Exception as e:
            print("Error checking workset '{}': {}".format(ws.Name, e))
    
    if non_editable:
        forms.alert("Non-editable worksets found:\n{}".format("\n".join(non_editable)))
        return False
    return True

def create_central():
    """Create central file from current document"""
    file_name = forms.save_file()
    if file_name:
        file_name_without_extension = file_name.rsplit(".", 1)[0]
        central_file_path = file_name_without_extension + ".rvt"
        try:
            options = WorksharingSaveAsOptions()
            options.SaveAsCentral = True
            saveas_option = SaveAsOptions()
            saveas_option.MaximumBackups = 10
            saveas_option.OverwriteExistingFile = True
            saveas_option.SetWorksharingOptions(options)
            doc.SaveAs(central_file_path, saveas_option)
            TaskDialog.Show("Central Created", "Central File was created at:\n{}".format(central_file_path))
        except Exception as e:
            print("Failed to create Central File: {}".format(e))
            TaskDialog.Show("Error", "Failed to create Central File:\n{}".format(e))
    else:
        TaskDialog.Show("Cancelled", "Please set path to create central file")

def can_delete_workset(ws_delete, ws_move):
    """Check if workset can be deleted by moving elements to another workset"""
    delete_settings = DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, ws_move.Id)
    return WorksetTable.CanDeleteWorkset(doc, ws_delete.Id, delete_settings)

def remove_workset(ws_delete_name, ws_move_name, worksets):
    """Remove workset and move its elements to another workset"""
    ws_delete = next((ws for ws in worksets if ws.Name == ws_delete_name), None)
    ws_move = next((ws for ws in worksets if ws.Name == ws_move_name), None)
    
    if not ws_delete or not ws_move:
        forms.alert("Workset '{}' or '{}' not found.".format(ws_delete_name, ws_move_name))
        return False

    t = Transaction(doc, "Delete Workset: {}".format(ws_delete_name))
    t.Start()
    try:
        delete_settings = DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, ws_move.Id)
        WorksetTable.DeleteWorkset(doc, ws_delete.Id, delete_settings)
        t.Commit()
        print("Deleted '{}' and moved content to '{}'.".format(ws_delete_name, ws_move_name))
        return True
    except Exception as e:
        t.RollBack()
        print("Failed to delete workset '{}': {}".format(ws_delete_name, e))
        forms.alert("Failed to delete workset '{}': {}".format(ws_delete_name, e))
        return False

# MAIN SCRIPT
# ==================================================

# Get dynamic workset list based on filename
workset_list = get_workset_list()

worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
existing_workset_names = [ws.Name for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)]

if __shiftclick__:
    # Delete unused worksets mode
    print("Running in delete mode (Shift+Click)")
    
    if not check_editable_worksets(worksets):
        forms.alert("Some worksets are not editable. Please make them editable first!")
    else:
        deleted_count = 0
        for ws in worksets:
            try:
                if ws.Name not in workset_list:
                    best_match = find_best_match(ws.Name, workset_list)
                    if best_match and remove_workset(ws.Name, best_match, worksets):
                        deleted_count += 1
            except Exception as e:
                print("Error processing workset '{}': {}".format(ws.Name, e))
        
        if deleted_count > 0:
            TaskDialog.Show("Cleanup Complete", "Removed {} unused worksets.".format(deleted_count))
        else:
            TaskDialog.Show("Cleanup Complete", "No unused worksets found to remove.")
else:
    # Create worksets mode
    print("Running in create mode")
    
    # Enable worksharing if not already enabled
    if not doc.IsWorkshared:
        if not enable_worksharing():
            forms.alert("Failed to enable worksharing. Cannot proceed.")
        else:
            TaskDialog.Show("Worksharing", "Worksharing has been enabled.")
    
    # Create missing worksets
    created_worksets = create_worksets(workset_list, existing_workset_names)
    
    if created_worksets:
        TaskDialog.Show("Worksets Created", "Created {} worksets:\n\n{}".format(
            len(created_worksets), "\n".join(created_worksets)))
    else:
        TaskDialog.Show("Worksets", "All worksets already exist.")