# -*- coding: utf-8 -*-
"""
Workset
    
Assign Workset to Project
and Remove Unused Workkset

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""
__author__ ="Tran Tien Thanh"
__title__ = "Workset"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import Workset, WorksetTable, Transaction, FilteredWorksetCollector, WorksetKind, DeleteWorksetSettings, DeleteWorksetOption, WorksharingSaveAsOptions, SaveAsOptions
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# workset_dict= \
#     {"Architect": "USE FOR ELEMENTS IN BUILDING (EXCEPT STRUCTURAL)",
#      "Civil":"USE FOR CIVIL",
#      "Facade":"USE FOR FACADE (EXCEPT STRUCTURE)",
#      "Landscape":"USE FOR LANDSCAPE",
#      "Unit":"USE FOR UNIT, GROUP UNIT, AND FURNITURE IN UNIT",
#      "Shared levels & grids":"USE FOR GRID & LEVEL",
#      "Structure":"USE FOR STRUCTURE",
#      "Ivisible":"USE FOR INVISIBLE ELEMENT"
#     }
workset_list = [
    "01_Shared Levels and Grid_CORE_OFF",
    "01_Shared Levels and Grids_PODIUM_OFF",
    "01_Shared Levels and Grids_RA_OFF",
    "01_Shared Levels and Grids_SA_OFF",
    "01_Shared Levels and Grids_TEMPORARY_OFF",
    "02_Link Architecture Models_OFF",
    "02_Link Architecture Models_Attachement",
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
# CLASS/FUNCTIONS
# ==================================================
# Find longest common subsequence
def lcs(str1, str2):
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
    max_lcs = ""
    best_match = None
    for s in string_list:
        current_lcs = lcs(target_str, s)
        if len(current_lcs) > len(max_lcs):
            max_lcs = current_lcs
            best_match = s
    return best_match

def enable_worksharing():
    try:
        doc.EnableWorksharing("_SHARED LEVELS & GRIDS","_ARCHITECT")
    except Exception as e:
        print("Failed to enable worksharing:", e)

def create_worksets():
    for workset_name in workset_list:
        if workset_name not in existing_workset_names:
            t = Transaction(doc,"Create workset: " + workset_name)
            t.Start()
            new_workse = Workset.Create(doc,workset_name)
            t.Commit()
            new_cre_worksets.append(workset_name)

def create_central():
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
            TaskDialog.Show("Central Created", "Central File was created")
        except Exception as e:
            print("Failed to create Central File:", e)
    else:
        TaskDialog.Show("Central Created", "Please set path")
        
def editable_workset():
    pass
def delete_check(workset_remove, workset_moveto):
    workset_to_delete = next((ws for ws in worksets if ws.Name == workset_remove), None)
    workset_to_move = next((ws for ws in worksets if ws.Name == workset_moveto), None)
    if workset_to_delete and workset_to_move:
        delete_settings = DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, workset_to_move.Id)
        check = WorksetTable.CanDeleteWorkset(doc, workset_to_delete.Id, delete_settings)
        print(workset_to_delete)
        return check
    
def remove_workset(workset_remove, workset_moveto):
    workset_to_delete = next((ws for ws in worksets if ws.Name == workset_remove), None)
    workset_to_move = next((ws for ws in worksets if ws.Name == workset_moveto), None)
    
    if not workset_to_delete:
        forms.alert("Workset '{}' not found.".format(workset_remove), exitscript=True)
        return
    if not workset_to_move:
        forms.alert("Workset '{}' not found.".format(workset_moveto), exitscript=True)
        return

    with Transaction(doc, "Delete Workset {}".format(workset_remove)) as t:
        t.Start()
        try:
            # Ensure workset_to_delete and workset_to_move have valid Ids before using them
            if workset_to_delete and workset_to_move:
                delete_settings = DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, workset_to_move.Id)
                WorksetTable.DeleteWorkset(doc, workset_to_delete.Id, delete_settings)
                
                t.Commit()
                forms.alert("Workset '{}' deleted successfully.".format(workset_remove))
            else:
                t.RollBack()
                forms.alert("Cannot delete workset '{}' because either the delete or move workset is invalid.".format(workset_remove), exitscript=True)
        except Exception as e:
            t.RollBack()
            forms.alert("Failed to delete workset: {}".format(str(e)), exitscript=True)

# MAIN SCRIPT
# ==================================================
#get elements
worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
existing_workset_names = list(ws.Name for ws in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
new_cre_worksets=[]


if __shiftclick__:
    for workset in worksets:
        if not (workset.Name in workset_list):
            
            workset_remove = workset.Name
            workset_moveto = find_best_match(workset_remove, workset_list)
            print(delete_check(workset_remove, workset_moveto))
            # remove_workset(workset_remove,workset_moveto)
else:
    if not doc.IsWorkshared:
        enable_worksharing()
    if not workset_list in existing_workset_names:
        create_worksets()
        if new_cre_worksets:
            TaskDialog.Show("Worksets Created",
                    "The following worksets were successfully created:\n" + "\n".join(new_cre_worksets))
        else:
            TaskDialog.Show("Worksets", "Worksets Existing")
    
# check update, which workset can remove, and move to the best match, else need to check the reason why can delete!
