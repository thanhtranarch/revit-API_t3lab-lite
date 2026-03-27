# -*- coding: utf-8 -*-
"""
Workset Management

Manage worksets in Revit projects.
- Click        : Create worksets or manage the workset list
- Shift+Click  : Quick remove unused worksets (select from checklist)

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__  = "Workset\nMgmt"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import (
    Workset, WorksetTable, Transaction,
    FilteredWorksetCollector, WorksetKind,
    DeleteWorksetSettings, DeleteWorksetOption,
    WorksharingSaveAsOptions, SaveAsOptions,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms, script
import os

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

SCRIPT_DIR        = os.path.dirname(__file__)
WORKSET_LIST_FILE = os.path.join(SCRIPT_DIR, "workset_list.txt")

# Hardcoded fallback list (used if workset_list.txt is missing/empty)
DEFAULT_WORKSET_LIST = [
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
    "Workset1",
]

# WORKSET LIST FILE HELPERS
# ==================================================

def load_workset_list():
    """Load workset list from workset_list.txt, falling back to defaults."""
    if os.path.isfile(WORKSET_LIST_FILE):
        with open(WORKSET_LIST_FILE, "r") as f:
            names = [
                line.strip()
                for line in f
                if line.strip() and not line.strip().startswith("#")
            ]
        if names:
            print("Loaded {} worksets from: {}".format(len(names), WORKSET_LIST_FILE))
            return names
    print("workset_list.txt not found or empty - using defaults")
    return list(DEFAULT_WORKSET_LIST)


def save_workset_list(names):
    """Save workset list to workset_list.txt (overwrites)."""
    with open(WORKSET_LIST_FILE, "w") as f:
        f.write("# Workset List for T3Lab Lite\n")
        f.write("# One workset name per line. Lines starting with '#' are comments.\n\n")
        for name in names:
            f.write(name + "\n")
    print("Saved {} worksets to: {}".format(len(names), WORKSET_LIST_FILE))


def import_from_txt(filepath):
    """Read workset names from any txt file (one per line, # = comment)."""
    with open(filepath, "r") as f:
        names = [
            line.strip()
            for line in f
            if line.strip() and not line.strip().startswith("#")
        ]
    return names


# REVIT HELPERS
# ==================================================

def lcs(str1, str2):
    """Longest common subsequence between two strings."""
    m, n = len(str1), len(str2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if str1[i - 1] == str2[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])
    result = ""
    i, j = m, n
    while i > 0 and j > 0:
        if str1[i - 1] == str2[j - 1]:
            result = str1[i - 1] + result
            i -= 1
            j -= 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1
    return result


def find_best_match(target, candidates):
    """Return the candidate with the longest common subsequence to target."""
    best, best_len = None, 0
    for c in candidates:
        length = len(lcs(target, c))
        if length > best_len:
            best_len = length
            best = c
    return best


def enable_worksharing():
    """Enable worksharing; returns True on success."""
    t = Transaction(doc, "Enable Worksharing")
    t.Start()
    try:
        doc.EnableWorksharing("_SHARED LEVELS & GRIDS", "_ARCHITECT")
        t.Commit()
        print("Worksharing enabled")
        return True
    except Exception as e:
        t.RollBack()
        print("Failed to enable worksharing: {}".format(e))
        return False


def create_worksets(workset_names, existing_names):
    """Create worksets that don't already exist; returns list of created names."""
    created = []
    for name in workset_names:
        if name not in existing_names:
            t = Transaction(doc, "Create Workset: {}".format(name))
            t.Start()
            try:
                Workset.Create(doc, name)
                t.Commit()
                created.append(name)
                print("Created: '{}'".format(name))
            except Exception as e:
                t.RollBack()
                print("Failed '{}': {}".format(name, e))
    return created


def remove_workset(ws_delete_name, ws_move_name, all_worksets):
    """Delete a workset and move its elements to another workset."""
    ws_delete = next((ws for ws in all_worksets if ws.Name == ws_delete_name), None)
    ws_move   = next((ws for ws in all_worksets if ws.Name == ws_move_name), None)

    if not ws_delete or not ws_move:
        print("Workset '{}' or '{}' not found".format(ws_delete_name, ws_move_name))
        return False

    t = Transaction(doc, "Delete Workset: {}".format(ws_delete_name))
    t.Start()
    try:
        settings = DeleteWorksetSettings(
            DeleteWorksetOption.MoveElementsToWorkset, ws_move.Id
        )
        WorksetTable.DeleteWorkset(doc, ws_delete.Id, settings)
        t.Commit()
        print("Deleted '{}' -> moved to '{}'".format(ws_delete_name, ws_move_name))
        return True
    except Exception as e:
        t.RollBack()
        print("Failed to delete '{}': {}".format(ws_delete_name, e))
        forms.alert("Failed to delete '{}':\n{}".format(ws_delete_name, e))
        return False


# MAIN SCRIPT
# ==================================================

# Load standard workset list
workset_list = load_workset_list()

# Add filename-specific worksets (CORE files)
try:
    fp = doc.PathName
    if fp:
        fname = os.path.splitext(os.path.basename(fp))[0].upper()
        if "_MD_CR_ZZ_" in fname:
            for extra in ["ARC_TRELLIS", "ARC_Staircase"]:
                if extra not in workset_list:
                    workset_list.append(extra)
            print("CORE file detected - added ARC_TRELLIS, ARC_Staircase")
except Exception as e:
    print("Error checking filename: {}".format(e))

# Collect current worksets from document
existing_worksets = list(
    FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
)
existing_names = [ws.Name for ws in existing_worksets]

# --------------------------------------------------
# SHIFT+CLICK  →  Quick remove unused worksets
# --------------------------------------------------
if __shiftclick__:
    unused = [ws for ws in existing_worksets if ws.Name not in workset_list]

    if not unused:
        TaskDialog.Show("Workset Management", "No unused worksets found.")
        script.exit()

    # Let user pick which ones to remove
    selected_names = forms.SelectFromList.show(
        sorted([ws.Name for ws in unused]),
        title="Remove Unused Worksets",
        button_name="Remove Selected",
        multiselect=True,
    )

    if not selected_names:
        script.exit()

    # For each selected workset, find the best destination in the remaining list
    keep_names = [n for n in existing_names if n not in selected_names]
    deleted = 0
    for name in selected_names:
        dest = find_best_match(name, keep_names)
        if dest:
            # Refresh workset objects after each deletion
            current_worksets = list(
                FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
            )
            if remove_workset(name, dest, current_worksets):
                deleted += 1
        else:
            print("No destination workset found for '{}', skipping".format(name))

    TaskDialog.Show(
        "Workset Management",
        "Removed {} of {} selected worksets.".format(deleted, len(selected_names)),
    )

# --------------------------------------------------
# NORMAL CLICK  →  Create / manage workset list
# --------------------------------------------------
else:
    action = forms.CommandSwitchWindow.show(
        [
            "Create Worksets",
            "Update List from TXT",
            "Reset to Default List",
        ],
        message="Workset Management\n({} worksets in current list)".format(
            len(workset_list)
        ),
    )

    if action == "Create Worksets":
        # Enable worksharing if needed
        if not doc.IsWorkshared:
            if not enable_worksharing():
                forms.alert("Failed to enable worksharing. Cannot proceed.")
                script.exit()
            TaskDialog.Show("Worksharing", "Worksharing has been enabled.")

        created = create_worksets(workset_list, existing_names)
        if created:
            TaskDialog.Show(
                "Worksets Created",
                "Created {} workset(s):\n\n{}".format(
                    len(created), "\n".join(created)
                ),
            )
        else:
            TaskDialog.Show("Worksets", "All worksets in the list already exist.")

    elif action == "Update List from TXT":
        filepath = forms.pick_file(file_ext="txt")
        if filepath:
            try:
                new_list = import_from_txt(filepath)
                if not new_list:
                    forms.alert("No workset names found in the selected file.")
                else:
                    save_workset_list(new_list)
                    TaskDialog.Show(
                        "List Updated",
                        "Workset list updated with {} entries.\n\nSaved to:\n{}".format(
                            len(new_list), WORKSET_LIST_FILE
                        ),
                    )
            except Exception as e:
                forms.alert("Failed to import file:\n{}".format(e))

    elif action == "Reset to Default List":
        save_workset_list(DEFAULT_WORKSET_LIST)
        TaskDialog.Show(
            "List Reset",
            "Workset list reset to {} default entries.\n\nFile:\n{}".format(
                len(DEFAULT_WORKSET_LIST), WORKSET_LIST_FILE
            ),
        )
