# -*- coding: utf-8 -*-
"""
Workset Manager

View and manage user worksets in the active Revit project.
- Click       : Open Workset Manager (list all worksets)
- Shift+Click : Quick remove unused worksets (select from checklist)

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Workset\nManagement"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==================================================
import os
import sys
import json

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from Autodesk.Revit.DB import (
    Workset, WorksetTable, Transaction,
    FilteredElementCollector, FilteredWorksetCollector, WorksetKind,
    DeleteWorksetSettings, DeleteWorksetOption,
    View3D, ViewFamilyType, ViewFamily, WorksetVisibility,
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from pyrevit import revit, forms, script

# Path setup
SCRIPT_DIR    = os.path.dirname(__file__)
EXT_DIR       = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
lib_dir       = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()
output = script.get_output()
REVIT_VERSION = int(revit.doc.Application.VersionNumber)

doc   = revit.doc
uidoc = revit.uidoc

XAML_FILE         = os.path.join(EXT_DIR, 'lib', 'GUI', 'Tools', 'WorksetManager.xaml')
WORKSET_LIST_FILE = os.path.join(SCRIPT_DIR, "workset_list.txt")

# ==================================================
# DEFAULT WORKSET LIST  (fallback if workset_list.txt missing)
# ==================================================
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

# CLASS/FUNCTIONS
# ==================================================

# WORKSET ITEM MODEL
class WorksetItem(object):
    """View-model for a single user workset row in the DataGrid."""

    def __init__(self, index, ws, active_id=None):
        self.Number   = index
        self.Name     = ws.Name
        self.IsOpen   = ws.IsOpen
        self.CanEdit  = ws.IsEditable
        self.Owner    = ws.Owner or ""
        self.IsActive = (
            active_id is not None
            and ws.Id.IntegerValue == active_id.IntegerValue
        )
        self._id = ws.Id


# ==================================================
# FILE HELPERS
# ==================================================

def load_workset_list():
    if os.path.isfile(WORKSET_LIST_FILE):
        with open(WORKSET_LIST_FILE, "r") as f:
            names = [
                line.strip()
                for line in f
                if line.strip() and not line.strip().startswith("#")
            ]
        if names:
            return names
    return list(DEFAULT_WORKSET_LIST)


def save_workset_list(names):
    with open(WORKSET_LIST_FILE, "w") as f:
        f.write("# Workset List for T3Lab Lite\n")
        f.write("# One workset name per line. Lines starting with '#' are comments.\n\n")
        for name in names:
            f.write(name + "\n")


# ==================================================
# REVIT HELPERS
# ==================================================

def get_user_worksets():
    return list(
        FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
    )


def get_workset_names():
    return [ws.Name for ws in get_user_worksets()]


def get_active_workset_id():
    try:
        return doc.GetWorksetTable().GetActiveWorksetId()
    except Exception:
        return None


def enable_worksharing():
    t = Transaction(doc, "Enable Worksharing")
    t.Start()
    try:
        doc.EnableWorksharing("_SHARED LEVELS & GRIDS", "_ARCHITECT")
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        forms.alert("Failed to enable worksharing:\n{}".format(e), title="Error")
        return False


def create_worksets(workset_names, existing_names):
    """Create worksets not already present; returns list of created names."""
    created = []
    for name in workset_names:
        if name not in existing_names:
            t = Transaction(doc, "Create Workset: {}".format(name))
            t.Start()
            try:
                Workset.Create(doc, name)
                t.Commit()
                created.append(name)
            except Exception as e:
                t.RollBack()
                print("Failed '{}': {}".format(name, e))
    return created


# LCS / fuzzy match helpers
def _lcs(str1, str2):
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


def _find_best_match(target, candidates):
    best, best_len = None, 0
    for c in candidates:
        length = len(_lcs(target, c))
        if length > best_len:
            best_len = length
            best = c
    return best


def _remove_workset(ws_delete_name, ws_move_name, all_worksets):
    ws_del  = next((ws for ws in all_worksets if ws.Name == ws_delete_name), None)
    ws_move = next((ws for ws in all_worksets if ws.Name == ws_move_name), None)
    if not ws_del or not ws_move:
        return False
    t = Transaction(doc, "Delete Workset: {}".format(ws_delete_name))
    t.Start()
    try:
        settings = DeleteWorksetSettings(
            DeleteWorksetOption.MoveElementsToWorkset, ws_move.Id
        )
        WorksetTable.DeleteWorkset(doc, ws_del.Id, settings)
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        forms.alert("Failed to delete '{}':\n{}".format(ws_delete_name, e))
        return False


def _get_3d_view_type_id():
    viewtypes = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    return next(
        (vt.Id for vt in viewtypes if vt.ViewFamily == ViewFamily.ThreeDimensional),
        None
    )


def create_workset_views():
    """Create one 3D isometric view per user workset, isolating visibility to that workset."""
    if not doc.IsWorkshared:
        return None, None, "Document is not workshared."

    type_id = _get_3d_view_type_id()
    if type_id is None:
        return None, None, "No 3D view family type found in document."

    worksets = get_user_worksets()
    if not worksets:
        return None, None, "No user worksets found."

    existing = set(v.Name for v in FilteredElementCollector(doc).OfClass(View3D).ToElements())
    created, skipped = [], []

    t = Transaction(doc, "Create Workset Views")
    t.Start()
    try:
        for ws in worksets:
            if ws.Name in existing:
                skipped.append(ws.Name)
                continue
            view3d = View3D.CreateIsometric(doc, type_id)
            view3d.Name = ws.Name
            for other in worksets:
                vis = (WorksetVisibility.Visible
                       if other.Id.IntegerValue == ws.Id.IntegerValue
                       else WorksetVisibility.Hidden)
                view3d.SetWorksetVisibility(other.Id, vis)
            created.append(ws.Name)
        t.Commit()
    except Exception as e:
        t.RollBack()
        return None, None, str(e)

    return created, skipped, None


def _confirm(message, title="Confirm"):
    td = TaskDialog(title)
    td.MainContent = message
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    return td.Show() == TaskDialogResult.Yes


# ==================================================
# WPF WINDOW
# ==================================================

class WorksetManagerWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._load_logo()
        try:
            fname = os.path.basename(doc.PathName) if doc.PathName else "Unsaved Document"
            self.doc_name.Text = fname
        except Exception:
            pass

        if not doc.IsWorkshared:
            self._set_worksharing_state(enabled=False)
        else:
            self._set_worksharing_state(enabled=True)
            self._refresh_worksets()
        self._update_status()

    def _load_logo(self):
        try:
            logo_path = os.path.join(EXT_DIR, 'lib', 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.Icon = bitmap
        except Exception as icon_ex:
            logger.warning("Could not set window icon: {}".format(icon_ex))

    # --------------------------------------------------
    # Worksharing state
    # --------------------------------------------------

    def _set_worksharing_state(self, enabled):
        self.ws_grid.IsEnabled          = enabled
        self.btn_create.IsEnabled       = enabled
        self.btn_delete.IsEnabled       = enabled
        self.btn_create_list.IsEnabled  = enabled
        self.btn_remove_unused.IsEnabled = enabled
        self.btn_create_views.IsEnabled = enabled
        self.btn_enable_ws.IsEnabled    = not enabled
        if not enabled:
            self.status_text.Text = "Worksharing not enabled - click 'Enable Worksharing' to continue"

    # --------------------------------------------------
    # Grid / status helpers
    # --------------------------------------------------

    def _refresh_worksets(self):
        active_id = get_active_workset_id()
        worksets  = get_user_worksets()
        items = [WorksetItem(i + 1, ws, active_id) for i, ws in enumerate(worksets)]
        self.ws_grid.ItemsSource = None
        self.ws_grid.ItemsSource = items

    def _update_status(self):
        if doc.IsWorkshared:
            count = len(get_user_worksets())
            self.status_text.Text  = "{} user workset(s) in document".format(count)
            self.ws_status.Text    = "Workshared"
        else:
            self.status_text.Text  = "Worksharing not enabled"
            self.ws_status.Text    = "Not Workshared"

    # --------------------------------------------------
    # Window chrome handlers
    # --------------------------------------------------

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()

    # --------------------------------------------------
    # Toolbar handlers
    # --------------------------------------------------

    def btn_enable_ws_click(self, sender, args):
        if not _confirm(
            "Enable worksharing on this document?\n\n"
            "This action cannot be undone.",
            title="Enable Worksharing"
        ):
            return
        if enable_worksharing():
            self._set_worksharing_state(enabled=True)
            self._refresh_worksets()
            self._update_status()

    def btn_create_click(self, sender, args):
        name = forms.ask_for_string(
            prompt="Enter workset name:",
            title="Create Workset",
        )
        if not name or not name.strip():
            return
        name = name.strip()
        existing = get_workset_names()
        if name in existing:
            forms.alert("'{}' already exists.".format(name), title="Create Workset")
            return
        created = create_worksets([name], existing)
        if created:
            self._refresh_worksets()
            self._update_status()

    def btn_delete_click(self, sender, args):
        selected = list(self.ws_grid.SelectedItems)
        if not selected:
            forms.alert("Select one or more worksets to delete.", title="Delete Workset")
            return

        names_to_delete = [item.Name for item in selected]
        all_names = get_workset_names()
        keep_names = [n for n in all_names if n not in names_to_delete]

        if not keep_names:
            forms.alert(
                "Cannot delete all worksets. At least one must remain.",
                title="Delete Workset"
            )
            return

        if not _confirm(
            "Delete {} workset(s)?\n\n{}\n\n"
            "Elements will be moved to the closest matching workset.".format(
                len(names_to_delete), "\n".join(names_to_delete)),
            title="Delete Worksets"
        ):
            return

        deleted = 0
        for name in names_to_delete:
            dest = _find_best_match(name, keep_names)
            if dest:
                if _remove_workset(name, dest, get_user_worksets()):
                    deleted += 1
                    keep_names = [n for n in keep_names if n != name]
            else:
                print("No destination found for '{}', skipping".format(name))

        forms.alert(
            "Deleted {} of {} workset(s).".format(deleted, len(names_to_delete)),
            title="Done"
        )
        self._refresh_worksets()
        self._update_status()

    def btn_create_list_click(self, sender, args):
        workset_list   = load_workset_list()
        existing_names = get_workset_names()
        created = create_worksets(workset_list, existing_names)
        if created:
            forms.alert(
                "Created {} workset(s):\n\n{}".format(
                    len(created), "\n".join(created)),
                title="Worksets Created"
            )
        else:
            forms.alert("All worksets in the list already exist.", title="Worksets")
        self._refresh_worksets()
        self._update_status()

    def btn_remove_unused_click(self, sender, args):
        workset_list = load_workset_list()
        all_ws       = get_user_worksets()
        all_names    = [ws.Name for ws in all_ws]
        unused       = [ws for ws in all_ws if ws.Name not in workset_list]

        if not unused:
            forms.alert("No unused worksets found.", title="Remove Unused")
            return

        selected_names = forms.SelectFromList.show(
            sorted([ws.Name for ws in unused]),
            title="Remove Unused Worksets",
            button_name="Remove Selected",
            multiselect=True,
        )
        if not selected_names:
            return

        keep_names = [n for n in all_names if n not in selected_names]
        deleted    = 0
        for name in selected_names:
            dest = _find_best_match(name, keep_names)
            if dest:
                if _remove_workset(name, dest, get_user_worksets()):
                    deleted += 1
            else:
                print("No destination found for '{}', skipping".format(name))

        forms.alert(
            "Removed {} of {} selected workset(s).".format(
                deleted, len(selected_names)),
            title="Done"
        )
        self._refresh_worksets()
        self._update_status()

    def btn_create_views_click(self, sender, args):
        created, skipped, err = create_workset_views()
        if err:
            forms.alert(err, title="Workset Views")
            return
        if created:
            msg = "Created {} view(s):\n{}".format(len(created), "\n".join(created))
            if skipped:
                msg += "\n\nSkipped {} existing view(s).".format(len(skipped))
            forms.alert(msg, title="Workset Views Created")
        else:
            forms.alert("All workset views already exist.", title="Workset Views")
        self._refresh_worksets()
        self._update_status()

    def btn_refresh_click(self, sender, args):
        self._refresh_worksets()
        self._update_status()


# ==================================================
# MAIN ENTRY POINT
# ==================================================

# SHIFT+CLICK  ->  Quick remove unused worksets
if __shiftclick__:
    if not doc.IsWorkshared:
        TaskDialog.Show("Workset Manager", "Document is not workshared.")
        script.exit()

    workset_list      = load_workset_list()
    existing_worksets = get_user_worksets()
    existing_names    = [ws.Name for ws in existing_worksets]
    unused = [ws for ws in existing_worksets if ws.Name not in workset_list]

    if not unused:
        TaskDialog.Show("Workset Manager", "No unused worksets found.")
        script.exit()

    selected_names = forms.SelectFromList.show(
        sorted([ws.Name for ws in unused]),
        title="Remove Unused Worksets",
        button_name="Remove Selected",
        multiselect=True,
    )
    if not selected_names:
        script.exit()

    keep_names = [n for n in existing_names if n not in selected_names]
    deleted    = 0
    for name in selected_names:
        dest = _find_best_match(name, keep_names)
        if dest:
            current = get_user_worksets()
            if _remove_workset(name, dest, current):
                deleted += 1
        else:
            print("No destination found for '{}', skipping".format(name))

    TaskDialog.Show(
        "Workset Manager",
        "Removed {} of {} selected workset(s).".format(deleted, len(selected_names)),
    )

# NORMAL CLICK  ->  Open Workset Manager window
else:
    WorksetManagerWindow().ShowDialog()
