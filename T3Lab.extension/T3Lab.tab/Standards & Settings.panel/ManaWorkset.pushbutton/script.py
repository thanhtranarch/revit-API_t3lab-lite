# -*- coding: utf-8 -*-
"""
Workset Manager
---------------
Manage Revit worksets using a rule-based assignment interface.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__ = "Workset\nManager"
__version__ = "1.0.0"

import os
import sys
import clr

clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState, Visibility
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FilteredWorksetCollector,
    Workset,
    WorksetKind,
    Transaction,
    WorksetTable,
    DeleteWorksetSettings,
    DeleteWorksetOption,
    View3D,
    ViewFamilyType,
    ViewFamily,
    WorksetVisibility,
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from pyrevit import revit, forms, script

from Snippets._compat import eid_value

# Path setup
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
lib_dir = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()
output = script.get_output()
REVIT_VERSION = int(revit.doc.Application.VersionNumber)

doc = revit.doc
uidoc = revit.uidoc

XAML_FILE = os.path.join(EXT_DIR, 'lib', 'GUI', 'Tools', 'ManaWorkset.xaml')
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
        self.Number = index
        self.Name = ws.Name
        self.IsOpen = ws.IsOpen
        self.CanEdit = ws.IsEditable
        self.Owner = ws.Owner or ""
        self.IsActive = (
            active_id is not None
            and eid_value(ws.Id) == eid_value(active_id)
        )
        self._id = ws.Id


# ==================================================
# FILE HELPERS
# ==================================================

def load_workset_list(list_file=WORKSET_LIST_FILE):
    if os.path.isfile(list_file):
        with open(list_file, "r") as f:
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
        f.write("# Workset List for T3Lab\n")
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
    # Document.EnableWorksharing manages its own transaction internally and
    # must be called outside of any explicit Transaction, otherwise Revit
    # raises "Operation is not permitted when there is any open sub-transaction,
    # transaction, or transaction group."
    try:
        doc.EnableWorksharing("_SHARED LEVELS & GRIDS", "_ARCHITECT")
        return True
    except Exception as e:
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
    ws_del = next((ws for ws in all_worksets if ws.Name == ws_delete_name), None)
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
                       if eid_value(other.Id) == eid_value(ws.Id)
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
        try:
            fname = os.path.basename(doc.PathName) if doc.PathName else "Unsaved Document"
            self.doc_name.Text = "  ·  {}".format(fname)
        except Exception:
            pass

        self.list_file_path = WORKSET_LIST_FILE
        self._update_list_path_display()

        if not doc.IsWorkshared:
            self._set_worksharing_state(enabled=False)
        else:
            self._set_worksharing_state(enabled=True)
            self._refresh_worksets()
        self._update_status()

    # ── Window chrome ──────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── Sidebar navigation ─────────────────────────────────────────────────────

    def nav_toggle_clicked(self, sender, e):
        """Switch the main content tab based on which sidebar RadioButton was checked."""
        try:
            if sender is self.nav_worksets:
                self.tab_control.SelectedIndex = 0
            elif sender is self.nav_bulk:
                self.tab_control.SelectedIndex = 1
            elif sender is self.nav_views:
                self.tab_control.SelectedIndex = 2
        except Exception:
            pass

    # ── Private helpers ────────────────────────────────────────────────────────

    def _set_worksharing_state(self, enabled):
        """Enable/disable controls based on whether worksharing is active."""
        self.btn_enable_ws.IsEnabled = not enabled
        for btn in [self.btn_create, self.btn_delete,
                    self.btn_create_list, self.btn_remove_unused,
                    self.btn_create_views, self.btn_refresh]:
            btn.IsEnabled = enabled
        try:
            self.worksharing_banner.Visibility = (
                Visibility.Collapsed if enabled else Visibility.Visible
            )
        except Exception:
            pass
        if not enabled:
            self.status_text.Text = "Worksharing is not enabled on this document."

    def _refresh_worksets(self):
        """Reload user worksets from the document into the DataGrid."""
        active_id = get_active_workset_id()
        worksets = get_user_worksets()
        items = [WorksetItem(i + 1, ws, active_id) for i, ws in enumerate(worksets)]
        self.ws_grid.ItemsSource = items
        count = len(items)
        self.ws_status.Text = "{} workset{}".format(count, "s" if count != 1 else "")

    def _update_status(self):
        """Refresh the status-bar message."""
        if not doc.IsWorkshared:
            self.status_text.Text = "Not workshared — enable worksharing first."
        else:
            count = len(get_user_worksets())
            self.status_text.Text = "Ready  —  {} user workset{} loaded.".format(
                count, "s" if count != 1 else "")

    def _update_list_path_display(self):
        """Refresh the workset-list source path shown in the Bulk Tools tab."""
        try:
            self.list_path_text.Text = self.list_file_path
        except Exception:
            pass

    # ── Toolbar button handlers ────────────────────────────────────────────────

    def btn_enable_ws_click(self, sender, e):
        if not _confirm(
            "Enable worksharing on this document?\n\n"
            "This will create two default worksets:\n"
            "  • _SHARED LEVELS & GRIDS\n"
            "  • _ARCHITECT",
            title="Enable Worksharing"
        ):
            return
        if enable_worksharing():
            self._set_worksharing_state(enabled=True)
            self._refresh_worksets()
            self.status_text.Text = "Worksharing enabled successfully."

    def btn_create_click(self, sender, e):
        name = forms.ask_for_string(
            prompt="Enter a name for the new workset:",
            title="Create Workset",
        )
        if not name:
            return
        name = name.strip()
        if not name:
            return
        existing = get_workset_names()
        if name in existing:
            forms.alert("Workset '{}' already exists.".format(name), title="Duplicate")
            return
        created = create_worksets([name], existing)
        if created:
            self._refresh_worksets()
            self.status_text.Text = "Created workset: {}.".format(name)

    def btn_delete_click(self, sender, e):
        selected = list(self.ws_grid.SelectedItems)
        if not selected:
            forms.alert("Select one or more worksets to delete.", title="No Selection")
            return
        names = [item.Name for item in selected]
        if not _confirm(
            "Delete {} workset(s)?\n\n{}\n\n"
            "Elements will be moved to the closest matching workset.".format(
                len(names), "\n".join("  • " + n for n in names)),
            title="Delete Workset(s)"
        ):
            return
        all_names = get_workset_names()
        deleted = 0
        for name in names:
            keep = [n for n in all_names if n != name]
            dest = _find_best_match(name, keep)
            if dest:
                current = get_user_worksets()
                if _remove_workset(name, dest, current):
                    all_names = keep
                    deleted += 1
            else:
                forms.alert(
                    "Cannot delete '{}': no other workset to move elements to.".format(name),
                    title="Delete Failed"
                )
        self._refresh_worksets()
        self.status_text.Text = "Deleted {} of {} workset(s).".format(deleted, len(names))

    def btn_import_list_click(self, sender, e):
        """Browse for a .txt file to use as the workset list source."""
        picked = forms.pick_file(file_ext="txt")
        if not picked:
            return
        self.list_file_path = picked
        self._update_list_path_display()
        self.status_text.Text = "Workset list source set to: {}".format(picked)

    def btn_create_list_click(self, sender, e):
        workset_list = load_workset_list(self.list_file_path)
        existing = get_workset_names()
        to_create = [n for n in workset_list if n not in existing]
        if not to_create:
            forms.alert(
                "All {} worksets in the list already exist.".format(len(workset_list)),
                title="Nothing to Create"
            )
            return
        if not _confirm(
            "Create {} new workset(s) from:\n{}?".format(len(to_create), self.list_file_path),
            title="Create from List"
        ):
            return
        created = create_worksets(to_create, existing)
        self._refresh_worksets()
        self.status_text.Text = "Created {} workset(s) from list.".format(len(created))

    def btn_remove_unused_click(self, sender, e):
        workset_list = load_workset_list(self.list_file_path)
        existing_ws = get_user_worksets()
        existing_names = [ws.Name for ws in existing_ws]
        unused = [ws for ws in existing_ws if ws.Name not in workset_list]
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
        keep_names = [n for n in existing_names if n not in selected_names]
        deleted = 0
        for name in selected_names:
            dest = _find_best_match(name, keep_names)
            if dest:
                current = get_user_worksets()
                if _remove_workset(name, dest, current):
                    deleted += 1
            else:
                print("No destination for '{}', skipping.".format(name))
        self._refresh_worksets()
        self.status_text.Text = "Removed {} workset(s).".format(deleted)

    def btn_create_views_click(self, sender, e):
        self.status_text.Text = "Creating workset views…"
        created, skipped, error = create_workset_views()
        if error:
            forms.alert("Error: {}".format(error), title="Create Workset Views")
            self.status_text.Text = "Error creating views."
            return
        msg = "Created {} view(s).".format(len(created))
        if skipped:
            msg += "  Skipped {} (already exist).".format(len(skipped))
        self.status_text.Text = msg

    def btn_refresh_click(self, sender, e):
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

    workset_list = load_workset_list()
    existing_worksets = get_user_worksets()
    existing_names = [ws.Name for ws in existing_worksets]
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
    deleted = 0
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
