# -*- coding: utf-8 -*-
"""
Workset Manager

Rule-based workset assignment for Revit projects.
- Click       : Open Workset Manager (rule grid - assign categories to worksets)
- Shift+Click : Quick remove unused worksets (select from checklist)

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__  = "Workset\nMgmt"

import os
import json
import copy

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
    BuiltInParameter, BuiltInCategory,
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
from pyrevit import forms, script

logger = script.get_logger()

uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

SCRIPT_DIR        = os.path.dirname(__file__)
RULES_FILE        = os.path.join(SCRIPT_DIR, "rules.json")
WORKSET_LIST_FILE = os.path.join(SCRIPT_DIR, "workset_list.txt")

# ==================================================
# CATEGORY MAP  (display name -> BuiltInCategory)
# Add entries here to support more filterable categories.
# ==================================================
CATEGORY_MAP = {
    "Areas":                BuiltInCategory.OST_Areas,
    "Casework":             BuiltInCategory.OST_Casework,
    "Ceilings":             BuiltInCategory.OST_Ceilings,
    "Columns":              BuiltInCategory.OST_Columns,
    "Curtain Mullions":     BuiltInCategory.OST_CurtainWallMullions,
    "Curtain Panels":       BuiltInCategory.OST_CurtainWallPanels,
    "Detail Items":         BuiltInCategory.OST_DetailComponents,
    "Doors":                BuiltInCategory.OST_Doors,
    "Electrical Equipment": BuiltInCategory.OST_ElectricalEquipment,
    "Electrical Fixtures":  BuiltInCategory.OST_ElectricalFixtures,
    "Entourage":            BuiltInCategory.OST_Entourage,
    "Floors":               BuiltInCategory.OST_Floors,
    "Furniture":            BuiltInCategory.OST_Furniture,
    "Furniture Systems":    BuiltInCategory.OST_FurnitureSystems,
    "Generic Models":       BuiltInCategory.OST_GenericModel,
    "Lighting Fixtures":    BuiltInCategory.OST_LightingFixtures,
    "Mass":                 BuiltInCategory.OST_Mass,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
    "Model Lines":          BuiltInCategory.OST_Lines,
    "Parking":              BuiltInCategory.OST_Parking,
    "Planting":             BuiltInCategory.OST_Planting,
    "Plumbing Fixtures":    BuiltInCategory.OST_PlumbingFixtures,
    "Railings":             BuiltInCategory.OST_StairsRailing,
    "Ramps":                BuiltInCategory.OST_Ramps,
    "Roofs":                BuiltInCategory.OST_Roofs,
    "Rooms":                BuiltInCategory.OST_Rooms,
    "Signage":              BuiltInCategory.OST_Signage,
    "Site":                 BuiltInCategory.OST_Site,
    "Specialty Equipment":  BuiltInCategory.OST_SpecialityEquipment,
    "Stairs":               BuiltInCategory.OST_Stairs,
    "Structural Columns":   BuiltInCategory.OST_StructuralColumns,
    "Structural Framing":   BuiltInCategory.OST_StructuralFraming,
    "Topography":           BuiltInCategory.OST_Topography,
    "Walls":                BuiltInCategory.OST_Walls,
    "Windows":              BuiltInCategory.OST_Windows,
}

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

# ==================================================
# DEFAULT RULES  (architecture-oriented presets)
# ==================================================
DEFAULT_RULES = [
    {"run_order":  1, "name": "Interior Walls",      "description": "Non-structural interior walls",     "category_filter": "Walls",      "workset": "ARC_WallInterior",          "enabled": True},
    {"run_order":  2, "name": "Floors",               "description": "Floor slab elements",               "category_filter": "Floors",     "workset": "ARC_Floor",                 "enabled": True},
    {"run_order":  3, "name": "Ceilings",             "description": "Ceiling elements",                  "category_filter": "Ceilings",   "workset": "ARC_Ceiling",               "enabled": True},
    {"run_order":  4, "name": "Doors",                "description": "Door families",                     "category_filter": "Doors",      "workset": "ARC_DoorAndWindow",         "enabled": True},
    {"run_order":  5, "name": "Windows",              "description": "Window families",                   "category_filter": "Windows",    "workset": "ARC_DoorAndWindow",         "enabled": True},
    {"run_order":  6, "name": "Furniture",            "description": "Furniture elements",                "category_filter": "Furniture",  "workset": "ARC_Furniture",             "enabled": True},
    {"run_order":  7, "name": "Stairs",               "description": "Stair elements",                    "category_filter": "Stairs",     "workset": "ARC_Ramp",                  "enabled": True},
    {"run_order":  8, "name": "Railings",             "description": "Railing elements",                  "category_filter": "Railings",   "workset": "ARC_Railing",               "enabled": True},
    {"run_order":  9, "name": "Roofs",                "description": "Roof elements",                     "category_filter": "Roofs",      "workset": "ARC_ExteriorRoofAndCanopy", "enabled": True},
    {"run_order": 10, "name": "Rooms",                "description": "Room elements",                     "category_filter": "Rooms",      "workset": "ARC_AreaRoomSpace",         "enabled": True},
    {"run_order": 11, "name": "Parking",              "description": "Parking elements",                  "category_filter": "Parking",    "workset": "ARC_ParkingLots",           "enabled": True},
    {"run_order": 12, "name": "Planting",             "description": "Planting and landscaping elements", "category_filter": "Planting",   "workset": "ARC_PlantingSoil",          "enabled": True},
    {"run_order": 13, "name": "Ramps",                "description": "Ramp elements",                     "category_filter": "Ramps",      "workset": "ARC_Ramp",                  "enabled": True},
    {"run_order": 14, "name": "Signage",              "description": "Signage elements",                  "category_filter": "Signage",    "workset": "ARC_Signage",               "enabled": True},
]

# ==================================================
# RULE MODEL
# ==================================================

class WorksetRule(object):
    """Single assignment rule: one category -> one workset."""

    def __init__(self, run_order=1, name="New Rule", description="",
                 category_filter="Walls", workset="", enabled=True):
        self.RunOrder       = run_order
        self.Name           = name
        self.Description    = description
        self.CategoryFilter = category_filter
        self.Workset        = workset
        self.Enabled        = enabled

    def to_dict(self):
        return {
            "run_order":       self.RunOrder,
            "name":            self.Name,
            "description":     self.Description,
            "category_filter": self.CategoryFilter,
            "workset":         self.Workset,
            "enabled":         self.Enabled,
        }

    @classmethod
    def from_dict(cls, d):
        return cls(
            run_order       = d.get("run_order", 1),
            name            = d.get("name", "Rule"),
            description     = d.get("description", ""),
            category_filter = d.get("category_filter", "Walls"),
            workset         = d.get("workset", ""),
            enabled         = bool(d.get("enabled", True)),
        )

    def clone(self):
        return WorksetRule.from_dict(self.to_dict())


# ==================================================
# FILE HELPERS
# ==================================================

def load_rules():
    """Load rules from rules.json; fall back to built-in defaults."""
    if os.path.isfile(RULES_FILE):
        try:
            with open(RULES_FILE, "r") as f:
                data = json.load(f)
            rules = [WorksetRule.from_dict(r) for r in data.get("rules", [])]
            if rules:
                print("Loaded {} rules from: {}".format(len(rules), RULES_FILE))
                return rules
        except Exception as e:
            print("Failed to load rules.json: {}".format(e))
    print("Using default rules ({})".format(len(DEFAULT_RULES)))
    return [WorksetRule.from_dict(r) for r in DEFAULT_RULES]


def save_rules(rules):
    """Persist rules to rules.json."""
    data = {"rules": [r.to_dict() for r in rules]}
    with open(RULES_FILE, "w") as f:
        json.dump(data, f, indent=2)
    print("Saved {} rules to: {}".format(len(rules), RULES_FILE))


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


def enable_worksharing():
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
                print("Created: '{}'".format(name))
            except Exception as e:
                t.RollBack()
                print("Failed '{}': {}".format(name, e))
    return created


def apply_rules(rules, view_id=None):
    """
    Execute enabled rules: collect elements by category and reassign their workset.

    Args:
        rules:   list of WorksetRule objects
        view_id: ElementId of a view to limit scope, or None for whole project

    Returns:
        (total_assigned, list_of_result_strings)
    """
    if not doc.IsWorkshared:
        return 0, ["ERROR: Document is not workshared. Cannot assign worksets."]

    ws_by_name = {ws.Name: ws for ws in get_user_worksets()}
    messages   = []
    total      = 0

    for rule in rules:
        if not rule.Enabled:
            continue

        if rule.Workset not in ws_by_name:
            messages.append(
                "SKIP  [{}] - workset '{}' not found in document".format(
                    rule.Name, rule.Workset)
            )
            continue

        bic = CATEGORY_MAP.get(rule.CategoryFilter)
        if bic is None:
            messages.append(
                "SKIP  [{}] - unknown category '{}' (check Category Filter column)".format(
                    rule.Name, rule.CategoryFilter)
            )
            continue

        target_id = ws_by_name[rule.Workset].Id.IntegerValue

        try:
            if view_id:
                collector = FilteredElementCollector(doc, view_id)
            else:
                collector = FilteredElementCollector(doc)
            elements = (collector
                        .OfCategory(bic)
                        .WhereElementIsNotElementType()
                        .ToElements())
        except Exception as e:
            messages.append("ERROR [{}] - collect failed: {}".format(rule.Name, e))
            continue

        assigned = 0
        t = Transaction(doc, "Workset Rule: {}".format(rule.Name))
        t.Start()
        try:
            for elem in elements:
                ws_param = elem.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
                if ws_param and not ws_param.IsReadOnly:
                    if ws_param.AsInteger() != target_id:
                        ws_param.Set(target_id)
                        assigned += 1
            t.Commit()
            total += assigned
            messages.append(
                "OK    [{}] - {} of {} elements -> '{}'".format(
                    rule.Name, assigned, len(list(elements)) if hasattr(elements, '__len__') else "?",
                    rule.Workset)
            )
        except Exception as e:
            t.RollBack()
            messages.append("ERROR [{}] - transaction failed: {}".format(rule.Name, e))

    return total, messages


# LCS / fuzzy match helpers (for shift+click remove)
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
        print("Workset '{}' or '{}' not found".format(ws_delete_name, ws_move_name))
        return False
    t = Transaction(doc, "Delete Workset: {}".format(ws_delete_name))
    t.Start()
    try:
        settings = DeleteWorksetSettings(
            DeleteWorksetOption.MoveElementsToWorkset, ws_move.Id
        )
        WorksetTable.DeleteWorkset(doc, ws_del.Id, settings)
        t.Commit()
        print("Deleted '{}' -> elements moved to '{}'".format(ws_delete_name, ws_move_name))
        return True
    except Exception as e:
        t.RollBack()
        forms.alert("Failed to delete '{}':\n{}".format(ws_delete_name, e))
        return False


def _confirm(message, title="Confirm"):
    """Show a Yes/No TaskDialog; returns True if Yes."""
    td = TaskDialog(title)
    td.MainContent = message
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    return td.Show() == TaskDialogResult.Yes


# ==================================================
# WPF WINDOW
# ==================================================

class WorksetManagerWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, "WorksetManager.xaml")
        self._rules = load_rules()
        self._dirty = False
        self._load_logo()
        self._refresh_grid()
        self._update_status()
        try:
            fname = os.path.basename(doc.PathName) if doc.PathName else "Unsaved Document"
            self.doc_name.Text = fname
        except Exception:
            pass

    def _load_logo(self):
        """Load T3Lab logo into the title bar image and window icon."""
        try:
            ext_dir   = os.path.dirname(os.path.dirname(
                            os.path.dirname(os.path.dirname(__file__))))
            logo_path = os.path.join(ext_dir, 'lib', 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.logo_image.Source = bitmap
                self.Icon = bitmap
        except Exception:
            pass

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
    # Grid / status helpers
    # --------------------------------------------------

    def _refresh_grid(self):
        self.rules_grid.ItemsSource = None
        self.rules_grid.ItemsSource = self._rules

    def _renumber(self):
        for i, r in enumerate(self._rules):
            r.RunOrder = i + 1

    def _update_status(self):
        n_enabled = sum(1 for r in self._rules if r.Enabled)
        ws_names  = get_workset_names()
        self.status_text.Text  = "{} rules  ({} enabled)".format(
            len(self._rules), n_enabled)
        self.workset_count.Text = "{} worksets in document".format(len(ws_names))

    def _selected_rule(self):
        return self.rules_grid.SelectedItem

    def _selected_rules(self):
        items = self.rules_grid.SelectedItems
        return list(items) if items else []

    # --------------------------------------------------
    # Toolbar handlers
    # --------------------------------------------------

    def btn_add_click(self, sender, args):
        ws_names = get_workset_names()
        new_rule = WorksetRule(
            run_order=len(self._rules) + 1,
            name="New Rule",
            workset=ws_names[0] if ws_names else "",
        )
        self._rules.append(new_rule)
        self._renumber()
        self._refresh_grid()
        self._update_status()
        self._dirty = True
        self.rules_grid.SelectedItem = new_rule
        self.rules_grid.ScrollIntoView(new_rule)

    def btn_delete_click(self, sender, args):
        selected = self._selected_rules()
        if not selected:
            return
        for r in selected:
            self._rules.remove(r)
        self._renumber()
        self._refresh_grid()
        self._update_status()
        self._dirty = True

    def btn_duplicate_click(self, sender, args):
        sel = self._selected_rule()
        if not sel:
            return
        idx   = self._rules.index(sel)
        clone = sel.clone()
        clone.Name = sel.Name + " (Copy)"
        self._rules.insert(idx + 1, clone)
        self._renumber()
        self._refresh_grid()
        self._update_status()
        self._dirty = True

    def btn_move_up_click(self, sender, args):
        sel = self._selected_rule()
        if not sel:
            return
        idx = self._rules.index(sel)
        if idx > 0:
            self._rules[idx], self._rules[idx - 1] = (
                self._rules[idx - 1], self._rules[idx]
            )
            self._renumber()
            self._refresh_grid()
            self.rules_grid.SelectedItem = sel
            self._dirty = True

    def btn_move_down_click(self, sender, args):
        sel = self._selected_rule()
        if not sel:
            return
        idx = self._rules.index(sel)
        if idx < len(self._rules) - 1:
            self._rules[idx], self._rules[idx + 1] = (
                self._rules[idx + 1], self._rules[idx]
            )
            self._renumber()
            self._refresh_grid()
            self.rules_grid.SelectedItem = sel
            self._dirty = True

    def btn_apply_view_click(self, sender, args):
        view  = uidoc.ActiveView
        total, msgs = apply_rules(self._rules, view_id=view.Id)
        self._show_results(total, msgs, scope="View: " + view.Name)
        self._update_status()

    def btn_apply_project_click(self, sender, args):
        n_enabled = sum(1 for r in self._rules if r.Enabled)
        if not _confirm(
            "Apply {} enabled rule(s) to the ENTIRE project?\n\n"
            "This will reassign worksets for all matching elements.".format(n_enabled),
            title="Apply to Project"
        ):
            return
        total, msgs = apply_rules(self._rules)
        self._show_results(total, msgs, scope="Entire Project")
        self._update_status()

    def btn_create_worksets_click(self, sender, args):
        workset_list = load_workset_list()
        if not doc.IsWorkshared:
            if not _confirm(
                "Worksharing is not enabled.\nEnable worksharing now?",
                title="Enable Worksharing"
            ):
                return
            if not enable_worksharing():
                forms.alert("Failed to enable worksharing.")
                return

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
        self._update_status()

    def btn_remove_unused_click(self, sender, args):
        workset_list  = load_workset_list()
        all_ws        = get_user_worksets()
        all_names     = [ws.Name for ws in all_ws]
        unused        = [ws for ws in all_ws if ws.Name not in workset_list]

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
                current = get_user_worksets()
                if _remove_workset(name, dest, current):
                    deleted += 1
            else:
                print("No destination found for '{}', skipping".format(name))

        forms.alert(
            "Removed {} of {} selected workset(s).".format(
                deleted, len(selected_names)),
            title="Done"
        )
        self._update_status()

    def btn_import_click(self, sender, args):
        filepath = forms.pick_file(file_ext="json")
        if not filepath:
            return
        try:
            with open(filepath, "r") as f:
                data = json.load(f)
            rules = [WorksetRule.from_dict(r) for r in data.get("rules", [])]
            if not rules:
                forms.alert("No rules found in the selected file.", title="Import")
                return
            self._rules = rules
            self._renumber()
            self._refresh_grid()
            self._update_status()
            self._dirty = True
            forms.alert(
                "Imported {} rule(s) from:\n{}".format(len(rules), filepath),
                title="Import Successful"
            )
        except Exception as e:
            forms.alert("Failed to import file:\n{}".format(e), title="Import Error")

    def btn_export_click(self, sender, args):
        # Save current rules first, then let user copy the file
        self._save_current()
        forms.alert(
            "Rules exported to:\n\n{}".format(RULES_FILE),
            title="Export"
        )

    def btn_reset_click(self, sender, args):
        if not _confirm(
            "Reset all rules to factory defaults?\nThis cannot be undone.",
            title="Reset Rules"
        ):
            return
        self._rules = [WorksetRule.from_dict(r) for r in DEFAULT_RULES]
        self._renumber()
        self._refresh_grid()
        self._update_status()
        self._dirty = True

    def btn_save_click(self, sender, args):
        self._save_current()
        forms.alert(
            "Rules saved to:\n{}".format(RULES_FILE),
            title="Saved"
        )

    # --------------------------------------------------
    # DataGrid cell edit
    # --------------------------------------------------

    def grid_cell_edit_ending(self, sender, args):
        self._dirty = True
        self._update_status()

    # --------------------------------------------------
    # Window closing
    # --------------------------------------------------

    def window_closing(self, sender, args):
        if self._dirty:
            save_rules(self._rules)

    # --------------------------------------------------
    # Internal helpers
    # --------------------------------------------------

    def _save_current(self):
        save_rules(self._rules)
        self._dirty = False

    def _show_results(self, total, msgs, scope=""):
        ok_lines  = [m for m in msgs if m.startswith("OK")]
        skip_lines = [m for m in msgs if m.startswith("SKIP")]
        err_lines  = [m for m in msgs if m.startswith("ERROR")]

        summary = (
            "Scope: {}\n"
            "Total elements reassigned: {}\n\n"
            "Rules applied:  {}\n"
            "Rules skipped:  {}\n"
            "Errors:         {}"
        ).format(scope, total, len(ok_lines), len(skip_lines), len(err_lines))

        if skip_lines or err_lines:
            detail = "\n\nDetails:\n" + "\n".join(skip_lines + err_lines)
            summary += detail

        TaskDialog.Show("Apply Results", summary)


# ==================================================
# MAIN ENTRY POINT
# ==================================================

# --------------------------------------------------
# SHIFT+CLICK  ->  Quick remove unused worksets
# --------------------------------------------------
if __shiftclick__:
    workset_list = load_workset_list()
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

# --------------------------------------------------
# NORMAL CLICK  ->  Open Workset Manager window
# --------------------------------------------------
else:
    WorksetManagerWindow().ShowDialog()
