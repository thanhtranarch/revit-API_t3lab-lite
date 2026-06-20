# -*- coding: utf-8 -*-
"""
FamilyCreatorDialog.py
Combined WPF dialog for Family Creator — CAD, JSON, and Batch modes.
"""

import os
import re
import sys
import math
import traceback
import json

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState, Visibility as WinVis, Clipboard
from System.Windows.Controls import DataGridComboBoxColumn, DataGridLength
from System.Windows.Data import Binding, BindingMode, UpdateSourceTrigger

from pyrevit import forms
import pyrevit.script as _pyrevit_script

logger = _pyrevit_script.get_logger()
_GUI_DIR = os.path.dirname(__file__)
_XAML = os.path.join(_GUI_DIR, 'Tools', 'FamilyCreator.xaml')

_LIB_DIR = os.path.dirname(_GUI_DIR)
if _LIB_DIR not in sys.path:
    sys.path.insert(0, _LIB_DIR)

from Autodesk.Revit.DB import (
    ImportInstance, FilteredElementCollector,
    Options, GeometryInstance,
    Line, Arc, XYZ, Plane,
    CurveArray, CurveArrArray,
    SketchPlane, SaveAsOptions,
    Transaction, ElementId,
    View, ViewType, ReferencePlane, ReferenceArray,
    PlanarFace, Solid, IFailuresPreprocessor, FailureProcessingResult, FailureSeverity,
    Transform,
)

from Utils.DWGFamilyHelpers import get_xy_bounds, _project_curve_to_z as _dwg_project_curve

# ==============================================================================
# CONSTANTS
# ==============================================================================
SCL = 1.0 / 304.8
MAX_DETAIL_CURVES = 150
MIN_CURVE_RATIO = 0.30

_DISCIPLINES = [
    "Architecture", "Structure", "Mechanical", "Electrical",
    "Plumbing", "Fire Protection", "General",
]

_CATEGORY_TEMPLATES = [
    ("Generic Model",        ["Generic Model.rft", "Metric Generic Model.rft"]),
    ("Door",                 ["Door.rft", "Metric Door.rft"]),
    ("Window",               ["Window.rft", "Metric Window.rft"]),
    ("Furniture",            ["Furniture.rft", "Metric Furniture.rft"]),
    ("Plumbing Fixture",     ["Plumbing Fixture.rft", "Metric Plumbing Fixture.rft"]),
    ("Electrical Equipment", ["Electrical Equipment.rft"]),
    ("Mechanical Equipment", ["Mechanical Equipment.rft"]),
    ("Specialty Equipment",  ["Specialty Equipment.rft", "Metric Specialty Equipment.rft"]),
    ("Casework",             ["Casework.rft", "Metric Casework.rft"]),
    ("Columns",              ["Column.rft", "Metric Column.rft"]),
    ("Lighting Fixture",     ["Lighting Fixture.rft", "Metric Lighting Fixture.rft"]),
    ("Site",                 ["Site.rft", "Metric Site.rft"]),
    ("Entourage",            ["Entourage.rft", "Metric Entourage.rft"]),
]

DOOR_PRESETS = [
    ("Single_Swing_700x2100",   700, 2100, 65, 25, 25, 40, 1),
    ("Single_Swing_810x2200",   810, 2200, 65, 25, 25, 40, 1),
    ("Single_Swing_900x2200",   900, 2200, 65, 25, 25, 40, 1),
    ("Single_Swing_1000x2200", 1000, 2200, 65, 25, 25, 40, 1),
    ("Single_Swing_810x2400",   810, 2400, 65, 25, 25, 40, 1),
    ("Single_Swing_900x2400",   900, 2400, 65, 25, 25, 40, 1),
    ("Double_Swing_1600x2200", 1600, 2200, 65, 25, 25, 40, 2),
    ("Double_Swing_1800x2200", 1800, 2200, 65, 25, 25, 40, 2),
    ("Double_Swing_2000x2200", 2000, 2200, 65, 25, 25, 40, 2),
    ("Double_Swing_1600x2400", 1600, 2400, 65, 25, 25, 40, 2),
]

WINDOW_PRESETS = [
    ("Fixed_600x1200",      600, 1200),
    ("Fixed_600x1500",      600, 1500),
    ("Fixed_900x1200",      900, 1200),
    ("Fixed_900x1500",      900, 1500),
    ("Fixed_1200x1500",    1200, 1500),
    ("Fixed_1500x1500",    1500, 1500),
    ("Fixed_1800x1500",    1800, 1500),
    ("Casement_1200x1500", 1200, 1500),
    ("Casement_1500x1500", 1500, 1500),
    ("Casement_1800x1500", 1800, 1500),
    ("Casement_2400x1500", 2400, 1500),
    ("Sliding_1800x1500",  1800, 1500),
    ("Sliding_2400x1500",  2400, 1500),
]

FURNITURE_PRESETS = [
    ("Chair_Office_600x600x900",    600,   600,  900),
    ("Chair_Dining_500x500x800",    500,   500,  800),
    ("Table_Dining_4pax_1200x800", 1200,   800,  750),
    ("Table_Dining_6pax_1600x800", 1600,   800,  750),
    ("Table_Conference_2400x1000", 2400,  1000,  750),
    ("Table_Conference_3600x1200", 3600,  1200,  750),
    ("Desk_Office_1200x600",       1200,   600,  750),
    ("Desk_Office_1600x700",       1600,   700,  750),
    ("Sofa_2Seat_1400x850",        1400,   850,  800),
    ("Sofa_3Seat_2000x850",        2000,   850,  800),
    ("Bed_Single_1000x2000",       1000,  2000,  500),
    ("Bed_Double_1600x2000",       1600,  2000,  500),
    ("Bed_King_1800x2000",         1800,  2000,  500),
    ("Wardrobe_2Door_1200x600",    1200,   600, 2200),
    ("Wardrobe_3Door_1800x600",    1800,   600, 2200),
    ("Bookcase_900x300x2100",       900,   300, 2100),
]

CASEWORK_PRESETS = [
    ("Cabinet_Base_600x600x850",       600,   600,  850),
    ("Cabinet_Base_800x600x850",       800,   600,  850),
    ("Cabinet_Wall_600x300x600",       600,   300,  600),
    ("Cabinet_Wall_800x300x600",       800,   300,  600),
    ("Counter_Kitchen_1200x600",      1200,   600,  900),
    ("Counter_Kitchen_1800x600",      1800,   600,  900),
    ("Counter_Kitchen_2400x600",      2400,   600,  900),
    ("Island_Kitchen_1200x900",       1200,   900,  900),
    ("Island_Kitchen_1500x900",       1500,   900,  900),
    ("Vanity_Unit_900x500x850",        900,   500,  850),
    ("Vanity_Unit_1200x500x850",      1200,   500,  850),
    ("Shelf_Unit_900x300x2100",        900,   300, 2100),
    ("Display_Cabinet_1200x400x2200", 1200,   400, 2200),
]

PLUMBING_PRESETS = [
    ("WC_Toilet_Std_380x700",      380,  700, 400),
    ("WC_Toilet_Compact_360x650",  360,  650, 400),
    ("WC_Wall_Hung_380x560",       380,  560, 390),
    ("Sink_Vanity_600x500",        600,  500, 150),
    ("Sink_Vanity_800x500",        800,  500, 150),
    ("Sink_Kitchen_800x500",       800,  500, 200),
    ("Sink_Kitchen_1000x500",     1000,  500, 200),
    ("Bath_Builtin_1500x700",     1500,  700, 600),
    ("Bath_Builtin_1700x700",     1700,  700, 600),
    ("Bath_Freestanding_1700x800",1700,  800, 600),
    ("Shower_Tray_900x900",        900,  900, 150),
    ("Shower_Tray_1200x800",      1200,  800, 150),
    ("Shower_Tray_1200x900",      1200,  900, 150),
    ("Urinal_Std_360x330",         360,  330, 560),
    ("Floor_Drain_150x150",        150,  150,  80),
]

LIGHTING_PRESETS = [
    ("Downlight_D100",        100,  100,  80),
    ("Downlight_D150",        150,  150,  80),
    ("Downlight_D200",        200,  200, 100),
    ("Panel_300x600",         300,  600,  80),
    ("Panel_300x1200",        300, 1200,  80),
    ("Panel_600x600",         600,  600,  80),
    ("Ceiling_Round_D300",    300,  300, 120),
    ("Ceiling_Round_D400",    400,  400, 150),
    ("Ceiling_Round_D600",    600,  600, 180),
    ("Linear_Strip_1200",     100, 1200,  60),
    ("Linear_Strip_2400",     100, 2400,  60),
    ("Pendant_D200",          200,  200, 300),
    ("Pendant_D400",          400,  400, 300),
    ("Wall_Sconce_200x100",   200,  100, 250),
    ("Floodlight_300x200",    300,  200, 150),
]

MECHANICAL_PRESETS = [
    ("FCU_Cassette_600x600",          600,   600,  280),
    ("FCU_Cassette_900x900",          900,   900,  280),
    ("FCU_Cassette_1200x1200",       1200,  1200,  280),
    ("FCU_Wall_Mount_900x300",        900,   300,  250),
    ("FCU_Wall_Mount_1200x300",      1200,   300,  250),
    ("AHU_Floor_800x600x1500",        800,   600, 1500),
    ("AHU_Floor_1200x800x1800",      1200,   800, 1800),
    ("AHU_Ceiling_1500x800x500",     1500,   800,  500),
    ("Chiller_2000x1000x1500",       2000,  1000, 1500),
    ("Cooling_Tower_2000x2000x3000", 2000,  2000, 3000),
    ("Pump_600x400x500",              600,   400,  500),
    ("Boiler_900x700x1200",           900,   700, 1200),
    ("Expansion_Tank_D600x900",       600,   600,  900),
]

ELECTRICAL_PRESETS = [
    ("Panel_DB_500x200x600",           500,  200,  600),
    ("Panel_DB_600x250x1000",          600,  250, 1000),
    ("Panel_MDB_800x400x1200",         800,  400, 1200),
    ("Cabinet_Control_800x400x1800",   800,  400, 1800),
    ("Cabinet_Control_1000x500x2000", 1000,  500, 2000),
    ("UPS_600x600x1000",               600,  600, 1000),
    ("UPS_800x700x1200",               800,  700, 1200),
    ("Transformer_1000x700x1400",     1000,  700, 1400),
    ("Switchgear_800x600x2000",        800,  600, 2000),
    ("Switchgear_1200x800x2200",      1200,  800, 2200),
    ("Socket_Outlet_86x86",             86,   86,   60),
    ("Junction_Box_150x150x100",       150,  150,  100),
]

SPECIALTY_PRESETS = [
    ("Counter_Reception_2000x800x1100",  2000,  800, 1100),
    ("Counter_Reception_3000x800x1100",  3000,  800, 1100),
    ("Counter_Service_1500x700x900",     1500,  700,  900),
    ("ATM_Machine_500x500x1800",          500,  500, 1800),
    ("Vending_Machine_700x800x1900",      700,  800, 1900),
    ("Safe_600x500x800",                  600,  500,  800),
    ("Server_Rack_600x1000x2000",         600, 1000, 2000),
    ("Server_Rack_800x1200x2200",         800, 1200, 2200),
    ("Kiosk_Self_Service_700x700x1500",   700,  700, 1500),
    ("Turnstile_1000x600x1000",          1000,  600, 1000),
    ("Fire_Extinguisher_D200x550",        200,  200,  550),
    ("Fire_Hose_Cabinet_700x250x900",     700,  250,  900),
]

COLUMN_PRESETS = [
    ("Col_Square_200x200x3000", 200, 200, 3000),
    ("Col_Square_250x250x3000", 250, 250, 3000),
    ("Col_Square_300x300x3000", 300, 300, 3000),
    ("Col_Square_400x400x3000", 400, 400, 3000),
    ("Col_Square_500x500x3000", 500, 500, 3000),
    ("Col_Square_600x600x3000", 600, 600, 3000),
    ("Col_Rect_200x400x3000",   200, 400, 3000),
    ("Col_Rect_250x500x3000",   250, 500, 3000),
    ("Col_Rect_300x600x3000",   300, 600, 3000),
    ("Col_Rect_350x700x3000",   350, 700, 3000),
    ("Col_Round_D300x3000",     300, 300, 3000),
    ("Col_Round_D400x3000",     400, 400, 3000),
    ("Col_Round_D500x3000",     500, 500, 3000),
]

GENERIC_PRESETS = [
    ("Box_100x100x100",     100,  100,  100),
    ("Box_200x200x200",     200,  200,  200),
    ("Box_500x500x500",     500,  500,  500),
    ("Box_1000x500x500",   1000,  500,  500),
    ("Box_1000x1000x500",  1000, 1000,  500),
    ("Box_1000x1000x1000", 1000, 1000, 1000),
    ("Slab_2000x1000x200", 2000, 1000,  200),
    ("Slab_3000x2000x300", 3000, 2000,  300),
    ("Wall_3000x200x3000", 3000,  200, 3000),
]

ENTOURAGE_PRESETS = [
    ("Tree_Small_D2000x4000",        2000, 2000, 4000),
    ("Tree_Medium_D3000x6000",       3000, 3000, 6000),
    ("Tree_Large_D5000x8000",        5000, 5000, 8000),
    ("Shrub_D1000x600",              1000, 1000,  600),
    ("Person_Standing_600x300x1750",  600,  300, 1750),
    ("Person_Seated_600x700x1200",    600,  700, 1200),
    ("Car_Sedan_4500x1900x1450",     4500, 1900, 1450),
    ("Car_SUV_4700x2000x1700",       4700, 2000, 1700),
    ("Bicycle_1800x600x1100",        1800,  600, 1100),
    ("Motorcycle_2200x800x1200",     2200,  800, 1200),
]

SITE_PRESETS = [
    ("Parking_Space_2500x5000",   2500, 5000,   50),
    ("Parking_Space_2700x5500",   2700, 5500,   50),
    ("Disabled_Parking_3500x5000",3500, 5000,   50),
    ("Bike_Stand_600x200x1000",    600,  200, 1000),
    ("Bollard_D200x800",           200,  200,  800),
    ("Planter_Box_1000x500x600",  1000,  500,  600),
    ("Planter_Box_2000x500x600",  2000,  500,  600),
    ("Bench_1800x500x450",        1800,  500,  450),
    ("Bench_1200x500x450",        1200,  500,  450),
    ("Sign_Post_100x100x3000",     100,  100, 3000),
    ("Waste_Bin_D400x800",         400,  400,  800),
    ("Light_Pole_D200x5000",       200,  200, 5000),
]

CATEGORY_PRESETS = {
    "Generic Model":        ("generic", GENERIC_PRESETS),
    "Door":                 ("door",    DOOR_PRESETS),
    "Window":               ("window",  WINDOW_PRESETS),
    "Furniture":            ("generic", FURNITURE_PRESETS),
    "Plumbing Fixture":     ("generic", PLUMBING_PRESETS),
    "Electrical Equipment": ("generic", ELECTRICAL_PRESETS),
    "Mechanical Equipment": ("generic", MECHANICAL_PRESETS),
    "Specialty Equipment":  ("generic", SPECIALTY_PRESETS),
    "Casework":             ("generic", CASEWORK_PRESETS),
    "Columns":              ("generic", COLUMN_PRESETS),
    "Lighting Fixture":     ("generic", LIGHTING_PRESETS),
    "Site":                 ("generic", SITE_PRESETS),
    "Entourage":            ("generic", ENTOURAGE_PRESETS),
}

_CAT_HINTS = [
    ("Door", [
        "door", "swing", "sliding door", "folding door",
        "cua di", "cuadi", "cua ra", "cua vao",
        "a-door", "kl-door", "arch-door", "-door",
    ]),
    ("Window", [
        "window", "casement", "skylight",
        "cua so", "cuaso",
        "a-wind", "kl-wind", "-wind",
    ]),
    ("Furniture", [
        "furnitur", "chair", "table", "desk", "sofa", "bed",
        "armchair", "bookcase", "wardrobe", "lounge", "seating",
        "noi that", "noithat", "ban ghe", "banghe",
        "ban lam viec", "ghe ngoi", "tu quan ao",
        "a-furn", "kl-furn", "-furn", "ff&e", "ff-e",
        "-furniture", "a-furniture",
    ]),
    ("Casework", [
        "casework", "counter", "kitchen", "shelv", "cabinet",
        "cupboard", "joinery",
        "tu bep", "tubep", "tu am tuong", "tu bep duoi",
        "quay bep", "quay le tan", "bep",
        "a-case", "kl-case", "-casework",
    ]),
    ("Plumbing Fixture", [
        "plumb", "sanitary", "toilet", "wc", "sink", "basin",
        "shower", "bath", "urinal", "lavatory", "bidet",
        "thiet bi ve sinh", "thietbivesinnh", "bon tam", "bon cau",
        "chau rua", "chau lavabo", "thiet bi nuoc", "ve sinh",
        "p-fixt", "m-plmb", "kl-plmb", "-plumb", "-sanitary",
        "eqpm-fixd", "eqpm-fix", "eqpm",
    ]),
    ("Lighting Fixture", [
        "light", "lamp", "luminaire", "led", "spotlight",
        "downlight", "pendant", "sconce", "chandelier",
        "den", "den chieu sang", "chieu sang", "den treo",
        "den am tran", "den tuong",
        "e-lite", "e-lght", "kl-lght", "-light", "-lite",
        "-lighting", "a-lighting",
    ]),
    ("Mechanical Equipment", [
        "mechanical", "hvac", "ahu", "fcu", "fahu", "chiller",
        "cooling", "boiler", "pump", "fan", "duct", "damper",
        "may lanh", "maylanh", "dieu hoa", "dieuhoa",
        "thong gio", "cap nhiet", "bom nhiet",
        "m-equip", "m-mech", "kl-mech", "-mech", "-hvac",
    ]),
    ("Electrical Equipment", [
        "electrical", "switchgear", "transformer", "ups",
        "panel", "mdb", "smdb", "db", "mcb", "busbar",
        "dien", "tu dien", "tudien", "bang dien", "thiet bi dien",
        "e-equip", "e-powr", "kl-elec", "-elec", "-electr",
    ]),
    ("Specialty Equipment", [
        "machine", "appliance", "kiosk", "atm",
        "vending", "server", "rack",
        "may moc",
        "a-equip", "kl-equip",
    ]),
    ("Columns", [
        "column", " col ", "pillar", "pier", "post",
        "struc", "structural", "ket cau", "ketcau",
        "beam", "slab", "footing", "foundation",
        "cot", "tru", "dam", "san", "mong",
        "s-col", "a-col", "kl-col", "-col-", "-column",
        "kc-", "s-beam", "s-slab", "s-wall", "s-str",
    ]),
    ("Site", [
        "site", "parking", "landscape", "paving", "bollard",
        "tree", "bench", "pavement",
        "san vuon", "cay xanh", "bai xe", "he thong ngoai that",
        "l-site", "a-site", "kl-site", "-site", "-land",
    ]),
    ("Entourage", [
        "entourage", "person", "people", "car", "vehicle",
        "bicycle", "human", "figure",
        "nguoi", "xe hoi", "xe dap",
        "-entour", "a-entour",
    ]),
    ("Generic Model", [
        "wall", "tuong", "a-wall", "kl-wall", "s-wall-",
        "glass", "glazing", "curtain", "kinh",
        "title block", "titleblock", "title blk", "title-blk",
        "khung ten", "khungten", "khung-ten",
        "border", "sheet border", "annotation", "detailitem",
        "detail item", "tb-", "-tblock",
    ]),
]


# ==============================================================================
# HELPER CLASSES / FUNCTIONS
# ==============================================================================

class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        fail_list = failuresAccessor.GetFailureMessages()
        if fail_list.Count == 0:
            return FailureProcessingResult.Continue
        for failure in fail_list:
            if failure.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue


def start_transaction(t):
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(WarningSwallower())
    t.SetFailureHandlingOptions(options)
    return t.Start()


def _suggest_category(name, arc_count, width_mm, depth_mm, layer=""):
    combined = (name + " " + layer).lower()
    for cat, keywords in _CAT_HINTS:
        if any(k in combined for k in keywords):
            return "Generic Model" if cat == "Door" else cat
    if arc_count == 0 and 0 < depth_mm < 350 and width_mm >= 400:
        return "Window"
    return "Generic Model"


def _graphicstyle_layer(geom_elem, doc):
    for item in geom_elem:
        try:
            sid = getattr(item, 'GraphicsStyleId', None)
            if sid and sid != ElementId.InvalidElementId:
                style = doc.GetElement(sid)
                if style:
                    try:
                        cat = style.GraphicsStyleCategory
                        if cat and cat.Name:
                            return cat.Name
                    except Exception:
                        pass
                    try:
                        if style.Name:
                            return style.Name
                    except Exception:
                        pass
        except Exception:
            pass
        if isinstance(item, GeometryInstance):
            try:
                nested = item.GetInstanceGeometry()
                if nested:
                    result = _graphicstyle_layer(nested, doc)
                    if result:
                        return result
            except Exception:
                pass
    return ""


class BlockItem(object):
    def __init__(self, name, curve_count, instance_count, curves,
                 layer_level="", placements=None, import_inst=None):
        self.IsSelected    = True
        self.BlockName     = name
        self.CurveCount    = curve_count
        self.InstanceCount = instance_count
        self.LayerLevel    = layer_level
        self._curves       = curves
        self._placements   = placements if placements is not None else []
        self._import_inst  = import_inst

        arc_count = sum(1 for c in curves if isinstance(c, Arc))
        self.ArcCount = arc_count

        try:
            min_x, max_x, min_y, max_y = get_xy_bounds(curves)
            w = (max_x - min_x) * 304.8
            d = (max_y - min_y) * 304.8
            self.WidthMM = "{:.0f}".format(w)
            self.DepthMM = "{:.0f}".format(d)
        except Exception:
            w, d = 0.0, 0.0
            self.WidthMM = "-"
            self.DepthMM = "-"

        self.SuggestedCat = _suggest_category(name, arc_count, w, d, layer=layer_level)
        self.Category     = self.SuggestedCat


# Aliases so methods copied verbatim from CAD script compile without change
DISCIPLINES       = _DISCIPLINES
CATEGORY_TEMPLATES = _CATEGORY_TEMPLATES


# ==============================================================================
# COMBINED DIALOG
# ==============================================================================

class FamilyCreatorDialog(forms.WPFWindow):

    def __init__(self, revit_doc, revit_app, initial_mode='cad'):
        forms.WPFWindow.__init__(self, _XAML)
        self._doc = revit_doc
        self._app = revit_app
        self._block_items      = []
        self._cad_instances    = []
        self._filter_text      = ""
        self._filter_cat       = ""
        self._cancel_requested = False
        self._pause_requested  = False

        self._init_cad_panel()

        self.mode_cad.Checked  += self.mode_cad_checked
        self.mode_json.Checked += self.mode_json_checked

        if initial_mode == 'json':
            self.mode_json.IsChecked = True
            self._show_panel('json')
        else:
            self.mode_cad.IsChecked = True
            self._show_panel('cad')

    # ── Window chrome ────────────────────────────────────────────────────────

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

    # ── Mode switching ───────────────────────────────────────────────────────

    def _show_panel(self, mode):
        self.panel_cad.Visibility  = WinVis.Visible if mode == 'cad'  else WinVis.Collapsed
        self.panel_json.Visibility = WinVis.Visible if mode == 'json' else WinVis.Collapsed

    def mode_cad_checked(self, sender, e):
        self._show_panel('cad')

    def mode_json_checked(self, sender, e):
        self._show_panel('json')


    # ── Status helpers ───────────────────────────────────────────────────────

    def _update_status(self, text):
        try:
            self.status_text.Text = text
        except Exception:
            pass

    # ── CAD panel initialisation ─────────────────────────────────────────────

    def _init_cad_panel(self):
        self._init_cad_files()
        self._init_disciplines()
        self._init_categories()
        self._init_filter_bar()
        self._update_status("Ready")

    def _init_cad_files(self):
        collector = FilteredElementCollector(self._doc).OfClass(ImportInstance)
        self.cad_file_combo.Items.Add("<All Imported CAD Files>")
        for inst in collector:
            name = self._get_cad_name(inst)
            self._cad_instances.append(inst)
            self.cad_file_combo.Items.Add(name)
        if self._cad_instances:
            self.cad_file_combo.SelectedIndex = 0

    @staticmethod
    def _read_symbol_name(inst):
        from Autodesk.Revit.DB import BuiltInParameter as BIP
        try:
            p = inst.get_Parameter(BIP.IMPORT_SYMBOL_NAME)
            if p and p.HasValue:
                val = p.AsString()
                if val:
                    return val
        except Exception:
            pass
        try:
            for p in inst.Parameters:
                if p.Definition.Name == "Name" and p.StorageType.ToString() == "String":
                    val = p.AsString()
                    if val:
                        return val
        except Exception:
            pass
        try:
            type_id = inst.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                elem_type = inst.Document.GetElement(type_id)
                if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
                    return elem_type.Name
        except Exception:
            pass
        return inst.Name if hasattr(inst, 'Name') else "Unknown"

    def _get_cad_name(self, inst):
        return self._read_symbol_name(inst)

    def _init_disciplines(self):
        for name in DISCIPLINES:
            self.discipline_combo.Items.Add(name)
        self.discipline_combo.SelectedIndex = 6

    def _init_categories(self):
        cat_names = [name for name, _ in CATEGORY_TEMPLATES]
        for name in cat_names:
            self.category_combo.Items.Add(name)
        self.category_combo.SelectedIndex = 0

        col = DataGridComboBoxColumn()
        col.Header = "Category"
        col.Width = DataGridLength(140)
        col.ItemsSource = cat_names
        b = Binding("Category")
        b.Mode = BindingMode.TwoWay
        b.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        col.SelectedItemBinding = b
        self.blocks_grid.Columns.Add(col)

    def _init_filter_bar(self):
        self.combo_filter_suggested.Items.Add("All Categories")
        self.combo_filter_suggested.SelectedIndex = 0

    # ── Filter ───────────────────────────────────────────────────────────────

    def _refresh_suggested_combo(self):
        prev = self._filter_cat
        self.combo_filter_suggested.SelectionChanged -= self.filter_suggested_changed
        self.combo_filter_suggested.Items.Clear()
        self.combo_filter_suggested.Items.Add("All Categories")
        cats = sorted(set(item.Category for item in self._block_items if item.Category))
        for c in cats:
            self.combo_filter_suggested.Items.Add(c)
        if prev and prev in cats:
            self.combo_filter_suggested.SelectedItem = prev
        else:
            self.combo_filter_suggested.SelectedIndex = 0
            self._filter_cat = ""
        self.combo_filter_suggested.SelectionChanged += self.filter_suggested_changed

    def _apply_filter(self):
        txt = self._filter_text.lower().strip()
        cat = self._filter_cat
        if not txt and not cat:
            visible = self._block_items
        else:
            visible = []
            for item in self._block_items:
                if txt and txt not in item.BlockName.lower() \
                        and txt not in item.LayerLevel.lower():
                    continue
                if cat and item.Category != cat:
                    continue
                visible.append(item)
        self.blocks_grid.ItemsSource = visible
        total   = len(self._block_items)
        showing = len(visible)
        if total == 0:
            self.txt_filter_count.Text = ""
        elif showing == total:
            self.txt_filter_count.Text = "{} items".format(total)
        else:
            self.txt_filter_count.Text = "{} / {} items".format(showing, total)

    def search_text_changed(self, sender, e):
        self._filter_text = self.txt_search.Text or ""
        has_text = bool(self._filter_text)
        self.txt_search_placeholder.Visibility = WinVis.Collapsed if has_text else WinVis.Visible
        self.btn_clear_search.Visibility       = WinVis.Visible   if has_text else WinVis.Collapsed
        self._apply_filter()

    def clear_search_clicked(self, sender, e):
        self.txt_search.Text = ""

    def filter_suggested_changed(self, sender, e):
        sel = self.combo_filter_suggested.SelectedItem
        self._filter_cat = "" if (sel is None or sel == "All Categories") else sel
        self._apply_filter()

    # ── Progress ─────────────────────────────────────────────────────────────

    def _do_events(self):
        try:
            from System.Windows.Threading import DispatcherPriority, DispatcherFrame
            import System
            frame = DispatcherFrame()
            def _stop(f=frame):
                f.Continue = False
            self.Dispatcher.BeginInvoke(DispatcherPriority.Background, System.Action(_stop))
            self.Dispatcher.PushFrame(frame)
        except Exception:
            pass

    def _update_progress(self, value, maximum):
        try:
            self.pb_export.Maximum = maximum
            self.pb_export.Value   = value
            self.progress_panel.Visibility = WinVis.Visible
        except Exception:
            pass
        self._do_events()
        while self._pause_requested and not self._cancel_requested:
            self._do_events()

    def _hide_progress(self):
        try:
            self.progress_panel.Visibility = WinVis.Collapsed
            self.pb_export.Value = 0
            self._cancel_requested = False
            self._pause_requested  = False
            self.btn_pause_export.Content  = u"⏸  Pause"
            self.btn_pause_export.IsEnabled = True
            self.btn_stop_export.IsEnabled  = True
        except Exception:
            pass

    def stop_export_clicked(self, sender, e):
        self._cancel_requested = True
        self._pause_requested  = False
        try:
            self.btn_stop_export.IsEnabled = False
            self._update_status(u"Stopping… finishing current block")
        except Exception:
            pass

    def pause_resume_clicked(self, sender, e):
        if self._pause_requested:
            self._pause_requested = False
            try:
                self.btn_pause_export.Content = u"⏸  Pause"
            except Exception:
                pass
        else:
            self._pause_requested = True
            try:
                self.btn_pause_export.Content = u"▶  Resume"
                self._update_status(u"Paused — click Resume to continue")
            except Exception:
                pass

    # ── Scanning ─────────────────────────────────────────────────────────────

    def scan_blocks_clicked(self, sender, e):
        if not self._cad_instances:
            forms.alert("No imported CAD files found in the document.")
            return
        idx = self.cad_file_combo.SelectedIndex - 1
        if idx < -1 or idx >= len(self._cad_instances):
            forms.alert("Please select a CAD file.")
            return
        self._update_status("Scanning blocks...")
        blocks = []
        try:
            if idx == -1:
                name_counts = {}
                for inst in self._cad_instances:
                    item = self._scan_entire_cad(inst)
                    if item:
                        base_name = item.BlockName
                        if base_name in name_counts:
                            name_counts[base_name] += 1
                            item.BlockName = "{}_{}".format(base_name, name_counts[base_name])
                        else:
                            name_counts[base_name] = 1
                        blocks.append(item)
            else:
                import_inst = self._cad_instances[idx]
                blocks = self._scan_blocks(import_inst)
                if not blocks:
                    item = self._scan_entire_cad(import_inst)
                    if item:
                        blocks.append(item)
        except Exception as ex:
            logger.error("Scan error:\n{}".format(traceback.format_exc()))
            forms.alert("Error scanning blocks:\n{}".format(str(ex)))
            self._update_status("Scan failed")
            return
        if not blocks:
            forms.alert("No blocks or curves found in the selected CAD file(s).")
            self._update_status("No geometry found")
            return
        self._block_items  = blocks
        self._filter_text  = ""
        self._filter_cat   = ""
        self.txt_search.Text = ""
        self._refresh_suggested_combo()
        self._apply_filter()
        self._update_status("Found {} unique item(s)".format(len(blocks)))
        self.block_count_text.Text = "{} items found".format(len(blocks))

    def _scan_entire_cad(self, import_inst):
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        geom = import_inst.get_Geometry(opt)
        if not geom:
            return None
        min_len = getattr(self._app, 'ShortCurveTolerance', 0.00256)

        def is_curve(item):
            try:
                from Autodesk.Revit.DB import Curve as _Curve
                return isinstance(item, _Curve) and item.IsBound and item.Length >= min_len
            except Exception:
                return False

        def collect_curves(geo_elem):
            from Autodesk.Revit.DB import PolyLine, Curve as _Curve
            curves = []
            for item in geo_elem:
                if is_curve(item):
                    curves.append(item)
                elif isinstance(item, _Curve) and not item.IsBound:
                    curves.append(item)
                elif isinstance(item, PolyLine):
                    pts = item.GetCoordinates()
                    for i in range(item.NumberOfCoordinates - 1):
                        try:
                            p1, p2 = pts[i], pts[i + 1]
                            if p1.DistanceTo(p2) >= min_len:
                                curves.append(Line.CreateBound(p1, p2))
                        except Exception:
                            pass
                elif isinstance(item, GeometryInstance):
                    try:
                        nested = item.GetInstanceGeometry()
                        if nested:
                            curves.extend(collect_curves(nested))
                    except Exception:
                        pass
                elif isinstance(item, Solid):
                    try:
                        for edge in item.Edges:
                            try:
                                ec = edge.AsCurve()
                                if is_curve(ec):
                                    curves.append(ec)
                            except Exception:
                                pass
                    except Exception:
                        pass
            return curves

        curves = collect_curves(geom)
        if curves:
            name = self._get_cad_name(import_inst)
            layer_name = _graphicstyle_layer(geom, self._doc)
            try:
                min_x, max_x, min_y, max_y = get_xy_bounds(curves)
                centroid = XYZ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, 0.0)
            except Exception:
                centroid = XYZ.Zero
            placements = [(centroid, 0.0)]
            return BlockItem(name, len(curves), 1, curves,
                             layer_level=layer_name, placements=placements,
                             import_inst=import_inst)
        return None

    def _scan_blocks(self, import_inst):
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        geom = import_inst.get_Geometry(opt)
        if not geom:
            return []
        min_len = getattr(self._app, 'ShortCurveTolerance', 0.00256)
        found   = {}
        counter = [0]

        def is_curve(item):
            try:
                from Autodesk.Revit.DB import Curve as _Curve
                return isinstance(item, _Curve) and item.IsBound and item.Length >= min_len
            except Exception:
                return False

        def collect_curves(geo_elem):
            curves = []
            for item in geo_elem:
                if is_curve(item):
                    curves.append(item)
                elif isinstance(item, GeometryInstance):
                    try:
                        nested = item.GetInstanceGeometry()
                        if nested:
                            curves.extend(collect_curves(nested))
                    except Exception:
                        pass
                elif isinstance(item, Solid):
                    try:
                        for edge in item.Edges:
                            try:
                                ec = edge.AsCurve()
                                if is_curve(ec):
                                    curves.append(ec)
                            except Exception:
                                pass
                    except Exception:
                        pass
            return curves

        def fingerprint(curves):
            return (len(curves), round(sum(c.Length for c in curves), 1))

        def style_name(geo_inst):
            try:
                sid = geo_inst.GraphicsStyleId
                if sid and sid != ElementId.InvalidElementId:
                    style = self._doc.GetElement(sid)
                    if style:
                        try:
                            cat = style.GraphicsStyleCategory
                            if cat and cat.Name:
                                return cat.Name
                        except Exception:
                            pass
                        try:
                            if style.Name:
                                return style.Name
                        except Exception:
                            pass
            except Exception:
                pass
            return None

        def _instance_placement(curves, geo_inst):
            try:
                min_x, max_x, min_y, max_y = get_xy_bounds(curves)
                centroid = XYZ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, 0.0)
            except Exception:
                centroid = XYZ.Zero
            try:
                bx    = geo_inst.Transform.BasisX
                angle = math.atan2(bx.Y, bx.X)
            except Exception:
                angle = 0.0
            return (centroid, angle)

        def register(curves, geo_inst):
            fp        = fingerprint(curves)
            placement = _instance_placement(curves, geo_inst)
            if fp in found:
                found[fp]['count'] += 1
                found[fp]['placements'].append(placement)
                return
            layer = style_name(geo_inst)
            counter[0] += 1
            block_name = ""
            try:
                if hasattr(geo_inst, 'Symbol') and geo_inst.Symbol:
                    block_name = (geo_inst.Symbol.Name or "").strip()
            except Exception:
                pass
            if not block_name:
                block_name = (FamilyCreatorDialog._read_symbol_name(import_inst) or "").strip()
            if block_name:
                name = block_name
            elif layer:
                name = "{}_Block_{:03d}".format(layer, counter[0])
            else:
                name = "Block_{:03d}".format(counter[0])
            found[fp] = {
                'name': name, 'curves': curves, 'count': 1,
                'layer': layer or "", 'placements': [placement],
            }

        def walk(geo_elem, depth):
            for item in geo_elem:
                if not isinstance(item, GeometryInstance):
                    continue
                inst_geom = item.GetInstanceGeometry()
                if not inst_geom:
                    continue
                if depth == 0:
                    walk(inst_geom, depth + 1)
                else:
                    curves = collect_curves(inst_geom)
                    if curves:
                        register(curves, item)

        walk(geom, 0)

        items = []
        for data in sorted(found.values(), key=lambda d: d['name']):
            items.append(BlockItem(
                data['name'], len(data['curves']), data['count'], data['curves'],
                layer_level=data.get('layer', ""),
                placements=data.get('placements', []),
                import_inst=import_inst))
        return items

    # ── CAD export UI ────────────────────────────────────────────────────────

    def browse_folder_clicked(self, sender, e):
        folder = forms.pick_folder()
        if folder:
            self.output_path.Text = folder

    def select_all_clicked(self, sender, e):
        for item in self._block_items:
            item.IsSelected = True
        self.blocks_grid.Items.Refresh()

    def deselect_all_clicked(self, sender, e):
        for item in self._block_items:
            item.IsSelected = False
        self.blocks_grid.Items.Refresh()

    def export_clicked(self, sender, e):
        output_folder = self.output_path.Text
        if not output_folder or not os.path.isdir(output_folder):
            forms.alert("Please select a valid output folder.")
            return
        selected = [b for b in self._block_items if b.IsSelected]
        if not selected:
            forms.alert("No blocks selected for export.")
            return
        disc_idx = self.discipline_combo.SelectedIndex
        if disc_idx < 0:
            forms.alert("Please select a discipline.")
            return
        discipline_name = DISCIPLINES[disc_idx]
        load_to_project = (self.chk_load_to_project.IsChecked == True)
        mode_2d = False
        try:
            mode_2d = bool(getattr(self, 'rb_2d_lines', None) and self.rb_2d_lines.IsChecked)
        except Exception:
            pass
        self._cancel_requested = False
        self._pause_requested  = False
        self._update_status("Exporting {} block(s)...".format(len(selected)))
        self._update_progress(0, len(selected))
        success, failed = 0, 0
        saved_paths = []
        import System as _System
        for i, item in enumerate(selected):
            if self._cancel_requested:
                break
            try:
                category_name = item.Category or "Generic Model"
                template_path = self._find_template_by_name(category_name)
                if not template_path:
                    logger.warning("No template for '{}', skipping '{}'".format(
                        category_name, item.BlockName))
                    failed += 1
                    self._update_progress(i + 1, len(selected))
                    continue
                self._update_status("Exporting [{}/{}]: {}".format(
                    i + 1, len(selected), item.BlockName))
                self._update_progress(i, len(selected))
                if self._cancel_requested:
                    break
                save_path = self._export_block(
                    item, template_path, output_folder,
                    discipline_name, category_name,
                    load_to_project=False,
                    mode_2d_only=mode_2d)
                if save_path:
                    saved_paths.append(save_path)
                    success += 1
                else:
                    failed += 1
            except Exception:
                logger.error("Export '{}' failed:\n{}".format(
                    item.BlockName, traceback.format_exc()))
                failed += 1
            self._update_progress(i + 1, len(selected))
            if (i + 1) % 10 == 0:
                try:
                    _System.GC.Collect()
                    _System.GC.WaitForPendingFinalizers()
                except Exception:
                    pass
        was_cancelled = self._cancel_requested
        loaded_count = 0
        if not was_cancelled and load_to_project and saved_paths:
            self._update_status("Loading {} families to project...".format(len(saved_paths)))
            loaded_count = self._batch_load_families(saved_paths)
        status = "Stopped" if was_cancelled else "Done"
        self._update_status("{}: {} exported, {} failed".format(status, success, failed))
        self._hide_progress()
        load_note = "\nLoaded to project: {}".format(loaded_count) if load_to_project else ""
        cancelled_note = "\n\nExport was stopped early." if was_cancelled else ""
        forms.alert(
            "Export complete!\n\nExported: {}\nFailed: {}{}{}\n\nOutput folder:\n{}".format(
                success, failed, load_note, cancelled_note, output_folder))

    def export_and_place_clicked(self, sender, e):
        output_folder = self.output_path.Text
        if not output_folder or not os.path.isdir(output_folder):
            forms.alert("Please select a valid output folder.")
            return
        selected = [b for b in self._block_items if b.IsSelected]
        if not selected:
            forms.alert("No blocks selected.")
            return
        disc_idx = self.discipline_combo.SelectedIndex
        discipline_name = DISCIPLINES[disc_idx] if disc_idx >= 0 else "General"
        place_level = None
        try:
            from pyrevit import revit as _revit
            from Autodesk.Revit.DB import Level
            place_level = self._doc.GetElement(_revit.uidoc.ActiveView.GenLevel.Id)
        except Exception:
            pass
        if not place_level:
            try:
                from Autodesk.Revit.DB import Level
                for lv in FilteredElementCollector(self._doc).OfClass(Level):
                    place_level = lv
                    break
            except Exception:
                pass
        mode_2d = False
        try:
            mode_2d = bool(getattr(self, 'rb_2d_lines', None) and self.rb_2d_lines.IsChecked)
        except Exception:
            pass
        exported, placed_total, failed = 0, 0, 0
        self._cancel_requested = False
        self._pause_requested  = False
        self._update_progress(0, len(selected))
        t_place = Transaction(self._doc, "T3Lab - Export & Place Families")
        start_transaction(t_place)
        try:
            for i, item in enumerate(selected):
                if self._cancel_requested:
                    break
                self._update_status(
                    "Exporting & placing [{}/{}]: {}".format(i + 1, len(selected), item.BlockName))
                self._update_progress(i, len(selected))
                if self._cancel_requested:
                    break
                try:
                    category_name = item.Category or "Generic Model"
                    template_path = self._find_template_by_name(category_name)
                    if not template_path:
                        failed += 1
                        self._update_progress(i + 1, len(selected))
                        continue
                    save_path = self._export_block(
                        item, template_path, output_folder,
                        discipline_name, category_name,
                        load_to_project=False, mode_2d_only=mode_2d)
                    if not save_path:
                        failed += 1
                        self._update_progress(i + 1, len(selected))
                        continue
                    exported += 1
                    n = self._place_family_instances(save_path, item, place_level)
                    placed_total += n
                except Exception:
                    logger.error("Export+Place '{}' failed:\n{}".format(
                        item.BlockName, traceback.format_exc()))
                    failed += 1
                self._update_progress(i + 1, len(selected))
            t_place.Commit()
        except Exception:
            try:
                t_place.RollBack()
            except Exception:
                pass
            self._hide_progress()
            logger.error("Export & Place failed:\n{}".format(traceback.format_exc()))
            forms.alert("Transaction failed - check the pyRevit log.")
            return
        was_cancelled = self._cancel_requested
        status = "Stopped" if was_cancelled else "Done"
        self._update_status("{}: {} exported, {} placed, {} failed".format(
            status, exported, placed_total, failed))
        self._hide_progress()
        forms.alert(
            "Export & Place complete!\n\n"
            "Families exported: {}\nInstances placed: {}\nFailed: {}\n\n"
            "Output folder:\n{}".format(exported, placed_total, failed, output_folder))

    # ── Template lookup ───────────────────────────────────────────────────────

    def _find_template(self, cat_idx):
        _, template_names = CATEGORY_TEMPLATES[cat_idx]
        search_dirs = []
        try:
            tdir = self._app.FamilyTemplatePath
            if tdir and os.path.isdir(tdir):
                search_dirs.append(tdir)
        except Exception:
            pass
        ver  = self._app.VersionNumber
        base = r"C:\ProgramData\Autodesk\RVT {}".format(ver)
        for sub in ("English", "", "English-Imperial", "English_I"):
            if sub:
                search_dirs.append(os.path.join(base, "Family Templates", sub))
            else:
                search_dirs.append(os.path.join(base, "Family Templates"))
        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for tname in template_names:
                fp = os.path.join(d, tname)
                if os.path.isfile(fp):
                    return fp
        return None

    def _find_template_by_name(self, cat_name):
        idx = next((i for i, (n, _) in enumerate(CATEGORY_TEMPLATES) if n == cat_name), 0)
        return self._find_template(idx)

    # ── Parametric reference planes ──────────────────────────────────────────

    def _find_family_views(self, fam_doc):
        plan_view = elev_view = None
        for v in FilteredElementCollector(fam_doc).OfClass(View):
            try:
                if v.IsTemplate:
                    continue
                vt = v.ViewType
                if vt == ViewType.FloorPlan and plan_view is None:
                    plan_view = v
                elif vt == ViewType.Elevation and elev_view is None:
                    try:
                        if abs(v.ViewDirection.Y) > 0.99:
                            elev_view = v
                    except Exception:
                        pass
                if plan_view and elev_view:
                    break
            except Exception:
                continue
        return plan_view, elev_view

    def _create_parametric_refs(self, fam_doc, half_w, height,
                                 plan_view, elev_view, param_width_fp, param_height_fp):
        rp_left = rp_right = rp_top = None
        if plan_view is not None:
            try:
                rp_left = fam_doc.FamilyCreate.NewReferencePlane(
                    XYZ(-half_w, -3, 0), XYZ(-half_w, 3, 0), XYZ.BasisZ, plan_view)
                rp_left.Name = "Edge_Left"
                rp_right = fam_doc.FamilyCreate.NewReferencePlane(
                    XYZ(half_w, -3, 0), XYZ(half_w, 3, 0), XYZ.BasisZ, plan_view)
                rp_right.Name = "Edge_Right"
                if param_width_fp is not None:
                    ref_arr = ReferenceArray()
                    ref_arr.Append(rp_left.GetReference())
                    ref_arr.Append(rp_right.GetReference())
                    dim_line = Line.CreateBound(
                        XYZ(-half_w * 1.5, 2, 0), XYZ(half_w * 1.5, 2, 0))
                    dim = fam_doc.FamilyCreate.NewDimension(plan_view, dim_line, ref_arr)
                    if dim:
                        dim.FamilyLabel = param_width_fp
            except Exception:
                pass
        if elev_view is not None:
            try:
                rp_top = fam_doc.FamilyCreate.NewReferencePlane(
                    XYZ(-3, 0, height), XYZ(3, 0, height), XYZ.BasisY, elev_view)
                rp_top.Name = "Top"
                if param_height_fp is not None:
                    rp_level = None
                    for rp in FilteredElementCollector(fam_doc).OfClass(ReferencePlane):
                        try:
                            n = rp.Name.lower()
                            if any(k in n for k in ("level", "floor", "bottom", "ref level")):
                                rp_level = rp
                                break
                        except Exception:
                            continue
                    if rp_level:
                        ref_arr = ReferenceArray()
                        ref_arr.Append(rp_level.GetReference())
                        ref_arr.Append(rp_top.GetReference())
                        dim_line = Line.CreateBound(XYZ(0, 0, -0.1), XYZ(0, 0, height + 0.1))
                        dim = fam_doc.FamilyCreate.NewDimension(elev_view, dim_line, ref_arr)
                        if dim:
                            dim.FamilyLabel = param_height_fp
            except Exception:
                pass
        return rp_left, rp_right, rp_top

    def _lock_faces_to_planes(self, fam_doc, solid_elem,
                               plan_view, elev_view, rp_left, rp_right, rp_top):
        try:
            geom_opt = Options()
            geom_opt.ComputeReferences = True
            geom_elem = solid_elem.get_Geometry(geom_opt)
            for geom_obj in geom_elem:
                if not isinstance(geom_obj, Solid):
                    continue
                for face in geom_obj.Faces:
                    if not isinstance(face, PlanarFace):
                        continue
                    n = face.FaceNormal
                    pairs = []
                    if rp_right and plan_view and n.X > 0.99:
                        pairs.append((rp_right, plan_view))
                    elif rp_left and plan_view and n.X < -0.99:
                        pairs.append((rp_left, plan_view))
                    elif rp_top and elev_view and n.Z > 0.99:
                        pairs.append((rp_top, elev_view))
                    for rp, view in pairs:
                        try:
                            align = fam_doc.FamilyCreate.NewAlignment(
                                view, rp.GetReference(), face.Reference)
                            if align:
                                align.IsLocked = True
                        except Exception:
                            pass
        except Exception:
            pass

    def _create_window_body(self, fam_doc, sketch_plane, half_w, half_depth, height,
                            param_height_fp, param_material):
        from Autodesk.Revit.DB import (
            FamilyElementVisibility, FamilyElementVisibilityType, BuiltInParameter,
        )
        FRAME_W = max(min(half_w * 0.12, 0.1312), 0.0492)
        half_d  = max(half_depth, 0.2461)

        def rect_loop(xmin, xmax, ymin, ymax):
            arr = CurveArray()
            pts = [XYZ(xmin, ymin, 0), XYZ(xmax, ymin, 0),
                   XYZ(xmax, ymax, 0), XYZ(xmin, ymax, 0)]
            for i in range(4):
                arr.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
            return arr

        outer = rect_loop(-half_w, half_w, -half_d, half_d)
        inner = rect_loop(-(half_w - FRAME_W), (half_w - FRAME_W),
                          -(half_d - FRAME_W), (half_d - FRAME_W))
        frame_profile = CurveArrArray()
        frame_profile.Append(outer)
        frame_profile.Append(inner)
        frame_ext = fam_doc.FamilyCreate.NewExtrusion(True, frame_profile, sketch_plane, height)
        try:
            if param_height_fp:
                end_p = frame_ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                if end_p:
                    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                        end_p, param_height_fp)
            if param_material:
                mat_p = frame_ext.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                if mat_p:
                    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                        mat_p, param_material)
        except Exception:
            pass
        glass_ext = None
        try:
            iw    = half_w - FRAME_W
            GLASS = 0.0082
            glass_rect = rect_loop(-iw, iw, -GLASS, GLASS)
            glass_profile = CurveArrArray()
            glass_profile.Append(glass_rect)
            glass_height = max(height - FRAME_W * 2, FRAME_W)
            glass_ext = fam_doc.FamilyCreate.NewExtrusion(
                True, glass_profile, sketch_plane, glass_height)
            glass_ext.StartOffset = FRAME_W
            vis = FamilyElementVisibility(FamilyElementVisibilityType.Model)
            vis.IsShownInTopBottom = False
            glass_ext.SetVisibility(vis)
        except Exception:
            glass_ext = None
        return frame_ext, glass_ext

    # ── Single block export ──────────────────────────────────────────────────

    def _embed_dwg_into_family(self, fam_doc, import_inst, cx, cy):
        pass  # fallback not implemented in combined dialog

    def _export_block(self, block_item, template_path, output_folder,
                      discipline_name, category_name, load_to_project=False,
                      mode_2d_only=False):
        from Autodesk.Revit.DB import (
            BuiltInParameter, FamilyElementVisibility,
            FamilyElementVisibilityType, GraphicsStyleType,
        )
        curves = block_item._curves
        if not curves:
            return None
        fam_doc = None
        fam_doc = self._app.NewFamilyDocument(template_path)
        try:
            min_x, max_x, min_y, max_y = get_xy_bounds(curves)
            is_door   = "door"   in category_name.lower()
            is_window = "window" in category_name.lower()
            door_width = None
            if is_door:
                frame_xs, frame_ys = [], []
                for curve in curves:
                    if isinstance(curve, Arc):
                        try:
                            C  = curve.Center
                            p0 = curve.GetEndPoint(0)
                            p1 = curve.GetEndPoint(1)
                            frame_xs.append(C.X)
                            frame_ys.append(C.Y)
                            if abs(p0.Y - C.Y) < abs(p1.Y - C.Y):
                                frame_xs.append(p0.X); frame_ys.append(p0.Y)
                            else:
                                frame_xs.append(p1.X); frame_ys.append(p1.Y)
                        except Exception:
                            pass
                if frame_xs:
                    cx = (min(frame_xs) + max(frame_xs)) / 2.0
                    cy = (min(frame_ys) + max(frame_ys)) / 2.0
                    calc_w = max(frame_xs) - min(frame_xs)
                    if calc_w > 0.01:
                        door_width = calc_w
                else:
                    cx = (min_x + max_x) / 2.0
                    cy = (min_y + max_y) / 2.0
                    door_width = max_x - min_x
            else:
                cx = (min_x + max_x) / 2.0
                cy = (min_y + max_y) / 2.0
            half_w = max((max_x - min_x) / 2.0, 0.01)
            half_h = max((max_y - min_y) / 2.0, 0.01)

            t = Transaction(fam_doc, 'Create Block Geometry')
            start_transaction(t)
            try:
                sketch_plane = None
                for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                    try:
                        if abs(sp.GetPlane().Normal.Z - 1.0) < 0.001:
                            sketch_plane = sp
                            break
                    except Exception:
                        pass
                if not sketch_plane:
                    sketch_plane = SketchPlane.Create(
                        fam_doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))

                if mode_2d_only:
                    _zvals = []
                    for _c in curves:
                        try:
                            _zvals.append(_c.GetEndPoint(0).Z)
                            _zvals.append(_c.GetEndPoint(1).Z)
                        except Exception:
                            pass
                    if _zvals:
                        _zvals.sort()
                        _mid = len(_zvals) // 2
                        cz = (_zvals[_mid] if len(_zvals) % 2 == 1
                              else (_zvals[_mid - 1] + _zvals[_mid]) / 2.0)
                    else:
                        cz = 0.0
                    translator = Transform.CreateTranslation(XYZ(-cx, -cy, -cz))

                    def _seg(pa, pb):
                        paf = XYZ(pa.X, pa.Y, 0.0)
                        pbf = XYZ(pb.X, pb.Y, 0.0)
                        if paf.DistanceTo(pbf) < 1e-4:
                            return False
                        try:
                            fam_doc.FamilyCreate.NewModelCurve(
                                Line.CreateBound(paf, pbf), sketch_plane)
                            return True
                        except Exception:
                            return False

                    ok_2d = fail_2d = 0
                    for curve in curves:
                        try:
                            new_c = curve.CreateTransformed(translator)
                            if isinstance(new_c, Line):
                                p0 = new_c.GetEndPoint(0)
                                p1 = new_c.GetEndPoint(1)
                                if _seg(p0, p1):
                                    ok_2d += 1
                                else:
                                    fail_2d += 1
                            else:
                                written = False
                                pts = new_c.Tessellate()
                                for i in range(len(pts) - 1):
                                    if _seg(pts[i], pts[i + 1]):
                                        written = True
                                if not new_c.IsBound:
                                    if _seg(pts[-1], pts[0]):
                                        written = True
                                if written:
                                    ok_2d += 1
                                else:
                                    fail_2d += 1
                        except Exception:
                            fail_2d += 1
                    logger.info("2D '{}': {} ok / {} skipped".format(
                        block_item.BlockName, ok_2d, fail_2d))
                else:
                    THICKNESS      = 0.1312
                    HEIGHT         = 7.2178
                    WINDOW_HEIGHT  = 4.9213
                    extrusion_depth = HEIGHT if is_door else (WINDOW_HEIGHT if is_window else 1.0)

                    swing_gs = frame_gs = None
                    if is_door:
                        try:
                            fam_cat = fam_doc.OwnerFamily.FamilyCategory
                            def get_or_create_subcat(name):
                                if fam_cat.SubCategories.Contains(name):
                                    return fam_cat.SubCategories.get_Item(name)
                                return fam_doc.Settings.Categories.NewSubcategory(fam_cat, name)
                            swing_subcat = get_or_create_subcat("Plan Swing")
                            frame_subcat = get_or_create_subcat("Frame/Mullion")
                            if swing_subcat:
                                swing_gs = swing_subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
                            if frame_subcat:
                                frame_gs = frame_subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
                        except Exception:
                            pass

                    param_height_fp = param_width_fp = param_material = None
                    try:
                        fam_mgr = fam_doc.FamilyManager
                        for param in fam_mgr.Parameters:
                            pname = param.Definition.Name.lower()
                            if pname in ("height", "chieu cao"):
                                try: fam_mgr.Set(param, extrusion_depth)
                                except Exception: pass
                                param_height_fp = param
                            elif pname in ("width", "chieu rong"):
                                if door_width:
                                    try: fam_mgr.Set(param, door_width)
                                    except Exception: pass
                                param_width_fp = param
                            elif pname in ("depth", "chieu sau", "length", "chieu dai"):
                                if not is_door and half_h * 2.0 > 0.01:
                                    try: fam_mgr.Set(param, half_h * 2.0)
                                    except Exception: pass
                            elif pname in ("material", "vat lieu"):
                                param_material = param
                    except Exception:
                        pass

                    ext_box = None
                    if is_window:
                        window_frame_ext, _ = self._create_window_body(
                            fam_doc, sketch_plane, half_w, half_h,
                            extrusion_depth, param_height_fp, param_material)
                        ext_box = window_frame_ext
                    elif not is_door:
                        c1 = XYZ(-half_w, -half_h, 0.0)
                        c2 = XYZ( half_w, -half_h, 0.0)
                        c3 = XYZ( half_w,  half_h, 0.0)
                        c4 = XYZ(-half_w,  half_h, 0.0)
                        rect = CurveArray()
                        rect.Append(Line.CreateBound(c1, c2))
                        rect.Append(Line.CreateBound(c2, c3))
                        rect.Append(Line.CreateBound(c3, c4))
                        rect.Append(Line.CreateBound(c4, c1))
                        profile = CurveArrArray()
                        profile.Append(rect)
                        ext_box = fam_doc.FamilyCreate.NewExtrusion(
                            True, profile, sketch_plane, extrusion_depth)
                        try:
                            if param_height_fp:
                                end_p = ext_box.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                if end_p:
                                    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                                        end_p, param_height_fp)
                            if param_material:
                                mat_p = ext_box.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                                if mat_p:
                                    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                                        mat_p, param_material)
                        except Exception:
                            pass
                        top_sp = SketchPlane.Create(
                            fam_doc,
                            Plane.CreateByNormalAndOrigin(
                                XYZ.BasisZ, XYZ(0.0, 0.0, extrusion_depth)))
                        ok_3d = fail_3d = 0
                        for curve in curves:
                            projected = _dwg_project_curve(curve, cx, cy, extrusion_depth)
                            if projected is None:
                                fail_3d += 1
                                continue
                            try:
                                fam_doc.FamilyCreate.NewModelCurve(projected, top_sp)
                                ok_3d += 1
                            except Exception:
                                try:
                                    pts = curve.Tessellate()
                                    for i in range(len(pts) - 1):
                                        pa = XYZ(pts[i].X - cx,   pts[i].Y - cy,   extrusion_depth)
                                        pb = XYZ(pts[i+1].X - cx, pts[i+1].Y - cy, extrusion_depth)
                                        if pa.DistanceTo(pb) > 1e-4:
                                            fam_doc.FamilyCreate.NewModelCurve(
                                                Line.CreateBound(pa, pb), top_sp)
                                            ok_3d += 1
                                except Exception:
                                    pass
                                fail_3d += 1
                        threshold = max(5, int(len(curves) * MIN_CURVE_RATIO))
                        if ok_3d < threshold:
                            self._embed_dwg_into_family(
                                fam_doc, block_item._import_inst, cx, cy)

                    panel_ext = None
                    for curve in curves:
                        if not is_door:
                            break
                        try:
                            translator = Transform.CreateTranslation(XYZ(-cx, -cy, 0.0))
                            new_c = curve.CreateTransformed(translator)
                            if isinstance(curve, Line):
                                sym_line = fam_doc.FamilyCreate.NewSymbolicCurve(new_c, sketch_plane)
                                if frame_gs:
                                    sym_line.Subcategory = frame_gs
                            elif isinstance(curve, Arc):
                                sym_arc = fam_doc.FamilyCreate.NewSymbolicCurve(new_c, sketch_plane)
                                if swing_gs:
                                    sym_arc.Subcategory = swing_gs
                                ctr  = curve.Center
                                nc   = ctr + XYZ(-cx, -cy, 0.0)
                                p0_orig = curve.GetEndPoint(0)
                                p1_orig = curve.GetEndPoint(1)
                                p_closed = (p0_orig if abs(p0_orig.Y - ctr.Y) < abs(p1_orig.Y - ctr.Y)
                                            else p1_orig)
                                np_closed = p_closed + XYZ(-cx, -cy, 0.0)
                                v_dir   = (np_closed - nc).Normalize()
                                v_ortho = XYZ(-v_dir.Y, v_dir.X, 0.0)
                                half_t  = THICKNESS / 2.0
                                pt1 = nc + v_ortho * half_t
                                pt2 = nc - v_ortho * half_t
                                pt3 = pt2 + v_dir * curve.Radius
                                pt4 = pt1 + v_dir * curve.Radius
                                p_rect = CurveArray()
                                p_rect.Append(Line.CreateBound(pt1, pt2))
                                p_rect.Append(Line.CreateBound(pt2, pt3))
                                p_rect.Append(Line.CreateBound(pt3, pt4))
                                p_rect.Append(Line.CreateBound(pt4, pt1))
                                p_profile = CurveArrArray()
                                p_profile.Append(p_rect)
                                panel_ext = fam_doc.FamilyCreate.NewExtrusion(
                                    True, p_profile, sketch_plane, HEIGHT)
                                try:
                                    vis = FamilyElementVisibility(
                                        FamilyElementVisibilityType.Model)
                                    vis.IsShownInTopBottom = False
                                    panel_ext.SetVisibility(vis)
                                    if param_height_fp:
                                        end_p = panel_ext.get_Parameter(
                                            BuiltInParameter.EXTRUSION_END_PARAM)
                                        if end_p:
                                            fam_doc.FamilyManager\
                                                .AssociateElementParameterToFamilyParameter(
                                                    end_p, param_height_fp)
                                    if param_material:
                                        mat_p = panel_ext.get_Parameter(
                                            BuiltInParameter.MATERIAL_ID_PARAM)
                                        if mat_p:
                                            fam_doc.FamilyManager\
                                                .AssociateElementParameterToFamilyParameter(
                                                    mat_p, param_material)
                                except Exception:
                                    pass
                        except Exception:
                            pass

                    if is_door or is_window:
                        fam_doc.Regenerate()
                        plan_view, elev_view = self._find_family_views(fam_doc)
                        rp_left, rp_right, rp_top = self._create_parametric_refs(
                            fam_doc,
                            half_w if not is_door else (door_width / 2.0 if door_width else half_w),
                            HEIGHT if is_door else extrusion_depth,
                            plan_view, elev_view, param_width_fp, param_height_fp)
                        fam_doc.Regenerate()
                        target_solid = panel_ext if is_door else ext_box
                        if target_solid and (rp_left or rp_right or rp_top):
                            self._lock_faces_to_planes(
                                fam_doc, target_solid,
                                plan_view, elev_view, rp_left, rp_right, rp_top)

                t.Commit()
            except Exception:
                try:
                    t.RollBack()
                except Exception:
                    pass
                raise

            safe_cad_name = block_item.BlockName.strip() or "Family"
            w_str = getattr(block_item, 'WidthMM', '-')
            d_str = getattr(block_item, 'DepthMM', '-')
            dim_suffix = "_{}x{}".format(w_str, d_str) if (w_str != '-' and d_str != '-') else ""
            base_name = "T3Lab_{}_{}{}".format(
                category_name.replace(" ", "_"),
                safe_cad_name.replace(" ", "_"),
                dim_suffix)
            base_name = re.sub(r'[\\/*?:"<>|]', "", base_name)
            save_path = os.path.join(output_folder, "{}.rfa".format(base_name))
            ctr = 1
            while os.path.exists(save_path):
                save_path = os.path.join(output_folder, "{}_{}.rfa".format(base_name, ctr))
                ctr += 1
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
            logger.info("Exported: {}".format(save_path))
            if load_to_project:
                try:
                    t_load = Transaction(self._doc, 'Load Family - {}'.format(safe_cad_name))
                    start_transaction(t_load)
                    try:
                        self._doc.LoadFamily(save_path)
                        t_load.Commit()
                    except Exception:
                        try: t_load.RollBack()
                        except Exception: pass
                        logger.warning("Could not load: {}".format(save_path))
                except Exception:
                    pass
            return save_path
        finally:
            if fam_doc is not None:
                try:
                    fam_doc.Close(False)
                except Exception:
                    pass

    # ── Preset generators ────────────────────────────────────────────────────

    def presets_clicked(self, sender, e):
        output_folder = self.output_path.Text
        if not output_folder or not os.path.isdir(output_folder):
            forms.alert("Please select a valid output folder (Browse...) first.")
            return
        cat_name = self.category_combo.SelectedItem
        if cat_name is None:
            forms.alert("Please select a category first.")
            return
        if cat_name not in CATEGORY_PRESETS:
            forms.alert("No presets for: {}".format(cat_name))
            return
        mode, preset_list = CATEGORY_PRESETS[cat_name]
        labels = [p[0] for p in preset_list]
        selected_labels = forms.SelectFromList.show(
            labels,
            title="T3Lab - {} Presets".format(cat_name),
            multiselect=True, button_name="Generate")
        if not selected_labels:
            return
        selected = [p for p in preset_list if p[0] in selected_labels]
        load_to_project = (self.chk_load_to_project.IsChecked == True)
        skip_existing   = (self.batch_skip_existing.IsChecked == True)
        cat_idx = next((i for i, (n, _) in enumerate(CATEGORY_TEMPLATES) if n == cat_name), 0)
        template_path = self._find_template(cat_idx)
        if not template_path:
            forms.alert("Template (.rft) for '{}' not found.".format(cat_name))
            return
        ok_count = fail_count = skipped = 0
        self._cancel_requested = False
        self._pause_requested  = False
        self._update_progress(0, len(selected))
        for i, preset in enumerate(selected):
            if self._cancel_requested:
                break
            self._update_status("[{}/{}] Generating: {}".format(i + 1, len(selected), preset[0]))
            self._update_progress(i, len(selected))
            if skip_existing:
                safe  = re.sub(r'[/*?:"<>|]', "_", preset[0])
                prefix = re.sub(r'\s+', '_', cat_name)
                check = os.path.join(output_folder, "T3Lab_{}_{}.rfa".format(prefix, safe))
                if os.path.exists(check):
                    skipped += 1
                    self._update_progress(i + 1, len(selected))
                    continue
            try:
                if mode == "door":
                    success = self._generate_door_from_preset(
                        preset, template_path, output_folder, load_to_project)
                elif mode == "window":
                    success = self._generate_window_from_preset(
                        preset, template_path, output_folder, load_to_project)
                else:
                    success = self._generate_generic_from_preset(
                        preset, template_path, output_folder, cat_name, load_to_project)
                if success:
                    ok_count += 1
                else:
                    fail_count += 1
            except Exception:
                logger.error("Preset error: {}\n{}".format(preset[0], traceback.format_exc()))
                fail_count += 1
            self._update_progress(i + 1, len(selected))
        was_cancelled = self._cancel_requested
        status = "Stopped" if was_cancelled else "Done"
        self._update_status("{}: {} ok, {} skipped, {} failed".format(status, ok_count, skipped, fail_count))
        self._hide_progress()
        forms.alert("{} Preset Export\n\nGenerated: {}\nSkipped: {}\nFailed: {}\n\nFolder: {}".format(
            cat_name, ok_count, skipped, fail_count, output_folder))

    def _generate_window_from_preset(self, preset, template_path, output_folder, load_to_project):
        label, width_mm, height_mm = preset
        half_w = (width_mm / 2.0) * SCL
        height = height_mm * SCL
        half_d = 0.3937
        fam_doc = self._app.NewFamilyDocument(template_path)
        t = Transaction(fam_doc, "T3Lab Window - " + label)
        start_transaction(t)
        try:
            sketch_plane = None
            for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                try:
                    if abs(sp.GetPlane().Normal.Z - 1.0) < 0.001:
                        sketch_plane = sp
                        break
                except Exception:
                    pass
            if not sketch_plane:
                sketch_plane = SketchPlane.Create(
                    fam_doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))
            param_width_fp = param_height_fp = param_material = None
            try:
                fm = fam_doc.FamilyManager
                for p in fm.Parameters:
                    pn = p.Definition.Name.lower()
                    if pn == "height":
                        try: fm.Set(p, height)
                        except Exception: pass
                        param_height_fp = p
                    elif pn == "width":
                        try: fm.Set(p, half_w * 2.0)
                        except Exception: pass
                        param_width_fp = p
                    elif "material" in pn:
                        param_material = p
            except Exception:
                pass
            self._create_window_body(fam_doc, sketch_plane, half_w, half_d, height,
                                     param_height_fp, param_material)
            fam_doc.Regenerate()
            plan_view, elev_view = self._find_family_views(fam_doc)
            self._create_parametric_refs(fam_doc, half_w, height, plan_view, elev_view,
                                         param_width_fp, param_height_fp)
            t.Commit()
        except Exception:
            try: t.RollBack()
            except Exception: pass
            fam_doc.Close(False)
            raise
        safe = re.sub(r'[/*?:"<>|]', "_", label)
        save_path = os.path.join(output_folder, "T3Lab_Window_{}.rfa".format(safe))
        ctr = 1
        while os.path.exists(save_path):
            save_path = os.path.join(output_folder, "T3Lab_Window_{}_{}.rfa".format(safe, ctr))
            ctr += 1
        try:
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
        finally:
            fam_doc.Close(False)
        logger.info("Saved: " + save_path)
        if load_to_project:
            try:
                t2 = Transaction(self._doc, "Load " + label)
                start_transaction(t2)
                try: self._doc.LoadFamily(save_path); t2.Commit()
                except Exception:
                    try: t2.RollBack()
                    except Exception: pass
            except Exception:
                pass
        return True

    def _generate_generic_from_preset(self, preset, template_path, output_folder,
                                       category_name, load_to_project):
        label, w_mm, d_mm, h_mm = preset
        w = w_mm * SCL
        d = d_mm * SCL
        h = h_mm * SCL
        fam_doc = self._app.NewFamilyDocument(template_path)
        t = Transaction(fam_doc, "T3Lab {} - {}".format(category_name, label))
        start_transaction(t)
        try:
            sketch_plane = None
            for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                try:
                    if abs(sp.GetPlane().Normal.Z - 1.0) < 0.001:
                        sketch_plane = sp
                        break
                except Exception:
                    pass
            if not sketch_plane:
                sketch_plane = SketchPlane.Create(
                    fam_doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))
            try:
                fm = fam_doc.FamilyManager
                for p in fm.Parameters:
                    pn = p.Definition.Name.lower()
                    if pn == "width":
                        try: fm.Set(p, w)
                        except Exception: pass
                    elif pn in ("depth", "length"):
                        try: fm.Set(p, d)
                        except Exception: pass
                    elif pn == "height":
                        try: fm.Set(p, h)
                        except Exception: pass
            except Exception:
                pass
            half_w = w / 2.0
            half_d = d / 2.0
            arr = CurveArray()
            pts = [XYZ(-half_w, -half_d, 0), XYZ(half_w, -half_d, 0),
                   XYZ(half_w, half_d, 0), XYZ(-half_w, half_d, 0)]
            for i in range(4):
                arr.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
            prof = CurveArrArray()
            prof.Append(arr)
            fam_doc.FamilyCreate.NewExtrusion(True, prof, sketch_plane, h)
            t.Commit()
        except Exception:
            try: t.RollBack()
            except Exception: pass
            fam_doc.Close(False)
            raise
        cat_prefix = re.sub(r'\s+', '_', category_name)
        safe = re.sub(r'[/*?:"<>|]', "_", label)
        save_path = os.path.join(output_folder, "T3Lab_{}_{}.rfa".format(cat_prefix, safe))
        ctr = 1
        while os.path.exists(save_path):
            save_path = os.path.join(
                output_folder, "T3Lab_{}_{}_{}.rfa".format(cat_prefix, safe, ctr))
            ctr += 1
        try:
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
        finally:
            fam_doc.Close(False)
        logger.info("Saved: " + save_path)
        if load_to_project:
            try:
                t2 = Transaction(self._doc, "Load " + label)
                start_transaction(t2)
                try: self._doc.LoadFamily(save_path); t2.Commit()
                except Exception:
                    try: t2.RollBack()
                    except Exception: pass
            except Exception:
                pass
        return True

    def _generate_door_from_preset(self, preset, template_path, output_folder, load_to_project):
        from Autodesk.Revit.DB import BuiltInParameter, GraphicsStyleType
        label, width_mm, height_mm, frame_w_mm, proj_ext_mm, proj_int_mm, leaf_t_mm, door_count = preset
        half_w   = (width_mm / 2.0) * SCL
        h        = height_mm * SCL
        fw       = frame_w_mm * SCL
        fpe      = proj_ext_mm * SCL
        fpi      = proj_int_mm * SCL
        dt       = leaf_t_mm * SCL
        half_fw  = half_w + fw
        total_fh = h + fw
        fam_doc = self._app.NewFamilyDocument(template_path)
        t = Transaction(fam_doc, "T3Lab Door - " + label)
        start_transaction(t)
        try:
            plan_sp = elev_sp = None
            for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                n, org = sp.GetPlane().Normal, sp.GetPlane().Origin
                if abs(n.Z - 1.0) < 0.001 and abs(org.X) < 0.01 and abs(org.Y) < 0.01 and plan_sp is None:
                    plan_sp = sp
                if abs(n.Y + 1.0) < 0.001 and abs(org.X) < 0.01 and abs(org.Y) < 0.01 and elev_sp is None:
                    elev_sp = sp
            if plan_sp is None:
                plan_sp = SketchPlane.Create(
                    fam_doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))
            if elev_sp is None:
                elev_sp = SketchPlane.Create(
                    fam_doc, Plane.CreateByNormalAndOrigin(XYZ(0.0, -1.0, 0.0), XYZ.Zero))
            param_width_fp = param_height_fp = None
            try:
                fm = fam_doc.FamilyManager
                for p in fm.Parameters:
                    pn = p.Definition.Name.lower()
                    if pn == "height":
                        try: fm.Set(p, h)
                        except Exception: pass
                        param_height_fp = p
                    elif pn == "width":
                        try: fm.Set(p, half_w * 2.0)
                        except Exception: pass
                        param_width_fp = p
                    elif pn == "frame width":
                        try: fm.Set(p, fw)
                        except Exception: pass
                    elif pn in ("frame projection ext.", "frame projection ext"):
                        try: fm.Set(p, fpe)
                        except Exception: pass
                    elif pn in ("frame projection int.", "frame projection int"):
                        try: fm.Set(p, fpi)
                        except Exception: pass
            except Exception:
                pass
            frame_gs = leaf_gs = None
            try:
                fam_cat = fam_doc.OwnerFamily.FamilyCategory
                def _sc(name):
                    return (fam_cat.SubCategories.get_Item(name)
                            if fam_cat.SubCategories.Contains(name)
                            else fam_doc.Settings.Categories.NewSubcategory(fam_cat, name))
                sc_f = _sc("Frame/Mullion")
                sc_p = _sc("Panel")
                if sc_f: frame_gs = sc_f.GetGraphicsStyle(GraphicsStyleType.Projection)
                if sc_p: leaf_gs  = sc_p.GetGraphicsStyle(GraphicsStyleType.Projection)
            except Exception:
                pass

            def _extrude_xz(x0, x1, z0, z1, depth, gs=None):
                arr = CurveArray()
                arr.Append(Line.CreateBound(XYZ(x0, 0, z0), XYZ(x1, 0, z0)))
                arr.Append(Line.CreateBound(XYZ(x1, 0, z0), XYZ(x1, 0, z1)))
                arr.Append(Line.CreateBound(XYZ(x1, 0, z1), XYZ(x0, 0, z1)))
                arr.Append(Line.CreateBound(XYZ(x0, 0, z1), XYZ(x0, 0, z0)))
                prof = CurveArrArray()
                prof.Append(arr)
                ext = fam_doc.FamilyCreate.NewExtrusion(True, prof, elev_sp, depth)
                if gs:
                    try: ext.Subcategory = gs
                    except Exception: pass
                return ext

            def _extrude_xy_leaf(x0, x1, y0, y1):
                arr = CurveArray()
                pts = [XYZ(x0, y0, 0), XYZ(x1, y0, 0), XYZ(x1, y1, 0), XYZ(x0, y1, 0)]
                for i in range(4): arr.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
                prof = CurveArrArray()
                prof.Append(arr)
                ext = fam_doc.FamilyCreate.NewExtrusion(True, prof, plan_sp, h)
                if leaf_gs:
                    try: ext.Subcategory = leaf_gs
                    except Exception: pass
                try:
                    if param_height_fp:
                        ep = ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                        if ep: fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                            ep, param_height_fp)
                except Exception:
                    pass
                return ext

            frame_pieces = [
                (-half_fw, -half_w, 0.0, total_fh),
                ( half_w,  half_fw, 0.0, total_fh),
                (-half_fw,  half_fw, h, total_fh),
            ]
            for x0, x1, z0, z1 in frame_pieces:
                _extrude_xz(x0, x1, z0, z1,  fpe, frame_gs)
                _extrude_xz(x0, x1, z0, z1, -fpi, frame_gs)
            GAP = 0.00328
            if door_count == 1:
                _extrude_xy_leaf(-half_w, half_w, 0.0, dt)
            else:
                _extrude_xy_leaf(-half_w, -GAP / 2.0, 0.0, dt)
                _extrude_xy_leaf( GAP / 2.0, half_w,  0.0, dt)
            fam_doc.Regenerate()
            plan_view, elev_view = self._find_family_views(fam_doc)
            self._create_parametric_refs(fam_doc, half_w, h, plan_view, elev_view,
                                         param_width_fp, param_height_fp)
            t.Commit()
        except Exception:
            try: t.RollBack()
            except Exception: pass
            fam_doc.Close(False)
            raise
        safe = re.sub(r'[/*?:"<>|]', "_", label)
        save_path = os.path.join(output_folder, "T3Lab_Door_{}.rfa".format(safe))
        ctr = 1
        while os.path.exists(save_path):
            save_path = os.path.join(output_folder, "T3Lab_Door_{}_{}.rfa".format(safe, ctr))
            ctr += 1
        try:
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
        finally:
            fam_doc.Close(False)
        logger.info("Saved: " + save_path)
        if load_to_project:
            try:
                t2 = Transaction(self._doc, "Load " + label)
                start_transaction(t2)
                try: self._doc.LoadFamily(save_path); t2.Commit()
                except Exception:
                    try: t2.RollBack()
                    except Exception: pass
            except Exception:
                pass
        return True

    # ── Place & batch load ────────────────────────────────────────────────────

    def _place_family_instances(self, rfa_path, block_item, level):
        from Autodesk.Revit.DB import ElementTransformUtils, Family, Line as DBLine
        from Autodesk.Revit.DB.Structure import StructuralType
        if not block_item._placements:
            return 0
        family = None
        try:
            loaded_ref = clr.Reference[Family]()
            if self._doc.LoadFamily(rfa_path, loaded_ref):
                family = loaded_ref.Value
        except Exception:
            pass
        if not family:
            stem = os.path.splitext(os.path.basename(rfa_path))[0]
            for f in FilteredElementCollector(self._doc).OfClass(Family):
                if f.Name == stem:
                    family = f
                    break
        if not family:
            logger.warning("Could not load family: {}".format(rfa_path))
            return 0
        symbol = None
        for sid in family.GetFamilySymbolIds():
            symbol = self._doc.GetElement(sid)
            break
        if not symbol:
            return 0
        if not symbol.IsActive:
            symbol.Activate()
            self._doc.Regenerate()
        placed = 0
        for (centroid, angle) in block_item._placements:
            try:
                z  = level.Elevation if level else centroid.Z
                pt = XYZ(centroid.X, centroid.Y, z)
                inst = self._doc.Create.NewFamilyInstance(
                    pt, symbol, level, StructuralType.NonStructural)
                if inst and abs(angle) > 0.001:
                    axis = DBLine.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 1.0))
                    ElementTransformUtils.RotateElement(self._doc, inst.Id, axis, angle)
                placed += 1
            except Exception:
                logger.warning("Could not place '{}': {}".format(
                    block_item.BlockName, traceback.format_exc()))
        return placed

    def _batch_load_families(self, save_paths):
        loaded = 0
        total  = len(save_paths)
        for i, path in enumerate(save_paths):
            self._update_status("Loading [{}/{}]: {}".format(
                i + 1, total, os.path.basename(path)))
            try:
                t = Transaction(self._doc, "T3Lab - Load {}".format(
                    os.path.splitext(os.path.basename(path))[0]))
                start_transaction(t)
                try:
                    self._doc.LoadFamily(path)
                    t.Commit()
                    loaded += 1
                except Exception:
                    try: t.RollBack()
                    except Exception: pass
                    logger.warning("Could not load: {}".format(path))
            except Exception:
                logger.warning("Transaction failed: {}".format(path))
        return loaded

    # ── JSON mode ────────────────────────────────────────────────────────────

    def copy_prompt_clicked(self, sender, e):
        try:
            text = self.json_tb.Text
            if text:
                Clipboard.SetText(text)
                self.lbl_status.Text = "JSON copied to clipboard."
        except Exception as ex:
            forms.alert("Could not copy: {}".format(ex))

    def cancel_clicked(self, sender, e):
        self.Close()

    def create_clicked(self, sender, e):
        raw = self.json_tb.Text
        if not raw or raw.strip() in ("", "Paste your JSON schema here..."):
            forms.alert("Please paste a valid JSON schema first.")
            return
        try:
            schema = json.loads(raw)
        except ValueError as ex:
            forms.alert("Invalid JSON:\n\n{}".format(ex), title="JSON Error")
            return
        self.lbl_status.Text = "Creating family..."
        if self._doc.IsFamilyDocument:
            try:
                t = Transaction(self._doc, "T3Lab - JSON to Family")
                start_transaction(t)
                try:
                    self._generate_json_family(self._doc, schema)
                    t.Commit()
                except Exception:
                    try: t.RollBack()
                    except Exception: pass
                    raise
                self.lbl_status.Text = "Family updated successfully."
                forms.alert("Family generated successfully!", title="Success")
            except Exception as ex:
                self.lbl_status.Text = "Error."
                forms.alert("Error:\n{}".format(ex), title="Error")
        else:
            cat_name = schema.get("family_category", "Generic Model")
            template_path = self._find_template_by_name(cat_name)
            if not template_path:
                forms.alert("Template for '{}' not found.".format(cat_name))
                self.lbl_status.Text = "Template not found."
                return
            output_folder = self.output_path.Text
            if not output_folder or not os.path.isdir(output_folder):
                output_folder = forms.pick_folder()
            if not output_folder:
                self.lbl_status.Text = "No output folder selected."
                return
            fam_doc = self._app.NewFamilyDocument(template_path)
            try:
                t = Transaction(fam_doc, "T3Lab - JSON to Family")
                start_transaction(t)
                try:
                    self._generate_json_family(fam_doc, schema)
                    t.Commit()
                except Exception:
                    try: t.RollBack()
                    except Exception: pass
                    raise
                family_name = schema.get("family_name", "T3Lab_JSONFamily")
                family_name = re.sub(r'[\\/*?:"<>|]', "_", family_name)
                save_path = os.path.join(output_folder, "{}.rfa".format(family_name))
                opts = SaveAsOptions()
                opts.OverwriteExistingFile = True
                fam_doc.SaveAs(save_path, opts)
                self.lbl_status.Text = "Saved: {}".format(os.path.basename(save_path))
                forms.alert("Family saved:\n{}".format(save_path), title="Success")
            except Exception as ex:
                self.lbl_status.Text = "Error."
                forms.alert("Error:\n{}".format(ex), title="Error")
            finally:
                try: fam_doc.Close(False)
                except Exception: pass

    def _json_curve(self, seg, offset_z=0.0):
        """Parse a JSON curve segment into a Revit Curve."""
        try:
            seg_type = seg.get("type", "Line")
            if seg_type == "Line":
                p0 = seg["start"]
                p1 = seg["end"]
                return Line.CreateBound(
                    XYZ(p0[0] * SCL, p0[1] * SCL, p0[2] * SCL + offset_z),
                    XYZ(p1[0] * SCL, p1[1] * SCL, p1[2] * SCL + offset_z))
            elif seg_type == "Arc":
                c = seg["center"]
                r = seg["radius"] * SCL
                a0 = seg.get("start_angle", 0.0)
                a1 = seg.get("end_angle", 6.283185307)
                nc = XYZ(c[0] * SCL, c[1] * SCL, c[2] * SCL + offset_z)
                return Arc.Create(nc, r, a0, a1, XYZ.BasisX, XYZ.BasisY)
        except Exception:
            pass
        return None

    def _generate_json_family(self, fam_doc, schema):
        """Apply a JSON schema to a (possibly new) family document."""
        fm = fam_doc.FamilyManager
        param_dict = {}
        for param_data in schema.get("parameters", []):
            name = param_data.get("name", "")
            for p in fm.Parameters:
                if p.Definition.Name == name:
                    param_dict[name] = p
                    try:
                        val = param_data.get("value")
                        if val is not None:
                            fm.Set(p, float(val) * SCL)
                    except Exception:
                        pass
                    break

        for geom_data in schema.get("geometry", []):
            try:
                is_solid  = geom_data.get("is_solid", True)
                geom_type = geom_data.get("type", "Extrusion")

                if "sketch_plane_x" in geom_data:
                    base_plane = Plane.CreateByNormalAndOrigin(
                        XYZ.BasisX, XYZ(geom_data["sketch_plane_x"] * SCL, 0, 0))
                elif "sketch_plane_y" in geom_data:
                    base_plane = Plane.CreateByNormalAndOrigin(
                        XYZ.BasisY, XYZ(0, geom_data["sketch_plane_y"] * SCL, 0))
                else:
                    z = geom_data.get("sketch_plane_z", 0.0) * SCL
                    base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, z))

                sketch_plane = SketchPlane.Create(fam_doc, base_plane)

                if geom_type == "Extrusion":
                    profile = CurveArrArray()
                    loop = CurveArray()
                    for seg in geom_data.get("profile", []):
                        c = self._json_curve(seg)
                        if c:
                            loop.Append(c)
                    if loop.Size > 0:
                        profile.Append(loop)
                        for inner in geom_data.get("inner_loops", []):
                            inner_loop = CurveArray()
                            for seg in inner:
                                c = self._json_curve(seg)
                                if c:
                                    inner_loop.Append(c)
                            if inner_loop.Size > 0:
                                profile.Append(inner_loop)
                        ext = fam_doc.FamilyCreate.NewExtrusion(
                            is_solid, profile, sketch_plane,
                            geom_data.get("extrusion_end", 1.0) * SCL)
                        ext.StartOffset = geom_data.get("extrusion_start", 0.0) * SCL
            except Exception:
                logger.warning("JSON geometry skip: {}".format(traceback.format_exc()))

    # ── Batch mode ───────────────────────────────────────────────────────────


# ==============================================================================
# ENTRY POINTS
# ==============================================================================

def show_family_creator(revit_doc, revit_app, initial_mode='cad'):
    FamilyCreatorDialog(revit_doc, revit_app, initial_mode).ShowDialog()

def show_family_creator_cad(revit_doc, revit_app):
    show_family_creator(revit_doc, revit_app, 'cad')

def show_family_creator_json(revit_doc, revit_app):
    show_family_creator(revit_doc, revit_app, 'json')

