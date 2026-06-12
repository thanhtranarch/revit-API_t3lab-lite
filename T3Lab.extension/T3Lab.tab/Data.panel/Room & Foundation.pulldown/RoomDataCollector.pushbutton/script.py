# -*- coding: utf-8 -*-
"""
Room Data Collector v1.0 - DQT
Collects parameter values from elements inside Rooms/Areas/Spaces
and aggregates them into spatial element parameters.

Reverse workflow of Contains Manager:
- Contains Manager: Room info -> Element parameters
- Room Data Collector: Element info -> Room parameters

Supports aggregation: Sum, Count, Average, Min, Max, First, Last, List (comma), Unique List

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Room Data\nCollector"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Collect element data inside Rooms/Areas/Spaces and aggregate into spatial parameters."

import clr
clr.AddReference('System')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Collections.Generic import List
from System.Windows import Window, WindowStartupLocation, Thickness, CornerRadius
from System.Windows import HorizontalAlignment, VerticalAlignment, ResizeMode, FontWeights
from System.Windows import Visibility, TextWrapping
from System.Windows.Controls import (StackPanel, DockPanel, Border, TextBlock, TextBox,
    Button, ComboBox, ComboBoxItem, CheckBox, Orientation, Dock,
    ScrollViewer, ScrollBarVisibility, SelectionMode, ListBox, ListBoxItem,
    RadioButton, GroupBox)
from System.Windows.Media import BrushConverter
import System.Windows
import System

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, StorageType, XYZ, SpatialElementBoundaryOptions, ElementId,
    SpatialElementBoundaryLocation, CurveLoop, AreaVolumeSettings)
from Autodesk.Revit.UI import TaskDialog

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# ================================================================
# CONSTANTS & HELPERS
# ================================================================
_conv = BrushConverter()
def brush(c): return _conv.ConvertFromString(c)

PRIMARY = "#0F172A"
SECONDARY = "#F8FAFC"
WHITE = "#FFFFFF"
BORDER = "#E0E0E0"
TEXT_DARK = "#0F172A"
TEXT_GRAY = "#475569"
TEXT_MUTED = "#94A3B8"
SUCCESS = "#10B981"
WARNING_CLR = "#F59E0B"
ERROR_CLR = "#EF4444"
ACCENT = "#C89650"

ROOMS = "Rooms"
AREAS = "Areas"
SPACES = "Spaces"

AGG_COUNT = "Count"
AGG_SUM = "Sum"
AGG_AVERAGE = "Average"
AGG_MIN = "Min"
AGG_MAX = "Max"
AGG_FIRST = "First"
AGG_LAST = "Last"
AGG_LIST = "List (Comma)"
AGG_UNIQUE = "Unique List"

ALL_AGG = [AGG_COUNT, AGG_SUM, AGG_AVERAGE, AGG_MIN, AGG_MAX,
           AGG_FIRST, AGG_LAST, AGG_LIST, AGG_UNIQUE]

# ================================================================
# SPATIAL GEOMETRY HELPERS (reused from Contains Manager)
# ================================================================
def safe_get_location_point(elem):
    try:
        loc = elem.Location
        if loc:
            try:
                return loc.Point
            except:
                try:
                    crv = loc.Curve
                    return crv.Evaluate(0.5, True)
                except:
                    pass
        bb = elem.get_BoundingBox(None)
        if bb:
            return XYZ((bb.Min.X + bb.Max.X) / 2.0,
                       (bb.Min.Y + bb.Max.Y) / 2.0,
                       (bb.Min.Z + bb.Max.Z) / 2.0)
    except:
        pass
    return None

def get_check_points_3d(elem):
    pts = []
    try:
        loc = elem.Location
        if loc:
            try:
                pts.append(loc.Point)
            except:
                try:
                    crv = loc.Curve
                    pts.append(crv.Evaluate(0.0, True))
                    pts.append(crv.Evaluate(0.5, True))
                    pts.append(crv.Evaluate(1.0, True))
                except:
                    pass
    except:
        pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            cx = (bb.Min.X + bb.Max.X) / 2.0
            cy = (bb.Min.Y + bb.Max.Y) / 2.0
            cz = (bb.Min.Z + bb.Max.Z) / 2.0
            pts.append(XYZ(cx, cy, cz))
            pts.append(XYZ(cx, cy, bb.Min.Z))
            pts.append(XYZ(cx, cy, bb.Max.Z))
    except:
        pass
    return pts

def in_room(room, pt):
    try:
        return room.IsPointInRoom(pt)
    except:
        return False

def in_space(space, pt):
    try:
        return space.IsPointInSpace(pt)
    except:
        return False

def get_spatial_bbox(elem):
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            return (bb.Min.X, bb.Min.Y, bb.Min.Z, bb.Max.X, bb.Max.Y, bb.Max.Z)
    except:
        pass
    return None

def get_elem_bbox(elem):
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            return (bb.Min.X, bb.Min.Y, bb.Min.Z, bb.Max.X, bb.Max.Y, bb.Max.Z)
    except:
        pass
    return None

def bbox_intersects_3d(bb1, bb2):
    if not bb1 or not bb2:
        return False
    return not (bb1[3] < bb2[0] or bb1[0] > bb2[3] or
                bb1[4] < bb2[1] or bb1[1] > bb2[4] or
                bb1[5] < bb2[2] or bb1[2] > bb2[5])

def get_check_points_2d(elem):
    pts = []
    try:
        loc = elem.Location
        if loc:
            try:
                p = loc.Point
                pts.append((p.X, p.Y))
            except:
                try:
                    crv = loc.Curve
                    for t in [0.0, 0.5, 1.0]:
                        p = crv.Evaluate(t, True)
                        pts.append((p.X, p.Y))
                except:
                    pass
    except:
        pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            cx = (bb.Min.X + bb.Max.X) / 2.0
            cy = (bb.Min.Y + bb.Max.Y) / 2.0
            pts.append((cx, cy))
    except:
        pass
    return pts

def pt_in_poly(pt_2d, poly):
    if not poly or len(poly) < 3:
        return False
    x, y = pt_2d
    n = len(poly)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = poly[i]
        xj, yj = poly[j]
        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside

def in_area_2d(pt_2d, polygons):
    if not polygons or len(polygons) == 0:
        return False
    try:
        if not pt_in_poly(pt_2d, polygons[0]):
            return False
        for i in range(1, len(polygons)):
            if pt_in_poly(pt_2d, polygons[i]):
                return False
        return True
    except:
        return False

def get_area_polygons(area):
    polygons = []
    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        segs_list = area.GetBoundarySegments(opts)
        if segs_list:
            for segs in segs_list:
                poly = []
                for seg in segs:
                    crv = seg.GetCurve()
                    p = crv.GetEndPoint(0)
                    poly.append((p.X, p.Y))
                if len(poly) >= 3:
                    polygons.append(poly)
    except:
        pass
    return polygons

def get_room_boundary_ids(room_elem):
    """Get element IDs of all boundary elements of a room"""
    ids = set()
    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        segs_list = room_elem.GetBoundarySegments(opts)
        if segs_list:
            for segs in segs_list:
                for seg in segs:
                    try:
                        eid = seg.ElementId
                        if eid and eid.IntegerValue != -1:
                            ids.add(eid.IntegerValue)
                    except:
                        pass
    except:
        pass
    return ids

def check_element_in_room(elem, room_elem, use_bbox=True, boundary_ids=None):
    elem_id = elem.Id.IntegerValue
    
    # 1. Check if element is a boundary element of the room
    if boundary_ids is not None:
        if elem_id in boundary_ids:
            return True
    
    # 2. Check using IsPointInRoom with multiple points
    if use_bbox:
        room_bb = get_spatial_bbox(room_elem)
        elem_bb = get_elem_bbox(elem)
        if not bbox_intersects_3d(room_bb, elem_bb):
            return False
        pts = get_check_points_3d(elem)
        for pt in pts:
            try:
                if in_room(room_elem, pt):
                    return True
            except:
                continue
        
        # 3. For linear elements (walls, beams etc), sample more points along curve
        try:
            loc = elem.Location
            if hasattr(loc, 'Curve'):
                crv = loc.Curve
                for t in [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]:
                    try:
                        pt = crv.Evaluate(t, True)
                        if in_room(room_elem, pt):
                            return True
                    except:
                        continue
        except:
            pass
        
        return False
    else:
        pt = safe_get_location_point(elem)
        if not pt:
            return False
        return in_room(room_elem, pt)

def check_element_in_area(elem, area_elem, polygons, use_bbox=True):
    if not polygons or len(polygons) == 0:
        return False
    pts_2d = get_check_points_2d(elem)
    for pt in pts_2d:
        try:
            if in_area_2d(pt, polygons):
                return True
        except:
            continue
    return False

def check_element_in_space(elem, space_elem, use_bbox=True):
    if use_bbox:
        space_bb = get_spatial_bbox(space_elem)
        elem_bb = get_elem_bbox(elem)
        if not bbox_intersects_3d(space_bb, elem_bb):
            return False
        pts = get_check_points_3d(elem)
        for pt in pts:
            try:
                if in_space(space_elem, pt):
                    return True
            except:
                continue
        return False
    else:
        pt = safe_get_location_point(elem)
        if not pt:
            return False
        return in_space(space_elem, pt)

# ================================================================
# DATA MODEL
# ================================================================
def get_rooms(view_only=False):
    items = []
    try:
        if view_only:
            col = FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            col = FilteredElementCollector(doc)
        for r in col.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
            try:
                if r.Area > 0:
                    items.append(r)
            except:
                pass
    except:
        pass
    return items

def get_areas(view_only=False):
    items = []
    try:
        if view_only:
            col = FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            col = FilteredElementCollector(doc)
        for a in col.OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements():
            try:
                if a.Area > 0:
                    items.append(a)
            except:
                pass
    except:
        pass
    return items

def get_spaces(view_only=False):
    items = []
    try:
        if view_only:
            col = FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            col = FilteredElementCollector(doc)
        for s in col.OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements():
            try:
                if s.Area > 0:
                    items.append(s)
            except:
                pass
    except:
        pass
    return items

def get_spatial_info(elem, stype):
    """Get number, name, level for spatial element"""
    number = ""
    name = ""
    level = ""
    try:
        if stype == ROOMS:
            number = elem.Number or ""
            p = elem.get_Parameter(BuiltInParameter.ROOM_NAME)
            if p: name = p.AsString() or ""
            lp = elem.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
            if lp:
                lid = lp.AsElementId()
                le = doc.GetElement(lid)
                if le: level = le.Name or ""
        elif stype == AREAS:
            p = elem.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            if p: number = p.AsString() or ""
            p2 = elem.get_Parameter(BuiltInParameter.ROOM_NAME)
            if p2: name = p2.AsString() or ""
            lp = elem.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
            if lp:
                lid = lp.AsElementId()
                le = doc.GetElement(lid)
                if le: level = le.Name or ""
        elif stype == SPACES:
            p = elem.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            if p: number = p.AsString() or ""
            p2 = elem.get_Parameter(BuiltInParameter.ROOM_NAME)
            if p2: name = p2.AsString() or ""
            lp = elem.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
            if lp:
                lid = lp.AsElementId()
                le = doc.GetElement(lid)
                if le: level = le.Name or ""
    except:
        pass
    return number, name, level

USEFUL_BICS = [
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_FurnitureSystems,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_SecurityDevices,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_TelephoneDevices,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_StairsRailing,
    BuiltInCategory.OST_Ramps,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_Parking,
    BuiltInCategory.OST_Planting,
    BuiltInCategory.OST_Site,
    BuiltInCategory.OST_Topography,
    BuiltInCategory.OST_Entourage,
]

def get_categories():
    """Get categories that have elements in the model from predefined useful list"""
    cats = []
    for bic in USEFUL_BICS:
        try:
            col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            count = col.GetElementCount()
            if count > 0:
                # Get category name from first element
                first = col.FirstElement()
                if first and first.Category:
                    cats.append((first.Category.Name, int(bic), count))
        except:
            pass
    cats.sort(key=lambda x: x[0])
    return cats

def get_element_params(elements):
    """Get all instance parameter names from a list of elements"""
    params = set()
    for elem in elements[:20]:  # Sample first 20
        try:
            for p in elem.Parameters:
                try:
                    if p and p.Definition and p.Definition.Name:
                        pname = p.Definition.Name
                        if not p.IsReadOnly:
                            params.add(pname)
                        else:
                            params.add(pname)
                except:
                    pass
            # Also get type parameters
            tid = elem.GetTypeId()
            if tid and tid.IntegerValue != -1:
                et = doc.GetElement(tid)
                if et:
                    for p in et.Parameters:
                        try:
                            if p and p.Definition and p.Definition.Name:
                                params.add("[Type] " + p.Definition.Name)
                        except:
                            pass
        except:
            pass
    return sorted(list(params))

def get_param_value_str(elem, param_name):
    """Get parameter value as string from element"""
    is_type = param_name.startswith("[Type] ")
    actual_name = param_name[7:] if is_type else param_name

    target = elem
    if is_type:
        tid = elem.GetTypeId()
        if tid and tid.IntegerValue != -1:
            target = doc.GetElement(tid)
        else:
            return ""

    if not target:
        return ""

    try:
        p = target.LookupParameter(actual_name)
        if not p:
            return ""
        st = p.StorageType
        if st == StorageType.String:
            return p.AsString() or ""
        elif st == StorageType.Integer:
            return str(p.AsInteger())
        elif st == StorageType.Double:
            val = p.AsDouble()
            # Try to get display value
            try:
                return p.AsValueString() or str(round(val, 4))
            except:
                return str(round(val, 4))
        elif st == StorageType.ElementId:
            eid = p.AsElementId()
            if eid and eid.IntegerValue != -1:
                e = doc.GetElement(eid)
                if e:
                    return e.Name or str(eid.IntegerValue)
            return ""
    except:
        pass
    return ""

def get_param_value_numeric(elem, param_name):
    """Get parameter value as float in DISPLAY UNITS for numeric aggregation"""
    is_type = param_name.startswith("[Type] ")
    actual_name = param_name[7:] if is_type else param_name

    target = elem
    if is_type:
        tid = elem.GetTypeId()
        if tid and tid.IntegerValue != -1:
            target = doc.GetElement(tid)
        else:
            return None

    if not target:
        return None

    try:
        p = target.LookupParameter(actual_name)
        if not p:
            return None
        st = p.StorageType
        if st == StorageType.Double:
            # Use AsValueString to get display units value, then parse number
            try:
                vs = p.AsValueString()
                if vs:
                    # Extract the numeric part: allow digits, dot, minus, comma
                    # Examples: "146.304 m2", "22200", "3.14", "-200.5 mm"
                    clean = ""
                    found_digit = False
                    for ch in vs:
                        if ch.isdigit() or ch == '.':
                            clean += ch
                            found_digit = True
                        elif ch == '-' and not found_digit:
                            clean += ch
                        elif ch == ',' and found_digit:
                            # Could be thousands separator or decimal
                            # Check if next chars are digits
                            clean += '.'  # Treat as decimal for now
                        elif ch == ' ' and found_digit:
                            # Space after number = unit separator, stop
                            break
                        elif found_digit and not ch.isdigit():
                            break
                    # Handle case where comma was thousands separator (e.g. "1.000,50")
                    # If multiple dots, keep only the last one as decimal
                    dot_count = clean.count('.')
                    if dot_count > 1:
                        parts = clean.split('.')
                        clean = "".join(parts[:-1]) + "." + parts[-1]
                    if clean and clean != '-' and clean != '.':
                        return float(clean)
            except:
                pass
            # Fallback to raw double
            return p.AsDouble()
        elif st == StorageType.Integer:
            return float(p.AsInteger())
        else:
            # Try parsing string
            s = get_param_value_str(elem, param_name)
            try:
                return float(s)
            except:
                return None
    except:
        return None

def aggregate_values(elements, param_name, agg_type):
    """Aggregate parameter values from elements"""
    if agg_type == AGG_COUNT:
        return str(len(elements))

    str_values = []
    num_values = []

    for elem in elements:
        sv = get_param_value_str(elem, param_name)
        if sv:
            str_values.append(sv)
        nv = get_param_value_numeric(elem, param_name)
        if nv is not None:
            num_values.append(nv)

    if agg_type == AGG_SUM:
        if num_values:
            v = sum(num_values)
            return str(v) if v == int(v) else str(v)
        return "0"
    elif agg_type == AGG_AVERAGE:
        if num_values:
            v = sum(num_values) / len(num_values)
            return str(v)
        return "0"
    elif agg_type == AGG_MIN:
        if num_values:
            v = min(num_values)
            return str(v)
        return ""
    elif agg_type == AGG_MAX:
        if num_values:
            v = max(num_values)
            return str(v)
        return ""
    elif agg_type == AGG_FIRST:
        return str_values[0] if str_values else ""
    elif agg_type == AGG_LAST:
        return str_values[-1] if str_values else ""
    elif agg_type == AGG_LIST:
        return ", ".join(str_values) if str_values else ""
    elif agg_type == AGG_UNIQUE:
        seen = []
        for v in str_values:
            if v not in seen:
                seen.append(v)
        return ", ".join(seen) if seen else ""

    return str(len(elements))

def get_writable_spatial_params(spatial_elem):
    """Get writable string/text parameters of a spatial element"""
    params = []
    try:
        for p in spatial_elem.Parameters:
            try:
                if p and p.Definition and p.Definition.Name:
                    if not p.IsReadOnly and p.StorageType == StorageType.String:
                        params.append(p.Definition.Name)
            except:
                pass
    except:
        pass
    return sorted(list(set(params)))

# ================================================================
# DATA CLASSES
# ================================================================
class SpatialData:
    def __init__(self, elem, stype):
        self.element = elem
        self.element_id = elem.Id.IntegerValue
        self.spatial_type = stype
        self.is_selected = False
        self.number, self.name, self.level = get_spatial_info(elem, stype)
        self.polygons = []
        self.boundary_ids = set()
        if stype == AREAS:
            self.polygons = get_area_polygons(elem)
        if stype in (ROOMS, SPACES):
            self.boundary_ids = get_room_boundary_ids(elem)
        self.contained_elements = []
        self.aggregated_value = ""

    @property
    def display_name(self):
        parts = []
        if self.number:
            parts.append(self.number)
        if self.name:
            parts.append(self.name)
        if not parts:
            return "ID: " + str(self.element_id)
        return " - ".join(parts)


class CatItem:
    def __init__(self, name, cat_id, count):
        self.name = name
        self.cat_id = cat_id
        self.count = count
        self.is_selected = False


class CollectResult:
    def __init__(self, spatial, elements, param_name, agg_type):
        self.spatial = spatial
        self.elements = elements
        self.element_count = len(elements)
        self.param_name = param_name
        self.agg_type = agg_type
        self.agg_value = aggregate_values(elements, param_name, agg_type) if param_name else str(len(elements))
        self.is_selected = False


# ================================================================
# MAIN WINDOW
# ================================================================
class RoomDataCollectorWindow(Window):
    def __init__(self):
        self.Title = "Room Data Collector - pyDQT"
        self.Width = 1200
        self.Height = 720
        self.MinWidth = 950
        self.MinHeight = 550
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = brush(WHITE)
        self.ResizeMode = ResizeMode.CanResize

        self.spatial_type = ROOMS
        self.view_only = False
        self.spatial_items = []
        self.cat_items = []
        self.results = []
        self.all_results = []
        self.elem_params = []  # Available element parameters

        self._build_ui()
        self._load_data()

    def _build_ui(self):
        root = DockPanel()
        root.LastChildFill = True

        header = self._make_header()
        DockPanel.SetDock(header, Dock.Top)
        root.Children.Add(header)

        footer = self._make_footer()
        DockPanel.SetDock(footer, Dock.Bottom)
        root.Children.Add(footer)

        content = self._make_content()
        root.Children.Add(content)

        self.Content = root

    def _make_header(self):
        bd = Border()
        bd.Background = brush(PRIMARY)
        bd.Padding = Thickness(20, 10, 20, 10)

        outer = DockPanel()
        outer.LastChildFill = True

        # Right side: cards
        cards = StackPanel()
        cards.Orientation = Orientation.Horizontal
        cards.Margin = Thickness(10, 0, 0, 0)
        cards.VerticalAlignment = VerticalAlignment.Center
        self.card_rooms = self._card("Rooms", "0")
        self.card_areas = self._card("Areas", "0")
        self.card_spaces = self._card("Spaces", "0")
        self.card_found = self._card("Results", "0")
        cards.Children.Add(self.card_rooms)
        cards.Children.Add(self.card_areas)
        cards.Children.Add(self.card_spaces)
        cards.Children.Add(self.card_found)
        DockPanel.SetDock(cards, Dock.Right)
        outer.Children.Add(cards)

        # Left side: title + copyright
        tp = StackPanel()
        tp.VerticalAlignment = VerticalAlignment.Center
        t = TextBlock()
        t.Text = "Room Data Collector"
        t.FontSize = 24
        t.FontWeight = FontWeights.Bold
        t.Foreground = brush(TEXT_DARK)
        tp.Children.Add(t)
        st = TextBlock()
        st.Text = "Copyright (c) 2026 by Dang Quoc Truong (DQT)"
        st.FontSize = 10
        st.Foreground = brush(TEXT_GRAY)
        tp.Children.Add(st)
        outer.Children.Add(tp)

        bd.Child = outer
        return bd

    def _card(self, label, value):
        bd = Border()
        bd.Background = brush(WHITE)
        bd.CornerRadius = CornerRadius(6)
        bd.Padding = Thickness(12, 4, 12, 4)
        bd.Margin = Thickness(4, 0, 4, 0)
        bd.MinWidth = 70

        sp = StackPanel()
        sp.HorizontalAlignment = HorizontalAlignment.Center
        vt = TextBlock()
        vt.Text = value
        vt.FontSize = 22
        vt.FontWeight = FontWeights.Bold
        vt.Foreground = brush(TEXT_DARK)
        vt.HorizontalAlignment = HorizontalAlignment.Center
        vt.Tag = "value"
        sp.Children.Add(vt)

        lt = TextBlock()
        lt.Text = label
        lt.FontSize = 12
        lt.Foreground = brush(TEXT_GRAY)
        lt.HorizontalAlignment = HorizontalAlignment.Center
        sp.Children.Add(lt)

        bd.Child = sp
        return bd

    def _update_card(self, card, value):
        sp = card.Child
        for child in sp.Children:
            if hasattr(child, 'Tag') and child.Tag == "value":
                child.Text = str(value)
                break

    def _make_footer(self):
        bd = Border()
        bd.Background = brush(SECONDARY)
        bd.Padding = Thickness(20, 8, 20, 8)

        dp = DockPanel()
        dp.LastChildFill = True

        # Right side: buttons
        btn_sp = StackPanel()
        btn_sp.Orientation = Orientation.Horizontal
        btn_sp.HorizontalAlignment = HorizontalAlignment.Right

        self.btn_collect = self._btn("Collect Data", self._on_collect, ACCENT, WHITE, 130)
        self.btn_apply = self._btn("Apply to Rooms", self._on_apply, SUCCESS, WHITE, 140)
        self.btn_select = self._btn("Select Elements", self._on_select, "#2196F3", WHITE, 140)
        self.btn_close = self._btn("Close", self._on_close, BORDER, TEXT_DARK, 80)

        btn_sp.Children.Add(self.btn_collect)
        btn_sp.Children.Add(self.btn_apply)
        btn_sp.Children.Add(self.btn_select)
        btn_sp.Children.Add(self.btn_close)

        DockPanel.SetDock(btn_sp, Dock.Right)
        dp.Children.Add(btn_sp)

        # Left side: copyright
        copy_txt = TextBlock()
        copy_txt.Text = "Copyright (c) 2026 by Dang Quoc Truong (DQT)"
        copy_txt.FontSize = 10
        copy_txt.Foreground = brush(TEXT_MUTED)
        copy_txt.VerticalAlignment = VerticalAlignment.Center
        dp.Children.Add(copy_txt)

        bd.Child = dp
        return bd

    def _btn(self, text, handler, bg, fg, width=100):
        b = Button()
        b.Content = text
        b.Width = width
        b.Height = 36
        b.Margin = Thickness(4, 0, 4, 0)
        b.Background = brush(bg)
        b.Foreground = brush(fg)
        b.FontWeight = FontWeights.SemiBold
        b.FontSize = 15
        b.BorderThickness = Thickness(0)
        b.Cursor = System.Windows.Input.Cursors.Hand
        b.Click += handler
        return b

    def _make_content(self):
        main_sp = DockPanel()
        main_sp.LastChildFill = True
        main_sp.Margin = Thickness(10, 8, 10, 0)

        # Left panel: Spatial + Categories + Config
        left = self._make_left_panel()
        DockPanel.SetDock(left, Dock.Left)
        main_sp.Children.Add(left)

        # Right panel: Results
        right = self._make_right_panel()
        main_sp.Children.Add(right)

        return main_sp

    def _make_left_panel(self):
        bd = Border()
        bd.Width = 320
        bd.Margin = Thickness(0, 0, 8, 0)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto

        sp = StackPanel()

        # --- Spatial Type ---
        sp.Children.Add(self._section_label("Spatial Type"))
        type_sp = StackPanel()
        type_sp.Orientation = Orientation.Horizontal
        type_sp.Margin = Thickness(0, 2, 0, 6)

        self.rb_rooms = RadioButton()
        self.rb_rooms.Content = "Rooms"
        self.rb_rooms.IsChecked = True
        self.rb_rooms.Margin = Thickness(0, 0, 12, 0)
        self.rb_rooms.Checked += self._on_type_changed
        type_sp.Children.Add(self.rb_rooms)

        self.rb_areas = RadioButton()
        self.rb_areas.Content = "Areas"
        self.rb_areas.Margin = Thickness(0, 0, 12, 0)
        self.rb_areas.Checked += self._on_type_changed
        type_sp.Children.Add(self.rb_areas)

        self.rb_spaces = RadioButton()
        self.rb_spaces.Content = "Spaces"
        self.rb_spaces.Checked += self._on_type_changed
        type_sp.Children.Add(self.rb_spaces)

        sp.Children.Add(type_sp)

        # View scope
        self.cb_view_only = CheckBox()
        self.cb_view_only.Content = "Active View Only"
        self.cb_view_only.Margin = Thickness(0, 0, 0, 6)
        self.cb_view_only.Checked += self._on_scope_changed
        self.cb_view_only.Unchecked += self._on_scope_changed
        sp.Children.Add(self.cb_view_only)

        # --- Spatial Elements List ---
        sp.Children.Add(self._section_label("Spatial Elements"))

        sp_btns = StackPanel()
        sp_btns.Orientation = Orientation.Horizontal
        sp_btns.Margin = Thickness(0, 0, 0, 4)
        btn_all_sp = self._small_btn("All", self._sel_all_spatial)
        btn_none_sp = self._small_btn("None", self._sel_none_spatial)
        btn_invert_sp = self._small_btn("Invert", self._sel_invert_spatial)
        sp_btns.Children.Add(btn_all_sp)
        sp_btns.Children.Add(btn_none_sp)
        sp_btns.Children.Add(btn_invert_sp)
        sp.Children.Add(sp_btns)

        self.spatial_search = TextBox()
        self.spatial_search.Height = 28
        self.spatial_search.Margin = Thickness(0, 0, 0, 4)
        self.spatial_search.Tag = "search_spatial"
        self.spatial_search.TextChanged += self._on_spatial_search
        sp.Children.Add(self.spatial_search)

        self.spatial_scroll = ScrollViewer()
        self.spatial_scroll.Height = 140
        self.spatial_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.spatial_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        bd_sp = Border()
        bd_sp.BorderBrush = brush(BORDER)
        bd_sp.BorderThickness = Thickness(1)
        bd_sp.CornerRadius = CornerRadius(4)
        self.spatial_panel = StackPanel()
        bd_sp.Child = self.spatial_panel
        self.spatial_scroll.Content = bd_sp
        sp.Children.Add(self.spatial_scroll)

        # --- Categories ---
        sp.Children.Add(self._section_label("Element Categories"))

        cat_btns = StackPanel()
        cat_btns.Orientation = Orientation.Horizontal
        cat_btns.Margin = Thickness(0, 0, 0, 4)
        btn_all_c = self._small_btn("All", self._sel_all_cats)
        btn_none_c = self._small_btn("None", self._sel_none_cats)
        cat_btns.Children.Add(btn_all_c)
        cat_btns.Children.Add(btn_none_c)
        sp.Children.Add(cat_btns)

        self.cat_search = TextBox()
        self.cat_search.Height = 28
        self.cat_search.Margin = Thickness(0, 0, 0, 4)
        self.cat_search.TextChanged += self._on_cat_search
        sp.Children.Add(self.cat_search)

        self.cat_scroll = ScrollViewer()
        self.cat_scroll.Height = 120
        self.cat_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.cat_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        bd_c = Border()
        bd_c.BorderBrush = brush(BORDER)
        bd_c.BorderThickness = Thickness(1)
        bd_c.CornerRadius = CornerRadius(4)
        self.cat_panel = StackPanel()
        bd_c.Child = self.cat_panel
        self.cat_scroll.Content = bd_c
        sp.Children.Add(self.cat_scroll)

        # --- Aggregation Config ---
        sp.Children.Add(self._section_label("Aggregation Settings"))

        # Source parameter
        lbl_src = TextBlock()
        lbl_src.Text = "Source Parameter (from elements):"
        lbl_src.FontSize = 14
        lbl_src.Foreground = brush(TEXT_GRAY)
        lbl_src.Margin = Thickness(0, 2, 0, 2)
        sp.Children.Add(lbl_src)

        self.cmb_source_param = ComboBox()
        self.cmb_source_param.Height = 32
        self.cmb_source_param.Margin = Thickness(0, 0, 0, 6)
        self.cmb_source_param.IsEditable = True
        sp.Children.Add(self.cmb_source_param)

        # Aggregation type
        lbl_agg = TextBlock()
        lbl_agg.Text = "Aggregation Method:"
        lbl_agg.FontSize = 14
        lbl_agg.Foreground = brush(TEXT_GRAY)
        lbl_agg.Margin = Thickness(0, 2, 0, 2)
        sp.Children.Add(lbl_agg)

        self.cmb_agg_type = ComboBox()
        self.cmb_agg_type.Height = 32
        self.cmb_agg_type.Margin = Thickness(0, 0, 0, 6)
        for agg in ALL_AGG:
            item = ComboBoxItem()
            item.Content = agg
            self.cmb_agg_type.Items.Add(item)
        self.cmb_agg_type.SelectedIndex = 0
        sp.Children.Add(self.cmb_agg_type)

        # Target parameter
        lbl_tgt = TextBlock()
        lbl_tgt.Text = "Target Parameter (on Room/Area/Space):"
        lbl_tgt.FontSize = 14
        lbl_tgt.Foreground = brush(TEXT_GRAY)
        lbl_tgt.Margin = Thickness(0, 2, 0, 2)
        sp.Children.Add(lbl_tgt)

        self.cmb_target_param = ComboBox()
        self.cmb_target_param.Height = 32
        self.cmb_target_param.Margin = Thickness(0, 0, 0, 6)
        self.cmb_target_param.IsEditable = True
        sp.Children.Add(self.cmb_target_param)

        scroll.Content = sp
        bd.Child = scroll
        return bd

    def _make_right_panel(self):
        bd = Border()
        bd.BorderBrush = brush(BORDER)
        bd.BorderThickness = Thickness(1)
        bd.CornerRadius = CornerRadius(6)

        dp = DockPanel()
        dp.LastChildFill = True

        # Top: Search + info
        top = StackPanel()
        top.Margin = Thickness(8, 6, 8, 4)

        top_row = StackPanel()
        top_row.Orientation = Orientation.Horizontal

        self.result_search = TextBox()
        self.result_search.Width = 250
        self.result_search.Height = 30
        self.result_search.Margin = Thickness(0, 0, 8, 0)
        self.result_search.TextChanged += self._on_result_search
        top_row.Children.Add(self.result_search)

        self.lbl_info = TextBlock()
        self.lbl_info.Text = "No results"
        self.lbl_info.FontSize = 14
        self.lbl_info.Foreground = brush(TEXT_GRAY)
        self.lbl_info.VerticalAlignment = VerticalAlignment.Center
        top_row.Children.Add(self.lbl_info)

        top.Children.Add(top_row)

        # Select all/none for results
        res_btns = StackPanel()
        res_btns.Orientation = Orientation.Horizontal
        res_btns.Margin = Thickness(0, 4, 0, 4)
        btn_ra = self._small_btn("All", self._sel_all_results)
        btn_rn = self._small_btn("None", self._sel_none_results)
        res_btns.Children.Add(btn_ra)
        res_btns.Children.Add(btn_rn)
        top.Children.Add(res_btns)

        DockPanel.SetDock(top, Dock.Top)
        dp.Children.Add(top)

        # Results list (manual rows)
        self.result_scroll = ScrollViewer()
        self.result_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.result_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto

        self.result_panel = StackPanel()
        self.result_scroll.Content = self.result_panel
        dp.Children.Add(self.result_scroll)

        bd.Child = dp
        return bd

    def _section_label(self, text):
        t = TextBlock()
        t.Text = text
        t.FontSize = 15
        t.FontWeight = FontWeights.SemiBold
        t.Foreground = brush(TEXT_DARK)
        t.Margin = Thickness(0, 8, 0, 4)
        return t

    def _small_btn(self, text, handler):
        b = Button()
        b.Content = text
        b.Width = 50
        b.Height = 28
        b.FontSize = 14
        b.Margin = Thickness(0, 0, 4, 0)
        b.Background = brush(SECONDARY)
        b.Foreground = brush(TEXT_DARK)
        b.BorderThickness = Thickness(1)
        b.BorderBrush = brush(BORDER)
        b.Cursor = System.Windows.Input.Cursors.Hand
        b.Click += handler
        return b

    # ================================================================
    # DATA LOADING
    # ================================================================
    def _load_data(self):
        self._load_spatial()
        self._load_cats()
        self._update_cards()

    def _load_spatial(self):
        self.spatial_items = []
        if self.spatial_type == ROOMS:
            for e in get_rooms(self.view_only):
                self.spatial_items.append(SpatialData(e, ROOMS))
        elif self.spatial_type == AREAS:
            for e in get_areas(self.view_only):
                self.spatial_items.append(SpatialData(e, AREAS))
        elif self.spatial_type == SPACES:
            for e in get_spaces(self.view_only):
                self.spatial_items.append(SpatialData(e, SPACES))

        self._refresh_spatial_list()
        self._load_target_params()

    def _load_cats(self):
        self.cat_items = []
        cats = get_categories()
        for name, cid, count in cats:
            self.cat_items.append(CatItem(name, cid, count))
        self._refresh_cat_list()

    def _load_target_params(self):
        """Load writable parameters from spatial elements"""
        self.cmb_target_param.Items.Clear()
        params = set()
        for si in self.spatial_items[:5]:
            for pn in get_writable_spatial_params(si.element):
                params.add(pn)
        for pn in sorted(params):
            item = ComboBoxItem()
            item.Content = pn
            self.cmb_target_param.Items.Add(item)

    def _refresh_spatial_list(self, filter_text=""):
        self.spatial_panel.Children.Clear()
        ft = filter_text.lower()
        for si in self.spatial_items:
            display = si.display_name
            if ft and ft not in display.lower():
                continue
            row = self._make_check_row(display, si, "spatial")
            self.spatial_panel.Children.Add(row)

    def _refresh_cat_list(self, filter_text=""):
        self.cat_panel.Children.Clear()
        ft = filter_text.lower()
        for ci in self.cat_items:
            if ft and ft not in ci.name.lower():
                continue
            row = self._make_check_row(ci.name + " (" + str(ci.count) + ")", ci, "cat")
            self.cat_panel.Children.Add(row)

    def _make_check_row(self, text, data_item, tag):
        cb = CheckBox()
        cb.Content = text
        cb.FontSize = 14
        cb.Margin = Thickness(4, 1, 4, 1)
        cb.IsChecked = data_item.is_selected
        cb.Tag = data_item
        cb.Checked += self._on_check_changed
        cb.Unchecked += self._on_check_changed
        return cb

    def _on_check_changed(self, sender, e):
        if sender.Tag:
            sender.Tag.is_selected = bool(sender.IsChecked)
            # If it's a category item, update source parameters
            if isinstance(sender.Tag, CatItem):
                self._update_source_params()

    def _update_source_params(self):
        """Populate source parameter combobox by sampling elements from selected categories"""
        sel_cats = [ci for ci in self.cat_items if ci.is_selected]
        if not sel_cats:
            self.cmb_source_param.Items.Clear()
            return

        # Save current selection
        old_text = ""
        if self.cmb_source_param.SelectedItem:
            old_text = self.cmb_source_param.SelectedItem.Content
        elif self.cmb_source_param.Text:
            old_text = self.cmb_source_param.Text

        sample_elems = []
        for ci in sel_cats:
            try:
                bic = System.Enum.ToObject(BuiltInCategory, ci.cat_id)
                col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
                count = 0
                for elem in col:
                    sample_elems.append(elem)
                    count += 1
                    if count >= 5:
                        break
            except:
                pass

        if sample_elems:
            self.elem_params = get_element_params(sample_elems)
            self.cmb_source_param.Items.Clear()
            for pn in self.elem_params:
                item = ComboBoxItem()
                item.Content = pn
                self.cmb_source_param.Items.Add(item)
            # Restore selection
            if old_text:
                self.cmb_source_param.Text = old_text

    def _update_cards(self):
        rooms = len([s for s in self.spatial_items if s.spatial_type == ROOMS]) if self.spatial_type == ROOMS else 0
        areas = len([s for s in self.spatial_items if s.spatial_type == AREAS]) if self.spatial_type == AREAS else 0
        spaces = len([s for s in self.spatial_items if s.spatial_type == SPACES]) if self.spatial_type == SPACES else 0

        if self.spatial_type == ROOMS:
            self._update_card(self.card_rooms, len(self.spatial_items))
        else:
            self._update_card(self.card_rooms, "0")
        if self.spatial_type == AREAS:
            self._update_card(self.card_areas, len(self.spatial_items))
        else:
            self._update_card(self.card_areas, "0")
        if self.spatial_type == SPACES:
            self._update_card(self.card_spaces, len(self.spatial_items))
        else:
            self._update_card(self.card_spaces, "0")

        self._update_card(self.card_found, len(self.results))

    # ================================================================
    # EVENT HANDLERS
    # ================================================================
    def _on_type_changed(self, sender, e):
        if self.rb_rooms.IsChecked:
            self.spatial_type = ROOMS
        elif self.rb_areas.IsChecked:
            self.spatial_type = AREAS
        elif self.rb_spaces.IsChecked:
            self.spatial_type = SPACES
        self._load_spatial()
        self._update_cards()

    def _on_scope_changed(self, sender, e):
        self.view_only = bool(self.cb_view_only.IsChecked)
        self._load_spatial()
        self._update_cards()

    def _on_spatial_search(self, sender, e):
        self._refresh_spatial_list(sender.Text or "")

    def _on_cat_search(self, sender, e):
        self._refresh_cat_list(sender.Text or "")

    def _on_result_search(self, sender, e):
        ft = (sender.Text or "").lower()
        self._refresh_results(ft)

    def _sel_all_spatial(self, s, e):
        for si in self.spatial_items:
            si.is_selected = True
        self._refresh_spatial_list(self.spatial_search.Text or "")

    def _sel_none_spatial(self, s, e):
        for si in self.spatial_items:
            si.is_selected = False
        self._refresh_spatial_list(self.spatial_search.Text or "")

    def _sel_invert_spatial(self, s, e):
        for si in self.spatial_items:
            si.is_selected = not si.is_selected
        self._refresh_spatial_list(self.spatial_search.Text or "")

    def _sel_all_cats(self, s, e):
        for ci in self.cat_items:
            ci.is_selected = True
        self._refresh_cat_list(self.cat_search.Text or "")
        self._update_source_params()

    def _sel_none_cats(self, s, e):
        for ci in self.cat_items:
            ci.is_selected = False
        self._refresh_cat_list(self.cat_search.Text or "")
        self._update_source_params()

    def _sel_all_results(self, s, e):
        for r in self.results:
            r.is_selected = True
        self._refresh_results()

    def _sel_none_results(self, s, e):
        for r in self.results:
            r.is_selected = False
        self._refresh_results()

    # ================================================================
    # COLLECT DATA
    # ================================================================
    def _on_collect(self, s, e):
        sel_spatial = [si for si in self.spatial_items if si.is_selected]
        sel_cats = [ci for ci in self.cat_items if ci.is_selected]

        if not sel_spatial:
            TaskDialog.Show("Room Data Collector", "Please select at least one spatial element.")
            return
        if not sel_cats:
            TaskDialog.Show("Room Data Collector", "Please select at least one element category.")
            return

        # Get aggregation settings
        src_param = ""
        if self.cmb_source_param.SelectedItem:
            src_param = self.cmb_source_param.SelectedItem.Content
        elif self.cmb_source_param.Text:
            src_param = self.cmb_source_param.Text

        agg_idx = self.cmb_agg_type.SelectedIndex
        agg_type = ALL_AGG[agg_idx] if agg_idx >= 0 else AGG_COUNT

        # Collect elements by category using FilteredElementCollector per BIC
        all_elements = []
        for ci in sel_cats:
            try:
                bic = System.Enum.ToObject(BuiltInCategory, ci.cat_id)
                col = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
                for elem in col:
                    all_elements.append(elem)
            except:
                pass

        # For each spatial element, find contained elements
        self.results = []
        self.all_results = []
        total_found = 0

        for si in sel_spatial:
            contained = []
            for elem in all_elements:
                try:
                    found = False
                    if si.spatial_type == ROOMS:
                        found = check_element_in_room(elem, si.element, boundary_ids=si.boundary_ids)
                    elif si.spatial_type == AREAS:
                        found = check_element_in_area(elem, si.element, si.polygons)
                    elif si.spatial_type == SPACES:
                        found = check_element_in_space(elem, si.element)
                    if found:
                        contained.append(elem)
                except:
                    pass

            si.contained_elements = contained
            result = CollectResult(si, contained, src_param, agg_type)
            self.results.append(result)
            self.all_results.append(result)
            total_found += len(contained)

        # Update source parameter combobox with actual params from found elements
        found_elems_all = []
        for r in self.results:
            found_elems_all.extend(r.elements)

        if found_elems_all:
            self.elem_params = get_element_params(found_elems_all)
            self.cmb_source_param.Items.Clear()
            for pn in self.elem_params:
                item = ComboBoxItem()
                item.Content = pn
                self.cmb_source_param.Items.Add(item)
            # Restore selection
            if src_param:
                self.cmb_source_param.Text = src_param

        self._refresh_results()
        self._update_card(self.card_found, len(self.results))

        TaskDialog.Show("Room Data Collector",
            "Collected data for " + str(len(sel_spatial)) + " spatial elements.\n" +
            "Total elements found: " + str(total_found))

    def _refresh_results(self, filter_text=""):
        self.result_panel.Children.Clear()
        ft = filter_text.lower() if filter_text else ""

        # Header row
        header = self._make_result_header()
        self.result_panel.Children.Add(header)

        visible = 0
        for r in self.results:
            display = r.spatial.display_name
            if ft and ft not in display.lower() and ft not in r.agg_value.lower():
                continue
            row = self._make_result_row(r)
            self.result_panel.Children.Add(row)
            visible += 1

        self.lbl_info.Text = str(visible) + " of " + str(len(self.results)) + " results"

    def _make_result_header(self):
        bd = Border()
        bd.Background = brush(PRIMARY)
        bd.Padding = Thickness(4, 6, 4, 6)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        cb_all = CheckBox()
        cb_all.Width = 28
        cb_all.Margin = Thickness(4, 0, 0, 0)
        cb_all.Checked += self._sel_all_results
        cb_all.Unchecked += self._sel_none_results
        sp.Children.Add(cb_all)

        cols = [("Spatial Element", 200), ("Level", 100), ("Elements", 70),
                ("Source Param", 150), ("Agg. Method", 90), ("Result Value", 200)]
        for label, w in cols:
            t = TextBlock()
            t.Text = label
            t.Width = w
            t.FontSize = 14
            t.FontWeight = FontWeights.SemiBold
            t.Foreground = brush(TEXT_DARK)
            t.Margin = Thickness(4, 0, 0, 0)
            sp.Children.Add(t)

        bd.Child = sp
        return bd

    def _make_result_row(self, result):
        bd = Border()
        bd.Padding = Thickness(4, 4, 4, 4)
        bd.BorderBrush = brush(BORDER)
        bd.BorderThickness = Thickness(0, 0, 0, 1)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        cb = CheckBox()
        cb.Width = 28
        cb.Margin = Thickness(4, 0, 0, 0)
        cb.IsChecked = result.is_selected
        cb.Tag = result
        cb.Checked += self._on_result_check
        cb.Unchecked += self._on_result_check
        sp.Children.Add(cb)

        # Spatial element name
        t1 = TextBlock()
        t1.Text = result.spatial.display_name
        t1.Width = 200
        t1.FontSize = 14
        t1.Foreground = brush(TEXT_DARK)
        t1.Margin = Thickness(4, 0, 0, 0)
        t1.TextWrapping = TextWrapping.NoWrap
        sp.Children.Add(t1)

        # Level
        t2 = TextBlock()
        t2.Text = result.spatial.level
        t2.Width = 100
        t2.FontSize = 14
        t2.Foreground = brush(TEXT_GRAY)
        t2.Margin = Thickness(4, 0, 0, 0)
        sp.Children.Add(t2)

        # Element count
        t3 = TextBlock()
        t3.Text = str(result.element_count)
        t3.Width = 70
        t3.FontSize = 14
        t3.FontWeight = FontWeights.SemiBold
        t3.Foreground = brush(ACCENT)
        t3.Margin = Thickness(4, 0, 0, 0)
        sp.Children.Add(t3)

        # Source param
        t4 = TextBlock()
        t4.Text = result.param_name or "(Count only)"
        t4.Width = 150
        t4.FontSize = 14
        t4.Foreground = brush(TEXT_GRAY)
        t4.Margin = Thickness(4, 0, 0, 0)
        sp.Children.Add(t4)

        # Aggregation method
        t5 = TextBlock()
        t5.Text = result.agg_type
        t5.Width = 90
        t5.FontSize = 14
        t5.Foreground = brush(TEXT_GRAY)
        t5.Margin = Thickness(4, 0, 0, 0)
        sp.Children.Add(t5)

        # Result value
        t6 = TextBlock()
        t6.Text = result.agg_value
        t6.Width = 200
        t6.FontSize = 14
        t6.FontWeight = FontWeights.SemiBold
        t6.Foreground = brush(SUCCESS)
        t6.Margin = Thickness(4, 0, 0, 0)
        t6.TextWrapping = TextWrapping.NoWrap
        sp.Children.Add(t6)

        bd.Child = sp
        return bd

    def _on_result_check(self, sender, e):
        if sender.Tag:
            sender.Tag.is_selected = bool(sender.IsChecked)

    # ================================================================
    # APPLY TO ROOMS
    # ================================================================
    def _on_apply(self, s, e):
        sel_results = [r for r in self.results if r.is_selected]
        if not sel_results:
            TaskDialog.Show("Apply", "Please select results to apply.")
            return

        # Get target parameter
        tgt_param = ""
        if self.cmb_target_param.SelectedItem:
            tgt_param = self.cmb_target_param.SelectedItem.Content
        elif self.cmb_target_param.Text:
            tgt_param = self.cmb_target_param.Text

        if not tgt_param:
            TaskDialog.Show("Apply", "Please select a Target Parameter.")
            return

        # Recalculate aggregation with current settings
        src_param = ""
        if self.cmb_source_param.SelectedItem:
            src_param = self.cmb_source_param.SelectedItem.Content
        elif self.cmb_source_param.Text:
            src_param = self.cmb_source_param.Text

        agg_idx = self.cmb_agg_type.SelectedIndex
        agg_type = ALL_AGG[agg_idx] if agg_idx >= 0 else AGG_COUNT

        t = Transaction(doc, "Room Data Collector - Apply")
        t.Start()
        try:
            ok, fail = 0, 0
            for result in sel_results:
                spatial_elem = result.spatial.element
                # Recalculate value with current settings
                val = aggregate_values(result.elements, src_param, agg_type)

                try:
                    p = spatial_elem.LookupParameter(tgt_param)
                    if p and not p.IsReadOnly:
                        if p.StorageType == StorageType.String:
                            p.Set(val)
                            ok += 1
                        elif p.StorageType == StorageType.Double:
                            try:
                                p.Set(float(val))
                                ok += 1
                            except:
                                fail += 1
                        elif p.StorageType == StorageType.Integer:
                            try:
                                p.Set(int(float(val)))
                                ok += 1
                            except:
                                fail += 1
                        else:
                            fail += 1
                    else:
                        fail += 1
                except:
                    fail += 1

            t.Commit()
            TaskDialog.Show("Apply", "Updated: " + str(ok) + "\nFailed: " + str(fail))

            # Update results display
            for result in sel_results:
                result.agg_value = aggregate_values(result.elements, src_param, agg_type)
            self._refresh_results()

        except Exception as ex:
            t.RollBack()
            TaskDialog.Show("Error", str(ex))

    # ================================================================
    # SELECT ELEMENTS
    # ================================================================
    def _on_select(self, s, e):
        sel_results = [r for r in self.results if r.is_selected]
        if not sel_results:
            TaskDialog.Show("Select", "Please select results first.")
            return
        ids = List[ElementId]()
        for r in sel_results:
            for elem in r.elements:
                ids.Add(elem.Id)
        if ids.Count > 0:
            uidoc.Selection.SetElementIds(ids)
            TaskDialog.Show("Select", "Selected " + str(ids.Count) + " element(s).")
        else:
            TaskDialog.Show("Select", "No elements to select.")

    def _on_close(self, s, e):
        self.Close()


# ================================================================
# ENTRY POINT
# ================================================================
def main():
    try:
        win = RoomDataCollectorWindow()
        win.ShowDialog()
    except Exception as ex:
        TaskDialog.Show("Error", str(ex))

if __name__ == "__main__":
    main()