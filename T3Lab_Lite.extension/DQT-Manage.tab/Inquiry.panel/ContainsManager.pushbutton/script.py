# -*- coding: utf-8 -*-
"""
Contains Manager - Find elements in Rooms/Areas/Spaces and assign parameters
Copyright (c) 2024 Dang Quoc Truong (DQT)
"""

__title__ = "Contains\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Find elements in Rooms/Areas/Spaces and assign parameter values"

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
from System.Windows.Controls import StackPanel, DockPanel, Border, TextBlock, TextBox
from System.Windows.Controls import Button, ListBox, ComboBox, CheckBox, RadioButton
from System.Windows.Controls import Orientation, Dock, ScrollViewer, ScrollBarVisibility, SelectionMode
from System.Windows.Media import BrushConverter
import System.Windows

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter
from Autodesk.Revit.DB import Transaction, StorageType, XYZ, SpatialElementBoundaryOptions, ElementId
from Autodesk.Revit.DB import SpatialElementBoundaryLocation
from Autodesk.Revit.UI import TaskDialog

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

_conv = BrushConverter()
def brush(c): return _conv.ConvertFromString(c)

PRIMARY = "#F0CC88"
SECONDARY = "#FEF8E7"
WHITE = "#FFFFFF"
BORDER = "#E0E0E0"
TEXT_DARK = "#333333"
TEXT_GRAY = "#666666"
TEXT_MUTED = "#999999"
SUCCESS = "#4CAF50"
HIGHLIGHT = "#C8E6C9"

ROOMS = "Rooms"
AREAS = "Areas"
SPACES = "Spaces"

BBOX_CATEGORIES = [
    BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Columns, BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_Railings, BuiltInCategory.OST_Stairs,
]

def get_room_parameters(room):
    """Get all parameters from a room element"""
    params = []
    if not room:
        return params
    try:
        for p in room.Parameters:
            if p.HasValue:
                name = p.Definition.Name
                if name and name not in params:
                    params.append(name)
    except:
        pass
    return sorted(params)

def get_param_value(elem, param_name):
    """Get parameter value from element"""
    try:
        p = elem.LookupParameter(param_name)
        if p and p.HasValue:
            if p.StorageType == StorageType.String:
                return p.AsString() or ""
            elif p.StorageType == StorageType.Integer:
                return str(p.AsInteger())
            elif p.StorageType == StorageType.Double:
                return str(round(p.AsDouble(), 2))
            elif p.StorageType == StorageType.ElementId:
                eid = p.AsElementId()
                if eid and eid.IntegerValue != -1:
                    el = doc.GetElement(eid)
                    if el:
                        return el.Name or str(eid.IntegerValue)
        return ""
    except:
        return ""

def get_room_boundary_elements(room):
    boundary_ids = set()
    try:
        opts = SpatialElementBoundaryOptions()
        opts.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        boundaries = room.GetBoundarySegments(opts)
        if boundaries:
            for boundary_loop in boundaries:
                for segment in boundary_loop:
                    elem_id = segment.ElementId
                    if elem_id and elem_id.IntegerValue > 0:
                        boundary_ids.add(elem_id.IntegerValue)
    except:
        pass
    return boundary_ids

def get_room_bbox_expanded(room, expand=2.0):
    try:
        bb = room.get_BoundingBox(None)
        if bb:
            return {
                'min_x': bb.Min.X - expand, 'min_y': bb.Min.Y - expand, 'min_z': bb.Min.Z - expand,
                'max_x': bb.Max.X + expand, 'max_y': bb.Max.Y + expand, 'max_z': bb.Max.Z + expand,
            }
    except:
        pass
    return None

def get_elem_bbox(elem):
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            return {
                'min_x': bb.Min.X, 'min_y': bb.Min.Y, 'min_z': bb.Min.Z,
                'max_x': bb.Max.X, 'max_y': bb.Max.Y, 'max_z': bb.Max.Z,
            }
    except:
        pass
    return None

def bbox_intersects(bb1, bb2):
    if not bb1 or not bb2:
        return False
    try:
        if bb1['max_x'] < bb2['min_x'] or bb1['min_x'] > bb2['max_x']:
            return False
        if bb1['max_y'] < bb2['min_y'] or bb1['min_y'] > bb2['max_y']:
            return False
        if bb1['max_z'] < bb2['min_z'] - 5 or bb1['min_z'] > bb2['max_z'] + 5:
            return False
        return True
    except:
        return False

def get_multiple_check_points(elem):
    pts = []
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            cx = (bb.Min.X + bb.Max.X) / 2
            cy = (bb.Min.Y + bb.Max.Y) / 2
            cz = (bb.Min.Z + bb.Max.Z) / 2
            pts.append(XYZ(cx, cy, cz))
            pts.append(XYZ(cx, cy, bb.Min.Z + 0.1))
            pts.append(XYZ(cx, cy, bb.Max.Z - 0.1))
            pts.append(XYZ(bb.Min.X + 0.1, bb.Min.Y + 0.1, cz))
            pts.append(XYZ(bb.Max.X - 0.1, bb.Min.Y + 0.1, cz))
            pts.append(XYZ(bb.Min.X + 0.1, bb.Max.Y - 0.1, cz))
            pts.append(XYZ(bb.Max.X - 0.1, bb.Max.Y - 0.1, cz))
        loc = elem.Location
        if loc:
            if hasattr(loc, 'Point') and loc.Point:
                pts.append(loc.Point)
            elif hasattr(loc, 'Curve') and loc.Curve:
                curve = loc.Curve
                pts.append(curve.GetEndPoint(0))
                pts.append(curve.Evaluate(0.5, True))
                pts.append(curve.GetEndPoint(1))
    except:
        pass
    return pts

def safe_get_location_point(elem):
    try:
        loc = elem.Location
        if loc:
            if hasattr(loc, 'Point') and loc.Point:
                return loc.Point
            elif hasattr(loc, 'Curve') and loc.Curve:
                return loc.Curve.Evaluate(0.5, True)
        bb = elem.get_BoundingBox(None)
        if bb:
            return XYZ((bb.Min.X+bb.Max.X)/2, (bb.Min.Y+bb.Max.Y)/2, (bb.Min.Z+bb.Max.Z)/2)
    except:
        pass
    return None

def in_room(r, pt):
    if not r or not pt: return False
    try: return r.IsPointInRoom(pt)
    except: return False

def in_space(s, pt):
    if not s or not pt: return False
    try: return s.IsPointInSpace(pt)
    except: return False

def in_area(a, pt):
    if not a or not pt: return False
    try:
        opts = SpatialElementBoundaryOptions()
        bounds = a.GetBoundarySegments(opts)
        if not bounds or bounds.Count == 0: return False
        lv = doc.GetElement(a.LevelId)
        if not lv: return False
        if abs(pt.Z - lv.Elevation) > 30: return False
        for b in bounds:
            poly = []
            for seg in b:
                c = seg.GetCurve()
                if c:
                    p = c.GetEndPoint(0)
                    poly.append((p.X, p.Y))
            if poly and len(poly) >= 3:
                if pt_in_poly((pt.X, pt.Y), poly): return True
        return False
    except: return False

def pt_in_poly(pt, poly):
    x, y = pt
    n = len(poly)
    inside = False
    px, py = poly[0]
    for i in range(1, n+1):
        nx, ny = poly[i % n]
        if y > min(py, ny) and y <= max(py, ny) and x <= max(px, nx):
            if py != ny: xi = (y - py) * (nx - px) / (ny - py) + px
            if px == nx or x <= xi: inside = not inside
        px, py = nx, ny
    return inside

def check_element_in_spatial_advanced(elem, sp_item, boundary_ids, use_bbox=False):
    elem_id = elem.Id.IntegerValue
    stype = sp_item.spatial_type
    spatial_elem = sp_item.element
    if elem_id in boundary_ids:
        return True
    if use_bbox:
        room_bb = get_room_bbox_expanded(spatial_elem, 1.0)
        elem_bb = get_elem_bbox(elem)
        if bbox_intersects(room_bb, elem_bb):
            pts = get_multiple_check_points(elem)
            for pt in pts:
                try:
                    inside = False
                    if stype == ROOMS: inside = in_room(spatial_elem, pt)
                    elif stype == AREAS: inside = in_area(spatial_elem, pt)
                    else: inside = in_space(spatial_elem, pt)
                    if inside: return True
                except: continue
        return False
    else:
        pt = safe_get_location_point(elem)
        if not pt: return False
        try:
            if stype == ROOMS: return in_room(spatial_elem, pt)
            elif stype == AREAS: return in_area(spatial_elem, pt)
            else: return in_space(spatial_elem, pt)
        except: return False

def get_rooms(vo=False):
    try:
        c = FilteredElementCollector(doc, doc.ActiveView.Id) if vo else FilteredElementCollector(doc)
        return [r for r in c.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements() if r.Location and r.Area > 0]
    except: return []

def get_areas(vo=False):
    try:
        c = FilteredElementCollector(doc, doc.ActiveView.Id) if vo else FilteredElementCollector(doc)
        return [a for a in c.OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToElements() if a.Location and a.Area > 0]
    except: return []

def get_spaces(vo=False):
    try:
        c = FilteredElementCollector(doc, doc.ActiveView.Id) if vo else FilteredElementCollector(doc)
        return [s for s in c.OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements() if s.Location and s.Area > 0]
    except: return []

def get_cats():
    cats = {}
    bics = [
        BuiltInCategory.OST_Casework, BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Columns,
        BuiltInCategory.OST_CurtainWallPanels, BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_Entourage, BuiltInCategory.OST_Floors, BuiltInCategory.OST_FoodServiceEquipment,
        BuiltInCategory.OST_Furniture, BuiltInCategory.OST_FurnitureSystems, BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_LightingFixtures, BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_Parking,
        BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_Railings, BuiltInCategory.OST_Signage,
        BuiltInCategory.OST_Site, BuiltInCategory.OST_SpecialityEquipment, BuiltInCategory.OST_Sprinklers,
        BuiltInCategory.OST_Stairs, BuiltInCategory.OST_StructuralColumns, BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_Walls, BuiltInCategory.OST_Windows,
    ]
    for bic in bics:
        try:
            elems = list(FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
            if elems:
                cat = doc.Settings.Categories.get_Item(bic)
                if cat: cats[cat.Name] = {"bic": bic, "count": len(elems)}
        except: pass
    return cats

def get_elems(bic, vo=False):
    try:
        c = FilteredElementCollector(doc, doc.ActiveView.Id) if vo else FilteredElementCollector(doc)
        return list(c.OfCategory(bic).WhereElementIsNotElementType().ToElements())
    except: return []

def get_str_params(elem):
    params = []
    try:
        for p in elem.Parameters:
            if not p.IsReadOnly and p.StorageType == StorageType.String:
                params.append((p.Definition.Name, "Shared" if p.IsShared else "Builtin"))
    except: pass
    return sorted(set(params), key=lambda x: x[0])

def get_family_type_name(elem):
    fname = ""
    tname = ""
    try:
        tid = elem.GetTypeId()
        if tid and tid.IntegerValue != -1:
            et = doc.GetElement(tid)
            if et:
                tname = et.Name or ""
                for bip in [BuiltInParameter.ALL_MODEL_FAMILY_NAME, 
                           BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM,
                           BuiltInParameter.ELEM_FAMILY_PARAM]:
                    fp = et.get_Parameter(bip)
                    if fp and fp.HasValue:
                        val = fp.AsString() if fp.StorageType == StorageType.String else ""
                        if val:
                            fname = val
                            break
                if not fname and hasattr(et, 'FamilyName'):
                    fname = et.FamilyName or ""
                if not fname:
                    if hasattr(et, 'Kind'):
                        fname = str(et.Kind)
                    elif hasattr(et, 'GetType'):
                        type_name = et.GetType().Name
                        if type_name == "WallType": fname = "Basic Wall"
                        elif type_name == "FloorType": fname = "Floor"
                        elif type_name == "CeilingType": fname = "Ceiling"
        if not fname:
            if hasattr(elem, 'Symbol') and elem.Symbol:
                symbol = elem.Symbol
                if hasattr(symbol, 'FamilyName'):
                    fname = symbol.FamilyName or ""
                if not tname and hasattr(symbol, 'Name'):
                    tname = symbol.Name or ""
        if not fname:
            cat = elem.Category
            if cat: fname = cat.Name or ""
    except:
        pass
    return fname, tname

class SpatialItem:
    def __init__(self, elem, stype):
        self.element = elem
        self.element_id = elem.Id.IntegerValue
        self.spatial_type = stype
        self.is_selected = False
        self.number = ""
        self.name = ""
        self.level = ""
        self.boundary_ids = set()
        self.all_params = []
        
        try:
            if stype == ROOMS:
                self.number = elem.Number or ""
                p = elem.get_Parameter(BuiltInParameter.ROOM_NAME)
                if p: self.name = p.AsString() or ""
                self.boundary_ids = get_room_boundary_elements(elem)
                self.all_params = get_room_parameters(elem)
            elif stype == AREAS:
                p = elem.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                if p: self.number = p.AsString() or ""
                p = elem.get_Parameter(BuiltInParameter.ROOM_NAME)
                if p: self.name = p.AsString() or ""
                self.all_params = get_room_parameters(elem)
            else:
                self.number = elem.Number or ""
                p = elem.get_Parameter(BuiltInParameter.ROOM_NAME)
                if p: self.name = p.AsString() or ""
                self.all_params = get_room_parameters(elem)
            
            lp = elem.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID)
            if lp:
                lid = lp.AsElementId()
                if lid and lid.IntegerValue != -1:
                    le = doc.GetElement(lid)
                    if le: self.level = le.Name
        except: pass

class CatItem:
    def __init__(self, name, bic, count):
        self.name = name
        self.bic = bic
        self.count = count
        self.is_selected = False

class ResultGroup:
    def __init__(self, cat_name, family_name, type_name, spatial_item):
        self.category_name = cat_name
        self.family_name = family_name
        self.type_name = type_name
        self.spatial_item = spatial_item
        self.elements = []
        self.is_selected = True
    
    @property
    def count(self):
        return len(self.elements)
    
    def add_element(self, elem):
        self.elements.append(elem)
    
    def get_define_value(self, params, separator="_"):
        if not self.spatial_item:
            return ""
        parts = []
        for p in params:
            val = get_param_value(self.spatial_item.element, p)
            if val:
                parts.append(val)
        return separator.join(parts) if parts else ""

class DefineValueDialog(Window):
    """Dialog to configure Define Value parameters"""
    def __init__(self, available_params, current_selected, current_separator):
        self.all_available = available_params
        self.selected_params = list(current_selected)
        self.separator = current_separator
        self.result = None
        
        self.Title = "Configure Define Value"
        self.Width = 580
        self.Height = 450
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.Background = brush(WHITE)
        self.ResizeMode = ResizeMode.NoResize
        self._build()
    
    def _build(self):
        root = StackPanel()
        root.Margin = Thickness(15)
        
        # Header
        hdr = TextBlock()
        hdr.Text = "Please move parameters from left to right list to build custom parameter value."
        hdr.Margin = Thickness(0, 0, 0, 15)
        hdr.TextWrapping = System.Windows.TextWrapping.Wrap
        root.Children.Add(hdr)
        
        # Main content - horizontal layout
        main_panel = StackPanel()
        main_panel.Orientation = Orientation.Horizontal
        main_panel.HorizontalAlignment = HorizontalAlignment.Center
        
        # Left panel - Available
        left_panel = StackPanel()
        left_panel.Width = 180
        left_panel.Margin = Thickness(0, 0, 5, 0)
        
        avail_lbl = TextBlock()
        avail_lbl.Text = "Available"
        avail_lbl.FontWeight = FontWeights.Bold
        avail_lbl.Margin = Thickness(0, 0, 0, 5)
        left_panel.Children.Add(avail_lbl)
        
        avail_scroll = ScrollViewer()
        avail_scroll.Height = 250
        avail_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        avail_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        self.avail_list = ListBox()
        self.avail_list.SelectionMode = SelectionMode.Single
        for p in self.all_available:
            if p not in self.selected_params:
                self.avail_list.Items.Add(p)
        avail_scroll.Content = self.avail_list
        left_panel.Children.Add(avail_scroll)
        
        main_panel.Children.Add(left_panel)
        
        # Arrow buttons panel
        arrow_panel = StackPanel()
        arrow_panel.VerticalAlignment = VerticalAlignment.Center
        arrow_panel.Margin = Thickness(5, 0, 5, 0)
        arrow_panel.Width = 40
        
        btn_right = Button()
        btn_right.Content = ">"
        btn_right.Width = 35
        btn_right.Height = 28
        btn_right.Margin = Thickness(0, 0, 0, 5)
        btn_right.Click += self._move_right
        arrow_panel.Children.Add(btn_right)
        
        btn_left = Button()
        btn_left.Content = "<"
        btn_left.Width = 35
        btn_left.Height = 28
        btn_left.Click += self._move_left
        arrow_panel.Children.Add(btn_left)
        
        main_panel.Children.Add(arrow_panel)
        
        # Right panel - Selected
        right_panel = StackPanel()
        right_panel.Width = 180
        right_panel.Margin = Thickness(5, 0, 5, 0)
        
        sel_lbl = TextBlock()
        sel_lbl.Text = "Selected"
        sel_lbl.FontWeight = FontWeights.Bold
        sel_lbl.Margin = Thickness(0, 0, 0, 5)
        right_panel.Children.Add(sel_lbl)
        
        sel_scroll = ScrollViewer()
        sel_scroll.Height = 250
        sel_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sel_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        self.sel_list = ListBox()
        self.sel_list.SelectionMode = SelectionMode.Single
        for p in self.selected_params:
            self.sel_list.Items.Add(p)
        sel_scroll.Content = self.sel_list
        right_panel.Children.Add(sel_scroll)
        
        main_panel.Children.Add(right_panel)
        
        # Up/Down buttons panel
        updown_panel = StackPanel()
        updown_panel.VerticalAlignment = VerticalAlignment.Center
        updown_panel.Margin = Thickness(5, 0, 0, 0)
        updown_panel.Width = 50
        
        btn_up = Button()
        btn_up.Content = "Up"
        btn_up.Width = 45
        btn_up.Height = 28
        btn_up.Margin = Thickness(0, 0, 0, 5)
        btn_up.Click += self._move_up
        updown_panel.Children.Add(btn_up)
        
        btn_down = Button()
        btn_down.Content = "Down"
        btn_down.Width = 45
        btn_down.Height = 28
        btn_down.Click += self._move_down
        updown_panel.Children.Add(btn_down)
        
        main_panel.Children.Add(updown_panel)
        
        root.Children.Add(main_panel)
        
        # Separator row
        sep_panel = StackPanel()
        sep_panel.Orientation = Orientation.Horizontal
        sep_panel.Margin = Thickness(0, 15, 0, 15)
        sep_panel.HorizontalAlignment = HorizontalAlignment.Left
        
        sep_cb = CheckBox()
        sep_cb.Content = "Field Separator"
        sep_cb.IsChecked = True
        sep_cb.Margin = Thickness(0, 0, 10, 0)
        sep_cb.VerticalAlignment = VerticalAlignment.Center
        sep_panel.Children.Add(sep_cb)
        
        self.sep_combo = ComboBox()
        self.sep_combo.Width = 60
        self.sep_combo.Items.Add("_")
        self.sep_combo.Items.Add("-")
        self.sep_combo.Items.Add(" ")
        self.sep_combo.Items.Add(",")
        self.sep_combo.Items.Add(".")
        idx = 0
        for i, item in enumerate(["_", "-", " ", ",", "."]):
            if item == self.separator:
                idx = i
                break
        self.sep_combo.SelectedIndex = idx
        sep_panel.Children.Add(self.sep_combo)
        
        root.Children.Add(sep_panel)
        
        # Footer buttons
        footer = StackPanel()
        footer.Orientation = Orientation.Horizontal
        footer.HorizontalAlignment = HorizontalAlignment.Right
        footer.Margin = Thickness(0, 10, 0, 0)
        
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 80
        btn_apply.Height = 28
        btn_apply.Margin = Thickness(0, 0, 8, 0)
        btn_apply.Background = brush(SUCCESS)
        btn_apply.Click += self._apply
        footer.Children.Add(btn_apply)
        
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 80
        btn_cancel.Height = 28
        btn_cancel.Click += self._cancel
        footer.Children.Add(btn_cancel)
        
        root.Children.Add(footer)
        
        self.Content = root
    
    def _move_right(self, s, e):
        sel = self.avail_list.SelectedItem
        if sel:
            self.avail_list.Items.Remove(sel)
            self.sel_list.Items.Add(sel)
    
    def _move_left(self, s, e):
        sel = self.sel_list.SelectedItem
        if sel:
            self.sel_list.Items.Remove(sel)
            self.avail_list.Items.Add(sel)
    
    def _move_up(self, s, e):
        idx = self.sel_list.SelectedIndex
        if idx > 0:
            item = self.sel_list.SelectedItem
            self.sel_list.Items.RemoveAt(idx)
            self.sel_list.Items.Insert(idx - 1, item)
            self.sel_list.SelectedIndex = idx - 1
    
    def _move_down(self, s, e):
        idx = self.sel_list.SelectedIndex
        if idx >= 0 and idx < self.sel_list.Items.Count - 1:
            item = self.sel_list.SelectedItem
            self.sel_list.Items.RemoveAt(idx)
            self.sel_list.Items.Insert(idx + 1, item)
            self.sel_list.SelectedIndex = idx + 1
    
    def _apply(self, s, e):
        params = []
        for i in range(self.sel_list.Items.Count):
            params.append(self.sel_list.Items[i])
        sep = self.sep_combo.SelectedItem or "_"
        self.result = {"params": params, "separator": sep}
        self.DialogResult = True
        self.Close()
    
    def _cancel(self, s, e):
        self.DialogResult = False
        self.Close()


class SetParamDialog(Window):
    def __init__(self, result_groups, spatial_type, define_params, separator):
        self.result_groups = result_groups
        self.spatial_type = spatial_type
        self.define_params = define_params
        self.separator = separator
        self.result = None
        self.Title = "Set Parameter Value"
        self.Width = 450
        self.Height = 500
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.Background = brush(WHITE)
        self.ResizeMode = ResizeMode.NoResize
        self._build()
        self._load_params()
    
    def _build(self):
        root = DockPanel()
        root.Margin = Thickness(15)
        root.LastChildFill = True
        
        hdr = TextBlock()
        hdr.Text = "Please select option for elements contained in multiple\n" + self.spatial_type + ", Spaces or Zones."
        hdr.Margin = Thickness(0,0,0,15)
        hdr.TextWrapping = System.Windows.TextWrapping.Wrap
        DockPanel.SetDock(hdr, Dock.Top)
        root.Children.Add(hdr)
        
        opts = StackPanel()
        opts.Margin = Thickness(0,0,0,15)
        DockPanel.SetDock(opts, Dock.Top)
        
        self.rb_comma = RadioButton()
        self.rb_comma.Content = "Set parameters values in comma separated format"
        self.rb_comma.IsChecked = True
        self.rb_comma.Margin = Thickness(0,0,0,6)
        self.rb_comma.GroupName = "Opts"
        opts.Children.Add(self.rb_comma)
        
        self.rb_first = RadioButton()
        self.rb_first.Content = "Set parameters values as the first alphabetically found container"
        self.rb_first.Margin = Thickness(0,0,0,6)
        self.rb_first.GroupName = "Opts"
        opts.Children.Add(self.rb_first)
        
        self.rb_none = RadioButton()
        self.rb_none.Content = "Don't set parameters values"
        self.rb_none.Margin = Thickness(0,0,0,6)
        self.rb_none.GroupName = "Opts"
        opts.Children.Add(self.rb_none)
        
        root.Children.Add(opts)
        
        plbl = TextBlock()
        plbl.Text = "Please select instance parameter to apply created parameter value"
        plbl.Margin = Thickness(0,0,0,8)
        DockPanel.SetDock(plbl, Dock.Top)
        root.Children.Add(plbl)
        
        self.psearch = TextBox()
        self.psearch.Height = 26
        self.psearch.Margin = Thickness(0,0,0,8)
        self.psearch.Text = "Search"
        self.psearch.Foreground = brush(TEXT_MUTED)
        self.psearch.Tag = "Search"
        self.psearch.GotFocus += self._sf
        self.psearch.LostFocus += self._sb
        self.psearch.TextChanged += self._ss
        DockPanel.SetDock(self.psearch, Dock.Top)
        root.Children.Add(self.psearch)
        
        footer = StackPanel()
        footer.Orientation = Orientation.Horizontal
        footer.HorizontalAlignment = HorizontalAlignment.Right
        footer.Margin = Thickness(0,15,0,0)
        DockPanel.SetDock(footer, Dock.Bottom)
        
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 80
        btn_apply.Height = 28
        btn_apply.Margin = Thickness(0,0,8,0)
        btn_apply.Background = brush(SUCCESS)
        btn_apply.Click += self._apply
        footer.Children.Add(btn_apply)
        
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 80
        btn_cancel.Height = 28
        btn_cancel.Click += self._cancel
        footer.Children.Add(btn_cancel)
        
        root.Children.Add(footer)
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.plist = ListBox()
        self.plist.BorderThickness = Thickness(1)
        self.plist.BorderBrush = brush(BORDER)
        scroll.Content = self.plist
        root.Children.Add(scroll)
        
        self.Content = root
    
    def _load_params(self):
        self.all_params = []
        if self.result_groups and len(self.result_groups) > 0 and self.result_groups[0].elements:
            self.all_params = get_str_params(self.result_groups[0].elements[0])
        self._refresh()
    
    def _refresh(self):
        self.plist.Items.Clear()
        search = "" if self.psearch.Text == self.psearch.Tag else self.psearch.Text.lower()
        for pn, pt in self.all_params:
            if search and search not in pn.lower(): continue
            sp = StackPanel()
            sp.Orientation = Orientation.Horizontal
            sp.Tag = pn
            n = TextBlock()
            n.Text = pn
            n.Width = 250
            sp.Children.Add(n)
            t = TextBlock()
            t.Text = pt
            t.Foreground = brush(TEXT_GRAY)
            sp.Children.Add(t)
            self.plist.Items.Add(sp)
        if self.plist.Items.Count > 0:
            self.plist.SelectedIndex = 0
    
    def _sf(self, s, e):
        if s.Text == s.Tag:
            s.Text = ""
            s.Foreground = brush(TEXT_DARK)
    
    def _sb(self, s, e):
        if not s.Text:
            s.Text = s.Tag
            s.Foreground = brush(TEXT_MUTED)
    
    def _ss(self, s, e):
        if s.Text != s.Tag: self._refresh()
    
    def _apply(self, s, e):
        if self.rb_none.IsChecked:
            self.result = {"mode": "none", "param": None}
        else:
            sel = self.plist.SelectedItem
            if not sel:
                TaskDialog.Show("Warning", "Please select a parameter.")
                return
            self.result = {"mode": "comma" if self.rb_comma.IsChecked else "first", "param": sel.Tag}
        self.DialogResult = True
        self.Close()
    
    def _cancel(self, s, e):
        self.DialogResult = False
        self.Close()


class ContainsWindow(Window):
    def __init__(self):
        self.Title = "Contains Manager - pyDQT"
        self.Width = 1300
        self.Height = 750
        self.MinWidth = 1000
        self.MinHeight = 600
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = brush(WHITE)
        self.ResizeMode = ResizeMode.CanResize
        self.spatial_items = []
        self.cat_items = []
        self.result_groups = []
        self.all_groups = []
        self.view_only = False
        self.spatial_type = ROOMS
        self.define_params = ["Number", "Name"]
        self.define_separator = "_"
        self.available_params = []
        self._build()
        self._load()
    
    def _build(self):
        root = DockPanel()
        root.LastChildFill = True
        hdr = self._header()
        DockPanel.SetDock(hdr, Dock.Top)
        root.Children.Add(hdr)
        ftr = self._footer()
        DockPanel.SetDock(ftr, Dock.Bottom)
        root.Children.Add(ftr)
        root.Children.Add(self._content())
        self.Content = root
    
    def _header(self):
        bd = Border()
        bd.Background = brush(PRIMARY)
        bd.Padding = Thickness(20,12,20,12)
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        tp = StackPanel()
        tp.Width = 250
        t = TextBlock()
        t.Text = "Contains Manager"
        t.FontSize = 20
        t.FontWeight = FontWeights.Bold
        t.Foreground = brush(TEXT_DARK)
        tp.Children.Add(t)
        st = TextBlock()
        st.Text = "Find elements in Rooms/Areas/Spaces"
        st.FontSize = 10
        st.Foreground = brush(TEXT_GRAY)
        tp.Children.Add(st)
        sp.Children.Add(tp)
        cards = StackPanel()
        cards.Orientation = Orientation.Horizontal
        cards.Margin = Thickness(30,0,30,0)
        self.card_rooms = self._card("Rooms", "0")
        self.card_areas = self._card("Areas", "0")
        self.card_spaces = self._card("Spaces", "0")
        self.card_found = self._card("Found", "0")
        cards.Children.Add(self.card_rooms)
        cards.Children.Add(self.card_areas)
        cards.Children.Add(self.card_spaces)
        cards.Children.Add(self.card_found)
        sp.Children.Add(cards)
        cp = TextBlock()
        cp.Text = "(c) DQT 2024"
        cp.FontSize = 9
        cp.Foreground = brush(TEXT_GRAY)
        cp.VerticalAlignment = VerticalAlignment.Center
        cp.Margin = Thickness(30,0,0,0)
        sp.Children.Add(cp)
        bd.Child = sp
        return bd
    
    def _card(self, label, val):
        bd = Border()
        bd.Background = brush(WHITE)
        bd.CornerRadius = CornerRadius(4)
        bd.Padding = Thickness(12,5,12,5)
        bd.Margin = Thickness(4,0,4,0)
        bd.BorderBrush = brush(BORDER)
        bd.BorderThickness = Thickness(1)
        sp = StackPanel()
        sp.HorizontalAlignment = HorizontalAlignment.Center
        vt = TextBlock()
        vt.Text = val
        vt.FontSize = 14
        vt.FontWeight = FontWeights.Bold
        vt.Foreground = brush(TEXT_DARK)
        vt.HorizontalAlignment = HorizontalAlignment.Center
        vt.Tag = "value"
        sp.Children.Add(vt)
        lt = TextBlock()
        lt.Text = label
        lt.FontSize = 9
        lt.Foreground = brush(TEXT_GRAY)
        lt.HorizontalAlignment = HorizontalAlignment.Center
        sp.Children.Add(lt)
        bd.Child = sp
        return bd
    
    def _ucard(self, card, val):
        sp = card.Child
        for i in range(sp.Children.Count):
            c = sp.Children[i]
            if hasattr(c, 'Tag') and c.Tag == "value":
                c.Text = str(val)
                break
    
    def _content(self):
        mp = DockPanel()
        mp.Margin = Thickness(10)
        mp.LastChildFill = True
        left = self._spatial_panel()
        DockPanel.SetDock(left, Dock.Left)
        mp.Children.Add(left)
        center = self._cat_panel()
        DockPanel.SetDock(center, Dock.Left)
        mp.Children.Add(center)
        mp.Children.Add(self._results_panel())
        return mp
    
    def _spatial_panel(self):
        bd = Border()
        bd.Background = brush(WHITE)
        bd.CornerRadius = CornerRadius(6)
        bd.BorderBrush = brush(BORDER)
        bd.BorderThickness = Thickness(1)
        bd.Margin = Thickness(0,0,8,0)
        bd.Padding = Thickness(10)
        bd.Width = 280
        sp = StackPanel()
        hdr = TextBlock()
        hdr.Text = "Select Spatial Elements"
        hdr.FontSize = 12
        hdr.FontWeight = FontWeights.Bold
        hdr.Margin = Thickness(0,0,0,8)
        sp.Children.Add(hdr)
        scope = StackPanel()
        scope.Orientation = Orientation.Horizontal
        scope.Margin = Thickness(0,0,0,8)
        self.rb_whole = RadioButton()
        self.rb_whole.Content = "Whole Model"
        self.rb_whole.IsChecked = True
        self.rb_whole.Margin = Thickness(0,0,12,0)
        self.rb_whole.GroupName = "Scope"
        self.rb_whole.Checked += self._scope
        scope.Children.Add(self.rb_whole)
        self.rb_view = RadioButton()
        self.rb_view.Content = "Active View"
        self.rb_view.GroupName = "Scope"
        self.rb_view.Checked += self._scope
        scope.Children.Add(self.rb_view)
        sp.Children.Add(scope)
        lb = TextBlock()
        lb.Text = "Spatial Type:"
        lb.Margin = Thickness(0,0,0,4)
        sp.Children.Add(lb)
        self.type_cb = ComboBox()
        self.type_cb.Margin = Thickness(0,0,0,8)
        self.type_cb.Items.Add(ROOMS)
        self.type_cb.Items.Add(AREAS)
        self.type_cb.Items.Add(SPACES)
        self.type_cb.SelectedIndex = 0
        self.type_cb.SelectionChanged += self._type
        sp.Children.Add(self.type_cb)
        self.sp_search = TextBox()
        self.sp_search.Height = 24
        self.sp_search.Margin = Thickness(0,0,0,6)
        self.sp_search.Text = "Search..."
        self.sp_search.Foreground = brush(TEXT_MUTED)
        self.sp_search.Tag = "Search..."
        self.sp_search.GotFocus += self._sf
        self.sp_search.LostFocus += self._sb
        self.sp_search.TextChanged += self._sp_ss
        sp.Children.Add(self.sp_search)
        scroll = ScrollViewer()
        scroll.Height = 280
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.sp_list = ListBox()
        self.sp_list.SelectionMode = SelectionMode.Extended
        self.sp_list.BorderThickness = Thickness(1)
        self.sp_list.BorderBrush = brush(BORDER)
        scroll.Content = self.sp_list
        sp.Children.Add(scroll)
        bp = StackPanel()
        bp.Orientation = Orientation.Horizontal
        bp.Margin = Thickness(0,6,0,0)
        ba = Button()
        ba.Content = "Select All"
        ba.Width = 75
        ba.Height = 24
        ba.Margin = Thickness(0,0,4,0)
        ba.Click += self._sp_all
        bp.Children.Add(ba)
        bc = Button()
        bc.Content = "Clear"
        bc.Width = 75
        bc.Height = 24
        bc.Click += self._sp_clr
        bp.Children.Add(bc)
        sp.Children.Add(bp)
        self.sp_hide = CheckBox()
        self.sp_hide.Content = "Hide unchecked"
        self.sp_hide.Margin = Thickness(0,6,0,0)
        self.sp_hide.Checked += self._sp_h
        self.sp_hide.Unchecked += self._sp_h
        sp.Children.Add(self.sp_hide)
        bd.Child = sp
        return bd
    
    def _cat_panel(self):
        bd = Border()
        bd.Background = brush(WHITE)
        bd.CornerRadius = CornerRadius(6)
        bd.BorderBrush = brush(BORDER)
        bd.BorderThickness = Thickness(1)
        bd.Margin = Thickness(0,0,8,0)
        bd.Padding = Thickness(10)
        bd.Width = 220
        sp = StackPanel()
        hdr = TextBlock()
        hdr.Text = "Select Category"
        hdr.FontSize = 12
        hdr.FontWeight = FontWeights.Bold
        hdr.Margin = Thickness(0,0,0,8)
        sp.Children.Add(hdr)
        self.cat_search = TextBox()
        self.cat_search.Height = 24
        self.cat_search.Margin = Thickness(0,0,0,6)
        self.cat_search.Text = "Search..."
        self.cat_search.Foreground = brush(TEXT_MUTED)
        self.cat_search.Tag = "Search..."
        self.cat_search.GotFocus += self._sf
        self.cat_search.LostFocus += self._sb
        self.cat_search.TextChanged += self._cat_ss
        sp.Children.Add(self.cat_search)
        scroll = ScrollViewer()
        scroll.Height = 350
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.cat_list = ListBox()
        self.cat_list.BorderThickness = Thickness(1)
        self.cat_list.BorderBrush = brush(BORDER)
        scroll.Content = self.cat_list
        sp.Children.Add(scroll)
        bp = StackPanel()
        bp.Orientation = Orientation.Horizontal
        bp.Margin = Thickness(0,6,0,0)
        ba = Button()
        ba.Content = "Select All"
        ba.Width = 70
        ba.Height = 24
        ba.Margin = Thickness(0,0,4,0)
        ba.Click += self._cat_all
        bp.Children.Add(ba)
        bc = Button()
        bc.Content = "Clear"
        bc.Width = 70
        bc.Height = 24
        bc.Click += self._cat_clr
        bp.Children.Add(bc)
        sp.Children.Add(bp)
        self.cat_hide = CheckBox()
        self.cat_hide.Content = "Hide unchecked"
        self.cat_hide.Margin = Thickness(0,6,0,0)
        self.cat_hide.Checked += self._cat_h
        self.cat_hide.Unchecked += self._cat_h
        sp.Children.Add(self.cat_hide)
        bd.Child = sp
        return bd
    
    def _results_panel(self):
        bd = Border()
        bd.Background = brush(WHITE)
        bd.CornerRadius = CornerRadius(6)
        bd.BorderBrush = brush(BORDER)
        bd.BorderThickness = Thickness(1)
        bd.Padding = Thickness(10)
        mp = DockPanel()
        mp.LastChildFill = True
        top = StackPanel()
        DockPanel.SetDock(top, Dock.Top)
        hr = StackPanel()
        hr.Orientation = Orientation.Horizontal
        hr.Margin = Thickness(0,0,0,6)
        h = TextBlock()
        h.Text = "Results"
        h.FontSize = 12
        h.FontWeight = FontWeights.Bold
        h.Width = 100
        hr.Children.Add(h)
        self.res_search = TextBox()
        self.res_search.Height = 24
        self.res_search.Width = 150
        self.res_search.Text = "Search..."
        self.res_search.Foreground = brush(TEXT_MUTED)
        self.res_search.Tag = "Search..."
        self.res_search.GotFocus += self._sf
        self.res_search.LostFocus += self._sb
        self.res_search.TextChanged += self._res_ss
        hr.Children.Add(self.res_search)
        btn_cfg = Button()
        btn_cfg.Content = "Configure Value..."
        btn_cfg.Height = 24
        btn_cfg.Margin = Thickness(10, 0, 0, 0)
        btn_cfg.Padding = Thickness(8, 0, 8, 0)
        btn_cfg.Click += self._cfg_define
        hr.Children.Add(btn_cfg)
        top.Children.Add(hr)
        col_hdr = StackPanel()
        col_hdr.Orientation = Orientation.Horizontal
        col_hdr.Margin = Thickness(0,4,0,4)
        col_hdr.Background = brush(SECONDARY)
        cb_all = CheckBox()
        cb_all.Width = 25
        cb_all.IsChecked = True
        cb_all.Checked += self._res_all_ck
        cb_all.Unchecked += self._res_all_uck
        col_hdr.Children.Add(cb_all)
        cols = [("Category", 100), ("Family Name", 150), ("Type Name", 180), ("Define Value", 120), ("Count", 50)]
        for txt, w in cols:
            tb = TextBlock()
            tb.Text = txt
            tb.Width = w
            tb.FontWeight = FontWeights.Bold
            tb.FontSize = 11
            if txt == "Define Value":
                tb.Foreground = brush(SUCCESS)
            col_hdr.Children.Add(tb)
        top.Children.Add(col_hdr)
        mp.Children.Add(top)
        bot = StackPanel()
        DockPanel.SetDock(bot, Dock.Bottom)
        self.status = TextBlock()
        self.status.Text = "Total number of elements found 0 | Selected 0"
        self.status.Margin = Thickness(0,6,0,0)
        self.status.Foreground = brush(TEXT_GRAY)
        bot.Children.Add(self.status)
        mp.Children.Add(bot)
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        self.res_list = ListBox()
        self.res_list.SelectionMode = SelectionMode.Extended
        self.res_list.BorderThickness = Thickness(1)
        self.res_list.BorderBrush = brush(BORDER)
        scroll.Content = self.res_list
        mp.Children.Add(scroll)
        bd.Child = mp
        return bd
    
    def _footer(self):
        bd = Border()
        bd.Background = brush(SECONDARY)
        bd.Padding = Thickness(15,8,15,8)
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.HorizontalAlignment = HorizontalAlignment.Right
        sp.Children.Add(self._btn("Reset", WHITE, self._reset, True))
        sp.Children.Add(self._btn("Visualize", "#9FC5E8", self._viz, False))
        sp.Children.Add(self._btn("Find", SUCCESS, self._find, False))
        sp.Children.Add(self._btn("Set Parameter Value", PRIMARY, self._set, False))
        sp.Children.Add(self._btn("Select", "#FFAB91", self._sel, False))
        sp.Children.Add(self._btn("Close", WHITE, self._close, True))
        bd.Child = sp
        return bd
    
    def _btn(self, txt, bg, handler, border):
        b = Button()
        b.Content = txt
        b.MinWidth = 85
        b.Height = 28
        b.Padding = Thickness(10,0,10,0)
        b.Margin = Thickness(3,0,3,0)
        b.Background = brush(bg)
        b.Foreground = brush(TEXT_DARK)
        if border:
            b.BorderBrush = brush(BORDER)
            b.BorderThickness = Thickness(1)
        else:
            b.BorderThickness = Thickness(0)
        b.Click += handler
        return b
    
    def _load(self):
        self._load_sp()
        self._load_cat()
        self._ucards()
        self._update_available_params()
    
    def _update_available_params(self):
        self.available_params = []
        if self.spatial_items:
            self.available_params = self.spatial_items[0].all_params
    
    def _load_sp(self):
        self.spatial_items = []
        if self.spatial_type == ROOMS:
            for e in get_rooms(self.view_only): self.spatial_items.append(SpatialItem(e, ROOMS))
        elif self.spatial_type == AREAS:
            for e in get_areas(self.view_only): self.spatial_items.append(SpatialItem(e, AREAS))
        else:
            for e in get_spaces(self.view_only): self.spatial_items.append(SpatialItem(e, SPACES))
        self._ref_sp()
        self._update_available_params()
    
    def _ref_sp(self):
        self.sp_list.Items.Clear()
        search = "" if self.sp_search.Text == self.sp_search.Tag else self.sp_search.Text.lower()
        hide = self.sp_hide.IsChecked
        for item in self.spatial_items:
            if search and search not in item.number.lower() and search not in item.name.lower(): continue
            if hide and not item.is_selected: continue
            self.sp_list.Items.Add(self._mk_sp(item))
    
    def _mk_sp(self, item):
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Tag = item
        cb = CheckBox()
        cb.IsChecked = item.is_selected
        cb.Margin = Thickness(0,0,6,0)
        cb.Tag = item
        cb.Checked += self._sp_ck
        cb.Unchecked += self._sp_uck
        sp.Children.Add(cb)
        n = TextBlock()
        n.Text = item.number
        n.Width = 50
        sp.Children.Add(n)
        nm = TextBlock()
        nm.Text = item.name
        nm.Width = 100
        sp.Children.Add(nm)
        lv = TextBlock()
        lv.Text = item.level
        lv.Foreground = brush(TEXT_GRAY)
        sp.Children.Add(lv)
        return sp
    
    def _load_cat(self):
        self.cat_items = []
        for name, data in get_cats().items():
            self.cat_items.append(CatItem(name, data["bic"], data["count"]))
        self.cat_items.sort(key=lambda x: x.name)
        self._ref_cat()
    
    def _ref_cat(self):
        self.cat_list.Items.Clear()
        search = "" if self.cat_search.Text == self.cat_search.Tag else self.cat_search.Text.lower()
        hide = self.cat_hide.IsChecked
        for item in self.cat_items:
            if search and search not in item.name.lower(): continue
            if hide and not item.is_selected: continue
            self.cat_list.Items.Add(self._mk_cat(item))
    
    def _mk_cat(self, item):
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Tag = item
        cb = CheckBox()
        cb.IsChecked = item.is_selected
        cb.Margin = Thickness(0,0,6,0)
        cb.Tag = item
        cb.Checked += self._cat_ck
        cb.Unchecked += self._cat_uck
        sp.Children.Add(cb)
        n = TextBlock()
        n.Text = item.name
        n.Width = 110
        sp.Children.Add(n)
        c = TextBlock()
        c.Text = "(" + str(item.count) + ")"
        c.Foreground = brush(TEXT_GRAY)
        sp.Children.Add(c)
        return sp
    
    def _ucards(self):
        self._ucard(self.card_rooms, len(get_rooms(self.view_only)))
        self._ucard(self.card_areas, len(get_areas(self.view_only)))
        self._ucard(self.card_spaces, len(get_spaces(self.view_only)))
        total = sum([g.count for g in self.result_groups])
        self._ucard(self.card_found, total)
    
    def _ref_res(self):
        self.res_list.Items.Clear()
        search = "" if self.res_search.Text == self.res_search.Tag else self.res_search.Text.lower()
        displayed = []
        for grp in self.all_groups:
            if search:
                dv = grp.get_define_value(self.define_params, self.define_separator)
                if search not in grp.category_name.lower() and search not in grp.family_name.lower() and search not in grp.type_name.lower() and search not in dv.lower():
                    continue
            displayed.append(grp)
            self.res_list.Items.Add(self._mk_res(grp))
        self.result_groups = displayed
        total = sum([g.count for g in self.all_groups])
        sel = sum([g.count for g in displayed if g.is_selected])
        self.status.Text = "Total number of elements found " + str(total) + " | Selected " + str(sel)
    
    def _mk_res(self, grp):
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Tag = grp
        if grp.is_selected: sp.Background = brush(HIGHLIGHT)
        cb = CheckBox()
        cb.IsChecked = grp.is_selected
        cb.Width = 25
        cb.Tag = grp
        cb.Checked += self._res_ck
        cb.Unchecked += self._res_uck
        sp.Children.Add(cb)
        cat = TextBlock()
        cat.Text = grp.category_name
        cat.Width = 100
        sp.Children.Add(cat)
        fam = TextBlock()
        fam.Text = grp.family_name
        fam.Width = 150
        sp.Children.Add(fam)
        typ = TextBlock()
        typ.Text = grp.type_name
        typ.Width = 180
        sp.Children.Add(typ)
        dv = TextBlock()
        dv.Text = grp.get_define_value(self.define_params, self.define_separator)
        dv.Width = 120
        dv.Foreground = brush(SUCCESS)
        dv.FontWeight = FontWeights.Bold
        sp.Children.Add(dv)
        cnt = TextBlock()
        cnt.Text = str(grp.count)
        cnt.Width = 50
        cnt.Foreground = brush(TEXT_GRAY)
        sp.Children.Add(cnt)
        return sp
    
    def _sf(self, s, e):
        if s.Text == s.Tag:
            s.Text = ""
            s.Foreground = brush(TEXT_DARK)
    
    def _sb(self, s, e):
        if not s.Text:
            s.Text = s.Tag
            s.Foreground = brush(TEXT_MUTED)
    
    def _scope(self, s, e):
        self.view_only = self.rb_view.IsChecked
        self._load_sp()
        self._ucards()
    
    def _type(self, s, e):
        self.spatial_type = self.type_cb.SelectedItem
        self._load_sp()
    
    def _sp_ss(self, s, e):
        if s.Text != s.Tag: self._ref_sp()
    def _cat_ss(self, s, e):
        if s.Text != s.Tag: self._ref_cat()
    def _res_ss(self, s, e):
        self._ref_res()
    def _sp_ck(self, s, e): s.Tag.is_selected = True
    def _sp_uck(self, s, e): s.Tag.is_selected = False
    def _cat_ck(self, s, e): s.Tag.is_selected = True
    def _cat_uck(self, s, e): s.Tag.is_selected = False
    def _sp_h(self, s, e): self._ref_sp()
    def _cat_h(self, s, e): self._ref_cat()
    
    def _res_ck(self, s, e):
        s.Tag.is_selected = True
        s.Parent.Background = brush(HIGHLIGHT)
        self._update_status()
    
    def _res_uck(self, s, e):
        s.Tag.is_selected = False
        s.Parent.Background = brush(WHITE)
        self._update_status()
    
    def _res_all_ck(self, s, e):
        for grp in self.result_groups: grp.is_selected = True
        self._ref_res()
    
    def _res_all_uck(self, s, e):
        for grp in self.result_groups: grp.is_selected = False
        self._ref_res()
    
    def _update_status(self):
        total = sum([g.count for g in self.all_groups])
        sel = sum([g.count for g in self.result_groups if g.is_selected])
        self.status.Text = "Total number of elements found " + str(total) + " | Selected " + str(sel)
    
    def _sp_all(self, s, e):
        for i in range(self.sp_list.Items.Count):
            p = self.sp_list.Items[i]
            p.Tag.is_selected = True
            for j in range(p.Children.Count):
                c = p.Children[j]
                if str(type(c).__name__) == "CheckBox":
                    c.IsChecked = True
                    break
    
    def _sp_clr(self, s, e):
        for item in self.spatial_items: item.is_selected = False
        self._ref_sp()
    
    def _cat_all(self, s, e):
        for i in range(self.cat_list.Items.Count):
            p = self.cat_list.Items[i]
            p.Tag.is_selected = True
            for j in range(p.Children.Count):
                c = p.Children[j]
                if str(type(c).__name__) == "CheckBox":
                    c.IsChecked = True
                    break
    
    def _cat_clr(self, s, e):
        for item in self.cat_items: item.is_selected = False
        self._ref_cat()
    
    def _reset(self, s, e):
        for item in self.spatial_items: item.is_selected = False
        for item in self.cat_items: item.is_selected = False
        self.result_groups = []
        self.all_groups = []
        self.sp_search.Text = self.sp_search.Tag
        self.sp_search.Foreground = brush(TEXT_MUTED)
        self.cat_search.Text = self.cat_search.Tag
        self.cat_search.Foreground = brush(TEXT_MUTED)
        self.res_search.Text = self.res_search.Tag
        self.res_search.Foreground = brush(TEXT_MUTED)
        self.sp_hide.IsChecked = False
        self.cat_hide.IsChecked = False
        self._ref_sp()
        self._ref_cat()
        self._ref_res()
        self._ucards()
    
    def _viz(self, s, e):
        sel = [i for i in self.spatial_items if i.is_selected]
        if not sel:
            TaskDialog.Show("Visualize", "Please select spatial element(s) first.")
            return
        ids = List[ElementId]()
        for i in sel: ids.Add(i.element.Id)
        uidoc.Selection.SetElementIds(ids)
        TaskDialog.Show("Visualize", "Selected " + str(len(sel)) + " element(s).")
    
    def _cfg_define(self, s, e):
        dlg = DefineValueDialog(self.available_params, self.define_params, self.define_separator)
        dlg.Owner = self
        result = dlg.ShowDialog()
        if result and dlg.result:
            self.define_params = dlg.result["params"]
            self.define_separator = dlg.result["separator"]
            self._ref_res()
    
    def _find(self, s, e):
        sel_sp = [i for i in self.spatial_items if i.is_selected]
        if not sel_sp:
            TaskDialog.Show("Find", "Please select spatial element(s) first.")
            return
        sel_cat = [i for i in self.cat_items if i.is_selected]
        if not sel_cat:
            TaskDialog.Show("Find", "Please select category(ies) first.")
            return
        
        self.result_groups = []
        self.all_groups = []
        groups = {}
        
        try:
            for cat in sel_cat:
                use_bbox = cat.bic in BBOX_CATEGORIES
                elems = get_elems(cat.bic, self.view_only)
                
                for elem in elems:
                    try:
                        fname, tname = get_family_type_name(elem)
                        
                        for sp_item in sel_sp:
                            try:
                                inside = check_element_in_spatial_advanced(elem, sp_item, sp_item.boundary_ids, use_bbox)
                                if inside:
                                    key = (cat.name, fname, tname, sp_item.element_id)
                                    if key not in groups:
                                        groups[key] = ResultGroup(cat.name, fname, tname, sp_item)
                                    groups[key].add_element(elem)
                                    break
                            except: continue
                    except: continue
            
            self.all_groups = list(groups.values())
            self.result_groups = list(groups.values())
            self._ref_res()
            self._ucards()
            
            total = sum([g.count for g in self.all_groups])
            TaskDialog.Show("Find Results", "Found " + str(total) + " element(s) in " + str(len(self.all_groups)) + " groups.")
        except Exception as ex:
            TaskDialog.Show("Error", "Error during find: " + str(ex))
    
    def _set(self, s, e):
        sel_groups = [g for g in self.result_groups if g.is_selected]
        if not sel_groups:
            TaskDialog.Show("Set Parameter", "No elements selected.")
            return
        dlg = SetParamDialog(sel_groups, self.spatial_type, self.define_params, self.define_separator)
        dlg.Owner = self
        result = dlg.ShowDialog()
        if not result or not dlg.result: return
        mode = dlg.result["mode"]
        param = dlg.result["param"]
        if mode == "none": return
        t = Transaction(doc, "Set Parameters - Contains Manager")
        t.Start()
        try:
            ok, fail = 0, 0
            for grp in sel_groups:
                val = grp.get_define_value(self.define_params, self.define_separator)
                for elem in grp.elements:
                    try:
                        p = elem.LookupParameter(param)
                        if p and not p.IsReadOnly:
                            p.Set(val)
                            ok += 1
                        else: fail += 1
                    except: fail += 1
            t.Commit()
            TaskDialog.Show("Set Parameter", "Updated: " + str(ok) + "\nFailed: " + str(fail))
        except Exception as ex:
            t.RollBack()
            TaskDialog.Show("Error", str(ex))
    
    def _sel(self, s, e):
        sel_groups = [g for g in self.result_groups if g.is_selected]
        if not sel_groups:
            TaskDialog.Show("Select", "No elements selected.")
            return
        ids = List[ElementId]()
        for grp in sel_groups:
            for elem in grp.elements: ids.Add(elem.Id)
        uidoc.Selection.SetElementIds(ids)
        TaskDialog.Show("Select", "Selected " + str(ids.Count) + " element(s).")
    
    def _close(self, s, e):
        self.Close()

def main():
    try:
        win = ContainsWindow()
        win.ShowDialog()
    except Exception as ex:
        TaskDialog.Show("Error", str(ex))

if __name__ == "__main__":
    main()