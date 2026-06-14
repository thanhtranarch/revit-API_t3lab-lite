# -*- coding: utf-8 -*-
"""
Room To Area v1.1 - DQT
Automatically creates Areas from selected Rooms.
Creates area boundary lines matching room boundaries and places area elements.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.

Author: Dang Quoc Truong (DQT) & T3Lab
Website: https://github.com/DangQuocTruong
License: All rights reserved - pyDQT Suite
"""

__title__ = "Room to Area"
__author__ = "Dang Quoc Truong & T3Lab"
__doc__ = "Auto-create Areas from selected Rooms with boundary lines.\nCollaborative tool by T3Lab & Dang Quoc Truong."

import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Collections.Generic import List
from System.Windows import (
    Window, WindowStartupLocation, Thickness, CornerRadius,
    HorizontalAlignment, VerticalAlignment, Visibility,
    GridLength, GridUnitType, FontWeights, TextWrapping,
    RoutedEventArgs
)
from System.Windows.Controls import *
from System.Windows.Controls.Primitives import *
from System.Windows.Media import BrushConverter
from System.Windows.Data import Binding

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, TransactionGroup,
    ElementId, XYZ, UV, Line, CurveLoop,
    SpatialElementBoundaryOptions, SpatialElementBoundaryLocation,
    AreaScheme, ViewPlan, ViewFamilyType, ViewFamily,
    Area, Level
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI import TaskDialog

import System
import math

# ============================================================================
# REVIT CONTEXT
# ============================================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

# ============================================================================
# CONFIGURATION (Slate/Gray Minimalism)
# ============================================================================
class Config:
    PRIMARY = "#0F172A"       # Slate-900
    SECONDARY = "#F8FAFC"     # Slate-50
    ACCENT = "#3B82F6"        # Refined Blue
    BACKGROUND = "#F8FAFC"    # Slate-50
    HEADER_BG = "#0F172A"     # Slate-900
    CARD_BG = "#FFFFFF"       # White
    BORDER = "#CBD5E1"        # Slate-200
    TEXT_DARK = "#0F172A"     # Slate-900
    TEXT_MEDIUM = "#475569"   # Slate-600
    TEXT_LIGHT = "#94A3B8"    # Slate-400
    SUCCESS = "#10B981"       # Emerald-500
    WARNING = "#F59E0B"       # Amber-500
    ERROR = "#EF4444"         # Red-500
    INFO = "#3B82F6"          # Blue-500
    ROW_ALT = "#F1F5F9"       # Slate-100

bc = BrushConverter()

# ============================================================================
# ROOM ITEM DATA MODEL
# ============================================================================
class RoomItem(object):
    """Data model for a room displayed in the DataGrid"""
    
    def __init__(self, room):
        self.element = room
        self.id = room.Id.IntegerValue
        self.is_checked = True
        self.check_display = "[x]"
        self.name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or "Unnamed"
        self.number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or ""
        self.level = self._get_level_name(room)
        self.area_sqm = self._get_area(room)
        self.area_display = str(round(self.area_sqm, 2)) if self.area_sqm > 0 else "Not Enclosed"
        self.department = self._get_department(room)
        self.status = "Ready"
        self.param_value = ""  # Dynamic parameter column value
        self._param_cache = {}  # Cache all parameter values
        self._build_param_cache(room)
    
    def _build_param_cache(self, room):
        """Cache all parameter names and string values"""
        try:
            for param in room.Parameters:
                try:
                    pname = param.Definition.Name
                    if param.HasValue:
                        if param.StorageType.ToString() == "String":
                            val = param.AsString() or ""
                        elif param.StorageType.ToString() == "Double":
                            val = str(round(param.AsDouble(), 4))
                        elif param.StorageType.ToString() == "Integer":
                            val = str(param.AsInteger())
                        elif param.StorageType.ToString() == "ElementId":
                            val = param.AsValueString() or str(param.AsElementId().IntegerValue)
                        else:
                            val = param.AsValueString() or ""
                    else:
                        val = ""
                    self._param_cache[pname] = val
                except:
                    pass
        except:
            pass
    
    def get_param_value(self, param_name):
        """Get cached parameter value by name"""
        return self._param_cache.get(param_name, "")
    
    def _get_level_name(self, room):
        try:
            level = doc.GetElement(room.LevelId)
            return level.Name if level else "N/A"
        except:
            return "N/A"
    
    def _get_area(self, room):
        try:
            area_param = room.get_Parameter(BuiltInParameter.ROOM_AREA)
            if area_param and area_param.HasValue:
                # Convert from sq ft to sq m
                return area_param.AsDouble() * 0.092903
            return 0.0
        except:
            return 0.0
    
    def _get_department(self, room):
        try:
            dept_param = room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT)
            if dept_param and dept_param.HasValue:
                return dept_param.AsString() or ""
            return ""
        except:
            return ""


# ============================================================================
# AREA CREATION ENGINE
# ============================================================================
class AreaCreationEngine(object):
    """Core engine for creating Areas from Rooms"""
    
    def __init__(self, doc, area_scheme, area_plan_view):
        self.doc = doc
        self.area_scheme = area_scheme
        self.area_plan_view = area_plan_view
    
    def get_room_boundaries(self, room):
        """Get boundary curves from a room"""
        options = SpatialElementBoundaryOptions()
        options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
        
        boundary_segments = room.GetBoundarySegments(options)
        if not boundary_segments:
            return None
        
        all_loops = []
        for seg_list in boundary_segments:
            curves = []
            for seg in seg_list:
                curve = seg.GetCurve()
                if curve:
                    curves.append(curve)
            if curves:
                all_loops.append(curves)
        
        return all_loops if all_loops else None
    
    def create_area_boundary_lines(self, view, curves):
        """Create area boundary lines in the given view from curve list"""
        created_lines = []
        for curve in curves:
            try:
                # Create area boundary line using the sketch plane of the view
                new_line = self.doc.Create.NewAreaBoundaryLine(
                    view.SketchPlane, curve, view
                )
                if new_line:
                    created_lines.append(new_line)
            except Exception as ex:
                # Try creating as a simple model line alternative
                pass
        return created_lines
    
    def get_room_center_point(self, room):
        """Get the center point (location) of a room"""
        try:
            location = room.Location
            if location:
                return location.Point
        except:
            pass
        
        # Fallback: calculate from boundaries
        try:
            options = SpatialElementBoundaryOptions()
            boundary_segments = room.GetBoundarySegments(options)
            if boundary_segments and len(boundary_segments) > 0:
                pts = []
                for seg in boundary_segments[0]:
                    curve = seg.GetCurve()
                    pts.append(curve.GetEndPoint(0))
                if pts:
                    avg_x = sum(p.X for p in pts) / len(pts)
                    avg_y = sum(p.Y for p in pts) / len(pts)
                    avg_z = sum(p.Z for p in pts) / len(pts)
                    return XYZ(avg_x, avg_y, avg_z)
        except:
            pass
        return None
    
    def create_area_from_room(self, room, transfer_params=None):
        """Create an Area from a Room: boundary lines + area placement
        
        Args:
            room: Revit Room element
            transfer_params: list of parameter names to copy from Room to Area
        """
        result = {
            'success': False,
            'message': '',
            'area_id': None,
            'boundary_count': 0
        }
        
        if transfer_params is None:
            transfer_params = ["Name", "Number"]
        
        # 1. Get room boundaries
        boundary_loops = self.get_room_boundaries(room)
        if not boundary_loops:
            result['message'] = "No boundary found for room"
            return result
        
        # 2. Create area boundary lines for the outer loop
        total_lines = 0
        for loop_curves in boundary_loops:
            lines = self.create_area_boundary_lines(self.area_plan_view, loop_curves)
            total_lines += len(lines)
        
        result['boundary_count'] = total_lines
        
        if total_lines == 0:
            result['message'] = "Failed to create boundary lines"
            return result
        
        # 3. Place area at room center
        center = self.get_room_center_point(room)
        if not center:
            result['message'] = "Cannot determine room center point"
            return result
        
        try:
            # Use UV point for area placement (XY plane)
            uv_point = UV(center.X, center.Y)
            new_area = self.doc.Create.NewArea(self.area_plan_view, uv_point)
            
            if new_area:
                # Transfer selected parameters from Room to Area
                transferred = 0
                for param_name in transfer_params:
                    try:
                        room_param = room.LookupParameter(param_name)
                        if not room_param or not room_param.HasValue:
                            continue
                        
                        area_param = new_area.LookupParameter(param_name)
                        if not area_param or area_param.IsReadOnly:
                            continue
                        
                        # Match storage types
                        if room_param.StorageType != area_param.StorageType:
                            continue
                        
                        storage = room_param.StorageType.ToString()
                        if storage == "String":
                            val = room_param.AsString()
                            if val:
                                area_param.Set(val)
                                transferred += 1
                        elif storage == "Double":
                            area_param.Set(room_param.AsDouble())
                            transferred += 1
                        elif storage == "Integer":
                            area_param.Set(room_param.AsInteger())
                            transferred += 1
                        elif storage == "ElementId":
                            area_param.Set(room_param.AsElementId())
                            transferred += 1
                    except:
                        pass
                
                result['success'] = True
                result['area_id'] = new_area.Id.IntegerValue
                result['message'] = "Created (ID: {}, {} params)".format(
                    new_area.Id.IntegerValue, transferred)
            else:
                result['message'] = "Area placement returned null"
        except Exception as ex:
            result['message'] = "Area placement failed: {}".format(str(ex))
        
        return result


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
def get_all_rooms(selected_only=False):
    """Get rooms from the model - either all or selected"""
    rooms = []
    
    if selected_only:
        selection = uidoc.Selection
        sel_ids = selection.GetElementIds()
        for eid in sel_ids:
            elem = doc.GetElement(eid)
            if elem and isinstance(elem, Room):
                # Only include placed, enclosed rooms
                if elem.Area > 0:
                    rooms.append(elem)
    else:
        collector = FilteredElementCollector(doc).OfCategory(
            BuiltInCategory.OST_Rooms
        ).WhereElementIsNotElementType()
        
        for room in collector:
            if isinstance(room, Room) and room.Area > 0:
                rooms.append(room)
    
    return rooms


def get_area_schemes():
    """Get all area schemes in the document"""
    collector = FilteredElementCollector(doc).OfClass(AreaScheme)
    schemes = []
    for scheme in collector:
        schemes.append(scheme)
    return schemes


def get_area_plan_views(scheme_id=None):
    """Get area plan views, optionally filtered by area scheme"""
    collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    views = []
    for view in collector:
        if not view.IsTemplate and view.AreaScheme is not None:
            if scheme_id is None or view.AreaScheme.Id == scheme_id:
                views.append(view)
    return views


def find_or_create_area_plan(scheme, level):
    """Find existing area plan or create a new one for the given scheme and level"""
    # Search existing area plans
    existing_views = get_area_plan_views(scheme.Id)
    for view in existing_views:
        if view.GenLevel and view.GenLevel.Id == level.Id:
            return view
    
    # Create new area plan
    try:
        # Find ViewFamilyType for AreaPlan
        vft_collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
        area_vft = None
        for vft in vft_collector:
            if vft.ViewFamily == ViewFamily.AreaPlan:
                area_vft = vft
                break
        
        if area_vft:
            new_view = ViewPlan.CreateAreaPlan(doc, scheme.Id, level.Id)
            if new_view:
                return new_view
    except Exception as ex:
        pass
    
    return None


def get_levels_from_rooms(rooms):
    """Get unique levels from a list of rooms"""
    level_dict = {}
    for room in rooms:
        try:
            level = doc.GetElement(room.LevelId)
            if level and level.Id.IntegerValue not in level_dict:
                level_dict[level.Id.IntegerValue] = level
        except:
            pass
    return level_dict.values()


# ============================================================================
# MAIN WINDOW
# ============================================================================
class RoomToAreaWindow(Window):
    """Main WPF Window for Room To Area tool"""
    
    def __init__(self):
        self.Title = "Room To Area - DQT"
        self.Width = 1050
        self.Height = 720
        self.MinWidth = 850
        self.MinHeight = 550
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = bc.ConvertFromString(Config.SECONDARY)
        
        self.room_items = []
        self.all_room_items = []
        self.area_schemes = []
        self.results_log = []
        
        self._build_ui()
        self._load_data()
    
    # ========================================================================
    # UI BUILDING
    # ========================================================================
    def _build_ui(self):
        """Build the complete UI"""
        root = DockPanel()
        root.LastChildFill = True
        
        # --- HEADER ---
        root.Children.Add(self._build_header())
        DockPanel.SetDock(root.Children[root.Children.Count - 1], Dock.Top)
        
        # --- FOOTER ---
        footer = self._build_footer()
        root.Children.Add(footer)
        DockPanel.SetDock(footer, Dock.Bottom)
        
        # --- MAIN CONTENT ---
        main_grid = Grid()
        main_grid.Margin = Thickness(15, 10, 15, 5)
        
        # Two columns: left filter panel + right content
        col_left = ColumnDefinition()
        col_left.Width = GridLength(240)
        main_grid.ColumnDefinitions.Add(col_left)
        
        col_right = ColumnDefinition()
        col_right.Width = GridLength(1, GridUnitType.Star)
        main_grid.ColumnDefinitions.Add(col_right)
        
        # Left panel
        left_panel = self._build_left_panel()
        Grid.SetColumn(left_panel, 0)
        main_grid.Children.Add(left_panel)
        
        # Right panel
        right_panel = self._build_right_panel()
        Grid.SetColumn(right_panel, 1)
        main_grid.Children.Add(right_panel)
        
        root.Children.Add(main_grid)
        self.Content = root
    
    def _build_header(self):
        """Build DQT branded header"""
        header = Border()
        header.Background = bc.ConvertFromString(Config.PRIMARY)
        header.Padding = Thickness(20, 12, 20, 12)
        
        header_grid = Grid()
        col1 = ColumnDefinition()
        col1.Width = GridLength(1, GridUnitType.Star)
        header_grid.ColumnDefinitions.Add(col1)
        col2 = ColumnDefinition()
        col2.Width = GridLength.Auto
        header_grid.ColumnDefinitions.Add(col2)
        
        # Title
        title_panel = StackPanel()
        title_panel.Orientation = Orientation.Horizontal
        title_panel.VerticalAlignment = VerticalAlignment.Center
        
        icon_text = TextBlock()
        icon_text.Text = u"\u25A0"
        icon_text.FontSize = 22
        icon_text.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        icon_text.Margin = Thickness(0, 0, 10, 0)
        icon_text.VerticalAlignment = VerticalAlignment.Center
        title_panel.Children.Add(icon_text)
        
        title = TextBlock()
        title.Text = "Room To Area"
        title.FontSize = 18
        title.FontWeight = FontWeights.Bold
        title.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        title.VerticalAlignment = VerticalAlignment.Center
        title_panel.Children.Add(title)
        
        subtitle = TextBlock()
        subtitle.Text = "  |  Auto-create Areas from Rooms  |  v1.1"
        subtitle.FontSize = 12
        subtitle.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        subtitle.VerticalAlignment = VerticalAlignment.Center
        title_panel.Children.Add(subtitle)
        
        Grid.SetColumn(title_panel, 0)
        header_grid.Children.Add(title_panel)
        
        # DQT branding
        brand = TextBlock()
        brand.Text = "DQT"
        brand.FontSize = 14
        brand.FontWeight = FontWeights.Bold
        brand.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        brand.VerticalAlignment = VerticalAlignment.Center
        brand.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(brand, 1)
        header_grid.Children.Add(brand)
        
        header.Child = header_grid
        return header
    
    def _build_left_panel(self):
        """Build left filter/settings panel"""
        border = Border()
        border.Background = bc.ConvertFromString(Config.CARD_BG)
        border.BorderBrush = bc.ConvertFromString(Config.BORDER)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(6)
        border.Margin = Thickness(0, 0, 10, 0)
        border.Padding = Thickness(12)
        
        panel = StackPanel()
        
        # -- SETTINGS SECTION --
        settings_title = TextBlock()
        settings_title.Text = "Settings"
        settings_title.FontSize = 14
        settings_title.FontWeight = FontWeights.Bold
        settings_title.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        settings_title.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(settings_title)
        
        # Area Scheme selector
        scheme_label = TextBlock()
        scheme_label.Text = "Area Scheme:"
        scheme_label.FontSize = 11
        scheme_label.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        scheme_label.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(scheme_label)
        
        self.scheme_combo = ComboBox()
        self.scheme_combo.Height = 28
        self.scheme_combo.Margin = Thickness(0, 0, 0, 12)
        self.scheme_combo.SelectionChanged += self.on_scheme_changed
        panel.Children.Add(self.scheme_combo)
        
        # Area Plan View selector
        view_label = TextBlock()
        view_label.Text = "Area Plan View:"
        view_label.FontSize = 11
        view_label.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        view_label.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(view_label)
        
        self.view_combo = ComboBox()
        self.view_combo.Height = 28
        self.view_combo.Margin = Thickness(0, 0, 0, 8)
        panel.Children.Add(self.view_combo)
        
        # Auto-create view checkbox
        self.auto_create_view_cb = CheckBox()
        self.auto_create_view_cb.Content = "Auto-create Area Plan if missing"
        self.auto_create_view_cb.IsChecked = True
        self.auto_create_view_cb.FontSize = 11
        self.auto_create_view_cb.Margin = Thickness(0, 0, 0, 15)
        panel.Children.Add(self.auto_create_view_cb)
        
        # Separator
        sep = Border()
        sep.Height = 1
        sep.Background = bc.ConvertFromString(Config.BORDER)
        sep.Margin = Thickness(0, 0, 0, 15)
        panel.Children.Add(sep)
        
        # -- TRANSFER PARAMETERS SECTION --
        options_title = TextBlock()
        options_title.Text = "Transfer Parameters"
        options_title.FontSize = 14
        options_title.FontWeight = FontWeights.Bold
        options_title.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        options_title.Margin = Thickness(0, 0, 0, 6)
        panel.Children.Add(options_title)
        
        options_hint = TextBlock()
        options_hint.Text = "Select parameters to copy from Room to Area:"
        options_hint.FontSize = 10
        options_hint.Foreground = bc.ConvertFromString(Config.TEXT_LIGHT)
        options_hint.TextWrapping = TextWrapping.Wrap
        options_hint.Margin = Thickness(0, 0, 0, 6)
        panel.Children.Add(options_hint)
        
        # Buttons row: Select All / Deselect All
        transfer_btn_panel = StackPanel()
        transfer_btn_panel.Orientation = Orientation.Horizontal
        transfer_btn_panel.Margin = Thickness(0, 0, 0, 4)
        
        btn_sel_all_params = Button()
        btn_sel_all_params.Content = "All"
        btn_sel_all_params.Width = 50
        btn_sel_all_params.Height = 22
        btn_sel_all_params.FontSize = 10
        btn_sel_all_params.Margin = Thickness(0, 0, 4, 0)
        btn_sel_all_params.Click += self.on_select_all_transfer_params
        transfer_btn_panel.Children.Add(btn_sel_all_params)
        
        btn_desel_all_params = Button()
        btn_desel_all_params.Content = "None"
        btn_desel_all_params.Width = 50
        btn_desel_all_params.Height = 22
        btn_desel_all_params.FontSize = 10
        btn_desel_all_params.Margin = Thickness(0, 0, 4, 0)
        btn_desel_all_params.Click += self.on_deselect_all_transfer_params
        transfer_btn_panel.Children.Add(btn_desel_all_params)
        
        panel.Children.Add(transfer_btn_panel)
        
        # Parameter list with checkboxes (ScrollViewer + StackPanel)
        transfer_border = Border()
        transfer_border.BorderBrush = bc.ConvertFromString(Config.BORDER)
        transfer_border.BorderThickness = Thickness(1)
        transfer_border.CornerRadius = CornerRadius(4)
        transfer_border.Background = bc.ConvertFromString(Config.BACKGROUND)
        transfer_border.Margin = Thickness(0, 0, 0, 15)
        transfer_border.Height = 160
        
        transfer_scroll = ScrollViewer()
        transfer_scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        transfer_scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        transfer_scroll.Padding = Thickness(4)
        
        self.transfer_params_panel = StackPanel()
        transfer_scroll.Content = self.transfer_params_panel
        transfer_border.Child = transfer_scroll
        panel.Children.Add(transfer_border)
        
        # Separator
        sep2 = Border()
        sep2.Height = 1
        sep2.Background = bc.ConvertFromString(Config.BORDER)
        sep2.Margin = Thickness(0, 0, 0, 15)
        panel.Children.Add(sep2)
        
        # -- FILTER SECTION --
        filter_title = TextBlock()
        filter_title.Text = "Filters"
        filter_title.FontSize = 14
        filter_title.FontWeight = FontWeights.Bold
        filter_title.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        filter_title.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(filter_title)
        
        # Filter by Level
        level_label = TextBlock()
        level_label.Text = "Level:"
        level_label.FontSize = 11
        level_label.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        level_label.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(level_label)
        
        self.level_combo = ComboBox()
        self.level_combo.Height = 28
        self.level_combo.Margin = Thickness(0, 0, 0, 12)
        self.level_combo.SelectionChanged += self.on_level_filter_changed
        panel.Children.Add(self.level_combo)
        
        # Parameter Column selector
        param_col_label = TextBlock()
        param_col_label.Text = "Parameter Column:"
        param_col_label.FontSize = 11
        param_col_label.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        param_col_label.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(param_col_label)
        
        self.param_col_combo = ComboBox()
        self.param_col_combo.Height = 28
        self.param_col_combo.Margin = Thickness(0, 0, 0, 12)
        self.param_col_combo.SelectionChanged += self.on_param_col_changed
        panel.Children.Add(self.param_col_combo)
        
        # Filter by Parameter Value
        param_val_label = TextBlock()
        param_val_label.Text = "Filter by Parameter Value:"
        param_val_label.FontSize = 11
        param_val_label.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        param_val_label.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(param_val_label)
        
        self.param_val_combo = ComboBox()
        self.param_val_combo.Height = 28
        self.param_val_combo.Margin = Thickness(0, 0, 0, 12)
        self.param_val_combo.SelectionChanged += self.on_param_val_filter_changed
        panel.Children.Add(self.param_val_combo)
        
        # Search box
        search_label = TextBlock()
        search_label.Text = "Search Rooms:"
        search_label.FontSize = 11
        search_label.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        search_label.Margin = Thickness(0, 5, 0, 4)
        panel.Children.Add(search_label)
        
        self.search_box = TextBox()
        self.search_box.Height = 28
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self.on_search_changed
        panel.Children.Add(self.search_box)
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll.Content = panel
        border.Child = scroll
        return border
    
    def _build_right_panel(self):
        """Build right content panel with summary + DataGrid"""
        panel = StackPanel()
        
        # Summary cards
        panel.Children.Add(self._build_summary_cards())
        
        # DataGrid
        panel.Children.Add(self._build_datagrid())
        
        return panel
    
    def _build_summary_cards(self):
        """Build summary statistics cards"""
        wrap = WrapPanel()
        wrap.Margin = Thickness(0, 0, 0, 10)
        
        self.card_total = self._create_card("Total Rooms", "0")
        wrap.Children.Add(self.card_total)
        
        self.card_checked = self._create_card("Selected", "0")
        wrap.Children.Add(self.card_checked)
        
        self.card_area = self._create_card("Total Area (m2)", "0.00")
        wrap.Children.Add(self.card_area)
        
        self.card_levels = self._create_card("Levels", "0")
        wrap.Children.Add(self.card_levels)
        
        return wrap
    
    def _create_card(self, title, value):
        """Create a single summary card"""
        border = Border()
        border.Background = bc.ConvertFromString(Config.CARD_BG)
        border.BorderBrush = bc.ConvertFromString(Config.BORDER)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(6)
        border.Padding = Thickness(15, 10, 15, 10)
        border.Margin = Thickness(0, 0, 8, 0)
        border.MinWidth = 140
        
        panel = StackPanel()
        
        title_tb = TextBlock()
        title_tb.Text = title
        title_tb.FontSize = 10
        title_tb.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        panel.Children.Add(title_tb)
        
        value_tb = TextBlock()
        value_tb.Text = value
        value_tb.FontSize = 20
        value_tb.FontWeight = FontWeights.Bold
        value_tb.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
        value_tb.Tag = "value"
        panel.Children.Add(value_tb)
        
        border.Child = panel
        return border
    
    def _update_card_value(self, card_border, new_value):
        """Update the value TextBlock inside a card"""
        panel = card_border.Child
        for i in range(panel.Children.Count):
            child = panel.Children[i]
            if hasattr(child, 'Tag') and child.Tag == "value":
                child.Text = str(new_value)
                break
    
    def _build_datagrid(self):
        """Build the main DataGrid for rooms"""
        border = Border()
        border.Background = bc.ConvertFromString(Config.CARD_BG)
        border.BorderBrush = bc.ConvertFromString(Config.BORDER)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(6)
        border.Padding = Thickness(2)
        border.Height = 420
        
        self.data_grid = DataGrid()
        self.data_grid.AutoGenerateColumns = False
        self.data_grid.CanUserAddRows = False
        self.data_grid.CanUserDeleteRows = False
        self.data_grid.IsReadOnly = True
        self.data_grid.SelectionUnit = DataGridSelectionUnit.FullRow
        self.data_grid.SelectionMode = DataGridSelectionMode.Extended
        self.data_grid.GridLinesVisibility = getattr(DataGridGridLinesVisibility, "None")
        self.data_grid.HeadersVisibility = DataGridHeadersVisibility.Column
        self.data_grid.RowHeight = 30
        self.data_grid.BorderThickness = Thickness(0)
        self.data_grid.Background = bc.ConvertFromString(Config.CARD_BG)
        self.data_grid.AlternatingRowBackground = bc.ConvertFromString(Config.ROW_ALT)
        self.data_grid.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        self.data_grid.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        
        # Columns - use text column for checkbox display (safe for IronPython)
        col_check = DataGridTextColumn()
        col_check.Header = "Sel"
        col_check.Binding = Binding("check_display")
        col_check.Width = DataGridLength(40)
        self.data_grid.Columns.Add(col_check)
        
        # ID
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Binding = Binding("id")
        col_id.Width = DataGridLength(70)
        col_id.IsReadOnly = True
        self.data_grid.Columns.Add(col_id)
        
        # Number
        col_num = DataGridTextColumn()
        col_num.Header = "Number"
        col_num.Binding = Binding("number")
        col_num.Width = DataGridLength(80)
        col_num.IsReadOnly = True
        self.data_grid.Columns.Add(col_num)
        
        # Name
        col_name = DataGridTextColumn()
        col_name.Header = "Room Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(1, DataGridLengthUnitType.Star)
        col_name.IsReadOnly = True
        self.data_grid.Columns.Add(col_name)
        
        # Level
        col_level = DataGridTextColumn()
        col_level.Header = "Level"
        col_level.Binding = Binding("level")
        col_level.Width = DataGridLength(120)
        col_level.IsReadOnly = True
        self.data_grid.Columns.Add(col_level)
        
        # Area
        col_area = DataGridTextColumn()
        col_area.Header = "Area (m2)"
        col_area.Binding = Binding("area_display")
        col_area.Width = DataGridLength(90)
        col_area.IsReadOnly = True
        self.data_grid.Columns.Add(col_area)
        
        # Department
        col_dept = DataGridTextColumn()
        col_dept.Header = "Department"
        col_dept.Binding = Binding("department")
        col_dept.Width = DataGridLength(100)
        col_dept.IsReadOnly = True
        self.data_grid.Columns.Add(col_dept)
        
        # Dynamic parameter column (hidden by default)
        self._param_col_ref = DataGridTextColumn()
        self._param_col_ref.Header = "Parameter"
        self._param_col_ref.Binding = Binding("param_value")
        self._param_col_ref.Width = DataGridLength(130)
        self._param_col_ref.IsReadOnly = True
        self._param_col_ref.Visibility = Visibility.Collapsed
        self.data_grid.Columns.Add(self._param_col_ref)
        
        # Status
        col_status = DataGridTextColumn()
        col_status.Header = "Status"
        col_status.Binding = Binding("status")
        col_status.Width = DataGridLength(120)
        col_status.IsReadOnly = True
        self.data_grid.Columns.Add(col_status)
        
        border.Child = self.data_grid
        
        # Double-click to toggle check
        self.data_grid.MouseDoubleClick += self.on_grid_double_click
        
        return border
    
    def _build_footer(self):
        """Build footer with action buttons"""
        footer = Border()
        footer.Background = bc.ConvertFromString(Config.CARD_BG)
        footer.BorderBrush = bc.ConvertFromString(Config.BORDER)
        footer.BorderThickness = Thickness(0, 1, 0, 0)
        footer.Padding = Thickness(15, 10, 15, 10)
        
        grid = Grid()
        col1 = ColumnDefinition()
        col1.Width = GridLength(1, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col1)
        col2 = ColumnDefinition()
        col2.Width = GridLength.Auto
        grid.ColumnDefinitions.Add(col2)
        
        # Left: Check All / Uncheck All
        left_panel = StackPanel()
        left_panel.Orientation = Orientation.Horizontal
        left_panel.VerticalAlignment = VerticalAlignment.Center
        
        btn_check_all = self._create_button("Check All", self.on_check_all)
        left_panel.Children.Add(btn_check_all)
        
        btn_uncheck_all = self._create_button("Uncheck All", self.on_uncheck_all)
        left_panel.Children.Add(btn_uncheck_all)
        
        btn_invert = self._create_button("Invert", self.on_invert_selection)
        left_panel.Children.Add(btn_invert)
        
        btn_select_revit = self._create_button("Select in Revit", self.on_select_in_revit)
        left_panel.Children.Add(btn_select_revit)
        
        # Status text
        self.status_text = TextBlock()
        self.status_text.Text = "Ready"
        self.status_text.FontSize = 11
        self.status_text.Foreground = bc.ConvertFromString(Config.TEXT_MEDIUM)
        self.status_text.VerticalAlignment = VerticalAlignment.Center
        self.status_text.Margin = Thickness(15, 0, 0, 0)
        left_panel.Children.Add(self.status_text)
        
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        # Right: Action buttons
        right_panel = StackPanel()
        right_panel.Orientation = Orientation.Horizontal
        right_panel.HorizontalAlignment = HorizontalAlignment.Right
        
        btn_create_selected = self._create_button("Create from Selected", self.on_create_from_selected)
        right_panel.Children.Add(btn_create_selected)
        
        btn_create = self._create_button("Create from Checked", self.on_create_areas, primary=True)
        right_panel.Children.Add(btn_create)
        
        btn_close = self._create_button("Close", self.on_close)
        right_panel.Children.Add(btn_close)
        
        Grid.SetColumn(right_panel, 1)
        grid.Children.Add(right_panel)
        
        # Footer layout: grid on top + copyright on bottom
        footer_stack = StackPanel()
        footer_stack.Children.Add(grid)
        
        # Copyright line
        copyright_text = TextBlock()
        copyright_text.Text = "Copyright (c) 2025 Dang Quoc Truong (DQT) | pyDQT Suite v1.1"
        copyright_text.FontSize = 9
        copyright_text.Foreground = bc.ConvertFromString(Config.TEXT_LIGHT)
        copyright_text.HorizontalAlignment = HorizontalAlignment.Right
        copyright_text.Margin = Thickness(0, 5, 5, 0)
        footer_stack.Children.Add(copyright_text)
        
        footer.Child = footer_stack
        return footer
    
    def _create_button(self, text, handler, primary=False):
        """Create a styled button"""
        btn = Button()
        btn.Content = text
        btn.Height = 32
        btn.MinWidth = 90
        btn.Padding = Thickness(15, 5, 15, 5)
        btn.Margin = Thickness(3)
        btn.FontWeight = FontWeights.SemiBold
        btn.FontSize = 12
        btn.Cursor = System.Windows.Input.Cursors.Hand
        
        if primary:
            btn.Background = bc.ConvertFromString(Config.PRIMARY)
            btn.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
            btn.BorderBrush = bc.ConvertFromString(Config.ACCENT)
        else:
            btn.Background = bc.ConvertFromString("#F5F5F5")
            btn.Foreground = bc.ConvertFromString(Config.TEXT_DARK)
            btn.BorderBrush = bc.ConvertFromString("#CCCCCC")
        
        btn.BorderThickness = Thickness(1)
        btn.Click += handler
        return btn
    
    # ========================================================================
    # DATA LOADING
    # ========================================================================
    def _load_data(self):
        """Load rooms, area schemes, levels, parameters"""
        # Load rooms - try selected first, then all
        selected_rooms = get_all_rooms(selected_only=True)
        if selected_rooms:
            rooms = selected_rooms
            self.status_text.Text = "Loaded {} selected room(s)".format(len(rooms))
        else:
            rooms = get_all_rooms(selected_only=False)
            self.status_text.Text = "Loaded all {} room(s)".format(len(rooms))
        
        self.all_room_items = [RoomItem(r) for r in rooms]
        self.room_items = list(self.all_room_items)
        
        # Load area schemes
        self.area_schemes = get_area_schemes()
        self.scheme_combo.Items.Clear()
        for scheme in self.area_schemes:
            item = ComboBoxItem()
            item.Content = scheme.Name
            item.Tag = scheme.Id
            self.scheme_combo.Items.Add(item)
        
        if self.scheme_combo.Items.Count > 0:
            self.scheme_combo.SelectedIndex = 0
        
        # Load level filter
        levels = set()
        for ri in self.all_room_items:
            levels.add(ri.level)
        
        self.level_combo.Items.Clear()
        all_item = ComboBoxItem()
        all_item.Content = "All Levels"
        self.level_combo.Items.Add(all_item)
        
        for lvl_name in sorted(levels):
            item = ComboBoxItem()
            item.Content = lvl_name
            self.level_combo.Items.Add(item)
        
        self.level_combo.SelectedIndex = 0
        
        # Load parameter names from all rooms
        self._load_parameter_names()
        
        # Refresh grid
        self._refresh_grid()
        self._update_summary()
    
    def _load_parameter_names(self):
        """Collect all unique parameter names from rooms"""
        param_names = set()
        for ri in self.all_room_items:
            for pname in ri._param_cache.keys():
                param_names.add(pname)
        
        self.param_col_combo.Items.Clear()
        none_item = ComboBoxItem()
        none_item.Content = "(None)"
        self.param_col_combo.Items.Add(none_item)
        
        for pname in sorted(param_names):
            item = ComboBoxItem()
            item.Content = pname
            self.param_col_combo.Items.Add(item)
        
        self.param_col_combo.SelectedIndex = 0
        
        # Clear param value filter
        self.param_val_combo.Items.Clear()
        all_val = ComboBoxItem()
        all_val.Content = "All Values"
        self.param_val_combo.Items.Add(all_val)
        self.param_val_combo.SelectedIndex = 0
        
        # Build transfer parameter checkboxes
        self._build_transfer_param_list()
    
    def _build_transfer_param_list(self):
        """Build checkboxes for parameters that can be transferred to Area"""
        self.transfer_params_panel.Children.Clear()
        self.transfer_param_checkboxes = {}
        
        # Collect parameter names from rooms
        # Find common transferable params (string/double/integer types)
        room_params = {}
        for ri in self.all_room_items:
            try:
                for param in ri.element.Parameters:
                    try:
                        pname = param.Definition.Name
                        if pname not in room_params:
                            storage = param.StorageType.ToString()
                            is_readonly = param.IsReadOnly
                            room_params[pname] = {
                                'storage': storage,
                                'readonly': is_readonly,
                                'is_builtin': param.Definition.BuiltInParameter.ToString() != "INVALID"
                            }
                    except:
                        pass
            except:
                pass
            # Only need one room to get parameter definitions
            if room_params:
                break
        
        # Default checked params
        default_checked = set(["Name", "Number", "Department", "Comments"])
        
        # Sort: common params first, then alphabetical
        common_first = ["Name", "Number", "Department", "Comments", "Occupancy", "Occupant"]
        sorted_params = []
        for p in common_first:
            if p in room_params:
                sorted_params.append(p)
        for p in sorted(room_params.keys()):
            if p not in sorted_params:
                sorted_params.append(p)
        
        for pname in sorted_params:
            info = room_params[pname]
            
            cb = CheckBox()
            cb.Content = pname
            cb.FontSize = 11
            cb.Margin = Thickness(2, 2, 2, 2)
            cb.Tag = pname
            
            # Default check Name and Number
            if pname in default_checked:
                cb.IsChecked = True
            else:
                cb.IsChecked = False
            
            # Mark read-only params with hint
            if info['readonly']:
                cb.Content = pname + " (read-only)"
                cb.IsEnabled = False
                cb.IsChecked = False
            
            self.transfer_params_panel.Children.Add(cb)
            self.transfer_param_checkboxes[pname] = cb
    
    def get_selected_transfer_params(self):
        """Get list of parameter names that user selected for transfer"""
        selected = []
        for pname, cb in self.transfer_param_checkboxes.items():
            if cb.IsChecked:
                selected.append(pname)
        return selected
    
    def on_select_all_transfer_params(self, sender, args):
        """Check all transfer parameter checkboxes"""
        for pname, cb in self.transfer_param_checkboxes.items():
            if cb.IsEnabled:
                cb.IsChecked = True
    
    def on_deselect_all_transfer_params(self, sender, args):
        """Uncheck all transfer parameter checkboxes"""
        for pname, cb in self.transfer_param_checkboxes.items():
            cb.IsChecked = False
    
    def _update_param_value_filter(self):
        """Update parameter value filter based on selected parameter column"""
        self.param_val_combo.Items.Clear()
        all_val = ComboBoxItem()
        all_val.Content = "All Values"
        self.param_val_combo.Items.Add(all_val)
        
        if self.param_col_combo.SelectedIndex > 0:
            param_name = self.param_col_combo.SelectedItem.Content
            values = set()
            for ri in self.all_room_items:
                val = ri.get_param_value(param_name)
                if val:
                    values.add(val)
            
            for val in sorted(values):
                item = ComboBoxItem()
                item.Content = val
                self.param_val_combo.Items.Add(item)
            
            # Add empty option
            if any(ri.get_param_value(param_name) == "" for ri in self.all_room_items):
                empty_item = ComboBoxItem()
                empty_item.Content = "(Empty)"
                self.param_val_combo.Items.Add(empty_item)
        
        self.param_val_combo.SelectedIndex = 0
    
    def _refresh_grid(self):
        """Refresh the DataGrid items"""
        self.data_grid.ItemsSource = None
        self.data_grid.Items.Clear()
        
        for item in self.room_items:
            self.data_grid.Items.Add(item)
    
    def _update_summary(self):
        """Update summary cards"""
        total = len(self.room_items)
        checked = sum(1 for ri in self.room_items if ri.is_checked)
        total_area = sum(ri.area_sqm for ri in self.room_items if ri.is_checked)
        levels = len(set(ri.level for ri in self.room_items))
        
        self._update_card_value(self.card_total, str(total))
        self._update_card_value(self.card_checked, str(checked))
        self._update_card_value(self.card_area, str(round(total_area, 2)))
        self._update_card_value(self.card_levels, str(levels))
    
    def _apply_filters(self):
        """Apply level, parameter value, and search filters"""
        filtered = list(self.all_room_items)
        
        # Level filter
        if self.level_combo.SelectedIndex > 0:
            sel_level = self.level_combo.SelectedItem.Content
            filtered = [ri for ri in filtered if ri.level == sel_level]
        
        # Parameter value filter
        if self.param_col_combo.SelectedIndex > 0 and self.param_val_combo.SelectedIndex > 0:
            param_name = self.param_col_combo.SelectedItem.Content
            sel_val = self.param_val_combo.SelectedItem.Content
            if sel_val == "(Empty)":
                filtered = [ri for ri in filtered if ri.get_param_value(param_name) == ""]
            else:
                filtered = [ri for ri in filtered if ri.get_param_value(param_name) == sel_val]
        
        # Search filter
        search_text = self.search_box.Text.strip().lower()
        if search_text:
            filtered = [
                ri for ri in filtered
                if search_text in ri.name.lower()
                or search_text in ri.number.lower()
                or search_text in ri.department.lower()
                or search_text in ri.param_value.lower()
                or search_text in str(ri.id)
            ]
        
        self.room_items = filtered
        self._refresh_grid()
        self._update_summary()
    
    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================
    def on_scheme_changed(self, sender, args):
        """When area scheme selection changes, update available views"""
        self.view_combo.Items.Clear()
        
        if self.scheme_combo.SelectedIndex < 0:
            return
        
        scheme_id = self.scheme_combo.SelectedItem.Tag
        views = get_area_plan_views(scheme_id)
        
        for view in views:
            item = ComboBoxItem()
            level_name = view.GenLevel.Name if view.GenLevel else "N/A"
            item.Content = "{} ({})".format(view.Name, level_name)
            item.Tag = view.Id
            self.view_combo.Items.Add(item)
        
        if self.view_combo.Items.Count > 0:
            self.view_combo.SelectedIndex = 0
    
    def on_level_filter_changed(self, sender, args):
        self._apply_filters()
    
    def on_param_col_changed(self, sender, args):
        """When parameter column selection changes"""
        if self.param_col_combo.SelectedIndex > 0:
            param_name = self.param_col_combo.SelectedItem.Content
            # Update param_value for all room items
            for ri in self.all_room_items:
                ri.param_value = ri.get_param_value(param_name)
            # Update the DataGrid column header
            self._update_param_column(param_name)
        else:
            for ri in self.all_room_items:
                ri.param_value = ""
            self._update_param_column(None)
        
        # Update parameter value filter dropdown
        self._update_param_value_filter()
        self._apply_filters()
    
    def on_param_val_filter_changed(self, sender, args):
        self._apply_filters()
    
    def _update_param_column(self, param_name):
        """Show or hide the parameter column in DataGrid"""
        # Find and update the param column (last column before Status)
        # The param column is at index 7 (after Dept, before Status)
        if hasattr(self, '_param_col_ref') and self._param_col_ref:
            if param_name:
                self._param_col_ref.Header = param_name
                self._param_col_ref.Visibility = Visibility.Visible
            else:
                self._param_col_ref.Visibility = Visibility.Collapsed
        self._refresh_grid()
    
    def on_search_changed(self, sender, args):
        self._apply_filters()
    
    def on_check_all(self, sender, args):
        for ri in self.room_items:
            ri.is_checked = True
            ri.check_display = "[x]"
        self._refresh_grid()
        self._update_summary()
    
    def on_uncheck_all(self, sender, args):
        for ri in self.room_items:
            ri.is_checked = False
            ri.check_display = "[ ]"
        self._refresh_grid()
        self._update_summary()
    
    def on_invert_selection(self, sender, args):
        for ri in self.room_items:
            ri.is_checked = not ri.is_checked
            ri.check_display = "[x]" if ri.is_checked else "[ ]"
        self._refresh_grid()
        self._update_summary()
    
    def on_close(self, sender, args):
        self.Close()
    
    def on_select_in_revit(self, sender, args):
        """Select checked rooms in Revit"""
        try:
            checked = [ri for ri in self.room_items if ri.is_checked]
            if not checked:
                TaskDialog.Show("Room To Area", "No rooms checked to select.")
                return
            
            ids = List[ElementId]()
            for ri in checked:
                ids.Add(ri.element.Id)
            
            uidoc.Selection.SetElementIds(ids)
            self.status_text.Text = "Selected {} room(s) in Revit".format(len(checked))
        except Exception as ex:
            TaskDialog.Show("Error", "Select failed: {}".format(str(ex)))
    
    def on_create_from_selected(self, sender, args):
        """Create areas from rows currently highlighted/selected in DataGrid"""
        selected_items = []
        for item in self.data_grid.SelectedItems:
            if item and hasattr(item, 'element'):
                selected_items.append(item)
        
        if not selected_items:
            TaskDialog.Show("Room To Area",
                "No rows selected in the table.\n"
                "Please click/highlight rows you want to create areas from.\n\n"
                "Tip: Hold Ctrl or Shift to select multiple rows.")
            return
        
        self._execute_create_areas(selected_items)
    
    def on_grid_double_click(self, sender, args):
        """Toggle check state on double-click"""
        try:
            item = self.data_grid.SelectedItem
            if item and hasattr(item, 'is_checked'):
                item.is_checked = not item.is_checked
                item.check_display = "[x]" if item.is_checked else "[ ]"
                self._refresh_grid()
                self._update_summary()
        except:
            pass
    
    # ========================================================================
    # CORE ACTION - CREATE AREAS
    # ========================================================================
    def on_create_areas(self, sender, args):
        """Create areas from checked rooms"""
        checked_rooms = [ri for ri in self.room_items if ri.is_checked]
        
        if not checked_rooms:
            TaskDialog.Show("Room To Area", "No rooms checked. Please check rooms to process.")
            return
        
        self._execute_create_areas(checked_rooms)
    
    def _execute_create_areas(self, room_item_list):
        """Core logic: create areas from a list of RoomItem objects"""
        if not room_item_list:
            return
        
        # Get selected area scheme
        if self.scheme_combo.SelectedIndex < 0 or not self.area_schemes:
            TaskDialog.Show("Room To Area", 
                "No Area Scheme found in the project.\n"
                "Please create an Area Scheme first:\n"
                "Architecture > Area > Area and Volume Computations")
            return
        
        scheme = self.area_schemes[self.scheme_combo.SelectedIndex]
        
        # Get selected transfer parameters
        transfer_params = self.get_selected_transfer_params()
        
        # Group rooms by level
        rooms_by_level = {}
        for ri in room_item_list:
            level_id = ri.element.LevelId.IntegerValue
            if level_id not in rooms_by_level:
                rooms_by_level[level_id] = []
            rooms_by_level[level_id].append(ri)
        
        success_count = 0
        fail_count = 0
        
        tg = TransactionGroup(doc, "DQT - Create Areas from Rooms")
        tg.Start()
        
        try:
            for level_id, room_items_in_level in rooms_by_level.items():
                level = doc.GetElement(ElementId(level_id))
                if not level:
                    for ri in room_items_in_level:
                        ri.status = "Level not found"
                        fail_count += 1
                    continue
                
                # Find or create area plan view for this level
                t_view = Transaction(doc, "Find/Create Area Plan")
                t_view.Start()
                
                try:
                    area_plan = find_or_create_area_plan(scheme, level)
                    
                    if not area_plan:
                        if self.auto_create_view_cb.IsChecked:
                            area_plan = ViewPlan.CreateAreaPlan(doc, scheme.Id, level.Id)
                        
                        if not area_plan:
                            t_view.RollBack()
                            for ri in room_items_in_level:
                                ri.status = "No Area Plan for level"
                                fail_count += 1
                            continue
                    
                    t_view.Commit()
                except Exception as ex:
                    t_view.RollBack()
                    for ri in room_items_in_level:
                        ri.status = "View error: {}".format(str(ex)[:40])
                        fail_count += 1
                    continue
                
                # Create areas for each room on this level
                engine = AreaCreationEngine(doc, scheme, area_plan)
                
                for ri in room_items_in_level:
                    t = Transaction(doc, "Create Area from Room {}".format(ri.number))
                    t.Start()
                    
                    try:
                        result = engine.create_area_from_room(ri.element, transfer_params)
                        
                        if result['success']:
                            ri.status = result['message'][:60]
                            success_count += 1
                            t.Commit()
                        else:
                            ri.status = result['message'][:50]
                            fail_count += 1
                            t.RollBack()
                    except Exception as ex:
                        ri.status = "Error: {}".format(str(ex)[:40])
                        fail_count += 1
                        t.RollBack()
            
            tg.Assimilate()
            
        except Exception as ex:
            tg.RollBack()
            TaskDialog.Show("Error", "Transaction failed:\n{}".format(str(ex)))
            return
        
        # Update UI
        self._refresh_grid()
        self._update_summary()
        
        self.status_text.Text = "Done: {} created, {} failed".format(success_count, fail_count)
        
        # Show summary dialog
        msg = "Room To Area completed!\n\n"
        msg += "  Successfully created: {} area(s)\n".format(success_count)
        msg += "  Failed: {} room(s)\n".format(fail_count)
        msg += "  Area Scheme: {}\n".format(scheme.Name)
        
        if transfer_params:
            msg += "  Transferred: {} parameter(s)\n".format(len(transfer_params))
        
        if fail_count > 0:
            msg += "\nCheck the Status column for details on failures."
        
        TaskDialog.Show("Room To Area - DQT", msg)


# ============================================================================
# MAIN ENTRY POINT
# Copyright (c) 2025 Dang Quoc Truong (DQT) - pyDQT Suite
# ============================================================================
def main():
    try:
        # Quick check: any rooms exist?
        room_count = FilteredElementCollector(doc).OfCategory(
            BuiltInCategory.OST_Rooms
        ).WhereElementIsNotElementType().GetElementCount()
        
        if room_count == 0:
            TaskDialog.Show("Room To Area", 
                "No rooms found in the current project.\n"
                "Please place rooms first before using this tool.")
            return
        
        # Check area schemes
        scheme_count = FilteredElementCollector(doc).OfClass(AreaScheme).GetElementCount()
        if scheme_count == 0:
            TaskDialog.Show("Room To Area",
                "No Area Scheme found in the project.\n\n"
                "To create one:\n"
                "Architecture tab > Room & Area panel > Area dropdown > "
                "Area and Volume Computations")
            return
        
        window = RoomToAreaWindow()
        window.ShowDialog()
        
    except Exception as ex:
        TaskDialog.Show("Room To Area - Error",
            "An error occurred:\n{}".format(str(ex)))

if __name__ == "__main__":
    main()