# -*- coding: utf-8 -*-
"""
Style Manager v1.0
Manage and rename Line Styles, Line Patterns, and Fill Patterns in Revit in a unified interface.

Copyright (c) 2026 Dang Quoc Truong & T3Lab
All rights reserved.
"""
__title__ = "Style\nManager"
__author__ = "Dang Quoc Truong & T3Lab"
__doc__ = "Manage and rename Fill Patterns, Line Styles, and Line Patterns in one unified dialog."

import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import (Window, Thickness, GridLength, GridUnitType,
                            HorizontalAlignment, VerticalAlignment, FontWeights,
                            MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult)
from System.Windows.Controls import (Grid, RowDefinition, ColumnDefinition, Border,
                                      StackPanel, TextBlock, TextBox, Button,
                                      ComboBox, ComboBoxItem, DataGrid, Orientation,
                                      DataGridTextColumn, DataGridCheckBoxColumn,
                                      ScrollViewer, TabControl, TabItem)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

import re
import os
import sys

from pyrevit import revit, forms
from Autodesk.Revit.DB import (FilteredElementCollector, FillPatternElement, 
                                FillPatternTarget, BuiltInParameter, ElementId,
                                Category, GraphicsStyleType, Transaction, TransactionGroup,
                                LinePatternElement, CurveElement)

doc = revit.doc

# ============================================================================
# REVIT API COMPATIBILITY HELPERS
# ============================================================================
def _eid_int(element_id):
    """Get integer value from ElementId - compatible with all Revit versions"""
    if not element_id:
        return -1
    try:
        return element_id.Value  # Revit 2025+
    except AttributeError:
        return element_id.IntegerValue  # Revit 2024 and earlier

def _get_invalid_element_id():
    return ElementId.InvalidElementId

# ============================================================================
# CONFIGURATION
# ============================================================================
class Config(object):
    PRIMARY_COLOR = "#0F172A"       # Dark slate blue header / cards
    SECONDARY_COLOR = "#E5B85C"     # Gold Accent
    BACKGROUND_COLOR = "#F8FAFC"    # Soft grey background
    BORDER_COLOR = "#CBD5E1"        # Light border
    TEXT_DARK = "#0F172A"
    TEXT_MUTED = "#64748B"
    SUCCESS_COLOR = "#10B981"
    WARNING_COLOR = "#F59E0B"
    ERROR_COLOR = "#EF4444"
    WHITE = "#FFFFFF"
    ROW_ALT_COLOR = "#F8FAFC"
    
    @staticmethod
    def hex_to_color(hex_color):
        hex_color = hex_color.replace("#", "")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return Color.FromArgb(255, r, g, b)

# ============================================================================
# DATA MODEL WRAPPERS
# ============================================================================

class FillPatternItem(INotifyPropertyChanged):
    def __init__(self, element):
        self._handlers = []
        self._element = element
        self._is_selected = False
        try:
            self._name = element.Name if element.Name else "Unnamed"
        except:
            self._name = "Unnamed"
        self._id = _eid_int(element.Id)
        
        # Get target type
        try:
            fill_pattern = element.GetFillPattern()
            if fill_pattern:
                target = fill_pattern.Target
                self._pattern_type = "Drafting" if target == FillPatternTarget.Drafting else "Model"
                grids = fill_pattern.GetFillGrids()
                self._grid_count = len(grids) if grids else 0
                
                if not grids or len(grids) == 0:
                    self._settings = "Solid fill"
                elif len(grids) == 1:
                    self._settings = "Parallel lines"
                else:
                    self._settings = "Crosshatch" if len(grids) == 2 else "Custom"
            else:
                self._pattern_type = "Unknown"
                self._grid_count = 0
                self._settings = "N/A"
        except:
            self._pattern_type = "Unknown"
            self._grid_count = 0
            self._settings = "N/A"
            
        self._is_system = "Solid fill" in self._name or self._name.startswith("<")

    @property
    def element(self): return self._element
    @property
    def id(self): return self._id
    @property
    def name(self): return self._name
    @name.setter
    def name(self, v):
        if self._name != v:
            self._name = v
            self._notify("name")
            
    @property
    def pattern_type(self): return self._pattern_type
    @property
    def grid_count(self): return self._grid_count
    @property
    def settings(self): return self._settings
    @property
    def is_system(self): return self._is_system
    
    @property
    def is_selected(self): return self._is_selected
    @is_selected.setter
    def is_selected(self, v):
        if self._is_selected != v:
            self._is_selected = v
            self._notify("is_selected")

    def add_PropertyChanged(self, h): self._handlers.append(h)
    def remove_PropertyChanged(self, h):
        if h in self._handlers: self._handlers.remove(h)
    def _notify(self, prop):
        for h in self._handlers: h(self, PropertyChangedEventArgs(prop))


class LineStyleItem(INotifyPropertyChanged):
    def __init__(self, category):
        self._handlers = []
        self._category = category
        self._is_selected = False
        try:
            self._name = category.Name if category.Name else "Unnamed"
        except:
            self._name = "Unnamed"
        self._id = _eid_int(category.Id)
        
        try:
            color = category.LineColor
            self._color = "RGB({},{},{})".format(color.Red, color.Green, color.Blue)
        except:
            self._color = "N/A"
            
        try:
            weight = category.GetLineWeight(GraphicsStyleType.Projection)
            self._weight = str(weight) if weight else "N/A"
        except:
            self._weight = "N/A"
            
        try:
            pattern_id = category.GetLinePatternId(GraphicsStyleType.Projection)
            if pattern_id and _eid_int(pattern_id) != _eid_int(_get_invalid_element_id()):
                pat = doc.GetElement(pattern_id)
                self._pattern = pat.Name if pat else "Solid"
            else:
                self._pattern = "Solid"
        except:
            self._pattern = "Solid"
            
        self._is_system = self._name.startswith('<') and self._name.endswith('>')
        self._usage_count = 0

    @property
    def category(self): return self._category
    @property
    def id(self): return self._id
    @property
    def name(self): return self._name
    @name.setter
    def name(self, v):
        if self._name != v:
            self._name = v
            self._notify("name")
            
    @property
    def color(self): return self._color
    @property
    def weight(self): return self._weight
    @property
    def pattern(self): return self._pattern
    @property
    def is_system(self): return self._is_system
    
    @property
    def usage_count(self): return self._usage_count
    @usage_count.setter
    def usage_count(self, v):
        if self._usage_count != v:
            self._usage_count = v
            self._notify("usage_count")
            
    @property
    def is_selected(self): return self._is_selected
    @is_selected.setter
    def is_selected(self, v):
        if self._is_selected != v:
            self._is_selected = v
            self._notify("is_selected")

    def add_PropertyChanged(self, h): self._handlers.append(h)
    def remove_PropertyChanged(self, h):
        if h in self._handlers: self._handlers.remove(h)
    def _notify(self, prop):
        for h in self._handlers: h(self, PropertyChangedEventArgs(prop))


class LinePatternItem(INotifyPropertyChanged):
    SYSTEM_PATTERNS = ["Solid", "Dash", "Dot", "Dash dot", "Dash dot dot"]
    
    def __init__(self, element):
        self._handlers = []
        self._element = element
        self._is_selected = False
        try:
            self._name = element.Name if element.Name else "Unnamed"
        except:
            self._name = "Unnamed"
        self._id = _eid_int(element.Id)
        
        try:
            line_pattern = element.GetLinePattern()
            if line_pattern:
                segments = line_pattern.GetSegments()
                self._segment_count = len(list(segments)) if segments else 0
                
                types = []
                values = []
                for seg in segments:
                    types.append(seg.Type.ToString())
                    values.append("{:.2f}mm".format(seg.Length * 304.8))
                self._segments_type = ", ".join(types) if types else "Solid"
                self._segments_value = ", ".join(values) if values else "-"
            else:
                self._segment_count = 0
                self._segments_type = "Solid"
                self._segments_value = "-"
        except:
            self._segment_count = 0
            self._segments_type = "Solid"
            self._segments_value = "-"
            
        self._is_system = self._name in self.SYSTEM_PATTERNS

    @property
    def element(self): return self._element
    @property
    def id(self): return self._id
    @property
    def name(self): return self._name
    @name.setter
    def name(self, v):
        if self._name != v:
            self._name = v
            self._notify("name")
            
    @property
    def segment_count(self): return self._segment_count
    @property
    def segments_type(self): return self._segments_type
    @property
    def segments_value(self): return self._segments_value
    @property
    def is_system(self): return self._is_system
    
    @property
    def is_selected(self): return self._is_selected
    @is_selected.setter
    def is_selected(self, v):
        if self._is_selected != v:
            self._is_selected = v
            self._notify("is_selected")

    def add_PropertyChanged(self, h): self._handlers.append(h)
    def remove_PropertyChanged(self, h):
        if h in self._handlers: self._handlers.remove(h)
    def _notify(self, prop):
        for h in self._handlers: h(self, PropertyChangedEventArgs(prop))


# ============================================================================
# MAIN WINDOW STYLE MANAGER
# ============================================================================

class StyleManagerWindow(Window):
    def __init__(self):
        self.Title = "Style Manager - Consolidated"
        self.Width = 1000
        self.Height = 700
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Config.hex_to_color(Config.BACKGROUND_COLOR))
        
        # Data caches
        self.fill_patterns = []
        self.line_styles = []
        self.line_patterns = []
        
        self.filtered_fill_patterns = ObservableCollection[object]()
        self.filtered_line_styles = ObservableCollection[object]()
        self.filtered_line_patterns = ObservableCollection[object]()
        
        self._create_ui()
        self._load_all_data()

    def _create_ui(self):
        main_grid = Grid()
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        
        # Header
        header_border = Border()
        header_border.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        header_border.Padding = Thickness(20, 10, 20, 10)
        
        header_text = TextBlock()
        header_text.Text = "T3Lab Revit Style Manager"
        header_text.FontSize = 22
        header_text.FontWeight = FontWeights.Bold
        header_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        header_text.VerticalAlignment = VerticalAlignment.Center
        header_border.Child = header_text
        Grid.SetRow(header_border, 0)
        main_grid.Children.Add(header_border)
        
        # Tab Control
        self.tab_control = TabControl()
        self.tab_control.Margin = Thickness(15)
        
        # Tab 1: Fill Patterns
        tab_fill = TabItem()
        tab_fill.Header = "  Fill Patterns  "
        tab_fill.Content = self._create_fill_patterns_tab()
        self.tab_control.Items.Add(tab_fill)
        
        # Tab 2: Line Styles
        tab_style = TabItem()
        tab_style.Header = "  Line Styles  "
        tab_style.Content = self._create_line_styles_tab()
        self.tab_control.Items.Add(tab_style)
        
        # Tab 3: Line Patterns
        tab_pattern = TabItem()
        tab_pattern.Header = "  Line Patterns  "
        tab_pattern.Content = self._create_line_patterns_tab()
        self.tab_control.Items.Add(tab_pattern)
        
        Grid.SetRow(self.tab_control, 1)
        main_grid.Children.Add(self.tab_control)
        
        self.Content = main_grid

    # ========================================================================
    # TAB UI GENERATORS
    # ========================================================================
    
    def _create_fill_patterns_tab(self):
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(55)))
        
        # Controls Row
        ctrl_panel = StackPanel()
        ctrl_panel.Orientation = Orientation.Horizontal
        ctrl_panel.Margin = Thickness(5)
        
        search_lbl = TextBlock()
        search_lbl.Text = "Search: "
        search_lbl.VerticalAlignment = VerticalAlignment.Center
        ctrl_panel.Children.Add(search_lbl)
        
        self.txt_search_fill = TextBox()
        self.txt_search_fill.Width = 250
        self.txt_search_fill.Height = 26
        self.txt_search_fill.VerticalContentAlignment = VerticalAlignment.Center
        self.txt_search_fill.TextChanged += self._on_fill_filter_changed
        ctrl_panel.Children.Add(self.txt_search_fill)
        
        type_lbl = TextBlock()
        type_lbl.Text = "  Type: "
        type_lbl.VerticalAlignment = VerticalAlignment.Center
        ctrl_panel.Children.Add(type_lbl)
        
        self.cmb_type_fill = ComboBox()
        self.cmb_type_fill.Width = 120
        self.cmb_type_fill.Height = 26
        self.cmb_type_fill.Items.Add("All")
        self.cmb_type_fill.Items.Add("Drafting")
        self.cmb_type_fill.Items.Add("Model")
        self.cmb_type_fill.SelectedIndex = 0
        self.cmb_type_fill.SelectionChanged += self._on_fill_filter_changed
        ctrl_panel.Children.Add(self.cmb_type_fill)
        
        self.txt_stats_fill = TextBlock()
        self.txt_stats_fill.VerticalAlignment = VerticalAlignment.Center
        self.txt_stats_fill.Margin = Thickness(20, 0, 0, 0)
        self.txt_stats_fill.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_MUTED))
        ctrl_panel.Children.Add(self.txt_stats_fill)
        
        Grid.SetRow(ctrl_panel, 0)
        grid.Children.Add(ctrl_panel)
        
        # DataGrid
        self.grid_fill = DataGrid()
        self.grid_fill.AutoGenerateColumns = False
        self.grid_fill.CanUserAddRows = False
        self.grid_fill.CanUserDeleteRows = False
        self.grid_fill.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.grid_fill.AlternatingRowBackground = SolidColorBrush(Config.hex_to_color(Config.ROW_ALT_COLOR))
        
        col_sel = DataGridCheckBoxColumn()
        col_sel.Header = "Sel"
        col_sel.Binding = Binding("is_selected")
        col_sel.Width = DataGridLength(40)
        self.grid_fill.Columns.Add(col_sel)
        
        col_name = DataGridTextColumn()
        col_name.Header = "Fill Pattern Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star)
        self.grid_fill.Columns.Add(col_name)
        
        col_type = DataGridTextColumn()
        col_type.Header = "Type"
        col_type.Binding = Binding("pattern_type")
        col_type.Width = DataGridLength(100)
        col_type.IsReadOnly = True
        self.grid_fill.Columns.Add(col_type)
        
        col_settings = DataGridTextColumn()
        col_settings.Header = "Settings"
        col_settings.Binding = Binding("settings")
        col_settings.Width = DataGridLength(150)
        col_settings.IsReadOnly = True
        self.grid_fill.Columns.Add(col_settings)
        
        col_grids = DataGridTextColumn()
        col_grids.Header = "Grids"
        col_grids.Binding = Binding("grid_count")
        col_grids.Width = DataGridLength(80)
        col_grids.IsReadOnly = True
        self.grid_fill.Columns.Add(col_grids)
        
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Binding = Binding("id")
        col_id.Width = DataGridLength(90)
        col_id.IsReadOnly = True
        self.grid_fill.Columns.Add(col_id)
        
        Grid.SetRow(self.grid_fill, 1)
        grid.Children.Add(self.grid_fill)
        
        # Footer
        footer = StackPanel()
        footer.Orientation = Orientation.Horizontal
        footer.Margin = Thickness(5)
        footer.VerticalAlignment = VerticalAlignment.Center
        
        btn_all = self._create_btn("Select All", self._on_fill_select_all)
        footer.Children.Add(btn_all)
        
        btn_none = self._create_btn("Clear All", self._on_fill_clear_all)
        footer.Children.Add(btn_none)
        
        btn_custom = self._create_btn("Select Custom", self._on_fill_select_custom)
        footer.Children.Add(btn_custom)
        
        btn_rename = self._create_btn("Rename", self._on_fill_rename, Config.SUCCESS_COLOR)
        footer.Children.Add(btn_rename)
        
        btn_dup = self._create_btn("Duplicate", self._on_fill_duplicate, Config.SECONDARY_COLOR)
        footer.Children.Add(btn_dup)
        
        btn_del = self._create_btn("Delete", self._on_fill_delete, Config.ERROR_COLOR)
        footer.Children.Add(btn_del)
        
        btn_ref = self._create_btn("Refresh", self._on_fill_refresh)
        footer.Children.Add(btn_ref)
        
        Grid.SetRow(footer, 2)
        grid.Children.Add(footer)
        return grid

    def _create_line_styles_tab(self):
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(55)))
        
        # Controls Row
        ctrl_panel = StackPanel()
        ctrl_panel.Orientation = Orientation.Horizontal
        ctrl_panel.Margin = Thickness(5)
        
        search_lbl = TextBlock()
        search_lbl.Text = "Search: "
        search_lbl.VerticalAlignment = VerticalAlignment.Center
        ctrl_panel.Children.Add(search_lbl)
        
        self.txt_search_style = TextBox()
        self.txt_search_style.Width = 250
        self.txt_search_style.Height = 26
        self.txt_search_style.VerticalContentAlignment = VerticalAlignment.Center
        self.txt_search_style.TextChanged += self._on_style_filter_changed
        ctrl_panel.Children.Add(self.txt_search_style)
        
        self.txt_stats_style = TextBlock()
        self.txt_stats_style.VerticalAlignment = VerticalAlignment.Center
        self.txt_stats_style.Margin = Thickness(20, 0, 0, 0)
        self.txt_stats_style.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_MUTED))
        ctrl_panel.Children.Add(self.txt_stats_style)
        
        Grid.SetRow(ctrl_panel, 0)
        grid.Children.Add(ctrl_panel)
        
        # DataGrid
        self.grid_style = DataGrid()
        self.grid_style.AutoGenerateColumns = False
        self.grid_style.CanUserAddRows = False
        self.grid_style.CanUserDeleteRows = False
        self.grid_style.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.grid_style.AlternatingRowBackground = SolidColorBrush(Config.hex_to_color(Config.ROW_ALT_COLOR))
        
        col_sel = DataGridCheckBoxColumn()
        col_sel.Header = "Sel"
        col_sel.Binding = Binding("is_selected")
        col_sel.Width = DataGridLength(40)
        self.grid_style.Columns.Add(col_sel)
        
        col_name = DataGridTextColumn()
        col_name.Header = "Line Style Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star)
        self.grid_style.Columns.Add(col_name)
        
        col_color = DataGridTextColumn()
        col_color.Header = "Color"
        col_color.Binding = Binding("color")
        col_color.Width = DataGridLength(120)
        col_color.IsReadOnly = True
        self.grid_style.Columns.Add(col_color)
        
        col_weight = DataGridTextColumn()
        col_weight.Header = "Weight"
        col_weight.Binding = Binding("weight")
        col_weight.Width = DataGridLength(80)
        col_weight.IsReadOnly = True
        self.grid_style.Columns.Add(col_weight)
        
        col_pattern = DataGridTextColumn()
        col_pattern.Header = "Line Pattern"
        col_pattern.Binding = Binding("pattern")
        col_pattern.Width = DataGridLength(150)
        col_pattern.IsReadOnly = True
        self.grid_style.Columns.Add(col_pattern)
        
        col_usage = DataGridTextColumn()
        col_usage.Header = "Usage (Model Curves)"
        col_usage.Binding = Binding("usage_count")
        col_usage.Width = DataGridLength(150)
        col_usage.IsReadOnly = True
        self.grid_style.Columns.Add(col_usage)
        
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Binding = Binding("id")
        col_id.Width = DataGridLength(90)
        col_id.IsReadOnly = True
        self.grid_style.Columns.Add(col_id)
        
        Grid.SetRow(self.grid_style, 1)
        grid.Children.Add(self.grid_style)
        
        # Footer
        footer = StackPanel()
        footer.Orientation = Orientation.Horizontal
        footer.Margin = Thickness(5)
        footer.VerticalAlignment = VerticalAlignment.Center
        
        btn_all = self._create_btn("Select All", self._on_style_select_all)
        footer.Children.Add(btn_all)
        
        btn_none = self._create_btn("Clear All", self._on_style_clear_all)
        footer.Children.Add(btn_none)
        
        btn_custom = self._create_btn("Select Custom", self._on_style_select_custom)
        footer.Children.Add(btn_custom)
        
        btn_rename = self._create_btn("Rename", self._on_style_rename, Config.SUCCESS_COLOR)
        footer.Children.Add(btn_rename)
        
        btn_del = self._create_btn("Delete", self._on_style_delete, Config.ERROR_COLOR)
        footer.Children.Add(btn_del)
        
        btn_calc_usage = self._create_btn("Calc Usage", self._on_style_calc_usage, Config.SECONDARY_COLOR)
        footer.Children.Add(btn_calc_usage)
        
        btn_ref = self._create_btn("Refresh", self._on_style_refresh)
        footer.Children.Add(btn_ref)
        
        Grid.SetRow(footer, 2)
        grid.Children.Add(footer)
        return grid

    def _create_line_patterns_tab(self):
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(55)))
        
        # Controls Row
        ctrl_panel = StackPanel()
        ctrl_panel.Orientation = Orientation.Horizontal
        ctrl_panel.Margin = Thickness(5)
        
        search_lbl = TextBlock()
        search_lbl.Text = "Search: "
        search_lbl.VerticalAlignment = VerticalAlignment.Center
        ctrl_panel.Children.Add(search_lbl)
        
        self.txt_search_pattern = TextBox()
        self.txt_search_pattern.Width = 250
        self.txt_search_pattern.Height = 26
        self.txt_search_pattern.VerticalContentAlignment = VerticalAlignment.Center
        self.txt_search_pattern.TextChanged += self._on_pattern_filter_changed
        ctrl_panel.Children.Add(self.txt_search_pattern)
        
        self.txt_stats_pattern = TextBlock()
        self.txt_stats_pattern.VerticalAlignment = VerticalAlignment.Center
        self.txt_stats_pattern.Margin = Thickness(20, 0, 0, 0)
        self.txt_stats_pattern.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_MUTED))
        ctrl_panel.Children.Add(self.txt_stats_pattern)
        
        Grid.SetRow(ctrl_panel, 0)
        grid.Children.Add(ctrl_panel)
        
        # DataGrid
        self.grid_pattern = DataGrid()
        self.grid_pattern.AutoGenerateColumns = False
        self.grid_pattern.CanUserAddRows = False
        self.grid_pattern.CanUserDeleteRows = False
        self.grid_pattern.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.grid_pattern.AlternatingRowBackground = SolidColorBrush(Config.hex_to_color(Config.ROW_ALT_COLOR))
        
        col_sel = DataGridCheckBoxColumn()
        col_sel.Header = "Sel"
        col_sel.Binding = Binding("is_selected")
        col_sel.Width = DataGridLength(40)
        self.grid_pattern.Columns.Add(col_sel)
        
        col_name = DataGridTextColumn()
        col_name.Header = "Line Pattern Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star)
        self.grid_pattern.Columns.Add(col_name)
        
        col_count = DataGridTextColumn()
        col_count.Header = "Segments"
        col_count.Binding = Binding("segment_count")
        col_count.Width = DataGridLength(80)
        col_count.IsReadOnly = True
        self.grid_pattern.Columns.Add(col_count)
        
        col_type = DataGridTextColumn()
        col_type.Header = "Segment Types"
        col_type.Binding = Binding("segments_type")
        col_type.Width = DataGridLength(200)
        col_type.IsReadOnly = True
        self.grid_pattern.Columns.Add(col_type)
        
        col_value = DataGridTextColumn()
        col_value.Header = "Segment Values"
        col_value.Binding = Binding("segments_value")
        col_value.Width = DataGridLength(200)
        col_value.IsReadOnly = True
        self.grid_pattern.Columns.Add(col_value)
        
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Binding = Binding("id")
        col_id.Width = DataGridLength(90)
        col_id.IsReadOnly = True
        self.grid_pattern.Columns.Add(col_id)
        
        Grid.SetRow(self.grid_pattern, 1)
        grid.Children.Add(self.grid_pattern)
        
        # Footer
        footer = StackPanel()
        footer.Orientation = Orientation.Horizontal
        footer.Margin = Thickness(5)
        footer.VerticalAlignment = VerticalAlignment.Center
        
        btn_all = self._create_btn("Select All", self._on_pattern_select_all)
        footer.Children.Add(btn_all)
        
        btn_none = self._create_btn("Clear All", self._on_pattern_clear_all)
        footer.Children.Add(btn_none)
        
        btn_custom = self._create_btn("Select Custom", self._on_pattern_select_custom)
        footer.Children.Add(btn_custom)
        
        btn_rename = self._create_btn("Rename", self._on_pattern_rename, Config.SUCCESS_COLOR)
        footer.Children.Add(btn_rename)
        
        btn_dup = self._create_btn("Duplicate", self._on_pattern_duplicate, Config.SECONDARY_COLOR)
        footer.Children.Add(btn_dup)
        
        btn_del = self._create_btn("Delete", self._on_pattern_delete, Config.ERROR_COLOR)
        footer.Children.Add(btn_del)
        
        btn_ref = self._create_btn("Refresh", self._on_pattern_refresh)
        footer.Children.Add(btn_ref)
        
        Grid.SetRow(footer, 2)
        grid.Children.Add(footer)
        return grid

    def _create_btn(self, text, handler, hex_color=None):
        btn = Button()
        btn.Content = text
        btn.Height = 28
        btn.Padding = Thickness(10, 0, 10, 0)
        btn.Margin = Thickness(0, 0, 8, 0)
        if hex_color:
            btn.Background = SolidColorBrush(Config.hex_to_color(hex_color))
            btn.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
            btn.BorderThickness = Thickness(0)
        else:
            btn.Background = SolidColorBrush(Color.FromRgb(240, 240, 240))
            btn.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        btn.Click += handler
        return btn

    # ========================================================================
    # DATA LOADING & FILTERING
    # ========================================================================
    
    def _load_all_data(self):
        self._load_fill_patterns()
        self._load_line_styles()
        self._load_line_patterns()

    def _load_fill_patterns(self):
        try:
            col = FilteredElementCollector(doc).OfClass(FillPatternElement)
            self.fill_patterns = [FillPatternItem(e) for e in col]
            self.fill_patterns.sort(key=lambda x: x.name)
            self._filter_fill_patterns()
        except Exception as ex:
            print("Error loading fill patterns: {}".format(str(ex)))
            
    def _filter_fill_patterns(self):
        query = self.txt_search_fill.Text.lower() if self.txt_search_fill.Text else ""
        type_filter = self.cmb_type_fill.SelectedItem if self.cmb_type_fill.SelectedItem else "All"
        if hasattr(type_filter, "Content"): type_filter = type_filter.Content
        
        self.filtered_fill_patterns.Clear()
        for item in self.fill_patterns:
            if query and query not in item.name.lower():
                continue
            if type_filter == "Drafting" and item.pattern_type != "Drafting":
                continue
            if type_filter == "Model" and item.pattern_type != "Model":
                continue
            self.filtered_fill_patterns.Add(item)
            
        self.grid_fill.ItemsSource = self.filtered_fill_patterns
        total = len(self.fill_patterns)
        shown = len(self.filtered_fill_patterns)
        self.txt_stats_fill.Text = "Showing {} / {} items".format(shown, total)

    def _load_line_styles(self):
        try:
            from Autodesk.Revit.DB import BuiltInCategory
            lines_category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
            
            self.line_styles = []
            if lines_category and lines_category.SubCategories:
                for subcat in lines_category.SubCategories:
                    self.line_styles.append(LineStyleItem(subcat))
            self.line_styles.sort(key=lambda x: x.name)
            self._filter_line_styles()
        except Exception as ex:
            print("Error loading line styles: {}".format(str(ex)))
            
    def _filter_line_styles(self):
        query = self.txt_search_style.Text.lower() if self.txt_search_style.Text else ""
        self.filtered_line_styles.Clear()
        for item in self.line_styles:
            if query and query not in item.name.lower():
                continue
            self.filtered_line_styles.Add(item)
            
        self.grid_style.ItemsSource = self.filtered_line_styles
        total = len(self.line_styles)
        shown = len(self.filtered_line_styles)
        self.txt_stats_style.Text = "Showing {} / {} items".format(shown, total)

    def _load_line_patterns(self):
        try:
            col = FilteredElementCollector(doc).OfClass(LinePatternElement)
            self.line_patterns = [LinePatternItem(e) for e in col]
            self.line_patterns.sort(key=lambda x: x.name)
            self._filter_line_patterns()
        except Exception as ex:
            print("Error loading line patterns: {}".format(str(ex)))
            
    def _filter_line_patterns(self):
        query = self.txt_search_pattern.Text.lower() if self.txt_search_pattern.Text else ""
        self.filtered_line_patterns.Clear()
        for item in self.line_patterns:
            if query and query not in item.name.lower():
                continue
            self.filtered_line_patterns.Add(item)
            
        self.grid_pattern.ItemsSource = self.filtered_line_patterns
        total = len(self.line_patterns)
        shown = len(self.filtered_line_patterns)
        self.txt_stats_pattern.Text = "Showing {} / {} items".format(shown, total)

    # Event filters TextChanged
    def _on_fill_filter_changed(self, s, e): self._filter_fill_patterns()
    def _on_style_filter_changed(self, s, e): self._filter_line_styles()
    def _on_pattern_filter_changed(self, s, e): self._filter_line_patterns()

    # ========================================================================
    # ACTIONS HANDLERS (FILL PATTERNS)
    # ========================================================================
    def _on_fill_select_all(self, s, e):
        for item in self.filtered_fill_patterns: item.is_selected = True
    def _on_fill_clear_all(self, s, e):
        for item in self.fill_patterns: item.is_selected = False
    def _on_fill_select_custom(self, s, e):
        for item in self.filtered_fill_patterns: item.is_selected = not item.is_system
    
    def _on_fill_rename(self, s, e):
        selected = [item for item in self.fill_patterns if item.is_selected]
        if len(selected) != 1:
            MessageBox.Show("Please select exactly one item to rename!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        if item.is_system:
            MessageBox.Show("Cannot rename system patterns!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
            
        new_name = forms.ask_for_string(prompt="Enter new name:", default=item.name, title="Rename Fill Pattern")
        if not new_name or new_name.strip() == "" or new_name == item.name:
            return
            
        new_name = re.sub(r'[\/:*?"<>|\\\[\]]', '', new_name).strip()
        
        # Conflict check
        for pat in self.fill_patterns:
            if pat.name == new_name:
                MessageBox.Show("A fill pattern with this name already exists!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                return
                
        t = Transaction(doc, "Rename Fill Pattern")
        t.Start()
        try:
            item.element.Name = new_name
            t.Commit()
            self._load_fill_patterns()
            MessageBox.Show("Fill Pattern renamed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_fill_duplicate(self, s, e):
        selected = [item for item in self.fill_patterns if item.is_selected]
        if not selected:
            MessageBox.Show("Please select at least one item to duplicate!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        t = Transaction(doc, "Duplicate Fill Patterns")
        t.Start()
        try:
            success = 0
            for item in selected:
                try:
                    new_name = "Copy of " + item.name
                    item.element.Duplicate(new_name)
                    success += 1
                except:
                    pass
            t.Commit()
            self._load_fill_patterns()
            MessageBox.Show("Duplicated {} patterns!".format(success), "Result", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_fill_delete(self, s, e):
        selected = [item for item in self.fill_patterns if item.is_selected]
        if not selected:
            MessageBox.Show("Please select at least one item to delete!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        custom_selected = [i for i in selected if not i.is_system]
        if not custom_selected:
            MessageBox.Show("None of the selected items can be deleted (system elements).", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
            
        res = MessageBox.Show("Delete {} custom fill pattern(s)?".format(len(custom_selected)), "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        if res != MessageBoxResult.Yes:
            return
            
        t = Transaction(doc, "Delete Fill Patterns")
        t.Start()
        try:
            deleted = 0
            failed = 0
            for item in custom_selected:
                try:
                    doc.Delete(item.element.Id)
                    deleted += 1
                except:
                    failed += 1
            t.Commit()
            self._load_fill_patterns()
            MessageBox.Show("Deleted: {}\nFailed: {} (may be in use)".format(deleted, failed), "Result", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_fill_refresh(self, s, e):
        self._load_fill_patterns()

    # ========================================================================
    # ACTIONS HANDLERS (LINE STYLES)
    # ========================================================================
    def _on_style_select_all(self, s, e):
        for item in self.filtered_line_styles: item.is_selected = True
    def _on_style_clear_all(self, s, e):
        for item in self.line_styles: item.is_selected = False
    def _on_style_select_custom(self, s, e):
        for item in self.filtered_line_styles: item.is_selected = not item.is_system

    def _on_style_rename(self, s, e):
        selected = [item for item in self.line_styles if item.is_selected]
        if len(selected) != 1:
            MessageBox.Show("Please select exactly one item to rename!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        item = selected[0]
        if item.is_system:
            MessageBox.Show("Cannot rename system line styles!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
            
        new_name = forms.ask_for_string(prompt="Enter new name:", default=item.name, title="Rename Line Style")
        if not new_name or new_name.strip() == "" or new_name == item.name:
            return
            
        new_name = re.sub(r'[\/:*?"<>|\\\[\]]', '', new_name).strip()
        
        # Conflict check
        for style in self.line_styles:
            if style.name == new_name:
                MessageBox.Show("A line style with this name already exists!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                return
                
        # Perform rename using copy-transfer-delete logic
        old_style = item.category
        from Autodesk.Revit.DB import BuiltInCategory
        lines_category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        
        tg = TransactionGroup(doc, "Rename Line Style")
        tg.Start()
        try:
            # 1. Create subcategory
            t1 = Transaction(doc, "Create New Line Style")
            t1.Start()
            new_subcategory = doc.Settings.Categories.NewSubcategory(lines_category, new_name)
            new_subcategory.LineColor = old_style.LineColor
            new_subcategory.SetLineWeight(old_style.GetLineWeight(GraphicsStyleType.Projection), GraphicsStyleType.Projection)
            pat_id = old_style.GetLinePatternId(GraphicsStyleType.Projection)
            if pat_id and _eid_int(pat_id) > 0:
                new_subcategory.SetLinePatternId(pat_id, GraphicsStyleType.Projection)
            t1.Commit()
            
            # 2. Find and transfer curves using old style
            lines_to_change = []
            col = FilteredElementCollector(doc).OfClass(CurveElement)
            for curve in col:
                try:
                    style_param = curve.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE)
                    if style_param:
                        style_id = style_param.AsElementId()
                        if style_id:
                            style_elem = doc.GetElement(style_id)
                            if style_elem and hasattr(style_elem, 'GraphicsStyleCategory'):
                                style_cat = style_elem.GraphicsStyleCategory
                                if style_cat and _eid_int(style_cat.Id) == _eid_int(old_style.Id):
                                    lines_to_change.append(curve)
                except:
                    pass
            
            lines_changed = 0
            if lines_to_change:
                t2 = Transaction(doc, "Transfer Lines to New Style")
                t2.Start()
                new_graphics_style = new_subcategory.GetGraphicsStyle(GraphicsStyleType.Projection)
                for line in lines_to_change:
                    try:
                        line.LineStyle = new_graphics_style
                        lines_changed += 1
                    except:
                        pass
                t2.Commit()
                
            # 3. Delete old subcategory
            t3 = Transaction(doc, "Delete Old Line Style")
            t3.Start()
            doc.Delete(old_style.Id)
            t3.Commit()
            
            tg.Assimilate()
            self._load_line_styles()
            MessageBox.Show("Line style renamed successfully!\nLines transferred: {}".format(lines_changed), "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            tg.RollBack()
            MessageBox.Show("Error renaming line style:\n\n{}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_style_delete(self, s, e):
        selected = [item for item in self.line_styles if item.is_selected]
        if not selected:
            MessageBox.Show("Please select at least one item to delete!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        custom_selected = [i for i in selected if not i.is_system]
        if not custom_selected:
            MessageBox.Show("None of the selected items can be deleted (system elements).", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
            
        res = MessageBox.Show("Delete {} custom line style(s)?".format(len(custom_selected)), "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        if res != MessageBoxResult.Yes:
            return
            
        t = Transaction(doc, "Delete Line Styles")
        t.Start()
        try:
            deleted = 0
            failed = 0
            for item in custom_selected:
                try:
                    doc.Delete(item.category.Id)
                    deleted += 1
                except:
                    failed += 1
            t.Commit()
            self._load_line_styles()
            MessageBox.Show("Deleted: {}\nFailed: {} (may be in use)".format(deleted, failed), "Result", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_style_calc_usage(self, s, e):
        # Reset
        for item in self.line_styles:
            item.usage_count = 0
            
        lookup = {}
        for item in self.line_styles:
            if item.category:
                lookup[_eid_int(item.category.Id)] = item
                
        # Count usage on CurveElements
        try:
            col = FilteredElementCollector(doc).OfClass(CurveElement)
            usage_dict = {}
            for curve in col:
                try:
                    style_param = curve.get_Parameter(BuiltInParameter.BUILDING_CURVE_GSTYLE)
                    if style_param:
                        style_id = style_param.AsElementId()
                        if style_id and _eid_int(style_id) > 0:
                            style_elem = doc.GetElement(style_id)
                            if style_elem and hasattr(style_elem, 'GraphicsStyleCategory'):
                                style_cat = style_elem.GraphicsStyleCategory
                                if style_cat:
                                    cat_id = _eid_int(style_cat.Id)
                                    if cat_id in lookup:
                                        usage_dict[cat_id] = usage_dict.get(cat_id, 0) + 1
                except:
                    pass
            
            for cat_id, count in usage_dict.items():
                item = lookup.get(cat_id)
                if item:
                    item.usage_count = count
            
            # Rebind to refresh UI
            self._filter_line_styles()
            MessageBox.Show("Usage calculation completed!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            MessageBox.Show("Error calculating usage: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_style_refresh(self, s, e):
        self._load_line_styles()

    # ========================================================================
    # ACTIONS HANDLERS (LINE PATTERNS)
    # ========================================================================
    def _on_pattern_select_all(self, s, e):
        for item in self.filtered_line_patterns: item.is_selected = True
    def _on_pattern_clear_all(self, s, e):
        for item in self.line_patterns: item.is_selected = False
    def _on_pattern_select_custom(self, s, e):
        for item in self.filtered_line_patterns: item.is_selected = not item.is_system

    def _on_pattern_rename(self, s, e):
        selected = [item for item in self.line_patterns if item.is_selected]
        if len(selected) != 1:
            MessageBox.Show("Please select exactly one item to rename!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        if item.is_system:
            MessageBox.Show("Cannot rename system line patterns!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
            
        new_name = forms.ask_for_string(prompt="Enter new name:", default=item.name, title="Rename Line Pattern")
        if not new_name or new_name.strip() == "" or new_name == item.name:
            return
            
        new_name = re.sub(r'[\/:*?"<>|\\\[\]]', '', new_name).strip()
        
        # Conflict check
        for pat in self.line_patterns:
            if pat.name == new_name:
                MessageBox.Show("A line pattern with this name already exists!", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                return
                
        t = Transaction(doc, "Rename Line Pattern")
        t.Start()
        try:
            item.element.Name = new_name
            t.Commit()
            self._load_line_patterns()
            MessageBox.Show("Line pattern renamed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_pattern_duplicate(self, s, e):
        selected = [item for item in self.line_patterns if item.is_selected]
        if not selected:
            MessageBox.Show("Please select at least one item to duplicate!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        t = Transaction(doc, "Duplicate Line Patterns")
        t.Start()
        try:
            success = 0
            for item in selected:
                try:
                    new_name = "Copy of " + item.name
                    item.element.Duplicate(new_name)
                    success += 1
                except:
                    pass
            t.Commit()
            self._load_line_patterns()
            MessageBox.Show("Duplicated {} line patterns!".format(success), "Result", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_pattern_delete(self, s, e):
        selected = [item for item in self.line_patterns if item.is_selected]
        if not selected:
            MessageBox.Show("Please select at least one item to delete!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        custom_selected = [i for i in selected if not i.is_system]
        if not custom_selected:
            MessageBox.Show("None of the selected items can be deleted (system elements).", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
            
        res = MessageBox.Show("Delete {} custom line pattern(s)?".format(len(custom_selected)), "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
        if res != MessageBoxResult.Yes:
            return
            
        t = Transaction(doc, "Delete Line Patterns")
        t.Start()
        try:
            deleted = 0
            failed = 0
            for item in custom_selected:
                try:
                    doc.Delete(item.element.Id)
                    deleted += 1
                except:
                    failed += 1
            t.Commit()
            self._load_line_patterns()
            MessageBox.Show("Deleted: {}\nFailed: {} (may be in use)".format(deleted, failed), "Result", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_pattern_refresh(self, s, e):
        self._load_line_patterns()


if __name__ == '__main__':
    try:
        window = StyleManagerWindow()
        window.ShowDialog()
    except Exception as e:
        import traceback
        MessageBox.Show("Fatal Error starting Style Manager:\n\n{}\n\n{}".format(str(e), traceback.format_exc()), "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error)
