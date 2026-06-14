# -*- coding: utf-8 -*-
"""
Line Pattern Manager v2.0
Manage Line Patterns with beautiful UI - Sheet Manager Style
Compatible with Revit 2024/2025/2026/2027

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""
__title__ = "Line Pattern\nManage"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Manage and rename Line Patterns - By DQT"

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
                                      ScrollViewer, ListBox, ListBoxItem,
                                      TabControl, TabItem, CheckBox)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

import re

from pyrevit import revit, forms
from Autodesk.Revit.DB import (FilteredElementCollector, LinePatternElement,
                                Transaction, ElementId)

doc = revit.doc


# ============================================================================
# REVIT API COMPATIBILITY HELPERS
# ============================================================================

def _eid_int(element_id):
    """Get integer value from ElementId - compatible with all Revit versions"""
    try:
        return element_id.Value  # Revit 2025+
    except AttributeError:
        return element_id.IntegerValue  # Revit 2024 and earlier


# ============================================================================
# CONFIGURATION
# ============================================================================

class Config(object):
    """Color scheme and settings"""
    PRIMARY_COLOR = "#0F172A"
    SECONDARY_COLOR = "#E5B85C"
    BACKGROUND_COLOR = "#F8FAFC"
    BORDER_COLOR = "#CBD5E1"
    TEXT_DARK = "#5D4E37"
    TEXT_LIGHT = "#888888"
    SUCCESS_COLOR = "#10B981"
    WARNING_COLOR = "#F59E0B"
    ERROR_COLOR = "#FF6B6B"
    ROW_ALT_COLOR = "#F1F5F9"
    WHITE = "#FFFFFF"
    
    @staticmethod
    def hex_to_color(hex_color):
        hex_color = hex_color.replace("#", "")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return Color.FromArgb(255, r, g, b)


# ============================================================================
# LINE PATTERN ITEM (Data Model)
# ============================================================================

class LinePatternItem(INotifyPropertyChanged):
    """Wrapper class for LinePatternElement with WPF binding support"""
    
    # System patterns that cannot be renamed/deleted
    SYSTEM_PATTERNS = ["Solid", "Dash", "Dot", "Dash dot", "Dash dot dot"]
    
    def __init__(self, element):
        self._property_changed_handlers = []
        self._element = element
        self._is_selected = False
        
        try:
            self._name = element.Name if element.Name else "Unnamed"
        except:
            self._name = "Unnamed"
        
        self._id = _eid_int(element.Id)  # FIXED: Use compatibility helper
        self._category = self._get_category()  # System or Custom
        self._segment_count = self._get_segment_count()
        self._segments_type = self._get_segments_type()  # Dash, Space, Dot sequence
        self._segments_value = self._get_segments_value()  # Length values in mm
        self._is_system = self._name in self.SYSTEM_PATTERNS
        self._usage_count = 0
    
    def _get_category(self):
        """Get pattern category (System or Custom)"""
        try:
            if self._name in self.SYSTEM_PATTERNS:
                return "System"
            return "Custom"
        except:
            return "Unknown"
    
    def _get_segment_count(self):
        """Get number of segments"""
        try:
            line_pattern = self._element.GetLinePattern()
            if line_pattern:
                segments = line_pattern.GetSegments()
                if segments:
                    return len(list(segments))
            return 0
        except:
            return 0
    
    def _get_segments_type(self):
        """Get line pattern segment types as display string (like Revit)"""
        try:
            line_pattern = self._element.GetLinePattern()
            if not line_pattern:
                return "Solid"
            
            segments = line_pattern.GetSegments()
            if not segments or len(list(segments)) == 0:
                return "Solid"
            
            types = []
            for seg in segments:
                seg_type = seg.Type.ToString()
                types.append(seg_type)
            
            return ", ".join(types) if types else "Solid"
        except:
            return "Solid"
    
    def _get_segments_value(self):
        """Get line pattern segment values (lengths) as display string"""
        try:
            line_pattern = self._element.GetLinePattern()
            if not line_pattern:
                return "-"
            
            segments = line_pattern.GetSegments()
            if not segments or len(list(segments)) == 0:
                return "-"
            
            values = []
            for seg in segments:
                length = seg.Length * 304.8  # to mm
                values.append("{:.2f}mm".format(length))
            
            return ", ".join(values) if values else "-"
        except:
            return "-"
    
    def _get_segments_type_short(self):
        """Get shortened segment type for naming (D=Dash, S=Space, Dot=Dot) - NO COMMAS"""
        try:
            line_pattern = self._element.GetLinePattern()
            if not line_pattern:
                return "Solid"
            
            segments = line_pattern.GetSegments()
            if not segments or len(list(segments)) == 0:
                return "Solid"
            
            types = []
            for seg in segments:
                seg_type = seg.Type.ToString()
                if seg_type == "Dash":
                    types.append("D")
                elif seg_type == "Space":
                    types.append("S")
                elif seg_type == "Dot":
                    types.append("Dot")
                else:
                    types.append(seg_type[0])
            
            return "-".join(types) if types else "Solid"
        except:
            return "Solid"
    
    def _get_segments_value_short(self):
        """Get shortened segment values for naming - NO COMMAS"""
        try:
            line_pattern = self._element.GetLinePattern()
            if not line_pattern:
                return ""
            
            segments = line_pattern.GetSegments()
            if not segments or len(list(segments)) == 0:
                return ""
            
            values = []
            for seg in segments:
                length = seg.Length * 304.8  # to mm
                values.append("{:.1f}".format(length))
            
            return "-".join(values) if values else ""
        except:
            return ""
    
    # Properties
    @property
    def element(self):
        return self._element
    
    @property
    def id(self):
        return self._id
    
    @property
    def name(self):
        return self._name
    
    @name.setter
    def name(self, value):
        if self._name != value:
            self._name = value
            self._notify_property_changed("name")
    
    @property
    def category(self):
        return self._category
    
    @property
    def segment_count(self):
        return self._segment_count
    
    @property
    def segments_type(self):
        return self._segments_type
    
    @property
    def segments_value(self):
        return self._segments_value
    
    @property
    def is_system(self):
        return self._is_system
    
    @property
    def is_selected(self):
        return self._is_selected
    
    @is_selected.setter
    def is_selected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify_property_changed("is_selected")
    
    @property
    def usage_count(self):
        return self._usage_count
    
    @usage_count.setter
    def usage_count(self, value):
        if self._usage_count != value:
            self._usage_count = value
            self._notify_property_changed("usage_count")
    
    # INotifyPropertyChanged implementation
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def _notify_property_changed(self, property_name):
        for handler in self._property_changed_handlers:
            handler(self, PropertyChangedEventArgs(property_name))


# ============================================================================
# BATCH RENAME DIALOG
# ============================================================================

class BatchRenameDialog(Window):
    """Dialog for batch renaming line patterns"""
    
    def __init__(self, items, parent=None):
        self.items = items
        self.parent = parent
        self.result_items = []
        
        self.Title = "Batch Rename Line Patterns - By DQT"
        self.Width = 900
        self.Height = 700
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Config.hex_to_color(Config.BACKGROUND_COLOR))
        
        self._create_ui()
        self._setup_data()
    
    def _create_ui(self):
        main_grid = Grid()
        main_grid.Margin = Thickness(0)
        
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(280)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        
        # Header
        header_border = Border()
        header_border.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        header_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        header_border.BorderThickness = Thickness(0, 0, 0, 2)
        
        header_text = TextBlock()
        header_text.Text = "Batch Rename Line Patterns"
        header_text.FontSize = 20
        header_text.FontWeight = FontWeights.Bold
        header_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        header_text.VerticalAlignment = VerticalAlignment.Center
        header_text.HorizontalAlignment = HorizontalAlignment.Center
        
        header_border.Child = header_text
        Grid.SetRow(header_border, 0)
        main_grid.Children.Add(header_border)
        
        # Content area
        content_grid = Grid()
        content_grid.Margin = Thickness(20)
        Grid.SetRow(content_grid, 1)
        main_grid.Children.Add(content_grid)
        
        # Data grid for current names
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        
        self.data_grid = DataGrid()
        self.data_grid.AutoGenerateColumns = False
        self.data_grid.CanUserAddRows = False
        self.data_grid.CanUserDeleteRows = False
        self.data_grid.IsReadOnly = True
        self.data_grid.SelectionMode = System.Windows.Controls.DataGridSelectionMode.Extended
        self.data_grid.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.data_grid.RowBackground = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.data_grid.AlternatingRowBackground = SolidColorBrush(Config.hex_to_color(Config.ROW_ALT_COLOR))
        self.data_grid.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        self.data_grid.BorderThickness = Thickness(1)
        
        col_name = DataGridTextColumn()
        col_name.Header = "Current Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star)
        self.data_grid.Columns.Add(col_name)
        
        col_type = DataGridTextColumn()
        col_type.Header = "Segments"
        col_type.Binding = Binding("segments_type")
        col_type.Width = DataGridLength(200)
        self.data_grid.Columns.Add(col_type)
        
        col_value = DataGridTextColumn()
        col_value.Header = "Values"
        col_value.Binding = Binding("segments_value")
        col_value.Width = DataGridLength(200)
        self.data_grid.Columns.Add(col_value)
        
        scroll.Content = self.data_grid
        content_grid.Children.Add(scroll)
        
        # Options panel
        options_border = Border()
        options_border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        options_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        options_border.BorderThickness = Thickness(1)
        options_border.Margin = Thickness(20, 0, 20, 0)
        Grid.SetRow(options_border, 2)
        
        options_stack = StackPanel()
        options_stack.Margin = Thickness(20)
        
        # Title
        title = TextBlock()
        title.Text = "Rename Options"
        title.FontSize = 16
        title.FontWeight = FontWeights.Bold
        title.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        title.Margin = Thickness(0, 0, 0, 15)
        options_stack.Children.Add(title)
        
        # Prefix
        prefix_panel = StackPanel()
        prefix_panel.Orientation = Orientation.Horizontal
        prefix_panel.Margin = Thickness(0, 0, 0, 10)
        
        self.chk_prefix = CheckBox()
        self.chk_prefix.Content = "Add Prefix:"
        self.chk_prefix.VerticalAlignment = VerticalAlignment.Center
        self.chk_prefix.Margin = Thickness(0, 0, 10, 0)
        prefix_panel.Children.Add(self.chk_prefix)
        
        self.txt_prefix = TextBox()
        self.txt_prefix.Width = 200
        self.txt_prefix.Padding = Thickness(5)
        self.txt_prefix.IsEnabled = False
        prefix_panel.Children.Add(self.txt_prefix)
        
        self.chk_prefix.Checked += self._on_prefix_checked
        self.chk_prefix.Unchecked += self._on_prefix_unchecked
        
        options_stack.Children.Add(prefix_panel)
        
        # Suffix
        suffix_panel = StackPanel()
        suffix_panel.Orientation = Orientation.Horizontal
        suffix_panel.Margin = Thickness(0, 0, 0, 10)
        
        self.chk_suffix = CheckBox()
        self.chk_suffix.Content = "Add Suffix:"
        self.chk_suffix.VerticalAlignment = VerticalAlignment.Center
        self.chk_suffix.Margin = Thickness(0, 0, 10, 0)
        suffix_panel.Children.Add(self.chk_suffix)
        
        self.txt_suffix = TextBox()
        self.txt_suffix.Width = 200
        self.txt_suffix.Padding = Thickness(5)
        self.txt_suffix.IsEnabled = False
        suffix_panel.Children.Add(self.txt_suffix)
        
        self.chk_suffix.Checked += self._on_suffix_checked
        self.chk_suffix.Unchecked += self._on_suffix_unchecked
        
        options_stack.Children.Add(suffix_panel)
        
        # Find/Replace
        find_panel = StackPanel()
        find_panel.Orientation = Orientation.Horizontal
        find_panel.Margin = Thickness(0, 0, 0, 10)
        
        self.chk_find = CheckBox()
        self.chk_find.Content = "Find & Replace:"
        self.chk_find.VerticalAlignment = VerticalAlignment.Center
        self.chk_find.Margin = Thickness(0, 0, 10, 0)
        find_panel.Children.Add(self.chk_find)
        
        self.txt_find = TextBox()
        self.txt_find.Width = 150
        self.txt_find.Padding = Thickness(5)
        self.txt_find.IsEnabled = False
        find_panel.Children.Add(self.txt_find)
        
        arrow = TextBlock()
        arrow.Text = " → "
        arrow.VerticalAlignment = VerticalAlignment.Center
        arrow.Margin = Thickness(5, 0, 5, 0)
        find_panel.Children.Add(arrow)
        
        self.txt_replace = TextBox()
        self.txt_replace.Width = 150
        self.txt_replace.Padding = Thickness(5)
        self.txt_replace.IsEnabled = False
        find_panel.Children.Add(self.txt_replace)
        
        self.chk_find.Checked += self._on_find_checked
        self.chk_find.Unchecked += self._on_find_unchecked
        
        options_stack.Children.Add(find_panel)
        
        # Auto-naming option
        auto_panel = StackPanel()
        auto_panel.Orientation = Orientation.Horizontal
        auto_panel.Margin = Thickness(0, 15, 0, 10)
        
        self.chk_auto = CheckBox()
        self.chk_auto.Content = "Auto-name based on segments (Format: Type-Values)"
        self.chk_auto.VerticalAlignment = VerticalAlignment.Center
        self.chk_auto.FontWeight = FontWeights.Bold
        self.chk_auto.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        auto_panel.Children.Add(self.chk_auto)
        
        self.chk_auto.Checked += self._on_auto_checked
        self.chk_auto.Unchecked += self._on_auto_unchecked
        
        options_stack.Children.Add(auto_panel)
        
        # Preview button
        preview_btn = Button()
        preview_btn.Content = "Preview New Names"
        preview_btn.Padding = Thickness(20, 8, 20, 8)
        preview_btn.Margin = Thickness(0, 15, 0, 0)
        preview_btn.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        preview_btn.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        preview_btn.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        preview_btn.Click += self._on_preview
        options_stack.Children.Add(preview_btn)
        
        options_border.Child = options_stack
        main_grid.Children.Add(options_border)
        
        # Footer buttons
        footer_border = Border()
        footer_border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        footer_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        footer_border.BorderThickness = Thickness(0, 1, 0, 0)
        Grid.SetRow(footer_border, 3)
        
        footer_stack = StackPanel()
        footer_stack.Orientation = Orientation.Horizontal
        footer_stack.HorizontalAlignment = HorizontalAlignment.Center
        footer_stack.VerticalAlignment = VerticalAlignment.Center
        
        btn_apply = Button()
        btn_apply.Content = "Apply Rename"
        btn_apply.Width = 120
        btn_apply.Height = 35
        btn_apply.Margin = Thickness(10, 0, 10, 0)
        btn_apply.Background = SolidColorBrush(Config.hex_to_color(Config.SUCCESS_COLOR))
        btn_apply.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_apply.BorderThickness = Thickness(0)
        btn_apply.Click += self._on_apply
        footer_stack.Children.Add(btn_apply)
        
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 120
        btn_cancel.Height = 35
        btn_cancel.Margin = Thickness(10, 0, 10, 0)
        btn_cancel.Background = SolidColorBrush(Config.hex_to_color(Config.TEXT_LIGHT))
        btn_cancel.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_cancel.BorderThickness = Thickness(0)
        btn_cancel.Click += self._on_cancel
        footer_stack.Children.Add(btn_cancel)
        
        footer_border.Child = footer_stack
        main_grid.Children.Add(footer_border)
        
        self.Content = main_grid
    
    def _setup_data(self):
        self.data_grid.ItemsSource = self.items
    
    def _on_prefix_checked(self, sender, args):
        self.txt_prefix.IsEnabled = True
    
    def _on_prefix_unchecked(self, sender, args):
        self.txt_prefix.IsEnabled = False
    
    def _on_suffix_checked(self, sender, args):
        self.txt_suffix.IsEnabled = True
    
    def _on_suffix_unchecked(self, sender, args):
        self.txt_suffix.IsEnabled = False
    
    def _on_find_checked(self, sender, args):
        self.txt_find.IsEnabled = True
        self.txt_replace.IsEnabled = True
    
    def _on_find_unchecked(self, sender, args):
        self.txt_find.IsEnabled = False
        self.txt_replace.IsEnabled = False
    
    def _on_auto_checked(self, sender, args):
        # Disable other options when auto is checked
        self.chk_prefix.IsEnabled = False
        self.chk_suffix.IsEnabled = False
        self.chk_find.IsEnabled = False
        self.txt_prefix.IsEnabled = False
        self.txt_suffix.IsEnabled = False
        self.txt_find.IsEnabled = False
        self.txt_replace.IsEnabled = False
    
    def _on_auto_unchecked(self, sender, args):
        # Re-enable other options
        self.chk_prefix.IsEnabled = True
        self.chk_suffix.IsEnabled = True
        self.chk_find.IsEnabled = True
    
    def _generate_new_name(self, item):
        """Generate new name based on options"""
        if self.chk_auto.IsChecked:
            # Auto-naming based on segments
            seg_type = item._get_segments_type_short()
            seg_value = item._get_segments_value_short()
            
            if seg_value:
                return "{}-{}".format(seg_type, seg_value)
            else:
                return seg_type
        else:
            # Manual prefix/suffix/find-replace
            new_name = item.name
            
            if self.chk_find.IsChecked:
                find_text = self.txt_find.Text if self.txt_find.Text else ""
                replace_text = self.txt_replace.Text if self.txt_replace.Text else ""
                if find_text:
                    new_name = new_name.replace(find_text, replace_text)
            
            if self.chk_prefix.IsChecked:
                prefix = self.txt_prefix.Text if self.txt_prefix.Text else ""
                new_name = prefix + new_name
            
            if self.chk_suffix.IsChecked:
                suffix = self.txt_suffix.Text if self.txt_suffix.Text else ""
                new_name = new_name + suffix
            
            return new_name
    
    def _on_preview(self, sender, args):
        """Show preview of new names"""
        preview_text = "New names preview:\n\n"
        
        for item in self.items:
            new_name = self._generate_new_name(item)
            preview_text += "{} → {}\n".format(item.name, new_name)
        
        MessageBox.Show(preview_text, "Preview", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_apply(self, sender, args):
        """Apply batch rename"""
        # Validate
        if not self.chk_prefix.IsChecked and not self.chk_suffix.IsChecked and not self.chk_find.IsChecked and not self.chk_auto.IsChecked:
            MessageBox.Show("Please select at least one rename option!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Generate result items
        self.result_items = []
        for item in self.items:
            new_name = self._generate_new_name(item)
            
            # Clean name
            invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
            for char in invalid_chars:
                new_name = new_name.replace(char, '')
            new_name = new_name.strip()
            
            if new_name and new_name != item.name:
                self.result_items.append((item, new_name))
        
        if not self.result_items:
            MessageBox.Show("No changes to apply!", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Check for duplicates
        new_names = [new_name for _, new_name in self.result_items]
        if len(new_names) != len(set(new_names)):
            MessageBox.Show("Duplicate names detected!\nPlease adjust your options.",
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Confirm
        result = MessageBox.Show(
            "Rename {} line patterns?".format(len(self.result_items)),
            "Confirm",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        
        if result == MessageBoxResult.Yes:
            self._apply_rename()
    
    def _apply_rename(self):
        """Apply the rename operation"""
        t = Transaction(doc, "DQT - Batch Rename Line Patterns")
        t.Start()
        
        try:
            success_count = 0
            error_count = 0
            
            for item, new_name in self.result_items:
                try:
                    item.element.Name = new_name
                    success_count += 1
                except Exception as ex:
                    print("Error renaming {}: {}".format(item.name, str(ex)))
                    error_count += 1
            
            t.Commit()
            
            msg = "Renamed: {}".format(success_count)
            if error_count > 0:
                msg += "\nFailed: {}".format(error_count)
            
            MessageBox.Show(msg, "Result", MessageBoxButton.OK, MessageBoxImage.Information)
            
            if self.parent:
                self.parent._load_data()
            
            self.Close()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_cancel(self, sender, args):
        self.Close()


# ============================================================================
# MAIN WINDOW
# ============================================================================

class LinePatternManagerWindow(Window):
    """Main window for Line Pattern Manager"""
    
    def __init__(self):
        self.Title = "Line Pattern Manager - By DQT"
        self.Width = 1200
        self.Height = 800
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Config.hex_to_color(Config.BACKGROUND_COLOR))
        
        self.all_items = []
        self.filtered_items = ObservableCollection[object]()
        
        self._create_ui()
        self._load_data()
    
    def _create_ui(self):
        """Create the main UI"""
        main_grid = Grid()
        main_grid.Margin = Thickness(0)
        
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(100)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        
        # Header
        header_border = Border()
        header_border.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        header_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        header_border.BorderThickness = Thickness(0, 0, 0, 2)
        
        header_text = TextBlock()
        header_text.Text = "Line Pattern Manager"
        header_text.FontSize = 24
        header_text.FontWeight = FontWeights.Bold
        header_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        header_text.VerticalAlignment = VerticalAlignment.Center
        header_text.HorizontalAlignment = HorizontalAlignment.Center
        
        header_border.Child = header_text
        Grid.SetRow(header_border, 0)
        main_grid.Children.Add(header_border)
        
        # Stats and filters panel
        stats_border = Border()
        stats_border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        stats_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        stats_border.BorderThickness = Thickness(0, 0, 0, 1)
        stats_border.Padding = Thickness(20, 15, 20, 15)
        Grid.SetRow(stats_border, 1)
        
        stats_stack = StackPanel()
        stats_stack.Orientation = Orientation.Vertical
        
        # Row 1: Stats
        stats_row = StackPanel()
        stats_row.Orientation = Orientation.Horizontal
        stats_row.Margin = Thickness(0, 0, 0, 10)
        
        self._add_stat_card(stats_row, "Total", "0", "txt_total")
        self._add_stat_card(stats_row, "Selected", "0", "txt_selected")
        self._add_stat_card(stats_row, "System", "0", "txt_system")
        self._add_stat_card(stats_row, "Custom", "0", "txt_custom")
        
        stats_stack.Children.Add(stats_row)
        
        # Row 2: Filters
        filter_row = StackPanel()
        filter_row.Orientation = Orientation.Horizontal
        filter_row.HorizontalAlignment = HorizontalAlignment.Left
        
        search_label = TextBlock()
        search_label.Text = "Search:"
        search_label.VerticalAlignment = VerticalAlignment.Center
        search_label.Margin = Thickness(0, 0, 10, 0)
        search_label.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        filter_row.Children.Add(search_label)
        
        self.txt_search = TextBox()
        self.txt_search.Width = 300
        self.txt_search.Padding = Thickness(5)
        self.txt_search.Margin = Thickness(0, 0, 20, 0)
        self.txt_search.TextChanged += self._on_filter_changed
        filter_row.Children.Add(self.txt_search)
        
        filter_label = TextBlock()
        filter_label.Text = "Category:"
        filter_label.VerticalAlignment = VerticalAlignment.Center
        filter_label.Margin = Thickness(0, 0, 10, 0)
        filter_label.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        filter_row.Children.Add(filter_label)
        
        self.cmb_category = ComboBox()
        self.cmb_category.Width = 150
        self.cmb_category.Margin = Thickness(0, 0, 20, 0)
        
        item_all = ComboBoxItem()
        item_all.Content = "All"
        self.cmb_category.Items.Add(item_all)
        
        item_system = ComboBoxItem()
        item_system.Content = "System"
        self.cmb_category.Items.Add(item_system)
        
        item_custom = ComboBoxItem()
        item_custom.Content = "Custom"
        self.cmb_category.Items.Add(item_custom)
        
        self.cmb_category.SelectedIndex = 0
        self.cmb_category.SelectionChanged += self._on_filter_changed
        filter_row.Children.Add(self.cmb_category)
        
        stats_stack.Children.Add(filter_row)
        
        stats_border.Child = stats_stack
        main_grid.Children.Add(stats_border)
        
        # Data grid area
        grid_border = Border()
        grid_border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        grid_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        grid_border.BorderThickness = Thickness(1)
        grid_border.Margin = Thickness(20, 20, 20, 10)
        Grid.SetRow(grid_border, 2)
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        
        self.data_grid = DataGrid()
        self.data_grid.AutoGenerateColumns = False
        self.data_grid.CanUserAddRows = False
        self.data_grid.CanUserDeleteRows = False
        self.data_grid.IsReadOnly = True
        self.data_grid.SelectionMode = System.Windows.Controls.DataGridSelectionMode.Extended
        self.data_grid.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.data_grid.RowBackground = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        self.data_grid.AlternatingRowBackground = SolidColorBrush(Config.hex_to_color(Config.ROW_ALT_COLOR))
        self.data_grid.BorderThickness = Thickness(0)
        self.data_grid.SelectionChanged += self._on_selection_changed
        
        # Columns
        col_name = DataGridTextColumn()
        col_name.Header = "Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star)
        self.data_grid.Columns.Add(col_name)
        
        col_category = DataGridTextColumn()
        col_category.Header = "Category"
        col_category.Binding = Binding("category")
        col_category.Width = DataGridLength(100)
        self.data_grid.Columns.Add(col_category)
        
        col_count = DataGridTextColumn()
        col_count.Header = "Segments"
        col_count.Binding = Binding("segment_count")
        col_count.Width = DataGridLength(100)
        self.data_grid.Columns.Add(col_count)
        
        col_type = DataGridTextColumn()
        col_type.Header = "Segment Types"
        col_type.Binding = Binding("segments_type")
        col_type.Width = DataGridLength(200)
        self.data_grid.Columns.Add(col_type)
        
        col_value = DataGridTextColumn()
        col_value.Header = "Segment Values"
        col_value.Binding = Binding("segments_value")
        col_value.Width = DataGridLength(250)
        self.data_grid.Columns.Add(col_value)
        
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Binding = Binding("id")
        col_id.Width = DataGridLength(100)
        self.data_grid.Columns.Add(col_id)
        
        scroll.Content = self.data_grid
        grid_border.Child = scroll
        main_grid.Children.Add(grid_border)
        
        # Footer with buttons
        footer_border = Border()
        footer_border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        footer_border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        footer_border.BorderThickness = Thickness(0, 1, 0, 0)
        Grid.SetRow(footer_border, 3)
        
        footer_stack = StackPanel()
        footer_stack.Orientation = Orientation.Horizontal
        footer_stack.HorizontalAlignment = HorizontalAlignment.Center
        footer_stack.VerticalAlignment = VerticalAlignment.Center
        
        btn_select_all = Button()
        btn_select_all.Content = "Select All"
        btn_select_all.Width = 100
        btn_select_all.Height = 35
        btn_select_all.Margin = Thickness(5, 0, 5, 0)
        btn_select_all.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        btn_select_all.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        btn_select_all.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        btn_select_all.Click += self._on_select_all
        footer_stack.Children.Add(btn_select_all)
        
        btn_clear_all = Button()
        btn_clear_all.Content = "Clear All"
        btn_clear_all.Width = 100
        btn_clear_all.Height = 35
        btn_clear_all.Margin = Thickness(5, 0, 5, 0)
        btn_clear_all.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        btn_clear_all.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        btn_clear_all.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        btn_clear_all.Click += self._on_clear_all
        footer_stack.Children.Add(btn_clear_all)
        
        btn_select_custom = Button()
        btn_select_custom.Content = "Select Custom"
        btn_select_custom.Width = 120
        btn_select_custom.Height = 35
        btn_select_custom.Margin = Thickness(5, 0, 20, 0)
        btn_select_custom.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        btn_select_custom.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        btn_select_custom.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        btn_select_custom.Click += self._on_select_custom
        footer_stack.Children.Add(btn_select_custom)
        
        btn_rename = Button()
        btn_rename.Content = "Rename"
        btn_rename.Width = 100
        btn_rename.Height = 35
        btn_rename.Margin = Thickness(5, 0, 5, 0)
        btn_rename.Background = SolidColorBrush(Config.hex_to_color(Config.SUCCESS_COLOR))
        btn_rename.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_rename.BorderThickness = Thickness(0)
        btn_rename.Click += self._on_rename
        footer_stack.Children.Add(btn_rename)
        
        btn_batch_rename = Button()
        btn_batch_rename.Content = "Batch Rename"
        btn_batch_rename.Width = 120
        btn_batch_rename.Height = 35
        btn_batch_rename.Margin = Thickness(5, 0, 5, 0)
        btn_batch_rename.Background = SolidColorBrush(Config.hex_to_color(Config.SUCCESS_COLOR))
        btn_batch_rename.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_batch_rename.BorderThickness = Thickness(0)
        btn_batch_rename.Click += self._on_batch_rename
        footer_stack.Children.Add(btn_batch_rename)
        
        btn_delete = Button()
        btn_delete.Content = "Delete"
        btn_delete.Width = 100
        btn_delete.Height = 35
        btn_delete.Margin = Thickness(5, 0, 5, 0)
        btn_delete.Background = SolidColorBrush(Config.hex_to_color(Config.ERROR_COLOR))
        btn_delete.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        btn_delete.BorderThickness = Thickness(0)
        btn_delete.Click += self._on_delete
        footer_stack.Children.Add(btn_delete)
        
        btn_refresh = Button()
        btn_refresh.Content = "Refresh"
        btn_refresh.Width = 100
        btn_refresh.Height = 35
        btn_refresh.Margin = Thickness(20, 0, 5, 0)
        btn_refresh.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        btn_refresh.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        btn_refresh.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        btn_refresh.Click += self._on_refresh
        footer_stack.Children.Add(btn_refresh)
        
        footer_border.Child = footer_stack
        main_grid.Children.Add(footer_border)
        
        self.Content = main_grid
    
    def _add_stat_card(self, parent, label, value, name):
        """Add a stat card to the parent panel"""
        card = Border()
        card.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        card.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        card.BorderThickness = Thickness(1)
        card.CornerRadius = System.Windows.CornerRadius(5)
        card.Padding = Thickness(15, 8, 15, 8)
        card.Margin = Thickness(0, 0, 15, 0)
        
        stack = StackPanel()
        stack.Orientation = Orientation.Horizontal
        
        label_text = TextBlock()
        label_text.Text = label + ": "
        label_text.FontWeight = FontWeights.Bold
        label_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        stack.Children.Add(label_text)
        
        value_text = TextBlock()
        value_text.Text = value
        value_text.FontWeight = FontWeights.Bold
        value_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        stack.Children.Add(value_text)
        
        setattr(self, name, value_text)
        
        card.Child = stack
        parent.Children.Add(card)
    
    # ========================================================================
    # DATA LOADING
    # ========================================================================
    
    def _load_data(self):
        """Load line patterns from document"""
        try:
            collector = FilteredElementCollector(doc).OfClass(LinePatternElement)
            
            self.all_items = []
            for elem in collector:
                try:
                    item = LinePatternItem(elem)
                    self.all_items.append(item)
                except Exception as ex:
                    print("Error creating item for {}: {}".format(elem.Id, str(ex)))
            
            self.all_items.sort(key=lambda x: x.name)
            
            self._apply_filters()
            self._update_stats()
            
        except Exception as ex:
            print("Error loading data: {}".format(str(ex)))
            MessageBox.Show("Error loading line patterns:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _apply_filters(self):
        """Apply search and category filters"""
        search_text = self.txt_search.Text.lower() if self.txt_search.Text else ""
        
        category_item = self.cmb_category.SelectedItem
        category_filter = category_item.Content if category_item else "All"
        
        self.filtered_items.Clear()
        
        for item in self.all_items:
            if search_text and search_text not in item.name.lower():
                continue
            
            if category_filter == "System" and item.category != "System":
                continue
            
            if category_filter == "Custom" and item.category != "Custom":
                continue
            
            self.filtered_items.Add(item)
        
        self.data_grid.ItemsSource = self.filtered_items
        self._update_stats()
    
    def _update_stats(self):
        """Update statistics display"""
        if self.txt_total:
            self.txt_total.Text = str(len(self.all_items))
        
        if self.txt_selected:
            selected = sum(1 for item in self.filtered_items if item.is_selected)
            self.txt_selected.Text = str(selected)
        
        if self.txt_system:
            system = sum(1 for item in self.all_items if item.category == "System")
            self.txt_system.Text = str(system)
        
        if self.txt_custom:
            custom = sum(1 for item in self.all_items if item.category == "Custom")
            self.txt_custom.Text = str(custom)
    
    def _get_selected_items(self):
        return [item for item in self.filtered_items if item.is_selected]
    
    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================
    
    def _on_filter_changed(self, sender, args):
        self._apply_filters()
    
    def _on_selection_changed(self, sender, args):
        if self.txt_selected:
            self.txt_selected.Text = str(self.data_grid.SelectedItems.Count)
        
        try:
            selected_items = list(self.data_grid.SelectedItems)
            for item in self.filtered_items:
                item.is_selected = item in selected_items
            self.data_grid.Items.Refresh()
        except:
            pass
    
    def _on_select_all(self, sender, args):
        self.data_grid.SelectAll()
    
    def _on_clear_all(self, sender, args):
        self.data_grid.UnselectAll()
        for item in self.filtered_items:
            item.is_selected = False
        self.data_grid.Items.Refresh()
        self._update_stats()
    
    def _on_select_custom(self, sender, args):
        """Select only custom patterns"""
        self.data_grid.UnselectAll()
        for item in self.filtered_items:
            item.is_selected = not item.is_system
        self.data_grid.Items.Refresh()
        self._update_stats()
        MessageBox.Show("Selected custom patterns only.\nSystem patterns cannot be renamed/deleted.",
                       "Info", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_refresh(self, sender, args):
        self._load_data()
        MessageBox.Show("Data refreshed!", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_rename(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select one line pattern to rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if len(selected) > 1:
            MessageBox.Show("Please select only one line pattern to rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        
        if item.is_system:
            MessageBox.Show("Cannot rename system line patterns!",
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        new_name = forms.ask_for_string(
            prompt="Enter new name for line pattern:",
            default=item.name,
            title="Rename Line Pattern"
        )
        
        if not new_name or new_name.strip() == "" or new_name == item.name:
            return
        
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            new_name = new_name.replace(char, '')
        new_name = new_name.strip()
        
        if not new_name:
            MessageBox.Show("Name cannot be empty!", "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        for other in self.all_items:
            if other.id != item.id and other.name == new_name:
                MessageBox.Show("A line pattern with name '{}' already exists!".format(new_name),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
                return
        
        t = Transaction(doc, "DQT - Rename Line Pattern")
        t.Start()
        
        try:
            item.element.Name = new_name
            t.Commit()
            
            MessageBox.Show("Line pattern renamed successfully!",
                          "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            self._load_data()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Failed to rename:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_batch_rename(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one line pattern to batch rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        renamable = [item for item in selected if not item.is_system]
        
        if not renamable:
            MessageBox.Show("Cannot rename system line patterns!\nPlease select custom patterns.",
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        if len(renamable) < len(selected):
            MessageBox.Show("{} system pattern(s) will be skipped.\n{} custom pattern(s) will be renamed.".format(
                len(selected) - len(renamable), len(renamable)),
                "Info", MessageBoxButton.OK, MessageBoxImage.Information)
        
        dialog = BatchRenameDialog(renamable, self)
        dialog.ShowDialog()
    
    def _on_delete(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one line pattern to delete!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        deletable = [item for item in selected if not item.is_system]
        
        if not deletable:
            MessageBox.Show("Cannot delete system line patterns!",
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        result = MessageBox.Show(
            "Delete {} line pattern(s)?\n\nNote: System patterns will be skipped.".format(len(deletable)),
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        
        if result != MessageBoxResult.Yes:
            return
        
        t = Transaction(doc, "DQT - Delete Line Patterns")
        t.Start()
        
        try:
            success_count = 0
            error_count = 0
            
            for item in deletable:
                try:
                    doc.Delete(item.element.Id)
                    success_count += 1
                except Exception as ex:
                    print("Error deleting {}: {}".format(item.name, str(ex)))
                    error_count += 1
            
            t.Commit()
            
            msg = "Deleted: {}".format(success_count)
            if error_count > 0:
                msg += "\nFailed: {} (may be in use)".format(error_count)
            
            MessageBox.Show(msg, "Result", MessageBoxButton.OK, MessageBoxImage.Information)
            self._load_data()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error deleting:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == '__main__':
    try:
        window = LinePatternManagerWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        MessageBox.Show(
            "Error starting Line Pattern Manager:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )