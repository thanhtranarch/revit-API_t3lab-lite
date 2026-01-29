# -*- coding: utf-8 -*-
"""
View Template Manager (Standalone)
Manage View Templates with beautiful UI - Sheet Manager Style

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""
__title__ = "View Template\nManage"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Manage View Templates - By DQT"

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
                                      ScrollViewer, TabControl, TabItem, CheckBox)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

import re

from pyrevit import revit, DB, forms

doc = revit.doc


# ============================================================================
# CONFIGURATION
# ============================================================================

class Config(object):
    """Color scheme and settings"""
    PRIMARY_COLOR = "#F0CC88"
    SECONDARY_COLOR = "#E5B85C"
    BACKGROUND_COLOR = "#FEF8E7"
    BORDER_COLOR = "#D4B87A"
    TEXT_DARK = "#5D4E37"
    TEXT_LIGHT = "#888888"
    SUCCESS_COLOR = "#4CAF50"
    WARNING_COLOR = "#FF9800"
    ERROR_COLOR = "#FF6B6B"
    ROW_ALT_COLOR = "#FFFDF5"
    WHITE = "#FFFFFF"
    BLUE_COLOR = "#5DADE2"
    
    @staticmethod
    def hex_to_color(hex_color):
        hex_color = hex_color.replace("#", "")
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return Color.FromArgb(255, r, g, b)


# ============================================================================
# VIEW TEMPLATE ITEM (Data Model)
# ============================================================================

class ViewTemplateItem(INotifyPropertyChanged):
    """Wrapper class for View Template with WPF binding support"""
    
    def __init__(self, view_template):
        self._property_changed_handlers = []
        self._view_template = view_template
        self._is_selected = False
        
        self._name = view_template.Name
        self._id = view_template.Id.IntegerValue
        
        # Get view type
        try:
            self._view_type = str(view_template.ViewType)
        except:
            self._view_type = "Unknown"
        
        # Get scale
        try:
            scale_param = view_template.get_Parameter(DB.BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
            if scale_param and scale_param.HasValue:
                self._scale = scale_param.AsValueString()
            else:
                self._scale = "N/A"
        except:
            self._scale = "N/A"
        
        self._usage_count = 0
        self._usage_percentage = 0.0
    
    # Properties
    @property
    def element(self):
        return self._view_template
    
    @property
    def view_template(self):
        return self._view_template
    
    @property
    def id(self):
        return self._id
    
    @property
    def name(self):
        return self._name
    
    @property
    def is_selected(self):
        return self._is_selected
    
    @is_selected.setter
    def is_selected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("is_selected")
    
    @property
    def view_type(self):
        return self._view_type
    
    @property
    def scale(self):
        return self._scale
    
    @property
    def usage_count(self):
        return self._usage_count
    
    @usage_count.setter
    def usage_count(self, value):
        self._usage_count = value
        self.OnPropertyChanged("usage_count")
    
    @property
    def usage_percentage(self):
        return "{:.1f}%".format(self._usage_percentage)
    
    @usage_percentage.setter
    def usage_percentage(self, value):
        self._usage_percentage = value
        self.OnPropertyChanged("usage_percentage")
    
    # INotifyPropertyChanged
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def OnPropertyChanged(self, prop_name):
        for handler in self._property_changed_handlers:
            try:
                handler(self, PropertyChangedEventArgs(prop_name))
            except:
                pass


# ============================================================================
# USAGE CALCULATION
# ============================================================================

def calculate_viewtemplate_usage(doc, template_items):
    """Calculate usage for view templates"""
    
    # Reset all counts
    for item in template_items:
        item.usage_count = 0
        item.usage_percentage = 0.0
    
    # Build lookup by template ID
    template_lookup = {}
    for item in template_items:
        template_lookup[item.id] = item
    
    total_views_with_template = 0
    
    try:
        # Get all views
        collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
        
        for view in collector:
            try:
                if view.IsTemplate:
                    continue
                
                template_id = view.ViewTemplateId
                if template_id and template_id.IntegerValue > 0:
                    if template_id.IntegerValue in template_lookup:
                        template_lookup[template_id.IntegerValue].usage_count += 1
                        total_views_with_template += 1
            except:
                pass
        
        # Calculate percentages
        if total_views_with_template > 0:
            for item in template_items:
                item.usage_percentage = (item.usage_count / float(total_views_with_template)) * 100.0
                
    except Exception as ex:
        print("Error calculating view template usage: {}".format(str(ex)))
    
    return total_views_with_template


# ============================================================================
# BATCH RENAME DIALOG
# ============================================================================

class BatchRenameDialog(Window):
    """Comprehensive Batch Rename Dialog with multiple tabs"""
    
    def __init__(self, items, parent_window):
        self.items = items
        self.parent_window = parent_window
        
        self.tab_control = None
        self.preview_text = None
        
        # Tab controls
        self.prefix_box = None
        self.suffix_box = None
        self.find_box = None
        self.replace_box = None
        self.case_check = None
        self.remove_numbers_check = None
        self.remove_special_check = None
        self.remove_spaces_check = None
        self.remove_custom_box = None
        self.case_combo = None
        self.numbering_check = None
        self.start_number_box = None
        self.padding_box = None
        self.number_position_combo = None
        self.number_separator_box = None
        
        # Template Info tab
        self.add_viewtype_check = None
        self.add_scale_check = None
        self.info_position_combo = None
        self.info_separator_box = None
        self.info_bracket_combo = None
        self.remove_current_name_check = None
        
        self._build_ui()
    
    def _build_ui(self):
        self.Title = "Batch Rename View Templates"
        self.Width = 600
        self.Height = 550
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Config.hex_to_color(Config.BACKGROUND_COLOR))
        
        main_grid = Grid()
        main_grid.Margin = Thickness(15)
        
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(150)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        
        header = self._create_header()
        Grid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        info = TextBlock()
        info.Text = "{} template(s) selected for renaming".format(len(self.items))
        info.FontSize = 12
        info.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_LIGHT))
        info.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(info, 1)
        main_grid.Children.Add(info)
        
        self.tab_control = self._create_tabs()
        Grid.SetRow(self.tab_control, 2)
        main_grid.Children.Add(self.tab_control)
        
        preview_section = self._create_preview_section()
        Grid.SetRow(preview_section, 3)
        main_grid.Children.Add(preview_section)
        
        buttons = self._create_buttons()
        Grid.SetRow(buttons, 4)
        main_grid.Children.Add(buttons)
        
        self.Content = main_grid
        self._update_preview()
    
    def _create_header(self):
        border = Border()
        border.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(10, 8, 10, 8)
        border.Margin = Thickness(0, 0, 0, 10)
        
        title = TextBlock()
        title.Text = "Batch Rename"
        title.FontSize = 16
        title.FontWeight = FontWeights.Bold
        
        border.Child = title
        return border
    
    def _create_tabs(self):
        tab_control = TabControl()
        tab_control.Margin = Thickness(0, 0, 0, 10)
        
        tab1 = TabItem()
        tab1.Header = "Prefix/Suffix"
        tab1.Content = self._create_prefix_suffix_tab()
        tab_control.Items.Add(tab1)
        
        tab2 = TabItem()
        tab2.Header = "Find/Replace"
        tab2.Content = self._create_find_replace_tab()
        tab_control.Items.Add(tab2)
        
        tab3 = TabItem()
        tab3.Header = "Remove"
        tab3.Content = self._create_remove_tab()
        tab_control.Items.Add(tab3)
        
        tab4 = TabItem()
        tab4.Header = "Change Case"
        tab4.Content = self._create_case_tab()
        tab_control.Items.Add(tab4)
        
        tab5 = TabItem()
        tab5.Header = "Numbering"
        tab5.Content = self._create_numbering_tab()
        tab_control.Items.Add(tab5)
        
        tab6 = TabItem()
        tab6.Header = "Template Info"
        tab6.Content = self._create_template_info_tab()
        tab_control.Items.Add(tab6)
        
        tab_control.SelectionChanged += lambda s, e: self._update_preview()
        
        return tab_control
    
    def _create_prefix_suffix_tab(self):
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        prefix_panel = StackPanel()
        prefix_panel.Orientation = Orientation.Horizontal
        prefix_panel.Margin = Thickness(0, 5, 0, 10)
        
        prefix_label = TextBlock()
        prefix_label.Text = "Prefix:"
        prefix_label.Width = 80
        prefix_label.VerticalAlignment = VerticalAlignment.Center
        
        self.prefix_box = TextBox()
        self.prefix_box.Width = 350
        self.prefix_box.Padding = Thickness(5)
        self.prefix_box.TextChanged += lambda s, e: self._update_preview()
        
        prefix_panel.Children.Add(prefix_label)
        prefix_panel.Children.Add(self.prefix_box)
        panel.Children.Add(prefix_panel)
        
        suffix_panel = StackPanel()
        suffix_panel.Orientation = Orientation.Horizontal
        
        suffix_label = TextBlock()
        suffix_label.Text = "Suffix:"
        suffix_label.Width = 80
        suffix_label.VerticalAlignment = VerticalAlignment.Center
        
        self.suffix_box = TextBox()
        self.suffix_box.Width = 350
        self.suffix_box.Padding = Thickness(5)
        self.suffix_box.TextChanged += lambda s, e: self._update_preview()
        
        suffix_panel.Children.Add(suffix_label)
        suffix_panel.Children.Add(self.suffix_box)
        panel.Children.Add(suffix_panel)
        
        return panel
    
    def _create_find_replace_tab(self):
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        find_panel = StackPanel()
        find_panel.Orientation = Orientation.Horizontal
        find_panel.Margin = Thickness(0, 5, 0, 10)
        
        find_label = TextBlock()
        find_label.Text = "Find:"
        find_label.Width = 80
        find_label.VerticalAlignment = VerticalAlignment.Center
        
        self.find_box = TextBox()
        self.find_box.Width = 350
        self.find_box.Padding = Thickness(5)
        self.find_box.TextChanged += lambda s, e: self._update_preview()
        
        find_panel.Children.Add(find_label)
        find_panel.Children.Add(self.find_box)
        panel.Children.Add(find_panel)
        
        replace_panel = StackPanel()
        replace_panel.Orientation = Orientation.Horizontal
        replace_panel.Margin = Thickness(0, 0, 0, 10)
        
        replace_label = TextBlock()
        replace_label.Text = "Replace:"
        replace_label.Width = 80
        replace_label.VerticalAlignment = VerticalAlignment.Center
        
        self.replace_box = TextBox()
        self.replace_box.Width = 350
        self.replace_box.Padding = Thickness(5)
        self.replace_box.TextChanged += lambda s, e: self._update_preview()
        
        replace_panel.Children.Add(replace_label)
        replace_panel.Children.Add(self.replace_box)
        panel.Children.Add(replace_panel)
        
        self.case_check = CheckBox()
        self.case_check.Content = "Case sensitive"
        self.case_check.Margin = Thickness(80, 0, 0, 0)
        self.case_check.Checked += lambda s, e: self._update_preview()
        self.case_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.case_check)
        
        return panel
    
    def _create_remove_tab(self):
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        self.remove_numbers_check = CheckBox()
        self.remove_numbers_check.Content = "Remove numbers (0-9)"
        self.remove_numbers_check.Margin = Thickness(0, 5, 0, 5)
        self.remove_numbers_check.Checked += lambda s, e: self._update_preview()
        self.remove_numbers_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.remove_numbers_check)
        
        self.remove_special_check = CheckBox()
        self.remove_special_check.Content = "Remove special characters (!@#$%^&*)"
        self.remove_special_check.Margin = Thickness(0, 0, 0, 5)
        self.remove_special_check.Checked += lambda s, e: self._update_preview()
        self.remove_special_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.remove_special_check)
        
        self.remove_spaces_check = CheckBox()
        self.remove_spaces_check.Content = "Remove spaces"
        self.remove_spaces_check.Margin = Thickness(0, 0, 0, 10)
        self.remove_spaces_check.Checked += lambda s, e: self._update_preview()
        self.remove_spaces_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.remove_spaces_check)
        
        custom_panel = StackPanel()
        custom_panel.Orientation = Orientation.Horizontal
        
        custom_label = TextBlock()
        custom_label.Text = "Remove custom:"
        custom_label.Width = 100
        custom_label.VerticalAlignment = VerticalAlignment.Center
        
        self.remove_custom_box = TextBox()
        self.remove_custom_box.Width = 250
        self.remove_custom_box.Padding = Thickness(5)
        self.remove_custom_box.TextChanged += lambda s, e: self._update_preview()
        
        custom_panel.Children.Add(custom_label)
        custom_panel.Children.Add(self.remove_custom_box)
        panel.Children.Add(custom_panel)
        
        return panel
    
    def _create_case_tab(self):
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        case_panel = StackPanel()
        case_panel.Orientation = Orientation.Horizontal
        case_panel.Margin = Thickness(0, 5, 0, 10)
        
        case_label = TextBlock()
        case_label.Text = "Change to:"
        case_label.Width = 80
        case_label.VerticalAlignment = VerticalAlignment.Center
        
        self.case_combo = ComboBox()
        self.case_combo.Width = 200
        self.case_combo.Padding = Thickness(5)
        
        for option in ["No Change", "UPPERCASE", "lowercase", "Title Case", "Sentence case"]:
            item = ComboBoxItem()
            item.Content = option
            self.case_combo.Items.Add(item)
        
        self.case_combo.SelectedIndex = 0
        self.case_combo.SelectionChanged += lambda s, e: self._update_preview()
        
        case_panel.Children.Add(case_label)
        case_panel.Children.Add(self.case_combo)
        panel.Children.Add(case_panel)
        
        return panel
    
    def _create_numbering_tab(self):
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        self.numbering_check = CheckBox()
        self.numbering_check.Content = "Add numbering to names"
        self.numbering_check.Margin = Thickness(0, 5, 0, 15)
        self.numbering_check.Checked += lambda s, e: self._update_preview()
        self.numbering_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.numbering_check)
        
        start_panel = StackPanel()
        start_panel.Orientation = Orientation.Horizontal
        start_panel.Margin = Thickness(0, 0, 0, 10)
        
        start_label = TextBlock()
        start_label.Text = "Start at:"
        start_label.Width = 100
        start_label.VerticalAlignment = VerticalAlignment.Center
        
        self.start_number_box = TextBox()
        self.start_number_box.Width = 80
        self.start_number_box.Text = "1"
        self.start_number_box.Padding = Thickness(5)
        self.start_number_box.TextChanged += lambda s, e: self._update_preview()
        
        start_panel.Children.Add(start_label)
        start_panel.Children.Add(self.start_number_box)
        panel.Children.Add(start_panel)
        
        padding_panel = StackPanel()
        padding_panel.Orientation = Orientation.Horizontal
        padding_panel.Margin = Thickness(0, 0, 0, 10)
        
        padding_label = TextBlock()
        padding_label.Text = "Padding (digits):"
        padding_label.Width = 100
        padding_label.VerticalAlignment = VerticalAlignment.Center
        
        self.padding_box = TextBox()
        self.padding_box.Width = 80
        self.padding_box.Text = "2"
        self.padding_box.Padding = Thickness(5)
        self.padding_box.TextChanged += lambda s, e: self._update_preview()
        
        padding_panel.Children.Add(padding_label)
        padding_panel.Children.Add(self.padding_box)
        panel.Children.Add(padding_panel)
        
        pos_panel = StackPanel()
        pos_panel.Orientation = Orientation.Horizontal
        pos_panel.Margin = Thickness(0, 0, 0, 10)
        
        pos_label = TextBlock()
        pos_label.Text = "Position:"
        pos_label.Width = 100
        pos_label.VerticalAlignment = VerticalAlignment.Center
        
        self.number_position_combo = ComboBox()
        self.number_position_combo.Width = 150
        self.number_position_combo.Padding = Thickness(5)
        
        for option in ["Prefix", "Suffix"]:
            item = ComboBoxItem()
            item.Content = option
            self.number_position_combo.Items.Add(item)
        
        self.number_position_combo.SelectedIndex = 0
        self.number_position_combo.SelectionChanged += lambda s, e: self._update_preview()
        
        pos_panel.Children.Add(pos_label)
        pos_panel.Children.Add(self.number_position_combo)
        panel.Children.Add(pos_panel)
        
        sep_panel = StackPanel()
        sep_panel.Orientation = Orientation.Horizontal
        
        sep_label = TextBlock()
        sep_label.Text = "Separator:"
        sep_label.Width = 100
        sep_label.VerticalAlignment = VerticalAlignment.Center
        
        self.number_separator_box = TextBox()
        self.number_separator_box.Width = 80
        self.number_separator_box.Text = "_"
        self.number_separator_box.Padding = Thickness(5)
        self.number_separator_box.TextChanged += lambda s, e: self._update_preview()
        
        sep_panel.Children.Add(sep_label)
        sep_panel.Children.Add(self.number_separator_box)
        panel.Children.Add(sep_panel)
        
        return panel
    
    def _create_template_info_tab(self):
        """Create Template Info tab - add view type/scale to name"""
        panel = StackPanel()
        panel.Margin = Thickness(10)
        
        title = TextBlock()
        title.Text = "Add template properties to name:"
        title.FontWeight = FontWeights.SemiBold
        title.Margin = Thickness(0, 0, 0, 10)
        panel.Children.Add(title)
        
        self.add_viewtype_check = CheckBox()
        self.add_viewtype_check.Content = "Add View Type (FloorPlan, Section...)"
        self.add_viewtype_check.Margin = Thickness(0, 5, 0, 5)
        self.add_viewtype_check.Checked += lambda s, e: self._update_preview()
        self.add_viewtype_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.add_viewtype_check)
        
        self.add_scale_check = CheckBox()
        self.add_scale_check.Content = "Add Scale (1:100, 1:50...)"
        self.add_scale_check.Margin = Thickness(0, 0, 0, 10)
        self.add_scale_check.Checked += lambda s, e: self._update_preview()
        self.add_scale_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.add_scale_check)
        
        self.remove_current_name_check = CheckBox()
        self.remove_current_name_check.Content = "Remove current name (use only template info)"
        self.remove_current_name_check.Margin = Thickness(0, 5, 0, 15)
        self.remove_current_name_check.Foreground = SolidColorBrush(Config.hex_to_color(Config.ERROR_COLOR))
        self.remove_current_name_check.Checked += lambda s, e: self._update_preview()
        self.remove_current_name_check.Unchecked += lambda s, e: self._update_preview()
        panel.Children.Add(self.remove_current_name_check)
        
        pos_panel = StackPanel()
        pos_panel.Orientation = Orientation.Horizontal
        pos_panel.Margin = Thickness(0, 0, 0, 10)
        
        pos_label = TextBlock()
        pos_label.Text = "Position:"
        pos_label.Width = 80
        pos_label.VerticalAlignment = VerticalAlignment.Center
        
        self.info_position_combo = ComboBox()
        self.info_position_combo.Width = 150
        self.info_position_combo.Padding = Thickness(5)
        
        for option in ["Prefix", "Suffix"]:
            item = ComboBoxItem()
            item.Content = option
            self.info_position_combo.Items.Add(item)
        
        self.info_position_combo.SelectedIndex = 1
        self.info_position_combo.SelectionChanged += lambda s, e: self._update_preview()
        
        pos_panel.Children.Add(pos_label)
        pos_panel.Children.Add(self.info_position_combo)
        panel.Children.Add(pos_panel)
        
        bracket_panel = StackPanel()
        bracket_panel.Orientation = Orientation.Horizontal
        bracket_panel.Margin = Thickness(0, 0, 0, 10)
        
        bracket_label = TextBlock()
        bracket_label.Text = "Brackets:"
        bracket_label.Width = 80
        bracket_label.VerticalAlignment = VerticalAlignment.Center
        
        self.info_bracket_combo = ComboBox()
        self.info_bracket_combo.Width = 150
        self.info_bracket_combo.Padding = Thickness(5)
        
        for option in ["None", "(Parentheses)", "_Underscore_"]:
            item = ComboBoxItem()
            item.Content = option
            self.info_bracket_combo.Items.Add(item)
        
        self.info_bracket_combo.SelectedIndex = 1
        self.info_bracket_combo.SelectionChanged += lambda s, e: self._update_preview()
        
        bracket_panel.Children.Add(bracket_label)
        bracket_panel.Children.Add(self.info_bracket_combo)
        panel.Children.Add(bracket_panel)
        
        sep_panel = StackPanel()
        sep_panel.Orientation = Orientation.Horizontal
        
        sep_label = TextBlock()
        sep_label.Text = "Separator:"
        sep_label.Width = 80
        sep_label.VerticalAlignment = VerticalAlignment.Center
        
        self.info_separator_box = TextBox()
        self.info_separator_box.Width = 80
        self.info_separator_box.Text = "_"
        self.info_separator_box.Padding = Thickness(5)
        self.info_separator_box.TextChanged += lambda s, e: self._update_preview()
        
        sep_panel.Children.Add(sep_label)
        sep_panel.Children.Add(self.info_separator_box)
        panel.Children.Add(sep_panel)
        
        return panel
    
    def _create_preview_section(self):
        border = Border()
        border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Margin = Thickness(0, 0, 0, 10)
        
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(25)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        
        label = TextBlock()
        label.Text = "PREVIEW"
        label.FontSize = 10
        label.FontWeight = FontWeights.SemiBold
        label.Padding = Thickness(10, 5, 10, 0)
        label.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        Grid.SetRow(label, 0)
        grid.Children.Add(label)
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        
        self.preview_text = TextBlock()
        self.preview_text.Text = "Preview will appear here..."
        self.preview_text.Padding = Thickness(10)
        self.preview_text.FontFamily = System.Windows.Media.FontFamily("Consolas")
        self.preview_text.FontSize = 11
        
        scroll.Content = self.preview_text
        Grid.SetRow(scroll, 1)
        grid.Children.Add(scroll)
        
        border.Child = grid
        return border
    
    def _create_buttons(self):
        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        panel.HorizontalAlignment = HorizontalAlignment.Right
        
        apply_btn = Button()
        apply_btn.Content = "Apply Rename"
        apply_btn.Width = 120
        apply_btn.Height = 35
        apply_btn.Margin = Thickness(0, 0, 10, 0)
        apply_btn.Background = SolidColorBrush(Config.hex_to_color(Config.SUCCESS_COLOR))
        apply_btn.Foreground = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        apply_btn.Click += self._on_apply
        panel.Children.Add(apply_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 100
        cancel_btn.Height = 35
        cancel_btn.Click += lambda s, e: self.Close()
        panel.Children.Add(cancel_btn)
        
        return panel
    
    def _sanitize_name(self, name):
        """Sanitize name for Revit view"""
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '{', '}', '[', ']', ';']
        for char in invalid_chars:
            name = name.replace(char, '')
        
        while '  ' in name:
            name = name.replace('  ', ' ')
        while '__' in name:
            name = name.replace('__', '_')
        
        return name.strip()
    
    def _apply_rules_to_name(self, name, index, item=None):
        """Apply all rename rules to a single name"""
        new_name = name
        selected_tab = self.tab_control.SelectedIndex
        
        # Tab 0: Prefix/Suffix
        if selected_tab == 0:
            prefix = self.prefix_box.Text if self.prefix_box.Text else ""
            suffix = self.suffix_box.Text if self.suffix_box.Text else ""
            new_name = prefix + new_name + suffix
        
        # Tab 1: Find/Replace
        elif selected_tab == 1:
            find_text = self.find_box.Text if self.find_box.Text else ""
            replace_text = self.replace_box.Text if self.replace_box.Text else ""
            
            if find_text:
                if self.case_check.IsChecked:
                    new_name = new_name.replace(find_text, replace_text)
                else:
                    pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                    new_name = pattern.sub(replace_text, new_name)
        
        # Tab 2: Remove Characters
        elif selected_tab == 2:
            if self.remove_numbers_check.IsChecked:
                new_name = re.sub(r'[0-9]', '', new_name)
            
            if self.remove_special_check.IsChecked:
                new_name = re.sub(r'[!@#$%^&*()+=\[\]{};:\'",.<>?/\\|`~]', '', new_name)
            
            if self.remove_spaces_check.IsChecked:
                new_name = new_name.replace(' ', '')
            
            custom = self.remove_custom_box.Text if self.remove_custom_box.Text else ""
            if custom:
                for char in custom:
                    new_name = new_name.replace(char, '')
        
        # Tab 3: Case Change
        elif selected_tab == 3:
            case_option = self.case_combo.SelectedIndex
            if case_option == 1:
                new_name = new_name.upper()
            elif case_option == 2:
                new_name = new_name.lower()
            elif case_option == 3:
                new_name = new_name.title()
            elif case_option == 4:
                new_name = new_name.capitalize()
        
        # Tab 4: Numbering
        elif selected_tab == 4:
            if self.numbering_check.IsChecked:
                try:
                    start = int(self.start_number_box.Text) if self.start_number_box.Text else 1
                    padding = int(self.padding_box.Text) if self.padding_box.Text else 2
                except:
                    start = 1
                    padding = 2
                
                number = str(start + index).zfill(padding)
                separator = self.number_separator_box.Text if self.number_separator_box.Text else "_"
                position = self.number_position_combo.SelectedIndex
                
                if position == 0:
                    new_name = number + separator + new_name
                else:
                    new_name = new_name + separator + number
        
        # Tab 5: Template Info
        elif selected_tab == 5:
            if item:
                info_parts = []
                
                if self.add_viewtype_check.IsChecked:
                    info_parts.append(item.view_type)
                
                if self.add_scale_check.IsChecked:
                    scale = item.scale.replace(":", "-") if item.scale else "N-A"
                    info_parts.append(scale)
                
                if info_parts:
                    bracket_index = self.info_bracket_combo.SelectedIndex if self.info_bracket_combo else 0
                    brackets = [("", ""), ("(", ")"), ("_", "_")]
                    left_b, right_b = brackets[bracket_index] if bracket_index < len(brackets) else ("", "")
                    
                    info_str = "".join(["{}{}{}".format(left_b, part, right_b) for part in info_parts])
                    
                    separator = self.info_separator_box.Text if self.info_separator_box.Text else "_"
                    
                    if self.remove_current_name_check.IsChecked:
                        new_name = info_str
                    else:
                        position = self.info_position_combo.SelectedIndex if self.info_position_combo else 1
                        
                        if position == 0:
                            new_name = info_str + separator + new_name
                        else:
                            new_name = new_name + separator + info_str
        
        return self._sanitize_name(new_name)
    
    def _update_preview(self, sender=None, args=None):
        if not self.preview_text:
            return
        
        preview_lines = []
        changed_count = 0
        
        for i, item in enumerate(self.items[:15]):
            old_name = item.name
            new_name = self._apply_rules_to_name(old_name, i, item)
            
            if old_name != new_name:
                preview_lines.append("{} -> {}".format(old_name, new_name))
                changed_count += 1
            else:
                preview_lines.append("{} (no change)".format(old_name))
        
        if len(self.items) > 15:
            preview_lines.append("... and {} more items".format(len(self.items) - 15))
        
        if changed_count > 0:
            self.preview_text.Text = "\n".join(preview_lines)
            self.preview_text.Foreground = SolidColorBrush(Config.hex_to_color("#333333"))
        else:
            self.preview_text.Text = "No changes will be made with current settings"
            self.preview_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_LIGHT))
    
    def _on_apply(self, sender, args):
        changes = []
        for i, item in enumerate(self.items):
            old_name = item.name
            new_name = self._apply_rules_to_name(old_name, i, item)
            if old_name != new_name and new_name:
                changes.append((item, new_name))
        
        if not changes:
            MessageBox.Show("No changes to apply!", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        t = DB.Transaction(doc, "Batch Rename View Templates")
        t.Start()
        
        try:
            success_count = 0
            for item, new_name in changes:
                try:
                    item.view_template.Name = new_name
                    success_count += 1
                except:
                    pass
            
            t.Commit()
            
            MessageBox.Show("Successfully renamed: {}".format(success_count), "Result",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            
            self.Close()
            self.parent_window._load_data()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# MAIN WINDOW
# ============================================================================

class ViewTemplateManagerWindow(Window):
    """View Template Manager with Sheet Manager style UI"""
    
    def __init__(self):
        self.all_items = []
        self.filtered_items = []
        
        self.txt_total = None
        self.txt_selected = None
        self.txt_used = None
        self.txt_unused = None
        self.search_box = None
        self.filter_combo = None
        self.viewtype_combo = None
        self.data_grid = None
        
        self._build_ui()
        self._load_data()
    
    def _build_ui(self):
        self.Title = "View Template Manager"
        self.Width = 1150
        self.Height = 700
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Config.hex_to_color(Config.BACKGROUND_COLOR))
        
        main_grid = Grid()
        main_grid.Margin = Thickness(12)
        
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        
        header = self._create_header()
        Grid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        cards = self._create_summary_cards()
        Grid.SetRow(cards, 1)
        main_grid.Children.Add(cards)
        
        content = self._create_main_content()
        Grid.SetRow(content, 2)
        main_grid.Children.Add(content)
        
        buttons = self._create_action_buttons()
        Grid.SetRow(buttons, 3)
        main_grid.Children.Add(buttons)
        
        tips = self._create_tips()
        Grid.SetRow(tips, 4)
        main_grid.Children.Add(tips)
        
        footer = self._create_footer()
        Grid.SetRow(footer, 5)
        main_grid.Children.Add(footer)
        
        self.Content = main_grid
    
    def _create_header(self):
        border = Border()
        border.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        border.CornerRadius = System.Windows.CornerRadius(5)
        border.Padding = Thickness(12, 8, 12, 8)
        border.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        
        title = TextBlock()
        title.Text = "View Template Manager v1.0"
        title.FontSize = 17
        title.FontWeight = FontWeights.Bold
        panel.Children.Add(title)
        
        subtitle = TextBlock()
        subtitle.Text = "by Dang Quoc Truong (DQT)"
        subtitle.FontSize = 10
        subtitle.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        subtitle.Margin = Thickness(0, 2, 0, 0)
        panel.Children.Add(subtitle)
        
        border.Child = panel
        return border
    
    def _create_summary_cards(self):
        grid = Grid()
        grid.Margin = Thickness(0, 0, 0, 10)
        
        for i in range(4):
            grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        card1, self.txt_total = self._create_card("TOTAL", "0", 0)
        Grid.SetColumn(card1, 0)
        grid.Children.Add(card1)
        
        card2, self.txt_selected = self._create_card("SELECTED", "0", 1, Config.SECONDARY_COLOR)
        Grid.SetColumn(card2, 1)
        grid.Children.Add(card2)
        
        card3, self.txt_used = self._create_card("IN USE", "0", 2, Config.SUCCESS_COLOR)
        Grid.SetColumn(card3, 2)
        grid.Children.Add(card3)
        
        card4, self.txt_unused = self._create_card("UNUSED", "0", 3, Config.WARNING_COLOR)
        Grid.SetColumn(card4, 3)
        grid.Children.Add(card4)
        
        return grid
    
    def _create_card(self, label, value, margin_index, value_color=None):
        border = Border()
        border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(10, 6, 10, 6)
        
        if margin_index == 0:
            border.Margin = Thickness(0, 0, 4, 0)
        elif margin_index == 3:
            border.Margin = Thickness(4, 0, 0, 0)
        else:
            border.Margin = Thickness(4, 0, 4, 0)
        
        panel = StackPanel()
        
        label_text = TextBlock()
        label_text.Text = label
        label_text.FontSize = 9
        label_text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_LIGHT))
        panel.Children.Add(label_text)
        
        value_text = TextBlock()
        value_text.Text = value
        value_text.FontSize = 22
        value_text.FontWeight = FontWeights.Bold
        
        if value_color:
            value_text.Foreground = SolidColorBrush(Config.hex_to_color(value_color))
        
        panel.Children.Add(value_text)
        border.Child = panel
        
        return border, value_text
    
    def _create_main_content(self):
        grid = Grid()
        
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(200)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        left_panel = self._create_left_panel()
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        self.data_grid = self._create_datagrid()
        Grid.SetColumn(self.data_grid, 1)
        grid.Children.Add(self.data_grid)
        
        return grid
    
    def _create_left_panel(self):
        border = Border()
        border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 0, 8, 0)
        
        panel = StackPanel()
        
        label1 = TextBlock()
        label1.Text = "SEARCH"
        label1.FontSize = 9
        label1.FontWeight = FontWeights.SemiBold
        label1.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label1)
        
        self.search_box = TextBox()
        self.search_box.Padding = Thickness(6, 4, 6, 4)
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self._on_filter_changed
        panel.Children.Add(self.search_box)
        
        label2 = TextBlock()
        label2.Text = "FILTER BY USAGE"
        label2.FontSize = 9
        label2.FontWeight = FontWeights.SemiBold
        label2.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label2)
        
        self.filter_combo = ComboBox()
        self.filter_combo.Padding = Thickness(6, 4, 6, 4)
        self.filter_combo.Margin = Thickness(0, 0, 0, 10)
        
        for item_text in ["All Templates", "In Use Only", "Unused Only"]:
            item = ComboBoxItem()
            item.Content = item_text
            self.filter_combo.Items.Add(item)
        
        self.filter_combo.SelectedIndex = 0
        self.filter_combo.SelectionChanged += self._on_filter_changed
        panel.Children.Add(self.filter_combo)
        
        label3 = TextBlock()
        label3.Text = "FILTER BY VIEW TYPE"
        label3.FontSize = 9
        label3.FontWeight = FontWeights.SemiBold
        label3.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label3)
        
        self.viewtype_combo = ComboBox()
        self.viewtype_combo.Padding = Thickness(6, 4, 6, 4)
        self.viewtype_combo.Margin = Thickness(0, 0, 0, 10)
        
        item = ComboBoxItem()
        item.Content = "All View Types"
        self.viewtype_combo.Items.Add(item)
        self.viewtype_combo.SelectedIndex = 0
        self.viewtype_combo.SelectionChanged += self._on_filter_changed
        panel.Children.Add(self.viewtype_combo)
        
        label4 = TextBlock()
        label4.Text = "QUICK SELECT"
        label4.FontSize = 9
        label4.FontWeight = FontWeights.SemiBold
        label4.Margin = Thickness(0, 10, 0, 4)
        panel.Children.Add(label4)
        
        btn_all = Button()
        btn_all.Content = "Select All"
        btn_all.Padding = Thickness(8, 4, 8, 4)
        btn_all.Margin = Thickness(0, 0, 0, 4)
        btn_all.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        btn_all.Click += self._on_select_all
        panel.Children.Add(btn_all)
        
        btn_clear = Button()
        btn_clear.Content = "Clear All"
        btn_clear.Padding = Thickness(8, 4, 8, 4)
        btn_clear.Margin = Thickness(0, 0, 0, 4)
        btn_clear.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        btn_clear.Click += self._on_clear_all
        panel.Children.Add(btn_clear)
        
        btn_unused = Button()
        btn_unused.Content = "Select Unused"
        btn_unused.Padding = Thickness(8, 4, 8, 4)
        btn_unused.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        btn_unused.Click += self._on_select_unused
        panel.Children.Add(btn_unused)
        
        border.Child = panel
        return border
    
    def _create_datagrid(self):
        grid = DataGrid()
        grid.AutoGenerateColumns = False
        grid.IsReadOnly = True
        grid.SelectionMode = System.Windows.Controls.DataGridSelectionMode.Extended
        grid.SelectionUnit = System.Windows.Controls.DataGridSelectionUnit.FullRow
        grid.CanUserSortColumns = True
        grid.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        grid.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        grid.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.Horizontal
        grid.HorizontalGridLinesBrush = SolidColorBrush(Color.FromArgb(255, 238, 238, 238))
        grid.RowBackground = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        grid.AlternatingRowBackground = SolidColorBrush(Config.hex_to_color(Config.ROW_ALT_COLOR))
        
        col_check = DataGridCheckBoxColumn()
        col_check.Header = u"[  ]"
        col_check.Binding = Binding("is_selected")
        col_check.Width = DataGridLength(50)
        col_check.IsReadOnly = False
        grid.Columns.Add(col_check)
        
        col_name = DataGridTextColumn()
        col_name.Header = "Template Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(300)
        grid.Columns.Add(col_name)
        
        col_viewtype = DataGridTextColumn()
        col_viewtype.Header = "View Type"
        col_viewtype.Binding = Binding("view_type")
        col_viewtype.Width = DataGridLength(120)
        grid.Columns.Add(col_viewtype)
        
        col_scale = DataGridTextColumn()
        col_scale.Header = "Scale"
        col_scale.Binding = Binding("scale")
        col_scale.Width = DataGridLength(100)
        grid.Columns.Add(col_scale)
        
        col_usage = DataGridTextColumn()
        col_usage.Header = "Usage"
        col_usage.Binding = Binding("usage_count")
        col_usage.Width = DataGridLength(70)
        grid.Columns.Add(col_usage)
        
        col_percent = DataGridTextColumn()
        col_percent.Header = "Usage %"
        col_percent.Binding = Binding("usage_percentage")
        col_percent.Width = DataGridLength(80)
        grid.Columns.Add(col_percent)
        
        col_id = DataGridTextColumn()
        col_id.Header = "ID"
        col_id.Binding = Binding("id")
        col_id.Width = DataGridLength(70)
        grid.Columns.Add(col_id)
        
        grid.SelectionChanged += self._on_selection_changed
        
        return grid
    
    def _create_action_buttons(self):
        border = Border()
        border.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        border.BorderBrush = SolidColorBrush(Config.hex_to_color(Config.BORDER_COLOR))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 10, 0, 0)
        
        grid = Grid()
        
        center_panel = StackPanel()
        center_panel.Orientation = Orientation.Horizontal
        center_panel.HorizontalAlignment = HorizontalAlignment.Center
        
        primary = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        blue = SolidColorBrush(Config.hex_to_color(Config.BLUE_COLOR))
        orange = SolidColorBrush(Config.hex_to_color(Config.WARNING_COLOR))
        red = SolidColorBrush(Config.hex_to_color(Config.ERROR_COLOR))
        white = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        
        btn_rename = self._create_button("Rename", primary)
        btn_rename.Click += self._on_rename
        center_panel.Children.Add(btn_rename)
        
        btn_batch = self._create_button("Batch Rename", primary)
        btn_batch.Click += self._on_batch_rename
        center_panel.Children.Add(btn_batch)
        
        btn_duplicate = self._create_button("Duplicate", orange, white_fg=True)
        btn_duplicate.Click += self._on_duplicate
        center_panel.Children.Add(btn_duplicate)
        
        btn_delete = self._create_button("Delete", red, white_fg=True)
        btn_delete.Click += self._on_delete
        center_panel.Children.Add(btn_delete)
        
        grid.Children.Add(center_panel)
        
        btn_refresh = Button()
        btn_refresh.Content = "Refresh"
        btn_refresh.Padding = Thickness(10, 5, 10, 5)
        btn_refresh.Margin = Thickness(2)
        btn_refresh.Background = white
        btn_refresh.HorizontalAlignment = HorizontalAlignment.Right
        btn_refresh.Click += self._on_refresh
        grid.Children.Add(btn_refresh)
        
        border.Child = grid
        return border
    
    def _create_button(self, text, bg_brush, white_fg=False):
        btn = Button()
        btn.Content = text
        btn.Padding = Thickness(10, 5, 10, 5)
        btn.Margin = Thickness(2)
        btn.Background = bg_brush
        
        if white_fg:
            btn.Foreground = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        
        return btn
    
    def _create_tips(self):
        grid = Grid()
        grid.Margin = Thickness(0, 8, 0, 0)
        
        tips = TextBlock()
        tips.Text = "Only UNUSED templates can be deleted | Use Batch Rename for multiple templates"
        tips.FontSize = 10
        tips.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_LIGHT))
        grid.Children.Add(tips)
        
        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Padding = Thickness(15, 5, 15, 5)
        btn_close.Background = SolidColorBrush(Config.hex_to_color(Config.WHITE))
        btn_close.HorizontalAlignment = HorizontalAlignment.Right
        btn_close.Click += lambda s, e: self.Close()
        grid.Children.Add(btn_close)
        
        return grid
    
    def _create_footer(self):
        border = Border()
        border.Background = SolidColorBrush(Config.hex_to_color(Config.PRIMARY_COLOR))
        border.CornerRadius = System.Windows.CornerRadius(3)
        border.Padding = Thickness(8, 5, 8, 5)
        border.Margin = Thickness(0, 8, 0, 0)
        
        text = TextBlock()
        text.Text = u"(c) 2025 Dang Quoc Truong (DQT) - All Rights Reserved"
        text.FontSize = 10
        text.FontWeight = FontWeights.SemiBold
        text.HorizontalAlignment = HorizontalAlignment.Center
        text.Foreground = SolidColorBrush(Config.hex_to_color(Config.TEXT_DARK))
        
        border.Child = text
        return border
    
    # ========================================================================
    # DATA METHODS
    # ========================================================================
    
    def _load_data(self):
        self.all_items = []
        
        try:
            collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
            
            view_types = set()
            
            for view in collector:
                try:
                    if view.IsTemplate:
                        item = ViewTemplateItem(view)
                        self.all_items.append(item)
                        view_types.add(item.view_type)
                except:
                    pass
            
            self.all_items.sort(key=lambda x: x.name)
            
            # Calculate usage
            calculate_viewtemplate_usage(doc, self.all_items)
            
            # Populate view type filter
            self.viewtype_combo.Items.Clear()
            item = ComboBoxItem()
            item.Content = "All View Types"
            self.viewtype_combo.Items.Add(item)
            
            for vt in sorted(view_types):
                item = ComboBoxItem()
                item.Content = vt
                self.viewtype_combo.Items.Add(item)
            
            self.viewtype_combo.SelectedIndex = 0
            
            self._apply_filters()
            self._update_stats()
            
        except Exception as ex:
            print("Error loading view templates: {}".format(str(ex)))
            MessageBox.Show("Failed to load view templates:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _apply_filters(self):
        search_text = ""
        if self.search_box and self.search_box.Text:
            search_text = self.search_box.Text.lower()
        
        filter_index = 0
        if self.filter_combo:
            filter_index = self.filter_combo.SelectedIndex
        
        viewtype_filter = ""
        if self.viewtype_combo and self.viewtype_combo.SelectedIndex > 0:
            selected = self.viewtype_combo.SelectedItem
            if selected:
                viewtype_filter = selected.Content
        
        self.filtered_items = []
        
        for item in self.all_items:
            if search_text and search_text not in item.name.lower():
                continue
            
            if filter_index == 1:  # In Use Only
                if item.usage_count == 0:
                    continue
            elif filter_index == 2:  # Unused Only
                if item.usage_count > 0:
                    continue
            
            if viewtype_filter and item.view_type != viewtype_filter:
                continue
            
            self.filtered_items.append(item)
        
        self.data_grid.ItemsSource = ObservableCollection[object](self.filtered_items)
    
    def _update_stats(self):
        if self.txt_total:
            self.txt_total.Text = str(len(self.all_items))
        
        if self.txt_selected:
            selected = sum(1 for item in self.all_items if item.is_selected)
            self.txt_selected.Text = str(selected)
        
        if self.txt_used:
            used = sum(1 for item in self.all_items if item.usage_count > 0)
            self.txt_used.Text = str(used)
        
        if self.txt_unused:
            unused = sum(1 for item in self.all_items if item.usage_count == 0)
            self.txt_unused.Text = str(unused)
    
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
    
    def _on_select_unused(self, sender, args):
        self.data_grid.UnselectAll()
        for item in self.filtered_items:
            item.is_selected = item.usage_count == 0
        self.data_grid.Items.Refresh()
        self._update_stats()
    
    def _on_refresh(self, sender, args):
        self._load_data()
        MessageBox.Show("Data refreshed!", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_rename(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select one template to rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if len(selected) > 1:
            MessageBox.Show("Please select only ONE template to rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        
        new_name = forms.ask_for_string(
            default=item.name,
            prompt="Enter new name:",
            title="Rename View Template"
        )
        
        if not new_name or new_name == item.name:
            return
        
        # Sanitize
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            new_name = new_name.replace(char, '')
        new_name = new_name.strip()
        
        if not new_name:
            return
        
        t = DB.Transaction(doc, "Rename View Template")
        t.Start()
        
        try:
            item.view_template.Name = new_name
            t.Commit()
            
            MessageBox.Show("Renamed successfully!", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            
            self._load_data()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_batch_rename(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one template to batch rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        dialog = BatchRenameDialog(selected, self)
        dialog.ShowDialog()
    
    def _on_duplicate(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one template to duplicate!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        t = DB.Transaction(doc, "Duplicate View Templates")
        t.Start()
        
        try:
            success_count = 0
            for item in selected:
                try:
                    new_id = item.view_template.Duplicate(DB.ViewDuplicateOption.Duplicate)
                    if new_id and new_id != DB.ElementId.InvalidElementId:
                        new_view = doc.GetElement(new_id)
                        new_view.Name = "Copy of " + item.name
                        success_count += 1
                except:
                    pass
            
            t.Commit()
            
            MessageBox.Show("Duplicated: {}".format(success_count), "Result",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            
            self._load_data()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_delete(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one template to delete!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Check usage - only unused can be deleted
        in_use = [item for item in selected if item.usage_count > 0]
        can_delete = [item for item in selected if item.usage_count == 0]
        
        if in_use:
            msg = "WARNING: {} template(s) are IN USE and CANNOT be deleted:\n\n".format(len(in_use))
            for item in in_use[:5]:
                msg += "  - '{}': {} views\n".format(item.name, item.usage_count)
            if len(in_use) > 5:
                msg += "  ... and {} more\n".format(len(in_use) - 5)
            
            MessageBox.Show(msg, "Templates In Use",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
        
        if not can_delete:
            MessageBox.Show("No deletable templates selected.\nOnly UNUSED templates can be deleted.",
                          "Info", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        result = MessageBox.Show(
            "Delete {} unused template(s)?".format(len(can_delete)),
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        
        if result != MessageBoxResult.Yes:
            return
        
        t = DB.Transaction(doc, "Delete View Templates")
        t.Start()
        
        try:
            success_count = 0
            error_count = 0
            
            for item in can_delete:
                try:
                    doc.Delete(item.view_template.Id)
                    success_count += 1
                except:
                    error_count += 1
            
            t.Commit()
            
            msg = "Deleted: {}".format(success_count)
            if error_count > 0:
                msg += "\nFailed: {}".format(error_count)
            
            MessageBox.Show(msg, "Result", MessageBoxButton.OK, MessageBoxImage.Information)
            self._load_data()
            
        except Exception as ex:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == '__main__':
    try:
        window = ViewTemplateManagerWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        MessageBox.Show(
            "Error starting View Template Manager:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )