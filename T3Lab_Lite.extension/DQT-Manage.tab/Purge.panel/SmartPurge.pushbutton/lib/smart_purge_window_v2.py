# -*- coding: utf-8 -*-
"""
Smart Purge Window v2.0 - Simplified Clean Version
UI with groups and batch operations

Copyright © 2025 Dang Quoc Truong (DQT)

FIXED in this version:
- on_scan_click now correctly filters selected groups
- Added debug output for troubleshooting
"""

__author__ = "Dang Quoc Truong (DQT)"
__version__ = "2.0.1"

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import Window, Application, WindowStartupLocation, Thickness
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows import FontWeights, Visibility, VerticalAlignment, TextAlignment
from System.Windows.Controls import (
    Grid, StackPanel, Border, TextBlock, Button, CheckBox, TextBox,
    ScrollViewer, ProgressBar, GroupBox, WrapPanel, ListBox, ListBoxItem,
    Orientation, ScrollBarVisibility, PanningMode
)
from System.Windows.Media import SolidColorBrush, Color
from System import Action

from config import Colors
# from purge_history_manager import PurgeHistoryManager  # DISABLED - causing issues
from purge_group import create_purge_groups, get_group_by_id
from purge_categories_v2 import create_purge_categories
from purge_scanner import create_scanner
from purge_executor import PurgeExecutor
from preview_window import PreviewWindow


class SmartPurgeWindowV2(Window):
    """Smart Purge Window v2.0 - Simplified"""
    
    def __init__(self, doc):
        """Initialize window"""
        self.doc = doc
        
        # Data
        self.groups = []
        self.categories = []
        
        # State
        self.is_scanning = False
        
        # UI elements
        self.group_checkboxes = {}
        self.category_checkboxes = {}
        self.category_labels = {}
        
        # Settings
        self.dry_run = True
        self.delete_system = False
        
        # History manager - DISABLED (causing issues in some environments)
        self.history_manager = None
        
        # Setup window
        self.Title = "Smart Purge v2.0"
        self.Width = 850
        self.Height = 950  # Increased from 850 to 950
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        
        # Build UI
        self.build_ui()
        
        # Load data
        self.load_data()
    
    def parse_color(self, hex_color):
        """Parse hex color to Color object"""
        hex_color = hex_color.replace('#', '')
        a = int(hex_color[0:2], 16) if len(hex_color) == 8 else 255
        r = int(hex_color[-6:-4], 16)
        g = int(hex_color[-4:-2], 16)
        b = int(hex_color[-2:], 16)
        return Color.FromArgb(a, r, g, b)
    
    def build_ui(self):
        """Build user interface"""
        main = Grid()
        
        # Rows
        main.RowDefinitions.Add(self.row_def(75))      # Header - increased height
        main.RowDefinitions.Add(self.row_def(120))     # Group selection
        main.RowDefinitions.Add(self.row_def(0, True)) # Progress (auto)
        main.RowDefinitions.Add(self.row_def(1, star=True))  # Categories
        main.RowDefinitions.Add(self.row_def(75))      # Settings
        main.RowDefinitions.Add(self.row_def(55))      # Summary
        main.RowDefinitions.Add(self.row_def(85))      # Actions - increased for copyright
        
        # Build panels
        row = 0
        main.Children.Add(self.header_panel(row)); row += 1
        main.Children.Add(self.group_panel(row)); row += 1
        main.Children.Add(self.progress_panel(row)); row += 1
        main.Children.Add(self.category_panel(row)); row += 1
        main.Children.Add(self.settings_panel(row)); row += 1
        main.Children.Add(self.summary_panel(row)); row += 1
        main.Children.Add(self.action_panel(row)); row += 1
        
        self.Content = main
    
    def row_def(self, height, auto=False, star=False):
        """Create row definition"""
        from System.Windows.Controls import RowDefinition
        from System.Windows import GridLength, GridUnitType
        rd = RowDefinition()
        if auto:
            rd.Height = GridLength.Auto
        elif star:
            rd.Height = GridLength(height, GridUnitType.Star)
        else:
            rd.Height = GridLength(height)
        return rd
    
    def header_panel(self, row):
        """Create header"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.HEADER))
        border.Padding = Thickness(15, 12, 15, 12)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        title = TextBlock()
        title.Text = "Smart Purge v2.0"
        title.FontSize = 22  # Larger title
        title.FontWeight = FontWeights.Bold
        stack.Children.Add(title)
        
        subtitle = TextBlock()
        subtitle.Text = u"Copyright \u00A9 2025 Dang Quoc Truong (DQT)"
        subtitle.FontSize = 12  # Larger copyright
        subtitle.FontWeight = FontWeights.Bold
        subtitle.Foreground = SolidColorBrush(self.parse_color("#FF555555"))
        subtitle.Margin = Thickness(0, 4, 0, 0)
        stack.Children.Add(subtitle)
        
        border.Child = stack
        return border
    
    def group_panel(self, row):
        """Create group selection panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color("#FFF5F5F5"))
        border.Padding = Thickness(15, 10, 15, 10)
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 0, 0, 1)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        # Title
        title = TextBlock()
        title.Text = u"\U0001F4E6 SELECT GROUPS TO SCAN:"
        title.FontSize = 14  # Larger
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 8)
        stack.Children.Add(title)
        
        # Checkboxes
        self.group_wrap = WrapPanel()
        stack.Children.Add(self.group_wrap)
        
        # Buttons
        btn_stack = StackPanel()
        btn_stack.Orientation = Orientation.Horizontal
        btn_stack.Margin = Thickness(0, 10, 0, 0)
        
        self.btn_scan_sel = Button()
        self.btn_scan_sel.Content = u"\U0001F50D Scan Selected"
        self.btn_scan_sel.Width = 140
        self.btn_scan_sel.Height = 34
        self.btn_scan_sel.FontSize = 13  # Larger
        self.btn_scan_sel.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_PRIMARY))
        self.btn_scan_sel.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        self.btn_scan_sel.Click += self.on_scan_click
        
        self.btn_scan_all = Button()
        self.btn_scan_all.Content = u"\U0001F50D Scan All"
        self.btn_scan_all.Width = 110
        self.btn_scan_all.Height = 34
        self.btn_scan_all.FontSize = 13  # Larger
        self.btn_scan_all.Margin = Thickness(5, 0, 0, 0)
        self.btn_scan_all.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_PRIMARY))
        self.btn_scan_all.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        self.btn_scan_all.Click += self.on_scan_all_click
        
        btn_stack.Children.Add(self.btn_scan_sel)
        btn_stack.Children.Add(self.btn_scan_all)
        stack.Children.Add(btn_stack)
        
        border.Child = stack
        return border
    
    def progress_panel(self, row):
        """Create progress panel"""
        self.progress_border = Border()
        self.progress_border.Background = SolidColorBrush(self.parse_color("#FFF5F5F5"))
        self.progress_border.Padding = Thickness(15, 10, 15, 10)
        self.progress_border.Visibility = Visibility.Collapsed
        Grid.SetRow(self.progress_border, row)
        
        stack = StackPanel()
        
        self.progress_text = TextBlock()
        self.progress_text.Text = "Scanning..."
        self.progress_text.FontSize = 12
        self.progress_text.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(self.progress_text)
        
        self.progress_bar = ProgressBar()
        self.progress_bar.Height = 20
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = 100
        stack.Children.Add(self.progress_bar)
        
        self.progress_border.Child = stack
        return self.progress_border
    
    def category_panel(self, row):
        """Create category list panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.BACKGROUND))
        border.Padding = Thickness(5)
        Grid.SetRow(border, row)
        
        main_stack = StackPanel()
        
        # Toolbar
        toolbar = StackPanel()
        toolbar.Orientation = Orientation.Horizontal
        toolbar.Margin = Thickness(5, 5, 5, 10)
        
        # Select All button
        btn_sel_all = Button()
        btn_sel_all.Content = u"\u2713 Select All"
        btn_sel_all.Width = 90
        btn_sel_all.Height = 28
        btn_sel_all.FontSize = 11
        btn_sel_all.Click += self.on_select_all
        
        # Deselect All button
        btn_desel_all = Button()
        btn_desel_all.Content = u"\u2717 Deselect All"
        btn_desel_all.Width = 90
        btn_desel_all.Height = 28
        btn_desel_all.FontSize = 11
        btn_desel_all.Margin = Thickness(5, 0, 0, 0)
        btn_desel_all.Click += self.on_deselect_all
        
        # Filter
        filter_label = TextBlock()
        filter_label.Text = u"\U0001F50D Filter:"
        filter_label.VerticalAlignment = VerticalAlignment.Center
        filter_label.Margin = Thickness(20, 0, 5, 0)
        filter_label.FontSize = 12
        
        self.txt_filter = TextBox()
        self.txt_filter.Width = 200
        self.txt_filter.Height = 26
        self.txt_filter.FontSize = 12
        self.txt_filter.TextChanged += self.on_filter_changed
        
        toolbar.Children.Add(btn_sel_all)
        toolbar.Children.Add(btn_desel_all)
        toolbar.Children.Add(filter_label)
        toolbar.Children.Add(self.txt_filter)
        
        main_stack.Children.Add(toolbar)
        
        # Scrollable category list
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll.PanningMode = PanningMode.VerticalOnly
        
        self.category_stack = StackPanel()
        scroll.Content = self.category_stack
        
        main_stack.Children.Add(scroll)
        border.Child = main_stack
        return border
    
    def settings_panel(self, row):
        """Create settings panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color("#FFF5F5F5"))
        border.Padding = Thickness(15, 10, 15, 10)
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 1, 0, 1)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        title = TextBlock()
        title.Text = u"\u2699 SETTINGS:"
        title.FontSize = 14  # Larger
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 8)
        stack.Children.Add(title)
        
        chk_stack = StackPanel()
        chk_stack.Orientation = Orientation.Horizontal
        
        self.chk_dry = CheckBox()
        self.chk_dry.Content = u"\u26A0 Dry Run (Preview only)"
        self.chk_dry.IsChecked = True
        self.chk_dry.FontSize = 13  # Larger
        self.chk_dry.ToolTip = "When enabled, shows what would be deleted without actually deleting"
        self.chk_dry.Checked += self.on_setting_changed
        self.chk_dry.Unchecked += self.on_setting_changed
        
        self.chk_system = CheckBox()
        self.chk_system.Content = u"\u26A0 Delete system types"
        self.chk_system.IsChecked = False
        self.chk_system.FontSize = 13  # Larger
        self.chk_system.Margin = Thickness(20, 0, 0, 0)
        self.chk_system.ToolTip = "Allow deletion of Revit system types (use with caution)"
        self.chk_system.Checked += self.on_setting_changed
        self.chk_system.Unchecked += self.on_setting_changed
        
        chk_stack.Children.Add(self.chk_dry)
        chk_stack.Children.Add(self.chk_system)
        stack.Children.Add(chk_stack)
        
        border.Child = stack
        return border
    
    def summary_panel(self, row):
        """Create summary panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.HEADER))
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        self.summary_text = TextBlock()
        self.summary_text.Text = u"\U0001F4CA SUMMARY: Categories: 0/0 scanned | Unused: 0 items | Selected: 0 items"
        self.summary_text.FontSize = 13  # Larger
        self.summary_text.FontWeight = FontWeights.Bold
        
        border.Child = self.summary_text
        return border
    
    def action_panel(self, row):
        """Create action buttons panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color("#FFF5F5F5"))
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        main_stack = StackPanel()
        
        grid = Grid()
        grid.ColumnDefinitions.Add(self.col_def(1, star=True))
        grid.ColumnDefinitions.Add(self.col_def(1, star=True))
        
        # Left
        left = StackPanel()
        left.Orientation = Orientation.Horizontal
        
        btn_safe = Button()
        btn_safe.Content = u"\u2713 Select Safe"
        btn_safe.Width = 100
        btn_safe.Height = 36
        btn_safe.FontSize = 13  # Larger
        btn_safe.Click += self.on_select_safe
        
        btn_none = Button()
        btn_none.Content = u"\u25A1 None"
        btn_none.Width = 80
        btn_none.Height = 36
        btn_none.FontSize = 13  # Larger
        btn_none.Margin = Thickness(5, 0, 0, 0)
        btn_none.Click += self.on_select_none
        
        left.Children.Add(btn_safe)
        left.Children.Add(btn_none)
        Grid.SetColumn(left, 0)
        
        # Right
        right = StackPanel()
        right.Orientation = Orientation.Horizontal
        
        btn_preview = Button()
        btn_preview.Content = u"\U0001F50D Preview"
        btn_preview.Width = 110
        btn_preview.Height = 36
        btn_preview.FontSize = 13  # Larger
        btn_preview.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_SUCCESS))
        btn_preview.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_preview.Click += self.on_preview
        
        self.btn_purge = Button()
        self.btn_purge.Content = u"\U0001F5D1 Purge"
        self.btn_purge.Width = 100
        self.btn_purge.Height = 36
        self.btn_purge.FontSize = 13  # Larger
        self.btn_purge.Margin = Thickness(5, 0, 0, 0)
        self.btn_purge.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_DANGER))
        self.btn_purge.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        self.btn_purge.Click += self.on_purge
        
        btn_close = Button()
        btn_close.Content = u"\u274C Close"
        btn_close.Width = 90
        btn_close.Height = 36
        btn_close.FontSize = 13  # Larger
        btn_close.Margin = Thickness(5, 0, 0, 0)
        btn_close.Background = SolidColorBrush(self.parse_color("#FF9E9E9E"))
        btn_close.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_close.Click += lambda s, e: self.Close()
        
        right.Children.Add(btn_preview)
        right.Children.Add(self.btn_purge)
        right.Children.Add(btn_close)
        Grid.SetColumn(right, 1)
        
        grid.Children.Add(left)
        grid.Children.Add(right)
        
        main_stack.Children.Add(grid)
        
        # Copyright at bottom
        copyright = TextBlock()
        copyright.Text = u"Copyright \u00A9 2025 Dang Quoc Truong (DQT) - All Rights Reserved"
        copyright.FontSize = 10
        copyright.Foreground = SolidColorBrush(self.parse_color("#FF999999"))
        copyright.TextAlignment = TextAlignment.Center
        copyright.Margin = Thickness(0, 8, 0, 0)
        main_stack.Children.Add(copyright)
        
        border.Child = main_stack
        return border
    
    def col_def(self, width, auto=False, star=False):
        """Create column definition"""
        from System.Windows.Controls import ColumnDefinition
        from System.Windows import GridLength, GridUnitType
        cd = ColumnDefinition()
        if auto:
            cd.Width = GridLength.Auto
        elif star:
            cd.Width = GridLength(width, GridUnitType.Star)
        else:
            cd.Width = GridLength(width)
        return cd
    
    def load_data(self):
        """Load groups and categories"""
        # Create groups
        self.groups = create_purge_groups()
        
        # Create categories
        categories = create_purge_categories()
        
        # Assign to groups
        for cat in categories:
            group = get_group_by_id(self.groups, cat.group_id)
            if group:
                group.add_category(cat)
        
        self.categories = categories
        
        # Populate UI
        self.populate_groups()
        self.populate_categories()
    
    def populate_groups(self):
        """Populate group checkboxes"""
        self.group_wrap.Children.Clear()
        
        for group in self.groups:
            chk = CheckBox()
            chk.Content = u"{} {} ({})".format(group.icon, group.name, group.total_categories)
            chk.IsChecked = group.default_checked
            chk.FontSize = 13  # Larger
            chk.ToolTip = group.description  # Add tooltip
            chk.Margin = Thickness(0, 0, 15, 5)
            chk.Tag = group.id
            
            self.group_checkboxes[group.id] = chk
            self.group_wrap.Children.Add(chk)
    
    def populate_categories(self):
        """Populate category list"""
        self.category_stack.Children.Clear()
        
        for group in self.groups:
            # Group box
            gb = GroupBox()
            gb.Header = u"{} {} ({} categories)".format(
                group.icon, group.name.upper(), group.total_categories
            )
            gb.FontSize = 14  # Larger
            gb.FontWeight = FontWeights.Bold
            gb.Margin = Thickness(5)
            gb.Padding = Thickness(10)
            
            # Categories
            stack = StackPanel()
            
            for cat in group.categories:
                cat_panel = self.create_category_row(cat)
                stack.Children.Add(cat_panel)
            
            gb.Content = stack
            self.category_stack.Children.Add(gb)
    
    def create_category_row(self, cat):
        """Create category row"""
        grid = Grid()
        grid.Margin = Thickness(0, 3, 0, 3)
        
        grid.ColumnDefinitions.Add(self.col_def(30))
        grid.ColumnDefinitions.Add(self.col_def(30))
        grid.ColumnDefinitions.Add(self.col_def(1, star=True))
        grid.ColumnDefinitions.Add(self.col_def(100))
        grid.ColumnDefinitions.Add(self.col_def(40))
        
        # Checkbox
        chk = CheckBox()
        chk.IsChecked = cat.default_checked
        chk.VerticalAlignment = VerticalAlignment.Center
        chk.Tag = cat.id
        chk.ToolTip = cat.description  # Add tooltip
        chk.Checked += self.on_cat_changed
        chk.Unchecked += self.on_cat_changed
        Grid.SetColumn(chk, 0)
        self.category_checkboxes[cat.id] = chk
        
        # Icon
        icon = TextBlock()
        icon.Text = cat.icon
        icon.FontSize = 15  # Slightly larger
        icon.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(icon, 1)
        
        # Name
        name = TextBlock()
        name.Text = cat.name
        name.FontSize = 13  # Larger
        name.FontWeight = FontWeights.Normal  # Normal weight, not bold
        name.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(name, 2)
        
        # Count
        count = TextBlock()
        count.Text = u"\u2014"  # em dash
        count.FontSize = 13  # Larger
        count.Foreground = SolidColorBrush(self.parse_color("#FF999999"))
        count.TextAlignment = TextAlignment.Right
        count.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(count, 3)
        self.category_labels[cat.id] = count
        
        # Safety
        safety = TextBlock()
        safety.Text = u"\u2713" if cat.is_safe() else u"\u26A0"
        safety.FontSize = 14
        safety.Foreground = SolidColorBrush(
            self.parse_color("#FF4CAF50" if cat.is_safe() else "#FFFFC107")
        )
        safety.TextAlignment = TextAlignment.Center
        safety.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(safety, 4)
        
        grid.Children.Add(chk)
        grid.Children.Add(icon)
        grid.Children.Add(name)
        grid.Children.Add(count)
        grid.Children.Add(safety)
        
        return grid
    
    # Event handlers
    
    def on_setting_changed(self, sender, e):
        """Settings changed"""
        self.dry_run = self.chk_dry.IsChecked
        self.delete_system = self.chk_system.IsChecked
    
    def on_cat_changed(self, sender, e):
        """Category selection changed"""
        self.update_summary()
    
    def on_select_all(self, sender, e):
        """Select all categories"""
        try:
            for cat in self.categories:
                if cat.is_scanned:  # Only select scanned categories
                    chk = self.category_checkboxes.get(cat.id)
                    if chk and chk.Visibility == Visibility.Visible:
                        chk.IsChecked = True
            self.update_summary()
        except Exception as ex:
            print("Error selecting all: {}".format(str(ex)))
    
    def on_deselect_all(self, sender, e):
        """Deselect all categories"""
        try:
            for cat in self.categories:
                chk = self.category_checkboxes.get(cat.id)
                if chk:
                    chk.IsChecked = False
            self.update_summary()
        except Exception as ex:
            print("Error deselecting all: {}".format(str(ex)))
    
    def on_filter_changed(self, sender, e):
        """Filter categories by name"""
        try:
            filter_text = self.txt_filter.Text.lower().strip()
            
            for cat in self.categories:
                # Get checkbox for this category
                chk = self.category_checkboxes.get(cat.id)
                if not chk:
                    continue
                
                # Show/hide based on filter
                if not filter_text:
                    # No filter - show all
                    chk.Visibility = Visibility.Visible
                else:
                    # Filter by name
                    if filter_text in cat.name.lower():
                        chk.Visibility = Visibility.Visible
                    else:
                        chk.Visibility = Visibility.Collapsed
        except Exception as ex:
            print("Error filtering: {}".format(str(ex)))
    
    def on_scan_click(self, sender, e):
        """Scan selected groups - FIXED VERSION"""
        print("="*60)
        print("DEBUG: on_scan_click (Scan Selected) called")
        print("="*60)
        
        # FIXED: Correctly build selected groups list
        selected = []
        
        # Debug: Print all group checkbox states
        print("\nDEBUG: Checking group checkboxes:")
        for group_id, chk in self.group_checkboxes.items():
            print("  Group '{}': IsChecked = {}".format(group_id, chk.IsChecked))
            
            if chk.IsChecked:
                # Find the group object by its ID
                for group in self.groups:
                    if group.id == group_id:
                        selected.append(group)
                        print("  -> Added group '{}' to selected (has {} categories)".format(
                            group.name, len(group.categories)))
                        break
        
        print("\nDEBUG: Total selected groups: {}".format(len(selected)))
        for g in selected:
            print("  - {} ({} categories):".format(g.name, len(g.categories)))
            for cat in g.categories:
                print("      - {}".format(cat.name))
        
        if not selected:
            MessageBox.Show("Please select at least one group.", "No Selection",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Scan ONLY selected groups
        print("\nDEBUG: Calling scan_groups with {} groups".format(len(selected)))
        self.scan_groups(selected)
    
    def on_scan_all_click(self, sender, e):
        """Scan all groups"""
        print("="*60)
        print("DEBUG: on_scan_all_click called")
        print("DEBUG: Will scan ALL {} groups".format(len(self.groups)))
        print("="*60)
        self.scan_groups(self.groups)
    
    def on_select_safe(self, sender, e):
        """Select all safe categories"""
        for cat in self.categories:
            if cat.is_safe() and cat.is_scanned and cat.unused_count > 0:
                chk = self.category_checkboxes.get(cat.id)
                if chk:
                    chk.IsChecked = True
        self.update_summary()
    
    def on_select_none(self, sender, e):
        """Deselect all"""
        for chk in self.category_checkboxes.values():
            chk.IsChecked = False
        self.update_summary()
    
    def on_preview(self, sender, e):
        """Preview selected"""
        items = self.get_selected_items()
        if not items:
            MessageBox.Show("No items selected.", "Nothing to Preview",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        preview = PreviewWindow(items)
        preview.ShowDialog()
    
    def on_purge(self, sender, e):
        """Purge selected"""
        items = self.get_selected_items()
        if not items:
            MessageBox.Show("No items selected.", "Nothing to Purge",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Confirm
        msg = "Purge {} items?\n\n".format(len(items))
        if self.dry_run:
            msg += "DRY RUN: No items will be deleted."
        else:
            msg += "WARNING: Items will be DELETED!"
        
        result = MessageBox.Show(msg, "Confirm Purge", MessageBoxButton.YesNo,
                                MessageBoxImage.Warning if not self.dry_run 
                                else MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        # Check document is writable
        if self.doc.IsReadOnly:
            MessageBox.Show("Document is read-only. Cannot purge.",
                          "Read-Only Document", MessageBoxButton.OK,
                          MessageBoxImage.Warning)
            return
        
        # Execute
        try:
            # Get selected categories (as objects, not dict)
            selected_categories = []
            for cat in self.categories:
                chk = self.category_checkboxes.get(cat.id)
                if chk and chk.IsChecked and cat.is_scanned and cat.unused_count > 0:
                    selected_categories.append(cat)
            
            if not selected_categories:
                MessageBox.Show("No items to purge.", "Nothing Selected",
                              MessageBoxButton.OK, MessageBoxImage.Information)
                return
            
            executor = PurgeExecutor(self.doc)
            executor.dry_run = self.dry_run
            
            # execute_purge returns tuple: (deleted_count, failed_count, deleted_items, failed_items)
            deleted_count, failed_count, deleted_items, failed_items = executor.execute_purge(selected_categories)
            
            # Log to history
            if self.history_manager:
                try:
                    self.history_manager.log_purge_operation(
                        deleted_items,
                        failed_items,
                        dry_run=self.dry_run
                    )
                except Exception as e:
                    print("Warning: Could not log to history: {}".format(str(e)))
            
            # Build message
            if self.dry_run:
                msg = u"\U0001F512 DRY RUN MODE - PREVIEW ONLY\n\n"
                msg += "NO ITEMS WERE ACTUALLY DELETED!\n"
                msg += "This is just a preview of what would be deleted.\n\n"
                msg += "Would delete: {}\n".format(deleted_count)
                msg += "Cannot delete: {}".format(failed_count)
            else:
                msg = "Purge Complete!\n\n"
                msg += "Deleted: {}\n".format(deleted_count)
                msg += "Failed: {}".format(failed_count)
            
            # Show failure reasons if any
            if failed_count > 0 and len(failed_items) > 0:
                msg += "\n\nFailed items:\n"
                # Show first 5 failure reasons
                for i, failed in enumerate(failed_items[:5]):
                    if isinstance(failed, dict):
                        reason = failed.get('reason', 'Unknown reason')
                        item_name = failed.get('item', {}).get('name', 'Unknown') if isinstance(failed.get('item'), dict) else 'Unknown'
                        msg += "- {}: {}\n".format(item_name, reason)
                    else:
                        msg += "- Failed item {}\n".format(i+1)
                
                if len(failed_items) > 5:
                    msg += "... and {} more\n".format(len(failed_items) - 5)
            
            # Show message
            MessageBox.Show(msg, 
                          "DRY RUN - Preview Results" if self.dry_run else "Purge Results", 
                          MessageBoxButton.OK,
                          MessageBoxImage.Information if failed_count == 0 else MessageBoxImage.Warning)
            
        except Exception as ex:
            MessageBox.Show("Error: {}".format(str(ex)), "Purge Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    # Scan logic
    
    def scan_groups(self, groups):
        """Scan selected groups"""
        print("\nDEBUG: scan_groups() called")
        print("DEBUG: Received {} groups to scan:".format(len(groups)))
        for g in groups:
            print("  - {} ({} categories)".format(g.name, len(g.categories)))
        
        if self.is_scanning:
            print("DEBUG: Already scanning, returning")
            return
        
        self.is_scanning = True
        self.btn_scan_sel.IsEnabled = False
        self.btn_scan_all.IsEnabled = False
        
        # Show progress
        self.progress_border.Visibility = Visibility.Visible
        self.progress_bar.Value = 0
        
        try:
            # Count total categories to scan
            total = sum(len(g.categories) for g in groups)
            print("DEBUG: Total categories to scan: {}".format(total))
            current = 0
            
            # IMPORTANT: Iterate over the 'groups' parameter, NOT self.groups!
            for group in groups:
                print("\nDEBUG: Scanning group: {}".format(group.name))
                for cat in group.categories:
                    # Scan
                    self.scan_category(cat)
                    
                    # Update progress
                    current += 1
                    pct = (current * 100.0) / total if total > 0 else 100
                    self.progress_bar.Value = pct
                    self.progress_text.Text = "Scanning: {} ({}/{})".format(
                        cat.name, current, total
                    )
            
            # Hide progress
            self.progress_border.Visibility = Visibility.Collapsed
            
            # Update summary
            self.update_summary()
            
            # Count actually scanned in this run
            scanned_this_run = sum(1 for g in groups for c in g.categories if c.is_scanned)
            
            # Show result
            total_unused = sum(c.unused_count for c in self.categories if c.is_scanned)
            
            print("\nDEBUG: Scan complete!")
            print("DEBUG: Scanned {} categories in this run".format(scanned_this_run))
            print("DEBUG: Total unused items found: {}".format(total_unused))
            
            if scanned_this_run == 0:
                MessageBox.Show("No scanners available for selected groups.\n\nThese categories are not implemented yet.",
                              "Scan Complete", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                MessageBox.Show("Scan complete!\n\nFound {} unused items.".format(total_unused),
                              "Scan Complete", MessageBoxButton.OK, MessageBoxImage.Information)
            
        except Exception as ex:
            MessageBox.Show("Scan error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
        finally:
            self.is_scanning = False
            self.btn_scan_sel.IsEnabled = True
            self.btn_scan_all.IsEnabled = True
    
    def scan_category(self, cat):
        """Scan single category"""
        try:
            print("DEBUG: Scanning category: {}".format(cat.name))
            print("DEBUG: Scanner class: {}".format(cat.scanner_class))
            
            scanner = create_scanner(cat.scanner_class, self.doc)
            if not scanner:
                cat.scan_error = "Scanner not found"
                print("ERROR: Scanner not found for {}".format(cat.scanner_class))
                return
            
            print("DEBUG: Scanner created: {}".format(type(scanner).__name__))
            
            result = scanner.scan()
            cat.is_scanned = True
            cat.unused_items = result if result else []
            
            print("DEBUG: Scan complete. Found {} items".format(len(cat.unused_items)))
            
            # Update label
            label = self.category_labels.get(cat.id)
            if label:
                if cat.unused_count > 0:
                    label.Text = "({})".format(cat.unused_count)
                    label.Foreground = SolidColorBrush(self.parse_color(Colors.HIGHLIGHT))
                else:
                    label.Text = "(0)"
                    label.Foreground = SolidColorBrush(self.parse_color("#FF999999"))
        
        except Exception as ex:
            cat.scan_error = str(ex)
            cat.is_scanned = True
            print("ERROR scanning {}: {}".format(cat.name, str(ex)))
            import traceback
            traceback.print_exc()
    
    def get_selected_items(self):
        """Get all selected items"""
        items = []
        for cat in self.categories:
            chk = self.category_checkboxes.get(cat.id)
            if chk and chk.IsChecked and cat.is_scanned:
                items.extend(cat.unused_items)
        return items
    
    def update_summary(self):
        """Update summary text"""
        scanned = sum(1 for c in self.categories if c.is_scanned)
        unused = sum(c.unused_count for c in self.categories if c.is_scanned)
        selected = len(self.get_selected_items())
        
        self.summary_text.Text = u"\U0001F4CA SUMMARY: Categories: {}/{} scanned | Unused: {} items | Selected: {} items".format(
            scanned, len(self.categories), unused, selected
        )