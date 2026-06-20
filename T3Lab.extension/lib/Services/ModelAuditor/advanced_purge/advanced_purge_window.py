# -*- coding: utf-8 -*-
"""
Advanced Purge Window
Main UI for Advanced Purge tool with red/orange danger theme

⚠️ WARNING: Contains dangerous operations!

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

# WPF imports
from System.Windows import Window, Application, WindowStartupLocation, Thickness
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
from System.Windows import FontWeights, Visibility, VerticalAlignment, HorizontalAlignment, TextAlignment, TextWrapping
from System.Windows.Controls import (
    Grid, StackPanel, Border, TextBlock, Button, CheckBox, TextBox,
    ScrollViewer, ProgressBar, GroupBox, WrapPanel, ListBox, ListBoxItem,
    Orientation, ScrollBarVisibility, PanningMode
)
from System.Windows.Media import SolidColorBrush, Color
from System import Action

# Revit imports
from Autodesk.Revit.DB import Transaction

# Local imports
from config_advanced import Colors, Fonts, Icons, Messages, Settings
from advanced_purge_group import (
    create_advanced_purge_groups, 
    get_group_by_id, 
    get_all_categories
)
from advanced_purge_executor import AdvancedPurgeExecutor
from preview_window import PreviewWindow

# Scanner imports
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'scanners'))
from unreferenced_views_scanner import UnreferencedViewsScanner
from workset_scanner import WorksetScanner
from model_deep_scanner import ModelDeepScanner
from dangerous_ops_scanner import DangerousOpsScanner


class AdvancedPurgeWindow(Window):
    """Advanced Purge main window"""
    
    def __init__(self, doc):
        """
        Initialize window
        
        Args:
            doc: Revit document
        """
        self.doc = doc
        
        # Create groups and categories (pass doc for dynamic workset categories)
        self.groups = create_advanced_purge_groups(doc)
        self.all_categories = get_all_categories(self.groups)
        
        # UI state
        self.group_checkboxes = {}
        self.category_checkboxes = {}
        self.category_panels = {}
        self.preview_shown = False
        self.last_clicked_category = None  # For Shift+Click multi-select
        
        # Settings - ALWAYS SAFE BY DEFAULT
        self.dry_run = True  # Cannot be False initially
        self.delete_system = False
        self.show_warnings = True  # Cannot be disabled
        
        # Setup window
        self.Title = "⚠️ Advanced Purge - Power User Tool"
        self.Width = Settings.WINDOW_WIDTH
        self.Height = Settings.WINDOW_HEIGHT
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        
        # Build UI
        self.build_ui()
    
    def build_ui(self):
        """Build the user interface"""
        main = Grid()
        
        # Rows
        main.RowDefinitions.Add(self.row_def(80))      # Header (warning colors)
        main.RowDefinitions.Add(self.row_def(150))     # Group selection (INCREASED from 120)
        main.RowDefinitions.Add(self.row_def(0, True)) # Progress (auto)
        main.RowDefinitions.Add(self.row_def(1, star=True))  # Categories (scrollable)
        main.RowDefinitions.Add(self.row_def(90))      # Settings (with warnings)
        main.RowDefinitions.Add(self.row_def(55))      # Summary
        main.RowDefinitions.Add(self.row_def(85))      # Actions
        
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
    
    def parse_color(self, hex_color):
        """Parse hex color string to Color (ARGB format)"""
        # hex_color format: #AARRGGBB
        # Color.FromArgb(alpha, red, green, blue)
        if len(hex_color) == 9:  # #AARRGGBB
            return Color.FromArgb(
                int(hex_color[1:3], 16),   # Alpha
                int(hex_color[3:5], 16),   # Red
                int(hex_color[5:7], 16),   # Green
                int(hex_color[7:9], 16)    # Blue
            )
        elif len(hex_color) == 7:  # #RRGGBB (no alpha)
            return Color.FromArgb(
                255,                       # Alpha = opaque
                int(hex_color[1:3], 16),   # Red
                int(hex_color[3:5], 16),   # Green
                int(hex_color[5:7], 16)    # Blue
            )
        else:
            # Default to white if invalid
            return Color.FromArgb(255, 255, 255, 255)

    
    def header_panel(self, row):
        """Create header with warning colors"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.HEADER))
        border.Padding = Thickness(15, 15, 15, 15)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        # Title with warning icon
        title = TextBlock()
        title.Text = u"{} Advanced Purge - Power User Tool".format(Icons.WARNING)
        title.FontSize = Fonts.TITLE
        title.FontWeight = FontWeights.Bold
        title.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
        
        # Copyright
        copyright = TextBlock()
        copyright.Text = u"Copyright \u00A9 2025 Dang Quoc Truong (DQT)"
        copyright.FontSize = Fonts.SMALL
        copyright.Foreground = SolidColorBrush(self.parse_color(Colors.TEXT_SECONDARY))
        copyright.Margin = Thickness(0, 5, 0, 0)
        
        # Warning message
        warning = TextBlock()
        warning.Text = u"{} DANGEROUS OPERATIONS - USE WITH EXTREME CAUTION!".format(Icons.DANGER)
        warning.FontSize = Fonts.NORMAL
        warning.FontWeight = FontWeights.Bold
        warning.Foreground = SolidColorBrush(self.parse_color(Colors.ERROR))
        warning.Margin = Thickness(0, 8, 0, 0)
        
        stack.Children.Add(title)
        stack.Children.Add(copyright)
        stack.Children.Add(warning)
        
        border.Child = stack
        return border
    
    def group_panel(self, row):
        """Create group selection panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.BACKGROUND))
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 0, 0, 1)
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        # Title
        title = TextBlock()
        title.Text = u"{} SELECT GROUPS TO SCAN:".format(Icons.SETTINGS)
        title.FontSize = Fonts.HEADER
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(title)
        
        # Groups in wrap panel
        wrap = WrapPanel()
        wrap.Orientation = Orientation.Horizontal
        
        for group in self.groups:
            cb = CheckBox()
            cb.Content = u"{} {} ({})".format(
                group.icon, group.name, len(group.categories)
            )
            cb.FontSize = Fonts.NORMAL
            cb.Margin = Thickness(0, 0, 20, 5)
            cb.ToolTip = group.description
            cb.Tag = group.id
            cb.Checked += self.on_group_checked
            cb.Unchecked += self.on_group_unchecked
            
            # Highlight dangerous groups
            if group.id == "dangerous":
                cb.Foreground = SolidColorBrush(self.parse_color(Colors.ERROR))
                cb.FontWeight = FontWeights.Bold
            elif group.id == "workset_cleanup":
                cb.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
            
            self.group_checkboxes[group.id] = cb
            wrap.Children.Add(cb)
        
        stack.Children.Add(wrap)
        
        # Scan buttons - separate from wrap panel
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.Margin = Thickness(0, 15, 0, 0)
        btn_panel.HorizontalAlignment = HorizontalAlignment.Left
        
        btn_scan_selected = Button()
        btn_scan_selected.Content = u"{} Scan Selected".format(Icons.SCAN)
        btn_scan_selected.Width = 150
        btn_scan_selected.Height = 36
        btn_scan_selected.FontSize = Fonts.NORMAL
        btn_scan_selected.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_PRIMARY))
        btn_scan_selected.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_scan_selected.Click += self.on_scan_selected
        
        btn_scan_all = Button()
        btn_scan_all.Content = u"{} Scan All".format(Icons.SCAN)
        btn_scan_all.Width = 130
        btn_scan_all.Height = 36
        btn_scan_all.FontSize = Fonts.NORMAL
        btn_scan_all.Margin = Thickness(10, 0, 0, 0)
        btn_scan_all.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_PRIMARY))
        btn_scan_all.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_scan_all.Click += self.on_scan_all
        
        btn_panel.Children.Add(btn_scan_selected)
        btn_panel.Children.Add(btn_scan_all)
        stack.Children.Add(btn_panel)
        
        border.Child = stack
        return border
    
    def progress_panel(self, row):
        """Create progress bar panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.PANEL))
        border.Padding = Thickness(15, 5, 15, 5)
        border.Visibility = Visibility.Collapsed
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        self.progress_text = TextBlock()
        self.progress_text.Text = "Scanning..."
        self.progress_text.FontSize = Fonts.SMALL
        self.progress_text.Margin = Thickness(0, 0, 0, 5)
        
        self.progress_bar = ProgressBar()
        self.progress_bar.Height = 20
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = 100
        self.progress_bar.Value = 0
        
        stack.Children.Add(self.progress_text)
        stack.Children.Add(self.progress_bar)
        
        self.progress_panel_border = border
        border.Child = stack
        return border
    
    def category_panel(self, row):
        """Create scrollable category list"""
        border = Border()
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 0, 0, 1)
        border.Padding = Thickness(10, 10, 10, 5)
        Grid.SetRow(border, row)
        
        # Grid layout for proper scrolling
        container_grid = Grid()
        container_grid.RowDefinitions.Add(self.row_def(0, auto=True))  # Filter
        container_grid.RowDefinitions.Add(self.row_def(1, star=True))  # Scrollable
        
        # Filter
        filter_panel = StackPanel()
        filter_panel.Orientation = Orientation.Horizontal
        filter_panel.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(filter_panel, 0)
        
        filter_label = TextBlock()
        filter_label.Text = u"{} Filter:".format(Icons.FILTER)
        filter_label.FontSize = Fonts.NORMAL
        filter_label.VerticalAlignment = VerticalAlignment.Center
        filter_label.Margin = Thickness(0, 0, 10, 0)
        
        self.txt_filter = TextBox()
        self.txt_filter.Width = 350
        self.txt_filter.Height = 28
        self.txt_filter.FontSize = Fonts.NORMAL
        self.txt_filter.TextChanged += self.on_filter_changed
        
        # Select All/None buttons
        btn_select_all = Button()
        btn_select_all.Content = "Select All"
        btn_select_all.Width = 90
        btn_select_all.Height = 28
        btn_select_all.FontSize = Fonts.SMALL
        btn_select_all.Margin = Thickness(15, 0, 5, 0)
        btn_select_all.Click += self.on_select_all
        
        btn_select_none = Button()
        btn_select_none.Content = "Select None"
        btn_select_none.Width = 95
        btn_select_none.Height = 28
        btn_select_none.FontSize = Fonts.SMALL
        btn_select_none.Margin = Thickness(5, 0, 0, 0)
        btn_select_none.Click += self.on_select_none
        
        filter_panel.Children.Add(filter_label)
        filter_panel.Children.Add(self.txt_filter)
        filter_panel.Children.Add(btn_select_all)
        filter_panel.Children.Add(btn_select_none)
        container_grid.Children.Add(filter_panel)
        
        # Scrollable content
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll.CanContentScroll = True
        scroll.PanningMode = PanningMode.VerticalOnly
        Grid.SetRow(scroll, 1)
        
        self.category_stack = StackPanel()
        scroll.Content = self.category_stack
        container_grid.Children.Add(scroll)
        
        border.Child = container_grid
        return border
    
    def settings_panel(self, row):
        """Create settings panel with warnings"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.PANEL))
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 0, 0, 1)
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        title = TextBlock()
        title.Text = u"{} SETTINGS:".format(Icons.SETTINGS)
        title.FontSize = Fonts.HEADER
        title.FontWeight = FontWeights.Bold
        title.Margin = Thickness(0, 0, 0, 8)
        stack.Children.Add(title)
        
        settings_wrap = WrapPanel()
        
        # Dry Run - ALWAYS ON by default
        self.cb_dry_run = CheckBox()
        self.cb_dry_run.Content = u"{} Dry Run (Preview only)".format(Icons.WARNING)
        self.cb_dry_run.IsChecked = True
        self.cb_dry_run.FontSize = Fonts.NORMAL
        self.cb_dry_run.FontWeight = FontWeights.Bold
        self.cb_dry_run.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
        self.cb_dry_run.Margin = Thickness(0, 0, 20, 0)
        self.cb_dry_run.Checked += self.on_dry_run_changed
        self.cb_dry_run.Unchecked += self.on_dry_run_changed
        
        # Delete system types
        self.cb_delete_system = CheckBox()
        self.cb_delete_system.Content = u"{} Delete system types".format(Icons.DANGER)
        self.cb_delete_system.IsChecked = False
        self.cb_delete_system.FontSize = Fonts.NORMAL
        self.cb_delete_system.Margin = Thickness(0, 0, 20, 0)
        self.cb_delete_system.Checked += self.on_delete_system_changed
        self.cb_delete_system.Unchecked += self.on_delete_system_changed
        
        settings_wrap.Children.Add(self.cb_dry_run)
        settings_wrap.Children.Add(self.cb_delete_system)
        stack.Children.Add(settings_wrap)
        
        # Warning text
        warning = TextBlock()
        warning.Text = Messages.DRY_RUN_REMINDER
        warning.FontSize = Fonts.SMALL
        warning.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
        warning.Margin = Thickness(0, 5, 0, 0)
        warning.TextWrapping = TextWrapping.Wrap
        self.warning_text = warning
        stack.Children.Add(warning)
        
        border.Child = stack
        return border
    
    def summary_panel(self, row):
        """Create summary panel"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.HEADER))
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        self.summary_text = TextBlock()
        self.summary_text.Text = u"{} SUMMARY: Categories: 0/0 scanned | Unused: 0 items | Selected: 0 items".format(Icons.INFO)
        self.summary_text.FontSize = Fonts.NORMAL
        self.summary_text.FontWeight = FontWeights.Bold
        
        border.Child = self.summary_text
        return border
    
    def action_panel(self, row):
        """Create action buttons"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.PANEL))
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        grid = Grid()
        grid.ColumnDefinitions.Add(self.col_def(1, star=True))
        grid.ColumnDefinitions.Add(self.col_def(0, auto=True))
        
        # Preview button
        self.btn_preview = Button()
        self.btn_preview.Content = u"{} Preview".format(Icons.PREVIEW)
        self.btn_preview.Width = 110
        self.btn_preview.Height = 36
        self.btn_preview.FontSize = Fonts.NORMAL
        self.btn_preview.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_SUCCESS))
        self.btn_preview.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        self.btn_preview.Click += self.on_preview
        Grid.SetColumn(self.btn_preview, 0)
        
        # Right buttons
        right_panel = StackPanel()
        right_panel.Orientation = Orientation.Horizontal
        Grid.SetColumn(right_panel, 1)
        
        self.btn_execute = Button()
        self.btn_execute.Content = u"{} Execute".format(Icons.EXECUTE)
        self.btn_execute.Width = 110
        self.btn_execute.Height = 36
        self.btn_execute.FontSize = Fonts.NORMAL
        self.btn_execute.Background = SolidColorBrush(self.parse_color(Colors.BUTTON_DANGER))
        self.btn_execute.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        self.btn_execute.Click += self.on_execute
        
        btn_close = Button()
        btn_close.Content = u"{} Close".format(Icons.CLOSE)
        btn_close.Width = 90
        btn_close.Height = 36
        btn_close.FontSize = Fonts.NORMAL
        btn_close.Margin = Thickness(5, 0, 0, 0)
        btn_close.Background = SolidColorBrush(self.parse_color("#FF9E9E9E"))
        btn_close.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_close.Click += lambda s, e: self.Close()
        
        right_panel.Children.Add(self.btn_preview)
        right_panel.Children.Add(self.btn_execute)
        right_panel.Children.Add(btn_close)
        
        grid.Children.Add(right_panel)
        border.Child = grid
        return border
    
    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================
    
    def on_group_checked(self, sender, e):
        """Group checkbox checked"""
        pass  # Auto-populate categories when scanned
    
    def on_group_unchecked(self, sender, e):
        """Group checkbox unchecked"""
        pass
    
    def on_filter_changed(self, sender, e):
        """Filter textbox changed"""
        filter_text = self.txt_filter.Text.lower()
        for panel in self.category_panels.values():
            category = panel.Tag
            if filter_text in category.name.lower():
                panel.Visibility = Visibility.Visible
            else:
                panel.Visibility = Visibility.Collapsed
    
    def on_dry_run_changed(self, sender, e):
        """Dry run checkbox changed"""
        self.dry_run = self.cb_dry_run.IsChecked
        if self.dry_run:
            self.warning_text.Text = Messages.DRY_RUN_REMINDER
        else:
            self.warning_text.Text = u"{} WARNING: Changes will be PERMANENT!".format(Icons.DANGER)
    
    def on_delete_system_changed(self, sender, e):
        """Delete system types checkbox changed"""
        self.delete_system = self.cb_delete_system.IsChecked
    
    def on_scan_selected(self, sender, e):
        """Scan selected groups"""
        selected_groups = [g for g in self.groups 
                          if self.group_checkboxes[g.id].IsChecked]
        if not selected_groups:
            MessageBox.Show(Messages.NO_SELECTION, "Warning", 
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        self.scan_groups(selected_groups)
    
    def on_scan_all(self, sender, e):
        """Scan all groups"""
        self.scan_groups(self.groups)
    
    def on_select_all(self, sender, e):
        """Select all visible categories"""
        for cb in self.category_checkboxes.values():
            if cb.Visibility == Visibility.Visible:
                cb.IsChecked = True
    
    def on_select_none(self, sender, e):
        """Deselect all categories"""
        for cb in self.category_checkboxes.values():
            cb.IsChecked = False
    
    def on_checkbox_mouse_down(self, sender, e):
        """Handle Shift+Click for multi-select"""
        from System.Windows.Input import Keyboard, Key
        
        # Check if Shift is pressed
        if not (Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)):
            # Normal click - just update last clicked
            self.last_clicked_category = sender.Tag
            return
        
        # Shift+Click - select range
        if self.last_clicked_category is None:
            self.last_clicked_category = sender.Tag
            return
        
        # Get all visible categories in order
        visible_categories = []
        for category in self.all_categories:
            cb = self.category_checkboxes.get(category.id)
            if cb and cb.Visibility == Visibility.Visible:
                visible_categories.append(category.id)
        
        # Find start and end indices
        try:
            start_idx = visible_categories.index(self.last_clicked_category)
            end_idx = visible_categories.index(sender.Tag)
            
            # Ensure start < end
            if start_idx > end_idx:
                start_idx, end_idx = end_idx, start_idx
            
            # Select range
            new_state = not sender.IsChecked  # Toggle to opposite
            for i in range(start_idx, end_idx + 1):
                cat_id = visible_categories[i]
                cb = self.category_checkboxes.get(cat_id)
                if cb:
                    cb.IsChecked = new_state
                    
        except (ValueError, IndexError):
            pass
        
        # Update last clicked
        self.last_clicked_category = sender.Tag
    
    def on_preview(self, sender, e):
        """Preview selected items"""
        try:
            selected_items = self.get_selected_items()
            if not selected_items:
                MessageBox.Show("Please select items to preview.", "Warning",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            # Convert dict {category: [items]} to flat list for PreviewWindow
            items_list = []
            for category, items in selected_items.items():
                for item in items:
                    items_list.append(item)
            
            if not items_list:
                MessageBox.Show("No items to preview.", "Warning",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            # Show preview window
            preview = PreviewWindow(items_list)
            preview.ShowDialog()
            self.preview_shown = True
            
        except Exception as ex:
            import traceback
            error_msg = "Error showing preview:\n\n{}\n\nTraceback:\n{}".format(
                str(ex), traceback.format_exc()
            )
            MessageBox.Show(error_msg, "Preview Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            print("=" * 80)
            print("PREVIEW ERROR:")
            print(error_msg)
            print("=" * 80)
    
    def on_execute(self, sender, e):
        """Execute purge"""
        if Settings.REQUIRE_PREVIEW and not self.preview_shown:
            MessageBox.Show(Messages.PREVIEW_REQUIRED, "Warning",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        selected_items = self.get_selected_items()
        if not selected_items:
            MessageBox.Show("No items selected!", "Warning",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Show confirmation
        if not self.dry_run:
            result = MessageBox.Show(
                Messages.CONFIRM_EXECUTE,
                "Confirm Execute",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning
            )
            if result != MessageBoxResult.Yes:
                return
        
        # Execute
        self.execute_purge(selected_items)
    
    # ========================================================================
    # HELPER METHODS
    # ========================================================================
    
    def scan_groups(self, groups):
        """Scan selected groups"""
        try:
            # Show progress panel
            self.progress_panel_border.Visibility = Visibility.Visible
            self.progress_bar.Value = 0
            
            # Clear existing categories
            self.category_stack.Children.Clear()
            self.category_checkboxes.clear()
            self.category_panels.clear()
            
            # Collect categories to scan
            categories_to_scan = []
            for group in groups:
                categories_to_scan.extend(group.categories)
            
            total_categories = len(categories_to_scan)
            
            # Scan each category
            for i, category in enumerate(categories_to_scan):
                # Update progress
                progress = int((i / float(total_categories)) * 100)
                self.progress_bar.Value = progress
                self.progress_text.Text = "Scanning {} ({}/{})...".format(
                    category.name, i+1, total_categories
                )
                self.Dispatcher.Invoke(Action(lambda: None))
                
                # Create scanner based on category type
                scanner = self._get_scanner_for_category(category)
                if scanner:
                    try:
                        # Scan category
                        items = scanner.scan(category)
                        category.count = len(items)
                        
                        # Store items for later
                        if not hasattr(category, 'items'):
                            category.items = []
                        category.items = items
                        
                        # Set status
                        if len(items) == 0:
                            category.status = 'safe'
                        elif category.is_dangerous:
                            category.status = 'error'
                        else:
                            category.status = 'warning'
                            
                    except Exception as e:
                        category.count = 0
                        category.status = 'error'
                        print("Error scanning {}: {}".format(category.name, str(e)))
                else:
                    # No scanner yet for this category
                    category.count = 0
                    category.status = None
            
            # Complete progress
            self.progress_bar.Value = 100
            self.progress_text.Text = "Scan complete!"
            self.Dispatcher.Invoke(Action(lambda: None))
            
            # Build category UI
            self._build_category_ui(groups)
            
            # Update summary
            self.update_summary()
            
            # Hide progress after delay
            import time
            time.sleep(1)
            self.progress_panel_border.Visibility = Visibility.Collapsed
            
        except Exception as e:
            self.progress_panel_border.Visibility = Visibility.Collapsed
            MessageBox.Show(
                "Error during scan: {}".format(str(e)),
                "Scan Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            )
    
    def _get_scanner_for_category(self, category):
        """Get appropriate scanner for category"""
        scanner_class = category.scanner_class
        
        if scanner_class == "UnreferencedViewsScanner":
            return UnreferencedViewsScanner(self.doc)
        elif scanner_class == "WorksetScanner":
            return WorksetScanner(self.doc)
        elif scanner_class == "ModelDeepScanner":
            return ModelDeepScanner(self.doc)
        elif scanner_class == "DangerousOpsScanner":
            return DangerousOpsScanner(self.doc)
        else:
            return None
    
    def _build_category_ui(self, groups):
        """Build category UI after scanning"""
        for group in groups:
            # Group header
            group_header = self._create_group_header(group)
            self.category_stack.Children.Add(group_header)
            
            # Category items
            for category in group.categories:
                if category.count is not None:  # Only show scanned categories
                    category_panel = self._create_category_panel(category)
                    self.category_stack.Children.Add(category_panel)
                    self.category_panels[category.id] = category_panel
    
    def _create_group_header(self, group):
        """Create group header UI"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.GROUP_EXPANDED))
        border.Padding = Thickness(10, 8, 10, 8)
        border.Margin = Thickness(0, 5, 0, 0)
        
        text = TextBlock()
        text.Text = u"{} {} ({} categories)".format(
            group.icon, group.name.upper(), len(group.categories)
        )
        text.FontSize = Fonts.HEADER
        text.FontWeight = FontWeights.Bold
        
        # Color dangerous groups
        if group.id == "dangerous":
            text.Foreground = SolidColorBrush(self.parse_color(Colors.ERROR))
        elif group.id == "workset_cleanup":
            text.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
        
        border.Child = text
        return border
    
    def _create_category_panel(self, category):
        """Create category panel UI"""
        grid = Grid()
        grid.Margin = Thickness(0, 2, 0, 2)
        grid.Tag = category
        
        # Columns
        grid.ColumnDefinitions.Add(self.col_def(30))     # Checkbox
        grid.ColumnDefinitions.Add(self.col_def(30))     # Icon
        grid.ColumnDefinitions.Add(self.col_def(1, star=True))  # Name
        grid.ColumnDefinitions.Add(self.col_def(80))     # Count
        grid.ColumnDefinitions.Add(self.col_def(30))     # Status
        
        # Checkbox
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Tag = category.id
        cb.IsChecked = (category.count > 0 and not category.is_dangerous)
        cb.PreviewMouseDown += self.on_checkbox_mouse_down  # Shift+Click support
        Grid.SetColumn(cb, 0)
        self.category_checkboxes[category.id] = cb
        
        # Icon (dangerous indicator)
        icon = TextBlock()
        if category.is_dangerous:
            icon.Text = Icons.DANGER
            icon.Foreground = SolidColorBrush(self.parse_color(Colors.ERROR))
        else:
            icon.Text = u"\u2022"  # Bullet point
        icon.FontSize = Fonts.NORMAL
        icon.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(icon, 1)
        
        # Name
        name = TextBlock()
        name.Text = category.name
        name.FontSize = Fonts.NORMAL
        name.VerticalAlignment = VerticalAlignment.Center
        name.ToolTip = category.description
        if category.is_dangerous:
            name.Foreground = SolidColorBrush(self.parse_color(Colors.ERROR))
        Grid.SetColumn(name, 2)
        
        # Count
        count = TextBlock()
        count.Text = "({})".format(category.count)
        count.FontSize = Fonts.NORMAL
        count.VerticalAlignment = VerticalAlignment.Center
        count.TextAlignment = TextAlignment.Right
        if category.count > 0:
            count.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
        Grid.SetColumn(count, 3)
        
        # Status icon
        status = TextBlock()
        if category.status == 'safe':
            status.Text = Icons.CHECK
            status.Foreground = SolidColorBrush(self.parse_color(Colors.SUCCESS))
        elif category.status == 'warning':
            status.Text = Icons.WARNING
            status.Foreground = SolidColorBrush(self.parse_color(Colors.WARNING))
        elif category.status == 'error':
            status.Text = Icons.DANGER
            status.Foreground = SolidColorBrush(self.parse_color(Colors.ERROR))
        status.FontSize = Fonts.NORMAL
        status.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(status, 4)
        
        grid.Children.Add(cb)
        grid.Children.Add(icon)
        grid.Children.Add(name)
        grid.Children.Add(count)
        grid.Children.Add(status)
        
        return grid
    
    def get_selected_items(self):
        """Get all selected items across categories"""
        items = {}
        for category in self.all_categories:
            if category.id in self.category_checkboxes:
                if self.category_checkboxes[category.id].IsChecked and category.count > 0:
                    # Get actual items from category
                    if hasattr(category, 'items'):
                        items[category] = category.items
        return items
    
    def execute_purge(self, items_by_category):
        """Execute purge operation"""
        try:
            # Show progress
            self.progress_panel_border.Visibility = Visibility.Visible
            self.progress_bar.Value = 0
            self.progress_text.Text = "Executing purge..."
            
            # Create executor
            executor = AdvancedPurgeExecutor(self.doc)
            executor.dry_run = self.dry_run
            
            # Execute
            def progress_callback(current, total, message):
                if total > 0:
                    self.progress_bar.Value = int((current / float(total)) * 100)
                self.progress_text.Text = message
                self.Dispatcher.Invoke(Action(lambda: None))
            
            deleted, failed = executor.execute_purge(items_by_category, progress_callback)
            
            # Show results
            result_msg = "Purge complete!\n\n"
            if self.dry_run:
                result_msg += "DRY RUN - No changes were made\n\n"
            result_msg += "Items processed: {}\n".format(len(deleted))
            if failed:
                result_msg += "Items failed: {}\n".format(len(failed))
            
            MessageBox.Show(
                result_msg,
                "Purge Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            )
            
            # Hide progress
            self.progress_panel_border.Visibility = Visibility.Collapsed
            
        except Exception as e:
            self.progress_panel_border.Visibility = Visibility.Collapsed
            MessageBox.Show(
                "Error during execution: {}".format(str(e)),
                "Execution Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            )
    
    def update_summary(self):
        """Update summary text"""
        scanned = sum(1 for cat in self.all_categories if cat.count is not None)
        total = len(self.all_categories)
        unused = sum(cat.count for cat in self.all_categories if cat.count)
        selected = sum(cat.count for cat in self.all_categories 
                      if cat.id in self.category_checkboxes 
                      and self.category_checkboxes[cat.id].IsChecked 
                      and cat.count)
        
        self.summary_text.Text = (
            u"{} SUMMARY: Categories: {}/{} scanned | Unused: {} items | Selected: {} items"
            .format(Icons.INFO, scanned, total, unused, selected)
        )