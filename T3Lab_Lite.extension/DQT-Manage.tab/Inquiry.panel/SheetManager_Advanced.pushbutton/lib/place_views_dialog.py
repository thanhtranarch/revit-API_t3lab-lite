# -*- coding: utf-8 -*-
"""
Sheet Manager - Place Views Dialog

Copyright (c) Dang Quoc Truong (DQT)
"""

from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, Thickness, GridLength
from System.Windows.Controls import (Grid, Label, Button, TextBox, ListBox, RadioButton, 
                                      ComboBox, StackPanel, RowDefinition, ColumnDefinition, GroupBox)
import System


class PlaceViewsDialog(Window):
    """Dialog for placing views on sheets"""
    
    def __init__(self, place_views_service, doc, selected_sheets):
        self.place_views_service = place_views_service
        self.doc = doc
        self.selected_sheets = selected_sheets
        self.all_views = []
        self.selected_views = []
        
        self._build_ui()
        self._load_views()
    
    def _build_ui(self):
        """Build dialog UI"""
        self.Title = "Place Views on Sheets"
        self.Width = 900
        self.Height = 650
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        # Main grid
        main_grid = Grid()
        main_grid.Margin = Thickness(20)
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))  # Info
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))  # Content
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))  # Buttons
        
        # Info label
        self.info_label = Label()
        self.info_label.Content = "Select views to place on {} sheet(s)".format(len(self.selected_sheets))
        self.info_label.FontSize = 14
        self.info_label.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(self.info_label, 0)
        main_grid.Children.Add(self.info_label)
        
        # Content grid
        content_grid = Grid()
        content_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(400)))
        content_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        content_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, System.Windows.GridUnitType.Star)))
        Grid.SetRow(content_grid, 1)
        
        # Left - Views list
        left_panel = Grid()
        left_panel.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        left_panel.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        left_panel.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        left_panel.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        Grid.SetColumn(left_panel, 0)
        
        views_label = Label()
        views_label.Content = "Available Views:"
        views_label.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(views_label, 0)
        left_panel.Children.Add(views_label)
        
        # Filter
        filter_panel = StackPanel()
        filter_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        Grid.SetRow(filter_panel, 1)
        
        filter_label = Label()
        filter_label.Content = "Filter:"
        filter_label.Width = 50
        filter_panel.Children.Add(filter_label)
        
        self.filter_combo = ComboBox()
        self.filter_combo.Width = 150
        self.filter_combo.Items.Add("All Views")
        self.filter_combo.Items.Add("Not on Sheets")
        self.filter_combo.Items.Add("Floor Plans")
        self.filter_combo.Items.Add("Sections")
        self.filter_combo.Items.Add("Elevations")
        self.filter_combo.SelectedIndex = 1  # Default: Not on sheets
        self.filter_combo.SelectionChanged += self.on_filter_changed
        filter_panel.Children.Add(self.filter_combo)
        
        left_panel.Children.Add(filter_panel)
        
        # Views listbox
        from System.Windows.Controls import ScrollViewer
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        Grid.SetRow(scroll, 2)
        
        self.views_panel = StackPanel()
        scroll.Content = self.views_panel
        left_panel.Children.Add(scroll)
        
        # Select buttons
        select_panel = StackPanel()
        select_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        select_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        Grid.SetRow(select_panel, 3)
        
        select_all_btn = Button()
        select_all_btn.Content = "Select All"
        select_all_btn.Width = 90
        select_all_btn.Height = 30
        select_all_btn.Margin = Thickness(5, 0, 5, 0)
        select_all_btn.Click += self.on_select_all_views
        select_panel.Children.Add(select_all_btn)
        
        select_none_btn = Button()
        select_none_btn.Content = "Select None"
        select_none_btn.Width = 90
        select_none_btn.Height = 30
        select_none_btn.Margin = Thickness(5, 0, 5, 0)
        select_none_btn.Click += self.on_select_none_views
        select_panel.Children.Add(select_none_btn)
        
        left_panel.Children.Add(select_panel)
        content_grid.Children.Add(left_panel)
        
        # Right - Options
        right_panel = Grid()
        right_panel.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        right_panel.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        Grid.SetColumn(right_panel, 2)
        
        # Placement mode
        mode_group = GroupBox()
        mode_group.Header = "Placement Mode"
        mode_group.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(mode_group, 0)
        
        mode_panel = StackPanel()
        mode_panel.Margin = Thickness(10)
        
        self.mode_one_per_sheet = RadioButton()
        self.mode_one_per_sheet.Content = "One view per sheet"
        self.mode_one_per_sheet.IsChecked = True
        self.mode_one_per_sheet.Margin = Thickness(0, 5, 0, 5)
        mode_panel.Children.Add(self.mode_one_per_sheet)
        
        desc1 = Label()
        desc1.Content = "  Place one view on each sheet"
        desc1.Foreground = System.Windows.Media.Brushes.Gray
        desc1.Margin = Thickness(20, 0, 0, 10)
        mode_panel.Children.Add(desc1)
        
        self.mode_all_on_each = RadioButton()
        self.mode_all_on_each.Content = "All views on each sheet"
        self.mode_all_on_each.Margin = Thickness(0, 5, 0, 5)
        mode_panel.Children.Add(self.mode_all_on_each)
        
        desc2 = Label()
        desc2.Content = "  Place all views on every sheet"
        desc2.Foreground = System.Windows.Media.Brushes.Gray
        desc2.Margin = Thickness(20, 0, 0, 10)
        mode_panel.Children.Add(desc2)
        
        self.mode_distribute = RadioButton()
        self.mode_distribute.Content = "Distribute evenly"
        self.mode_distribute.Margin = Thickness(0, 5, 0, 5)
        mode_panel.Children.Add(self.mode_distribute)
        
        desc3 = Label()
        desc3.Content = "  Distribute views across sheets"
        desc3.Foreground = System.Windows.Media.Brushes.Gray
        desc3.Margin = Thickness(20, 0, 0, 10)
        mode_panel.Children.Add(desc3)
        
        mode_group.Content = mode_panel
        right_panel.Children.Add(mode_group)
        
        # Arrangement
        arrange_group = GroupBox()
        arrange_group.Header = "Auto-Arrange"
        Grid.SetRow(arrange_group, 1)
        
        arrange_panel = StackPanel()
        arrange_panel.Margin = Thickness(10)
        
        grid_label = Label()
        grid_label.Content = "Grid Layout:"
        grid_label.FontWeight = System.Windows.FontWeights.Bold
        arrange_panel.Children.Add(grid_label)
        
        grid_options = StackPanel()
        grid_options.Orientation = System.Windows.Controls.Orientation.Horizontal
        grid_options.Margin = Thickness(0, 10, 0, 10)
        
        rows_label = Label()
        rows_label.Content = "Rows:"
        rows_label.Width = 50
        grid_options.Children.Add(rows_label)
        
        self.rows_combo = ComboBox()
        self.rows_combo.Width = 60
        for i in range(1, 5):
            self.rows_combo.Items.Add(i)
        self.rows_combo.SelectedIndex = 1  # Default: 2 rows
        grid_options.Children.Add(self.rows_combo)
        
        cols_label = Label()
        cols_label.Content = "  Columns:"
        cols_label.Width = 80
        grid_options.Children.Add(cols_label)
        
        self.cols_combo = ComboBox()
        self.cols_combo.Width = 60
        for i in range(1, 5):
            self.cols_combo.Items.Add(i)
        self.cols_combo.SelectedIndex = 1  # Default: 2 cols
        grid_options.Children.Add(self.cols_combo)
        
        arrange_panel.Children.Add(grid_options)
        
        preview_label = Label()
        preview_label.Content = "Views will be arranged in a 2x2 grid\non each sheet"
        preview_label.Foreground = System.Windows.Media.Brushes.Gray
        arrange_panel.Children.Add(preview_label)
        
        arrange_group.Content = arrange_panel
        right_panel.Children.Add(arrange_group)
        
        content_grid.Children.Add(right_panel)
        main_grid.Children.Add(content_grid)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        Grid.SetRow(btn_panel, 2)
        
        place_btn = Button()
        place_btn.Content = "Place Views"
        place_btn.Width = 120
        place_btn.Height = 35
        place_btn.Margin = Thickness(0, 0, 10, 0)
        place_btn.Click += self.on_place_click
        btn_panel.Children.Add(place_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 100
        cancel_btn.Height = 35
        cancel_btn.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(cancel_btn)
        
        main_grid.Children.Add(btn_panel)
        self.Content = main_grid
    
    def _load_views(self):
        """Load available views"""
        self.all_views = self.place_views_service.get_placeable_views()
        self.update_views_list()
    
    def update_views_list(self):
        """Update views list based on filter"""
        from System.Windows.Controls import CheckBox
        
        self.views_panel.Children.Clear()
        
        filter_mode = self.filter_combo.SelectedIndex
        
        for view_info in self.all_views:
            # Apply filter
            if filter_mode == 1 and view_info['on_sheet']:  # Not on sheets
                continue
            elif filter_mode == 2 and 'Floor' not in view_info['type']:  # Floor plans
                continue
            elif filter_mode == 3 and 'Section' not in view_info['type']:  # Sections
                continue
            elif filter_mode == 4 and 'Elevation' not in view_info['type']:  # Elevations
                continue
            
            cb = CheckBox()
            cb.Content = "{} [{}]{}".format(
                view_info['name'],
                view_info['type'],
                " (on sheet)" if view_info['on_sheet'] else ""
            )
            cb.Tag = view_info
            cb.Margin = Thickness(5, 2, 5, 2)
            if view_info['on_sheet']:
                cb.Foreground = System.Windows.Media.Brushes.Gray
            self.views_panel.Children.Add(cb)
    
    def on_filter_changed(self, sender, args):
        """Filter changed"""
        self.update_views_list()
    
    def on_select_all_views(self, sender, args):
        """Select all views"""
        from System.Windows.Controls import CheckBox
        for child in self.views_panel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = True
    
    def on_select_none_views(self, sender, args):
        """Deselect all views"""
        from System.Windows.Controls import CheckBox
        for child in self.views_panel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = False
    
    def on_place_click(self, sender, args):
        """Place views on sheets"""
        from System.Windows.Controls import CheckBox
        
        # Get selected views
        selected_views = []
        for child in self.views_panel.Children:
            if isinstance(child, CheckBox) and child.IsChecked:
                selected_views.append(child.Tag['element'])
        
        if not selected_views:
            MessageBox.Show("Please select at least one view", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Get placement mode
        if self.mode_one_per_sheet.IsChecked:
            mode = 'one_per_sheet'
        elif self.mode_all_on_each.IsChecked:
            mode = 'all_on_each'
        else:
            mode = 'distribute'
        
        # Confirm
        result = MessageBox.Show(
            "Place {} view(s) on {} sheet(s)?\n\nMode: {}".format(
                len(selected_views), len(self.selected_sheets), mode),
            "Confirm",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question)
        
        if result != System.Windows.MessageBoxResult.Yes:
            return
        
        # Place views
        from Autodesk.Revit.DB import Transaction
        
        t = Transaction(self.doc, "Place Views on Sheets")
        t.Start()
        
        try:
            rows = int(self.rows_combo.SelectedItem)
            cols = int(self.cols_combo.SelectedItem)
            
            if mode == 'one_per_sheet':
                success_count = 0
                for i, sheet in enumerate(self.selected_sheets):
                    if i < len(selected_views):
                        viewport = self.place_views_service.place_view_on_sheet(
                            sheet, selected_views[i])
                        if viewport:
                            success_count += 1
            
            elif mode == 'all_on_each':
                success_count = 0
                for sheet in self.selected_sheets:
                    viewports = self.place_views_service.auto_arrange_views_on_sheet(
                        sheet, selected_views, rows, cols)
                    success_count += len(viewports)
            
            else:  # distribute
                placements = self.place_views_service.batch_place_views(
                    self.selected_sheets, selected_views, mode='distribute')
                success_count = len(placements)
            
            t.Commit()
            
            MessageBox.Show(
                "Successfully placed {} viewport(s)!".format(success_count),
                "Success",
                MessageBoxButton.OK,
                MessageBoxImage.Information)
            
            self.Close()
            
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
