# -*- coding: utf-8 -*-
"""
Sheet Manager - Sheet List Tab (In-Place Manager Style)
Beautiful UI with summary cards, left panel, and organized layout

Copyright © Dang Quoc Truong (DQT)
"""

import sys
import os

from System.Windows.Controls import (Grid, DataGrid, StackPanel, TextBlock, Border, 
                                      TextBox, ComboBox, ComboBoxItem, Button,
                                      RowDefinition, ColumnDefinition, Orientation,
                                      DataGridSelectionMode, DataGridSelectionUnit,
                                      ContextMenu, MenuItem, ListBox, ListBoxItem,
                                      ScrollViewer)
from System.Windows import Thickness, GridLength, GridUnitType, HorizontalAlignment, VerticalAlignment, FontWeights
from System.Windows.Media import SolidColorBrush, Color
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Input import Key
import System

from components import DataGridHelper


class SheetListTab(object):
    """Sheet list tab with In-Place Manager style UI"""
    
    def __init__(self, parent_window, doc):
        self.parent_window = parent_window
        self.doc = doc
        self.revit_service = parent_window.revit_service
        self.change_tracker = parent_window.change_tracker
        
        # Services from parent
        self.excel_service = parent_window.excel_service
        self.sheet_sets_service = parent_window.sheet_sets_service
        self.place_views_service = parent_window.place_views_service
        self.params_service = parent_window.params_service
        
        # Data
        self.all_items = []
        self.filtered_items = []
        self.last_clicked_index = -1
        
        # UI elements for updating
        self.txt_total = None
        self.txt_selected = None
        self.txt_categories = None
        self.search_box = None
        self.filter_combo = None
        self.category_list = None
        self.data_grid = None
        
        # Custom parameter columns tracking
        self.custom_columns = {}  # {column_name: parameter_name}
        
        # Build content
        self.content = self._build_content()
    
    def _build_content(self):
        """Build main content with In-Place Manager style layout"""
        main_grid = Grid()
        main_grid.Margin = Thickness(12)
        
        # Row definitions
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # Header
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # Summary Cards
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # Main Content
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # Action Buttons
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # Tips
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Auto)))  # Copyright
        
        # Row 0: Header
        header = self._create_header()
        Grid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        # Row 1: Summary Cards
        cards = self._create_summary_cards()
        Grid.SetRow(cards, 1)
        main_grid.Children.Add(cards)
        
        # Row 2: Main Content (Left Panel + DataGrid)
        content = self._create_main_content()
        Grid.SetRow(content, 2)
        main_grid.Children.Add(content)
        
        # Row 3: Action Buttons
        buttons = self._create_action_buttons()
        Grid.SetRow(buttons, 3)
        main_grid.Children.Add(buttons)
        
        # Row 4: Tips
        tips = self._create_tips()
        Grid.SetRow(tips, 4)
        main_grid.Children.Add(tips)
        
        # Row 5: Copyright Footer
        footer = self._create_footer()
        Grid.SetRow(footer, 5)
        main_grid.Children.Add(footer)
        
        # Load data
        self.refresh_data()
        
        return main_grid
    
    def _create_header(self):
        """Create header with title and settings button"""
        border = Border()
        border.Background = SolidColorBrush(Color.FromArgb(255, 240, 204, 136))  # #F0CC88
        border.CornerRadius = System.Windows.CornerRadius(5)
        border.Padding = Thickness(12, 8, 12, 8)
        border.Margin = Thickness(0, 0, 0, 10)
        
        grid = Grid()
        
        # Left side - Title
        left_panel = StackPanel()
        
        title = TextBlock()
        title.Text = "Sheet Manager v1.0"
        title.FontSize = 17
        title.FontWeight = FontWeights.Bold
        
        subtitle = TextBlock()
        subtitle.Text = "by Dang Quoc Truong (DQT)"
        subtitle.FontSize = 10
        subtitle.Foreground = SolidColorBrush(Color.FromArgb(255, 93, 78, 55))  # #5D4E37
        subtitle.Margin = Thickness(0, 2, 0, 0)
        
        left_panel.Children.Add(title)
        left_panel.Children.Add(subtitle)
        
        grid.Children.Add(left_panel)
        
        border.Child = grid
        return border
    
    def _create_summary_cards(self):
        """Create summary cards (Total, Selected, Categories, Issues)"""
        grid = Grid()
        grid.Margin = Thickness(0, 0, 0, 10)
        
        # 4 columns
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        # Card 1: Total
        card1, self.txt_total = self._create_card("TOTAL", "0", 0)
        Grid.SetColumn(card1, 0)
        grid.Children.Add(card1)
        
        # Card 2: Selected
        card2, self.txt_selected = self._create_card("SELECTED", "0", 1, "#E5B85C")
        Grid.SetColumn(card2, 1)
        grid.Children.Add(card2)
        
        # Card 3: Categories
        card3, self.txt_categories = self._create_card("CATEGORIES", "0", 2, "#4CAF50")
        Grid.SetColumn(card3, 2)
        grid.Children.Add(card3)
        
        # Card 4: Filters
        card4, txt_filters = self._create_card("FILTERS", "Active", 3, "#666666")
        Grid.SetColumn(card4, 3)
        grid.Children.Add(card4)
        
        return grid
    
    def _create_card(self, label, value, margin_index, value_color=None):
        """Create a summary card"""
        border = Border()
        border.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))  # White
        border.BorderBrush = SolidColorBrush(Color.FromArgb(255, 212, 184, 122))  # #D4B87A
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(10, 6, 10, 6)
        
        # Margin based on position
        if margin_index == 0:
            border.Margin = Thickness(0, 0, 4, 0)
        elif margin_index == 3:
            border.Margin = Thickness(4, 0, 0, 0)
        else:
            border.Margin = Thickness(4, 0, 4, 0)
        
        panel = StackPanel()
        
        # Label
        label_text = TextBlock()
        label_text.Text = label
        label_text.FontSize = 9
        label_text.Foreground = SolidColorBrush(Color.FromArgb(255, 102, 102, 102))  # #666
        panel.Children.Add(label_text)
        
        # Value
        value_text = TextBlock()
        value_text.Text = value
        value_text.FontSize = 22
        value_text.FontWeight = FontWeights.Bold
        
        if value_color:
            color_int = int(value_color.replace("#", ""), 16)
            r = (color_int >> 16) & 255
            g = (color_int >> 8) & 255
            b = color_int & 255
            value_text.Foreground = SolidColorBrush(Color.FromArgb(255, r, g, b))
        
        panel.Children.Add(value_text)
        
        border.Child = panel
        
        # Return both border and value_text for reference
        return border, value_text
    
    def _create_main_content(self):
        """Create main content with left panel and datagrid"""
        grid = Grid()
        
        # 2 columns: Left panel (180px) + DataGrid (*)
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(180)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        # Left Panel
        left_panel = self._create_left_panel()
        Grid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        # DataGrid
        self.data_grid = self._create_datagrid()
        Grid.SetColumn(self.data_grid, 1)
        grid.Children.Add(self.data_grid)
        
        return grid
    
    def _create_left_panel(self):
        """Create left panel with search and filters"""
        border = Border()
        border.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))  # White
        border.BorderBrush = SolidColorBrush(Color.FromArgb(255, 212, 184, 122))  # #D4B87A
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 0, 8, 0)
        
        panel = StackPanel()
        
        # SEARCH label
        label1 = TextBlock()
        label1.Text = "SEARCH"
        label1.FontSize = 9
        label1.FontWeight = FontWeights.SemiBold
        label1.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label1)
        
        # Search TextBox
        self.search_box = TextBox()
        self.search_box.Padding = Thickness(6, 4, 6, 4)
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self.on_filter_changed
        panel.Children.Add(self.search_box)
        
        # FILTER label
        label2 = TextBlock()
        label2.Text = "FILTER"
        label2.FontSize = 9
        label2.FontWeight = FontWeights.SemiBold
        label2.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label2)
        
        # Filter ComboBox
        self.filter_combo = ComboBox()
        self.filter_combo.Padding = Thickness(6, 4, 6, 4)
        self.filter_combo.Margin = Thickness(0, 0, 0, 10)
        
        # Add items
        for item_text in ["All Sheets", "Has Revision", "No Revision"]:
            item = ComboBoxItem()
            item.Content = item_text
            self.filter_combo.Items.Add(item)
        
        self.filter_combo.SelectedIndex = 0
        self.filter_combo.SelectionChanged += self.on_filter_changed
        panel.Children.Add(self.filter_combo)
        
        # CATEGORY label (placeholder)
        label3 = TextBlock()
        label3.Text = "QUICK SELECT"
        label3.FontSize = 9
        label3.FontWeight = FontWeights.SemiBold
        label3.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label3)
        
        # Quick action buttons
        btn_all = Button()
        btn_all.Content = "Select All"
        btn_all.Padding = Thickness(8, 4, 8, 4)
        btn_all.Margin = Thickness(0, 0, 0, 4)
        btn_all.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        btn_all.Click += self.on_select_all_click
        panel.Children.Add(btn_all)
        
        btn_clear = Button()
        btn_clear.Content = "Clear All"
        btn_clear.Padding = Thickness(8, 4, 8, 4)
        btn_clear.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        btn_clear.Click += self.on_clear_all_click
        panel.Children.Add(btn_clear)
        
        border.Child = panel
        return border
    
    def _create_datagrid(self):
        """Create data grid with Excel-like selection"""
        grid = DataGrid()
        grid.AutoGenerateColumns = False
        grid.IsReadOnly = True
        grid.SelectionMode = DataGridSelectionMode.Extended
        grid.SelectionUnit = DataGridSelectionUnit.FullRow
        grid.CanUserSortColumns = True
        grid.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        grid.BorderBrush = SolidColorBrush(Color.FromArgb(255, 212, 184, 122))
        grid.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.Horizontal
        grid.HorizontalGridLinesBrush = SolidColorBrush(Color.FromArgb(255, 238, 238, 238))
        grid.RowBackground = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        grid.AlternatingRowBackground = SolidColorBrush(Color.FromArgb(255, 255, 253, 245))  # #FFFDF5
        
        # Columns
        grid.Columns.Add(DataGridHelper.create_checkbox_column(u"☐", "is_selected", 100))
        grid.Columns.Add(DataGridHelper.create_text_column("Sheet Number", "sheet_number", 150))
        grid.Columns.Add(DataGridHelper.create_text_column("Sheet Name", "sheet_name", 350))
        grid.Columns.Add(DataGridHelper.create_text_column("Current Revision", "current_revision", 150))
        
        # Events
        grid.SelectionChanged += self.on_selection_changed
        grid.PreviewKeyDown += self.on_key_down
        grid.PreviewMouseRightButtonDown += self.on_header_right_click
        
        return grid
    
    def _create_action_buttons(self):
        """Create action buttons row"""
        border = Border()
        border.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        border.BorderBrush = SolidColorBrush(Color.FromArgb(255, 212, 184, 122))
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 10, 0, 0)
        
        grid = Grid()
        
        # Center panel with main buttons
        center_panel = StackPanel()
        center_panel.Orientation = Orientation.Horizontal
        center_panel.HorizontalAlignment = HorizontalAlignment.Center
        
        # Create buttons
        primary_color = SolidColorBrush(Color.FromArgb(255, 240, 204, 136))  # #F0CC88
        white_color = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        green_color = SolidColorBrush(Color.FromArgb(255, 76, 175, 80))  # #4CAF50
        red_color = SolidColorBrush(Color.FromArgb(255, 255, 107, 107))  # #FF6B6B
        
        # Create Sheet
        btn_create = self._create_button("Create Sheet", primary_color)
        btn_create.Click += self.on_create_click
        center_panel.Children.Add(btn_create)
        
        # Duplicate
        btn_dup = self._create_button("Duplicate", primary_color)
        btn_dup.Click += self.on_duplicate_click
        center_panel.Children.Add(btn_dup)
        
        # Rename
        btn_rename = self._create_button("Rename", primary_color)
        btn_rename.Click += self.on_rename_click
        center_panel.Children.Add(btn_rename)
        
        # Export Excel
        btn_excel = self._create_button("Excel", primary_color)
        btn_excel.Click += self.on_export_excel_click
        center_panel.Children.Add(btn_excel)
        
        # Export HTML
        btn_html = self._create_button("Report", primary_color)
        btn_html.Click += self.on_export_excel_click  # Will add HTML export
        center_panel.Children.Add(btn_html)
        
        # ViewSheet Sets
        
        # Place Views
        btn_place = self._create_button("Place Views", green_color, white_fg=True)
        btn_place.Click += self.on_place_views_click
        center_panel.Children.Add(btn_place)
        
        # Parameters
        btn_params = self._create_button("Parameters", primary_color)
        btn_params.Click += self.on_parameters_click
        center_panel.Children.Add(btn_params)
        
        # Delete
        btn_delete = self._create_button("Delete", red_color, white_fg=True)
        btn_delete.Click += self.on_delete_click
        center_panel.Children.Add(btn_delete)
        
        grid.Children.Add(center_panel)
        
        # Refresh button on right
        btn_refresh = Button()
        btn_refresh.Content = "Refresh"
        btn_refresh.Padding = Thickness(10, 5, 10, 5)
        btn_refresh.Margin = Thickness(2)
        btn_refresh.Background = white_color
        btn_refresh.HorizontalAlignment = HorizontalAlignment.Right
        btn_refresh.Click += self.on_refresh_click
        grid.Children.Add(btn_refresh)
        
        border.Child = grid
        return border
    
    def _create_button(self, text, bg_brush, white_fg=False):
        """Helper to create a button"""
        btn = Button()
        btn.Content = text
        btn.Padding = Thickness(10, 5, 10, 5)
        btn.Margin = Thickness(2)
        btn.Background = bg_brush
        
        if white_fg:
            btn.Foreground = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        
        return btn
    
    def _create_tips(self):
        """Create tips row"""
        grid = Grid()
        grid.Margin = Thickness(0, 8, 0, 0)
        
        # Tips text on left
        tips = TextBlock()
        tips.Text = "Click row to select | Shift+Click for range | Space Bar to tick selected | Right-click header to add columns"
        tips.FontSize = 10
        tips.Foreground = SolidColorBrush(Color.FromArgb(255, 136, 136, 136))  # #888
        grid.Children.Add(tips)
        
        # Close button on right
        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Padding = Thickness(15, 5, 15, 5)
        btn_close.Background = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
        btn_close.HorizontalAlignment = HorizontalAlignment.Right
        btn_close.Click += self.on_close_click
        grid.Children.Add(btn_close)
        
        return grid
    
    def _create_footer(self):
        """Create copyright footer"""
        border = Border()
        border.Background = SolidColorBrush(Color.FromArgb(255, 240, 204, 136))  # #F0CC88
        border.CornerRadius = System.Windows.CornerRadius(3)
        border.Padding = Thickness(8, 5, 8, 5)
        border.Margin = Thickness(0, 8, 0, 0)
        
        text = TextBlock()
        text.Text = u"© 2024 Dang Quoc Truong (DQT) - All Rights Reserved"
        text.FontSize = 10
        text.FontWeight = FontWeights.SemiBold
        text.HorizontalAlignment = HorizontalAlignment.Center
        text.Foreground = SolidColorBrush(Color.FromArgb(255, 93, 78, 55))  # #5D4E37
        
        border.Child = text
        return border
    
    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================
    
    def on_filter_changed(self, sender, args):
        """Handle filter change"""
        self.apply_filters()
    
    def on_selection_changed(self, sender, args):
        """Update selected count AND auto-sync checkboxes with row selection"""
        if self.txt_selected:
            self.txt_selected.Text = str(self.data_grid.SelectedItems.Count)
        
        # AUTO-SYNC: Update checkboxes to match row selection
        try:
            # Get all items
            all_items = []
            if self.data_grid.Items:
                for item in self.data_grid.Items:
                    all_items.append(item)
            
            # Get selected items
            selected_items = []
            for item in self.data_grid.SelectedItems:
                selected_items.append(item)
            
            # Update is_selected property
            changes_made = False
            for item in all_items:
                new_state = item in selected_items
                if item.is_selected != new_state:
                    item.is_selected = new_state
                    changes_made = True
            
            # Refresh UI if changes were made
            if changes_made:
                try:
                    self.data_grid.Items.Refresh()
                except:
                    from System.Collections.ObjectModel import ObservableCollection
                    new_source = ObservableCollection[object](self.filtered_items)
                    self.data_grid.ItemsSource = new_source
        except Exception as e:
            print("DEBUG: Auto-sync error: {}".format(str(e)))
    
    def on_select_all_click(self, sender, args):
        """Select all rows"""
        self.data_grid.SelectAll()
    
    def on_clear_all_click(self, sender, args):
        """Clear all selections"""
        self.data_grid.UnselectAll()
    
    def on_close_click(self, sender, args):
        """Close window"""
        self.parent_window.Close()
    
    def on_key_down(self, sender, args):
        """Handle Space bar to toggle selected rows"""
        from System.Windows.Input import Key
        
        if args.Key != Key.Space:
            return
        
        # Get selected items
        try:
            selected_items = []
            for item in self.data_grid.SelectedItems:
                selected_items.append(item)
            
            if len(selected_items) == 0:
                return
            
            # Determine new state from first item
            new_state = not selected_items[0].is_selected
            
            # Mark event as handled
            args.Handled = True
            
            # Toggle all selected items
            items_to_toggle = list(selected_items)
            
            def toggle_items():
                try:
                    for item in items_to_toggle:
                        item.is_selected = new_state
                    
                    # Try to refresh UI
                    try:
                        self.data_grid.Items.Refresh()
                    except:
                        from System.Collections.ObjectModel import ObservableCollection
                        new_source = ObservableCollection[object](self.filtered_items)
                        self.data_grid.ItemsSource = new_source
                    
                    self.update_status()
                except Exception as e:
                    print("ERROR: Toggle failed: {}".format(str(e)))
            
            from System.Windows.Threading import DispatcherPriority
            from System import Action
            self.data_grid.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                Action(toggle_items)
            )
        except Exception as e:
            print("ERROR: Key handler failed: {}".format(str(e)))
    
    def on_header_right_click(self, sender, args):
        """Show context menu on right-click"""
        try:
            print("DEBUG: Right-click detected")
            
            # Get click position and element
            pos = args.GetPosition(self.data_grid)
            element = self.data_grid.InputHitTest(pos)
            
            print("DEBUG: Clicked element type: {}".format(type(element).__name__))
            
            # Check if clicked on header - walk up visual tree
            from System.Windows.Media import VisualTreeHelper
            
            current = element
            is_header = False
            depth = 0
            
            while current is not None and depth < 10:
                element_type = type(current).__name__
                print("DEBUG: Level {}: {}".format(depth, element_type))
                
                # Check if it's a header element by type name
                if element_type in ["DataGridColumnHeader", "DataGridHeaderBorder"]:
                    is_header = True
                    print("DEBUG: Found header type '{}' at level {}".format(element_type, depth))
                    break
                
                try:
                    current = VisualTreeHelper.GetParent(current)
                except:
                    print("DEBUG: Cannot get parent, stopping")
                    break
                
                depth += 1
            
            if not is_header:
                print("DEBUG: Not a header click, ignoring")
                return
            
            print("DEBUG: Creating context menu")
            
            # Create context menu
            menu = ContextMenu()
            
            # Add parameter column
            add_item = MenuItem()
            add_item.Header = "Add Parameter Column..."
            add_item.Click += self.on_add_parameter_column
            menu.Items.Add(add_item)
            
            print("DEBUG: Added 'Add Parameter Column' item")
            
            # Remove custom columns (if any exist)
            if len(self.custom_columns) > 0:
                print("DEBUG: Adding {} remove items".format(len(self.custom_columns)))
                
                # Add separator
                separator = MenuItem()
                separator.Header = "-"
                separator.IsEnabled = False
                menu.Items.Add(separator)
                
                for col_name in self.custom_columns.keys():
                    remove_item = MenuItem()
                    remove_item.Header = "Remove '{}'".format(col_name)
                    remove_item.Tag = col_name
                    remove_item.Click += self.on_remove_parameter_column
                    menu.Items.Add(remove_item)
            
            # Show menu
            print("DEBUG: Showing menu")
            menu.PlacementTarget = self.data_grid
            menu.IsOpen = True
            
            # Mark as handled
            args.Handled = True
            
            print("DEBUG: Menu shown successfully")
            
        except Exception as e:
            print("ERROR in header right-click: {}".format(str(e)))
            import traceback
            traceback.print_exc()
    
    def on_add_parameter_column(self, sender, args):
        """Add a custom parameter column"""
        print("DEBUG: on_add_parameter_column called!")
        
        from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, Thickness, GridLength
        from System.Windows.Controls import Grid, Label, ComboBox, ComboBoxItem, Button, RowDefinition
        import System
        
        try:
            print("DEBUG: Getting sheets for parameters...")
            
            # Get all parameters from a sample sheet
            from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet
            collector = FilteredElementCollector(self.doc).OfClass(ViewSheet)
            sample_sheet = None
            for sheet in collector:
                if not sheet.IsTemplate:
                    sample_sheet = sheet
                    break
            
            if not sample_sheet:
                MessageBox.Show("No sheets found in project", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
                return
            
            # Get all parameters
            params = []
            for param in sample_sheet.Parameters:
                if param.Definition and param.Definition.Name:
                    param_name = param.Definition.Name
                    if param_name not in params:
                        params.append(param_name)
            
            params.sort()
            
            if not params:
                MessageBox.Show("No parameters found", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
                return
            
            # Create selection dialog
            dialog = Window()
            dialog.Title = "Add Parameter Column"
            dialog.Width = 500
            dialog.Height = 400
            dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
            
            main_grid = Grid()
            main_grid.Margin = Thickness(20)
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))  # Title
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # List
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))  # Buttons
            
            # Title and instruction
            title_panel = StackPanel()
            Grid.SetRow(title_panel, 0)
            
            title = TextBlock()
            title.Text = "Select Parameter to Add as Column"
            title.FontSize = 14
            title.FontWeight = FontWeights.Bold
            title.Margin = Thickness(0, 0, 0, 5)
            title_panel.Children.Add(title)
            
            instruction = TextBlock()
            instruction.Text = "Choose a sheet parameter from the list below:"
            instruction.FontSize = 11
            instruction.Foreground = SolidColorBrush(Color.FromArgb(255, 100, 100, 100))
            title_panel.Children.Add(instruction)
            
            main_grid.Children.Add(title_panel)
            
            # ListBox with search
            list_container = Grid()
            list_container.Margin = Thickness(0, 10, 0, 10)
            Grid.SetRow(list_container, 1)
            
            list_container.RowDefinitions.Add(RowDefinition(Height=GridLength(35)))  # Search
            list_container.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # List
            
            # Search box
            search_label = TextBlock()
            search_label.Text = "Search:"
            search_label.Margin = Thickness(0, 0, 0, 5)
            search_label.FontSize = 10
            Grid.SetRow(search_label, 0)
            list_container.Children.Add(search_label)
            
            search_box = TextBox()
            search_box.Margin = Thickness(50, 0, 0, 5)
            search_box.Padding = Thickness(5)
            Grid.SetRow(search_box, 0)
            
            # ListBox instead of ComboBox
            scroll = ScrollViewer()
            scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
            scroll.BorderBrush = SolidColorBrush(Color.FromArgb(255, 200, 200, 200))
            scroll.BorderThickness = Thickness(1)
            Grid.SetRow(scroll, 1)
            
            param_listbox = ListBox()
            param_listbox.Padding = Thickness(5)
            
            # Add parameters to ListBox
            for param_name in params:
                item = ListBoxItem()
                item.Content = param_name
                item.Padding = Thickness(8, 6, 8, 6)
                item.FontSize = 12
                param_listbox.Items.Add(item)
            
            if param_listbox.Items.Count > 0:
                param_listbox.SelectedIndex = 0
            
            scroll.Content = param_listbox
            list_container.Children.Add(scroll)
            
            # Search functionality
            def on_search_changed(s, e):
                search_text = search_box.Text.lower()
                param_listbox.Items.Clear()
                
                for param_name in params:
                    if not search_text or search_text in param_name.lower():
                        item = ListBoxItem()
                        item.Content = param_name
                        item.Padding = Thickness(8, 6, 8, 6)
                        item.FontSize = 12
                        param_listbox.Items.Add(item)
                
                if param_listbox.Items.Count > 0:
                    param_listbox.SelectedIndex = 0
            
            search_box.TextChanged += on_search_changed
            list_container.Children.Add(search_box)
            
            main_grid.Children.Add(list_container)
            
            # Info text
            info_text = TextBlock()
            info_text.Text = "{} parameters available".format(len(params))
            info_text.FontSize = 10
            info_text.Foreground = SolidColorBrush(Color.FromArgb(255, 120, 120, 120))
            info_text.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            info_text.Margin = Thickness(0, 0, 0, 10)
            Grid.SetRow(info_text, 2)
            main_grid.Children.Add(info_text)
            
            # Buttons
            btn_panel = StackPanel()
            btn_panel.Orientation = Orientation.Horizontal
            btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            btn_panel.VerticalAlignment = System.Windows.VerticalAlignment.Bottom
            Grid.SetRow(btn_panel, 2)
            
            result_holder = [False]
            
            def on_ok(s, e):
                if param_listbox.SelectedIndex < 0:
                    MessageBox.Show("Please select a parameter", "Info",
                                  MessageBoxButton.OK, MessageBoxImage.Information)
                    return
                result_holder[0] = True
                dialog.Close()
            
            def on_cancel(s, e):
                result_holder[0] = False
                dialog.Close()
            
            ok_btn = Button()
            ok_btn.Content = "Add Column"
            ok_btn.Width = 100
            ok_btn.Height = 32
            ok_btn.Margin = Thickness(5, 0, 5, 0)
            ok_btn.Background = SolidColorBrush(Color.FromArgb(255, 76, 175, 80))  # Green
            ok_btn.Foreground = SolidColorBrush(Color.FromArgb(255, 255, 255, 255))
            ok_btn.FontWeight = FontWeights.SemiBold
            ok_btn.Click += on_ok
            btn_panel.Children.Add(ok_btn)
            
            cancel_btn = Button()
            cancel_btn.Content = "Cancel"
            cancel_btn.Width = 100
            cancel_btn.Height = 32
            cancel_btn.Click += on_cancel
            btn_panel.Children.Add(cancel_btn)
            
            main_grid.Children.Add(btn_panel)
            dialog.Content = main_grid
            
            dialog.ShowDialog()
            
            if not result_holder[0] or param_listbox.SelectedIndex < 0:
                return
            
            param_name = param_listbox.SelectedItem.Content
            
            # Check if already exists
            if param_name in self.custom_columns.values():
                MessageBox.Show("This parameter is already displayed", "Info",
                              MessageBoxButton.OK, MessageBoxImage.Information)
                return
            
            # Add column
            col_index = len(self.custom_columns)
            col_name = param_name
            binding_name = "param_{}".format(col_index)
            
            from components import DataGridHelper
            new_col = DataGridHelper.create_text_column(col_name, binding_name, 150)
            self.data_grid.Columns.Add(new_col)
            
            # Track it
            self.custom_columns[col_name] = param_name
            
            # Update all items with parameter value
            for item in self.all_items:
                try:
                    param = item.element.LookupParameter(param_name)
                    if param and param.HasValue:
                        from Autodesk.Revit.DB import StorageType
                        if param.StorageType == StorageType.String:
                            value = param.AsString() or ""
                        elif param.StorageType == StorageType.Integer:
                            value = str(param.AsInteger())
                        elif param.StorageType == StorageType.Double:
                            value = str(param.AsDouble())
                        elif param.StorageType == StorageType.ElementId:
                            value = str(param.AsValueString() or param.AsElementId().IntegerValue)
                        else:
                            value = ""
                    else:
                        value = ""
                    setattr(item, binding_name, value)
                except:
                    setattr(item, binding_name, "")
            
            # Refresh display
            self.update_display()
            
        except Exception as e:
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def on_remove_parameter_column(self, sender, args):
        """Remove a custom parameter column"""
        try:
            col_name = sender.Tag
            if col_name not in self.custom_columns:
                return
            
            # Find and remove column
            for col in self.data_grid.Columns:
                if col.Header == col_name:
                    self.data_grid.Columns.Remove(col)
                    break
            
            # Remove from tracking
            del self.custom_columns[col_name]
            
            # Refresh display
            self.update_display()
            
        except Exception as e:
            from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    # ========================================================================
    # DATA METHODS
    # ========================================================================
    
    def refresh_data(self):
        """Refresh sheet data from Revit"""
        print("DEBUG: Loading sheets...")
        self.all_items = self.get_all_sheets()
        print("DEBUG: Total sheets loaded: {}".format(len(self.all_items)))
        
        # Re-apply custom columns if any
        self._reload_custom_columns()
        
        # Update display
        self.apply_filters()
        self.update_summary()
    
    def get_all_sheets(self):
        """Get all sheets from Revit"""
        from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet, BuiltInParameter
        # Try multiple import paths for SheetItemModel
        try:
            from models.sheet_item_model import SheetItemModel
        except ImportError:
            try:
                from lib.models.sheet_item_model import SheetItemModel
            except ImportError:
                # Define inline if import fails
                class SheetItemModel(object):
                    def __init__(self, element):
                        self.element = element
                        self.is_selected = False
                        self.status = ""
                        self.sheet_number = element.SheetNumber if element else ""
                        self.sheet_name = element.Name if element else ""
                        try:
                            from Autodesk.Revit.DB import BuiltInParameter
                            rev_param = element.get_Parameter(BuiltInParameter.SHEET_CURRENT_REVISION)
                            if rev_param and rev_param.HasValue:
                                self.current_revision = rev_param.AsString() or ""
                            else:
                                self.current_revision = ""
                        except:
                            self.current_revision = ""

        
        items = []
        collector = FilteredElementCollector(self.doc).OfClass(ViewSheet)
        
        for sheet in collector:
            if sheet.IsTemplate:
                continue
            
            try:
                item = SheetItemModel(sheet)
                items.append(item)
            except Exception as e:
                print("ERROR loading sheet {}: {}".format(sheet.SheetNumber, str(e)))
        
        return items
    
    def apply_filters(self):
        """Apply search and filter"""
        # Get filter values
        search_text = self.search_box.Text.lower() if self.search_box.Text else ""
        filter_index = self.filter_combo.SelectedIndex
        
        # Filter items
        self.filtered_items = []
        for item in self.all_items:
            # Search filter
            if search_text:
                if search_text not in item.sheet_number.lower() and \
                   search_text not in item.sheet_name.lower():
                    continue
            
            # Type filter
            if filter_index == 1:  # Has Revision
                if not item.current_revision:
                    continue
            elif filter_index == 2:  # No Revision
                if item.current_revision:
                    continue
            
            self.filtered_items.append(item)
        
        # Update display
        self.update_display()
    
    def update_display(self, items=None):
        """Update DataGrid display"""
        if items is None:
            items = self.filtered_items
        
        from System.Collections.ObjectModel import ObservableCollection
        new_source = ObservableCollection[object](items)
        self.data_grid.ItemsSource = new_source
    
    def update_summary(self):
        """Update summary cards"""
        if self.txt_total:
            self.txt_total.Text = str(len(self.all_items))
        if self.txt_categories:
            self.txt_categories.Text = "1"  # Only sheets for now
    
    def update_status(self):
        """Update status (selected count)"""
        if self.txt_selected:
            self.txt_selected.Text = str(self.data_grid.SelectedItems.Count)
    
    def _reload_custom_columns(self):
        """Reload values for custom parameter columns"""
        if not self.custom_columns:
            return
        
        print("DEBUG: Re-loading {} custom columns: {}".format(
            len(self.custom_columns), 
            list(self.custom_columns.keys())))
        
        # Get column indices
        col_index = 0
        for col_name, param_name in self.custom_columns.items():
            binding_name = "param_{}".format(col_index)
            
            for item in self.all_items:
                value = self._get_parameter_value(item.element, param_name)
                setattr(item, binding_name, value)
            
            col_index += 1


    def on_create_click(self, sender, args):
        """Create new sheet(s) with dialog"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, Window
        from System.Windows.Controls import Label, TextBox, Button, ComboBox, CheckBox, StackPanel, Grid, RowDefinition
        from System.Windows import Thickness, GridLength
        from Autodesk.Revit.DB import Transaction
        
        # Get titleblocks with debug
        print("DEBUG: Getting titleblocks...")
        titleblocks = self.revit_service.get_all_titleblocks()
        print("DEBUG: Found {} titleblocks".format(len(titleblocks)))
        
        # Use list to store result
        result = [False, "", "", None, False, 1]  # [result, number, name, titleblock_id, is_batch, count]
        
        # Create dialog
        dialog = Window()
        dialog.Title = "Create New Sheet(s)"
        dialog.Width = 450
        dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        main_grid = Grid()
        main_grid.Margin = Thickness(20)
        
        if not titleblocks:
            # No titleblocks - show error
            dialog.Height = 200
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(100)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
            
            # Error message
            error_label = Label()
            error_label.Content = "No titleblocks found!\n\nCannot create sheet without titleblock.\nPlease load a titleblock family first."
            error_label.Foreground = System.Windows.Media.Brushes.Red
            error_label.FontSize = 14
            Grid.SetRow(error_label, 0)
            main_grid.Children.Add(error_label)
            
            # Close button
            close_btn = Button()
            close_btn.Content = "Close"
            close_btn.Width = 80
            close_btn.Height = 30
            close_btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            close_btn.Click += lambda s, e: dialog.Close()
            Grid.SetRow(close_btn, 1)
            main_grid.Children.Add(close_btn)
            
            dialog.Content = main_grid
            dialog.ShowDialog()
            return
        
        # Has titleblocks - full dialog
        dialog.Height = 380
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))  # Number
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))  # Name
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(80)))  # Titleblock
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))  # Batch create
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))  # Count
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))  # Buttons
        
        # Sheet Number
        number_label = Label()
        number_label.Content = "Sheet Number:"
        Grid.SetRow(number_label, 0)
        main_grid.Children.Add(number_label)
        
        number_box = TextBox()
        number_box.Margin = Thickness(0, 25, 0, 0)
        number_box.Text = "A000"
        Grid.SetRow(number_box, 0)
        main_grid.Children.Add(number_box)
        
        # Sheet Name
        name_label = Label()
        name_label.Content = "Sheet Name:"
        Grid.SetRow(name_label, 1)
        main_grid.Children.Add(name_label)
        
        name_box = TextBox()
        name_box.Margin = Thickness(0, 25, 0, 0)
        name_box.Text = "New Sheet"
        Grid.SetRow(name_box, 1)
        main_grid.Children.Add(name_box)
        
        # Titleblock
        tb_label = Label()
        tb_label.Content = "Titleblock:"
        Grid.SetRow(tb_label, 2)
        main_grid.Children.Add(tb_label)
        
        tb_combo = ComboBox()
        tb_combo.Margin = Thickness(0, 25, 0, 0)
        for tb_id, tb_name in titleblocks:
            tb_combo.Items.Add(tb_name)
        tb_combo.SelectedIndex = 0
        Grid.SetRow(tb_combo, 2)
        main_grid.Children.Add(tb_combo)
        
        # Batch create checkbox
        batch_check = CheckBox()
        batch_check.Content = "Create multiple sheets"
        batch_check.Margin = Thickness(0, 20, 0, 0)
        batch_check.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(batch_check, 3)
        main_grid.Children.Add(batch_check)
        
        # Count (only enabled when batch checked)
        count_label = Label()
        count_label.Content = "Number of sheets:"
        count_label.IsEnabled = False
        Grid.SetRow(count_label, 4)
        main_grid.Children.Add(count_label)
        
        count_box = TextBox()
        count_box.Margin = Thickness(0, 25, 0, 0)
        count_box.Text = "10"
        count_box.IsEnabled = False
        Grid.SetRow(count_box, 4)
        main_grid.Children.Add(count_box)
        
        # Enable/disable count based on checkbox
        def on_batch_changed(s, e):
            is_checked = batch_check.IsChecked
            count_label.IsEnabled = is_checked
            count_box.IsEnabled = is_checked
        
        batch_check.Checked += on_batch_changed
        batch_check.Unchecked += on_batch_changed
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        Grid.SetRow(btn_panel, 5)
        
        def on_ok(s, e):
            if not number_box.Text:
                MessageBox.Show("Sheet number required", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            result[0] = True
            result[1] = number_box.Text
            result[2] = name_box.Text
            result[3] = titleblocks[tb_combo.SelectedIndex][0]
            result[4] = batch_check.IsChecked == True
            
            if result[4]:  # Batch mode
                try:
                    result[5] = int(count_box.Text)
                    if result[5] < 1 or result[5] > 100:
                        MessageBox.Show("Please enter 1-100 sheets", "Error",
                                      MessageBoxButton.OK, MessageBoxImage.Warning)
                        return
                except:
                    MessageBox.Show("Please enter a valid number", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Warning)
                    return
            
            dialog.Close()
        
        def on_cancel(s, e):
            dialog.Close()
        
        ok_btn = Button()
        ok_btn.Content = "Create"
        ok_btn.Width = 80
        ok_btn.Height = 30
        ok_btn.Margin = Thickness(0, 0, 10, 0)
        ok_btn.Click += on_ok
        btn_panel.Children.Add(ok_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Height = 30
        cancel_btn.Click += on_cancel
        btn_panel.Children.Add(cancel_btn)
        
        main_grid.Children.Add(btn_panel)
        dialog.Content = main_grid
        
        # Show dialog
        dialog.ShowDialog()
        
        if not result[0]:
            return
        
        # Create sheet(s)
        t = Transaction(self.doc, "Create Sheet(s)")
        t.Start()
        
        try:
            if result[4]:  # Batch mode
                success_count = 0
                for i in range(result[5]):
                    # Auto-increment sheet number
                    sheet_num = result[1]
                    if i > 0:
                        # Try to increment number (A001 -> A002, etc)
                        try:
                            # Find last digits
                            base = sheet_num.rstrip('0123456789')
                            num_str = sheet_num[len(base):]
                            if num_str:
                                num = int(num_str) + i
                                sheet_num = base + str(num).zfill(len(num_str))
                            else:
                                sheet_num = sheet_num + "_" + str(i + 1)
                        except:
                            sheet_num = result[1] + "_" + str(i + 1)
                    
                    sheet_name = result[2]
                    if result[5] > 1:
                        sheet_name = "{} - {}".format(result[2], i + 1)
                    
                    new_sheet = self.revit_service.create_sheet(sheet_num, sheet_name, result[3])
                    if new_sheet:
                        success_count += 1
                
                t.Commit()
                self.refresh_data()
                MessageBox.Show("Created {} sheet(s) successfully".format(success_count),
                              "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                # Single sheet
                new_sheet = self.revit_service.create_sheet(result[1], result[2], result[3])
                
                if new_sheet:
                    t.Commit()
                    self.refresh_data()
                    MessageBox.Show("Sheet created successfully", "Success",
                                  MessageBoxButton.OK, MessageBoxImage.Information)
                else:
                    t.RollBack()
                    MessageBox.Show("Failed to create sheet", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error)
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
        """Create new sheet with dialog"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, Window
        from System.Windows.Controls import Label, TextBox, Button, ComboBox, StackPanel, Grid, RowDefinition
        from System.Windows import Thickness, GridLength
        from Autodesk.Revit.DB import Transaction
        
        # Get titleblocks
        titleblocks = self.revit_service.get_all_titleblocks()
        
        # Use list to store result
        result = [False, "", "", None]  # [result, number, name, titleblock_id]
        
        # Create dialog
        dialog = Window()
        dialog.Title = "Create New Sheet"
        dialog.Width = 450
        dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        main_grid = Grid()
        main_grid.Margin = Thickness(20)
        
        if not titleblocks:
            # No titleblocks - simple dialog
            dialog.Height = 200
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
            
            # Warning
            warning_label = Label()
            warning_label.Content = "No titleblocks found in project.\nSheet will be created without titleblock."
            warning_label.Foreground = System.Windows.Media.Brushes.Red
            Grid.SetRow(warning_label, 0)
            main_grid.Children.Add(warning_label)
            
            # Sheet Number
            number_label = Label()
            number_label.Content = "Sheet Number:"
            Grid.SetRow(number_label, 1)
            main_grid.Children.Add(number_label)
            
            number_box = TextBox()
            number_box.Margin = Thickness(0, 25, 0, 0)
            number_box.Text = "A000"
            Grid.SetRow(number_box, 1)
            main_grid.Children.Add(number_box)
            
            # Buttons
            btn_panel = StackPanel()
            btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
            btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            Grid.SetRow(btn_panel, 2)
            
            def on_ok_no_tb(s, e):
                MessageBox.Show(
                    "Cannot create sheet without titleblock.\n\nPlease load a titleblock family first.",
                    "No Titleblocks",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                )
                dialog.Close()
            
            def on_cancel(s, e):
                dialog.Close()
            
            ok_btn = Button()
            ok_btn.Content = "OK"
            ok_btn.Width = 80
            ok_btn.Height = 30
            ok_btn.Margin = Thickness(0, 0, 10, 0)
            ok_btn.Click += on_ok_no_tb
            btn_panel.Children.Add(ok_btn)
            
            cancel_btn = Button()
            cancel_btn.Content = "Cancel"
            cancel_btn.Width = 80
            cancel_btn.Height = 30
            cancel_btn.Click += on_cancel
            btn_panel.Children.Add(cancel_btn)
            
            main_grid.Children.Add(btn_panel)
            dialog.Content = main_grid
            dialog.ShowDialog()
            return
        
        # Has titleblocks - full dialog
        dialog.Height = 280
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(80)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        
        # Sheet Number
        number_label = Label()
        number_label.Content = "Sheet Number:"
        Grid.SetRow(number_label, 0)
        main_grid.Children.Add(number_label)
        
        number_box = TextBox()
        number_box.Margin = Thickness(0, 25, 0, 0)
        number_box.Text = "A000"
        Grid.SetRow(number_box, 0)
        main_grid.Children.Add(number_box)
        
        # Sheet Name
        name_label = Label()
        name_label.Content = "Sheet Name:"
        Grid.SetRow(name_label, 1)
        main_grid.Children.Add(name_label)
        
        name_box = TextBox()
        name_box.Margin = Thickness(0, 25, 0, 0)
        name_box.Text = "New Sheet"
        Grid.SetRow(name_box, 1)
        main_grid.Children.Add(name_box)
        
        # Titleblock
        tb_label = Label()
        tb_label.Content = "Titleblock:"
        Grid.SetRow(tb_label, 2)
        main_grid.Children.Add(tb_label)
        
        tb_combo = ComboBox()
        tb_combo.Margin = Thickness(0, 25, 0, 0)
        for tb_id, tb_name in titleblocks:
            tb_combo.Items.Add(tb_name)
        tb_combo.SelectedIndex = 0
        Grid.SetRow(tb_combo, 2)
        main_grid.Children.Add(tb_combo)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        Grid.SetRow(btn_panel, 3)
        
        def on_ok(s, e):
            if not number_box.Text:
                MessageBox.Show("Sheet number required", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            
            result[0] = True
            result[1] = number_box.Text
            result[2] = name_box.Text
            result[3] = titleblocks[tb_combo.SelectedIndex][0]
            dialog.Close()
        
        def on_cancel(s, e):
            dialog.Close()
        
        ok_btn = Button()
        ok_btn.Content = "Create"
        ok_btn.Width = 80
        ok_btn.Height = 30
        ok_btn.Margin = Thickness(0, 0, 10, 0)
        ok_btn.Click += on_ok
        btn_panel.Children.Add(ok_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Height = 30
        cancel_btn.Click += on_cancel
        btn_panel.Children.Add(cancel_btn)
        
        main_grid.Children.Add(btn_panel)
        dialog.Content = main_grid
        
        # Show dialog
        dialog.ShowDialog()
        
        if not result[0]:
            return
        
        # Create sheet
        t = Transaction(self.doc, "Create Sheet")
        t.Start()
        
        try:
            new_sheet = self.revit_service.create_sheet(result[1], result[2], result[3])
            
            if new_sheet:
                t.Commit()
                self.refresh_data()
                MessageBox.Show("Sheet created successfully", "Success",
                              MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                t.RollBack()
                MessageBox.Show("Failed to create sheet", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_duplicate_click(self, sender, args):
        """Duplicate selected sheets with batch option"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, Window
        from System.Windows.Controls import Label, TextBox, Button, CheckBox, StackPanel, Grid, RowDefinition
        from System.Windows import Thickness, GridLength
        from Autodesk.Revit.DB import Transaction
        
        selected = [item for item in self.all_items if item.is_selected]
        
        if not selected:
            MessageBox.Show("No sheets selected", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Use list to store result
        result = [False, False, 1]  # [proceed, is_batch, count]
        
        # Create dialog
        dialog = Window()
        dialog.Title = "Duplicate Sheets"
        dialog.Width = 400
        dialog.Height = 280
        dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        main_grid = Grid()
        main_grid.Margin = Thickness(20)
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        
        # Info label
        info_label = Label()
        info_label.Content = "{} sheet(s) selected".format(len(selected))
        info_label.FontSize = 14
        info_label.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(info_label, 0)
        main_grid.Children.Add(info_label)
        
        # Batch checkbox
        batch_check = CheckBox()
        batch_check.Content = "Duplicate each sheet multiple times"
        batch_check.Margin = Thickness(0, 10, 0, 0)
        batch_check.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(batch_check, 1)
        main_grid.Children.Add(batch_check)
        
        # Count input
        count_label = Label()
        count_label.Content = "Number of copies per sheet:"
        count_label.IsEnabled = False
        Grid.SetRow(count_label, 2)
        main_grid.Children.Add(count_label)
        
        count_box = TextBox()
        count_box.Margin = Thickness(0, 25, 0, 0)
        count_box.Text = "5"
        count_box.IsEnabled = False
        Grid.SetRow(count_box, 2)
        main_grid.Children.Add(count_box)
        
        # Enable/disable count
        def on_batch_changed(s, e):
            is_checked = batch_check.IsChecked
            count_label.IsEnabled = is_checked
            count_box.IsEnabled = is_checked
        
        batch_check.Checked += on_batch_changed
        batch_check.Unchecked += on_batch_changed
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        Grid.SetRow(btn_panel, 3)
        
        def on_ok(s, e):
            result[0] = True
            result[1] = batch_check.IsChecked == True
            
            if result[1]:
                try:
                    result[2] = int(count_box.Text)
                    if result[2] < 1 or result[2] > 20:
                        MessageBox.Show("Please enter 1-20 copies", "Error",
                                      MessageBoxButton.OK, MessageBoxImage.Warning)
                        return
                except:
                    MessageBox.Show("Please enter a valid number", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Warning)
                    return
            
            dialog.Close()
        
        def on_cancel(s, e):
            dialog.Close()
        
        ok_btn = Button()
        ok_btn.Content = "Duplicate"
        ok_btn.Width = 90
        ok_btn.Height = 30
        ok_btn.Margin = Thickness(0, 0, 10, 0)
        ok_btn.Click += on_ok
        btn_panel.Children.Add(ok_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Height = 30
        cancel_btn.Click += on_cancel
        btn_panel.Children.Add(cancel_btn)
        
        main_grid.Children.Add(btn_panel)
        dialog.Content = main_grid
        
        # Show dialog
        dialog.ShowDialog()
        
        if not result[0]:
            return
        
        # Duplicate sheets
        t = Transaction(self.doc, "Duplicate Sheets")
        t.Start()
        
        try:
            success_count = 0
            
            if result[1]:  # Batch mode
                for sheet_model in selected:
                    for i in range(result[2]):
                        if self.revit_service.duplicate_sheet(sheet_model.element):
                            success_count += 1
            else:  # Normal mode
                for sheet_model in selected:
                    if self.revit_service.duplicate_sheet(sheet_model.element):
                        success_count += 1
            
            t.Commit()
            self.refresh_data()
            
            MessageBox.Show("Duplicated {} sheet(s)".format(success_count), "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_rename_click(self, sender, args):
        """Rename selected sheets with inline dialog"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, Window
        from System.Windows.Controls import Label, TextBox, Button, StackPanel, Grid, RowDefinition
        from System.Windows import Thickness, GridLength
        from Autodesk.Revit.DB import Transaction
        
        selected = [item for item in self.all_items if item.is_selected]
        
        if not selected:
            MessageBox.Show("No sheets selected", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Use list to store result
        result = [False, "", "", ""]  # [result, find_text, replace_text, prefix]
        
        # Create dialog
        dialog = Window()
        dialog.Title = "Rename Sheets"
        dialog.Width = 450
        dialog.Height = 250
        dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        main_grid = Grid()
        main_grid.Margin = Thickness(20)
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(60)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        
        # Find/Replace
        find_label = Label()
        find_label.Content = "Find:"
        Grid.SetRow(find_label, 0)
        main_grid.Children.Add(find_label)
        
        find_box = TextBox()
        find_box.Margin = Thickness(0, 25, 0, 0)
        Grid.SetRow(find_box, 0)
        main_grid.Children.Add(find_box)
        
        replace_label = Label()
        replace_label.Content = "Replace with:"
        Grid.SetRow(replace_label, 1)
        main_grid.Children.Add(replace_label)
        
        replace_box = TextBox()
        replace_box.Margin = Thickness(0, 25, 0, 0)
        Grid.SetRow(replace_box, 1)
        main_grid.Children.Add(replace_box)
        
        # Prefix
        prefix_label = Label()
        prefix_label.Content = "Add Prefix:"
        Grid.SetRow(prefix_label, 2)
        main_grid.Children.Add(prefix_label)
        
        prefix_box = TextBox()
        prefix_box.Margin = Thickness(0, 25, 0, 0)
        Grid.SetRow(prefix_box, 2)
        main_grid.Children.Add(prefix_box)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        Grid.SetRow(btn_panel, 3)
        
        def on_ok(s, e):
            result[0] = True
            result[1] = find_box.Text
            result[2] = replace_box.Text
            result[3] = prefix_box.Text
            dialog.Close()
        
        def on_cancel(s, e):
            dialog.Close()
        
        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 80
        ok_btn.Height = 30
        ok_btn.Margin = Thickness(0, 0, 10, 0)
        ok_btn.Click += on_ok
        btn_panel.Children.Add(ok_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Height = 30
        cancel_btn.Click += on_cancel
        btn_panel.Children.Add(cancel_btn)
        
        main_grid.Children.Add(btn_panel)
        dialog.Content = main_grid
        
        # Show dialog
        dialog.ShowDialog()
        
        if not result[0]:
            return
        
        # Apply rename
        t = Transaction(self.doc, "Rename Sheets")
        t.Start()
        
        try:
            count = 0
            for sheet_model in selected:
                new_number = sheet_model.sheet_number
                new_name = sheet_model.sheet_name
                
                # Find/Replace in number
                if result[1]:
                    new_number = new_number.replace(result[1], result[2])
                    new_name = new_name.replace(result[1], result[2])
                
                # Prefix
                if result[3]:
                    new_number = result[3] + new_number
                
                # Update
                sheet_model.sheet_number = new_number
                sheet_model.sheet_name = new_name
                sheet_model.check_if_modified()
                
                # Track change
                if sheet_model.is_modified:
                    self.change_tracker.track_modification(sheet_model)
                    count += 1
            
            t.Commit()
            self.data_grid.Items.Refresh()
            self.update_status()
            self.parent_window.update_status_with_counts()
            
            MessageBox.Show("Renamed {} sheet(s)\n\nClick Apply to save changes".format(count),
                          "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_delete_click(self, sender, args):
        """Delete selected sheets"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
        from Autodesk.Revit.DB import Transaction
        
        selected = [item for item in self.all_items if item.is_selected]
        
        if not selected:
            MessageBox.Show("No sheets selected", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        msg = "Delete {} sheet(s)?".format(len(selected))
        result = MessageBox.Show(msg, "Confirm Delete", MessageBoxButton.YesNo,
                                MessageBoxImage.Warning)
        
        if result != MessageBoxResult.Yes:
            return
        
        t = Transaction(self.doc, "Delete Sheets")
        t.Start()
        
        try:
            success_count = 0
            for sheet_model in selected:
                if self.revit_service.delete_sheet(sheet_model.id):
                    success_count += 1
            
            t.Commit()
            self.refresh_data()
            
            MessageBox.Show("Deleted {} sheet(s)".format(success_count), "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as e:
            t.RollBack()
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_select_none_click(self, sender, args):
        """Deselect all sheets"""
        for item in self.all_items:
            item.is_selected = False
        self.data_grid.Items.Refresh()
        self.update_status()
    
    def on_export_excel_click(self, sender, args):
        """Export sheets to CSV (no openpyxl needed!)"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
        import clr
        clr.AddReference('System.Windows.Forms')
        from System.Windows.Forms import SaveFileDialog, DialogResult
        import os
        import csv
        
        try:
            # Select file location
            dialog = SaveFileDialog()
            dialog.Filter = "CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx"
            dialog.FileName = "Sheet_List.csv"
            dialog.Title = "Export Sheet List"
            
            if dialog.ShowDialog() != DialogResult.OK:
                return
            
            filepath = dialog.FileName
            
            # Export to CSV (always works!)
            if filepath.lower().endswith('.csv') or not hasattr(self.parent_window, 'excel_service') or self.parent_window.excel_service is None:
                # Use CSV export
                try:
                    with open(filepath, 'w') as f:
                        writer = csv.writer(f)
                        
                        # Header
                        writer.writerow([
                            'Sheet Number',
                            'Sheet Name',
                            'Current Revision'
                        ])
                        
                        # Data
                        for sheet_model in self.all_items:
                            # Get current revision safely
                            try:
                                current_rev = getattr(sheet_model, 'current_revision', '-')
                            except:
                                current_rev = '-'
                            
                            writer.writerow([
                                sheet_model.sheet_number,
                                sheet_model.sheet_name,
                                current_rev
                            ])
                    
                    result = MessageBox.Show(
                        "Exported {} sheets successfully!\n\nFile: {}\n\nOpen file?".format(
                            len(self.all_items), filepath),
                        "Success",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information)
                    
                    if result == MessageBoxResult.Yes:
                        os.startfile(filepath)
                    
                except Exception as e:
                    MessageBox.Show("CSV Export Error: {}".format(str(e)), "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error)
                    import traceback
                    traceback.print_exc()
            else:
                # Use Excel service (if available)
                success = self.parent_window.excel_service.export_sheets_to_excel(
                    self.all_items, filepath)
                
                if success:
                    result = MessageBox.Show(
                        "Exported {} sheets successfully!\n\nOpen file?".format(len(self.all_items)),
                        "Success",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information)
                    
                    if result == MessageBoxResult.Yes:
                        os.startfile(filepath)
                else:
                    MessageBox.Show("Export failed", "Error",
                                  MessageBoxButton.OK, MessageBoxImage.Error)
                
        except Exception as e:
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def on_import_excel_click(self, sender, args):
        """Import sheets from Excel"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        
        try:
            # Check if services available
            if not hasattr(self.parent_window, 'excel_service') or self.parent_window.excel_service is None:
                MessageBox.Show(
                    "Excel service not available.\n\nPlease ensure openpyxl is installed:\npip install openpyxl --break-system-packages",
                    "Service Not Available",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning)
                return
            
            # Import dialog
            from import_excel_dialog import ImportExcelDialog
            
            dialog = ImportExcelDialog(
                self.parent_window.excel_service,
                self.parent_window.revit_service,
                self.doc
            )
            dialog.ShowDialog()
            
            # Refresh if import was successful
            if dialog.result:
                self.refresh_data()
                
        except Exception as e:
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def on_place_views_click(self, sender, args):
        """Place views on sheets"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        
        try:
            # Check if service available
            if not hasattr(self.parent_window, 'place_views_service') or self.parent_window.place_views_service is None:
                MessageBox.Show(
                    "Place Views service not available.",
                    "Service Not Available",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning)
                return
            
            # Get selected sheets
            selected_sheets = [item.element for item in self.all_items if item.is_selected]
            
            if not selected_sheets:
                MessageBox.Show(
                    "Please select at least one sheet first.",
                    "No Sheets Selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information)
                return
            
            # Show dialog
            from place_views_dialog import PlaceViewsDialog
            
            dialog = PlaceViewsDialog(
                self.parent_window.place_views_service,
                self.doc,
                selected_sheets
            )
            dialog.ShowDialog()
            
        except Exception as e:
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def on_parameters_click(self, sender, args):
        """Custom parameters management"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        
        try:
            # Check if service available
            if not hasattr(self.parent_window, 'params_service') or self.parent_window.params_service is None:
                MessageBox.Show(
                    "Custom Parameters service not available.",
                    "Service Not Available",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning)
                return
            
            # Get selected sheets
            selected_sheets = [item.element for item in self.all_items if item.is_selected]
            
            if not selected_sheets:
                MessageBox.Show(
                    "Please select at least one sheet first.",
                    "No Sheets Selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information)
                return
            
            # Show dialog
            from custom_parameters_dialog import CustomParametersDialog
            
            dialog = CustomParametersDialog(
                self.parent_window.params_service,
                self.doc,
                selected_sheets
            )
            dialog.ShowDialog()
            
        except Exception as e:
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def on_refresh_click(self, sender, args):
        """Refresh data grid to show updated parameter values"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        
        try:
            # Reload ALL sheet data from Revit
            print("DEBUG: Refreshing sheet data...")
            self.all_items = self.revit_service.get_all_sheets()
            print("DEBUG: Total sheets loaded: {}".format(len(self.all_items)))
            
            # Re-add custom parameter columns to new data
            custom_columns = []
            for col in self.data_grid.Columns:
                if hasattr(col, 'Header'):
                    header = str(col.Header)
                    # Skip default columns
                    if header not in ['', u'☐', u'✓', 'Sheet Number', 'Sheet Name', 'Current Revision']:
                        custom_columns.append(header)
            
            # Add parameter values for custom columns
            if custom_columns:
                print("DEBUG: Re-loading {} custom columns: {}".format(len(custom_columns), custom_columns))
                for item in self.all_items:
                    for param_name in custom_columns:
                        # Use same property naming as add_parameter_column
                        property_name = "param_" + param_name.replace(" ", "_")
                        param_value = self._get_parameter_value(item.element, param_name)
                        setattr(item, property_name, param_value)
                        print("DEBUG: Set {} for sheet {} = {}".format(property_name, item.sheet_number, param_value))
            
            # Update display with new items (don't call refresh_data which reloads again!)
            print("DEBUG: Updating display with refreshed data...")
            self.filtered_items = list(self.all_items)
            
            # Create NEW ObservableCollection for proper binding
            from System.Collections.ObjectModel import ObservableCollection
            new_source = ObservableCollection[object](self.filtered_items)
            self.data_grid.ItemsSource = new_source
            print("DEBUG: New ObservableCollection created!")
            
            # Update status
            self.update_status()
            
            MessageBox.Show("Data refreshed successfully!\n\n{} sheets loaded\n{} custom columns updated".format(
                len(self.all_items), len(custom_columns)), 
                "Success",
                MessageBoxButton.OK, MessageBoxImage.Information)
            
        except Exception as e:
            MessageBox.Show("Error refreshing: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def show_add_column_dialog(self):
        """Show dialog to select parameter to add as column"""
        from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, Thickness, GridLength
        from System.Windows.Controls import Grid, Label, ListBox, Button, StackPanel, RowDefinition, ScrollViewer
        import System
        
        try:
            # Get all sheet parameters
            if not self.all_items:
                MessageBox.Show("No sheets loaded", "Info",
                              MessageBoxButton.OK, MessageBoxImage.Information)
                return
            
            # Get first sheet to get parameters
            first_sheet = self.all_items[0].element
            params = []
            
            for param in first_sheet.Parameters:
                param_name = param.Definition.Name
                # Skip parameters already displayed
                if param_name not in ['Sheet Number', 'Sheet Name', 'Current Revision']:
                    params.append(param_name)
            
            if not params:
                MessageBox.Show("No additional parameters available", "Info",
                              MessageBoxButton.OK, MessageBoxImage.Information)
                return
            
            # Create dialog
            dialog = Window()
            dialog.Title = "Add Parameter Column"
            dialog.Width = 500
            dialog.Height = 600
            dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
            
            main_grid = Grid()
            main_grid.Margin = Thickness(20)
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
            main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
            
            # Info
            info = Label()
            info.Content = "Available Parameters:"
            info.FontSize = 14
            info.FontWeight = System.Windows.FontWeights.Bold
            Grid.SetRow(info, 0)
            main_grid.Children.Add(info)
            
            # ListBox with parameters
            scroll = ScrollViewer()
            scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
            Grid.SetRow(scroll, 1)
            
            param_list = ListBox()
            param_list.Margin = Thickness(0, 10, 0, 10)
            
            for param in sorted(params):
                param_list.Items.Add(param)
            
            scroll.Content = param_list
            main_grid.Children.Add(scroll)
            
            # Buttons
            btn_panel = StackPanel()
            btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
            btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            Grid.SetRow(btn_panel, 2)
            
            add_btn = Button()
            add_btn.Content = "Add Column"
            add_btn.Width = 120
            add_btn.Height = 35
            add_btn.Margin = Thickness(0, 0, 10, 0)
            
            # Add click handler
            def on_add_click(s, e):
                if param_list.SelectedItem:
                    selected_param = str(param_list.SelectedItem)
                    self.add_parameter_column(selected_param)
                    dialog.Close()
                else:
                    MessageBox.Show("Please select a parameter", "Info",
                                  MessageBoxButton.OK, MessageBoxImage.Information)
            
            add_btn.Click += on_add_click
            btn_panel.Children.Add(add_btn)
            
            cancel_btn = Button()
            cancel_btn.Content = "Cancel"
            cancel_btn.Width = 100
            cancel_btn.Height = 35
            cancel_btn.Click += lambda s, e: dialog.Close()
            btn_panel.Children.Add(cancel_btn)
            
            main_grid.Children.Add(btn_panel)
            dialog.Content = main_grid
            dialog.ShowDialog()
            
        except Exception as e:
            MessageBox.Show("Error: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def add_parameter_column(self, param_name):
        """Add a parameter as a new column"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        from System.Windows.Data import Binding
        from System.Windows.Controls import DataGridTextColumn
        
        try:
            # Check if column already exists
            for col in self.data_grid.Columns:
                if hasattr(col, 'Header') and col.Header == param_name:
                    MessageBox.Show("Column '{}' already exists!".format(param_name),
                                  "Info", MessageBoxButton.OK, MessageBoxImage.Information)
                    return
            
            # Create property name (replace spaces with underscores for WPF binding)
            property_name = "param_" + param_name.replace(" ", "_")
            
            print("DEBUG: Adding column '{}' with property '{}'".format(param_name, property_name))
            
            # Create new column
            from ui.components import DataGridHelper
            new_col = DataGridHelper.create_text_column(param_name, property_name, 150)
            
            # Add to grid
            self.data_grid.Columns.Add(new_col)
            
            # Update all items to have this property
            success_count = 0
            for item in self.all_items:
                param_value = self._get_parameter_value(item.element, param_name)
                setattr(item, property_name, param_value)
                if param_value != "-":
                    success_count += 1
                print("DEBUG: Set {} for sheet {} = '{}'".format(property_name, item.sheet_number, param_value))
            
            # Also update filtered items if different
            if self.filtered_items != self.all_items:
                for item in self.filtered_items:
                    param_value = self._get_parameter_value(item.element, param_name)
                    setattr(item, property_name, param_value)
            
            # CRITICAL FIX: Create NEW ObservableCollection to force complete rebinding!
            print("DEBUG: Creating new ObservableCollection for rebinding...")
            from System.Collections.ObjectModel import ObservableCollection
            new_source = ObservableCollection[object](self.filtered_items)
            self.data_grid.ItemsSource = new_source
            print("DEBUG: New ObservableCollection created and set!")
            
            # Force grid refresh
            self.data_grid.Items.Refresh()
            print("DEBUG: Grid items refreshed!")
            
            MessageBox.Show(
                "Column '{}' added successfully!\n\n{} sheets with values\n{} sheets empty".format(
                    param_name, success_count, len(self.all_items) - success_count),
                "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            
        except Exception as e:
            MessageBox.Show("Error adding column: {}".format(str(e)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
            import traceback
            traceback.print_exc()
    
    def _get_parameter_value(self, element, param_name):
        """Get parameter value from element"""
        from Autodesk.Revit.DB import StorageType
        
        try:
            param = element.LookupParameter(param_name)
            if not param:
                print("DEBUG: Parameter '{}' not found on element".format(param_name))
                return "-"
            
            if not param.HasValue:
                print("DEBUG: Parameter '{}' has no value".format(param_name))
                return "-"
            
            if param.StorageType == StorageType.String:
                value = param.AsString()
                print("DEBUG: Got String value '{}' for '{}'".format(value, param_name))
                return value or "-"
            elif param.StorageType == StorageType.Integer:
                value = str(param.AsInteger())
                print("DEBUG: Got Integer value '{}' for '{}'".format(value, param_name))
                return value
            elif param.StorageType == StorageType.Double:
                value = str(param.AsDouble())
                print("DEBUG: Got Double value '{}' for '{}'".format(value, param_name))
                return value
            elif param.StorageType == StorageType.ElementId:
                value = str(param.AsElementId().IntegerValue)
                print("DEBUG: Got ElementId value '{}' for '{}'".format(value, param_name))
                return value
            else:
                print("DEBUG: Unsupported StorageType {} for '{}'".format(param.StorageType, param_name))
                return "-"
        except Exception as e:
            print("ERROR: Exception getting parameter '{}': {}".format(param_name, str(e)))
            import traceback
            traceback.print_exc()
            return "-"
    
    def remove_custom_columns(self):
        """Remove all custom columns (keep only default 5)"""
        from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
        
        # Keep only first 5 columns (checkbox, status, number, name, revision)
        while self.data_grid.Columns.Count > 5:
            self.data_grid.Columns.RemoveAt(5)