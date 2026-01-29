# -*- coding: utf-8 -*-
"""
Preview Window - Simplified for Smart Purge v2
Shows items to be purged before deletion

Copyright (c) 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')

import System
from System.Windows import Window, WindowStartupLocation, Thickness
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows import FontWeights
from System.Windows.Controls import (
    Grid, StackPanel, Border, TextBlock, Button, 
    DataGrid, DataGridTextColumn, ScrollViewer
)
from System.Windows.Media import SolidColorBrush, Color
from System.Collections.ObjectModel import ObservableCollection

from config_advanced import Colors


class PreviewWindow(Window):
    """Preview window for items to purge"""
    
    def __init__(self, items):
        """
        Initialize preview window
        
        Args:
            items: List of item dictionaries to preview
        """
        self.items = items
        
        # Setup window
        self.Title = "Preview - Items to Purge"
        self.Width = 800
        self.Height = 600
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        
        # Build UI
        self.build_ui()
    
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
        main.RowDefinitions.Add(self.row_def(60))      # Header
        main.RowDefinitions.Add(self.row_def(1, star=True))  # List
        main.RowDefinitions.Add(self.row_def(60))      # Actions
        
        # Panels
        main.Children.Add(self.header_panel(0))
        main.Children.Add(self.list_panel(1))
        main.Children.Add(self.action_panel(2))
        
        self.Content = main
    
    def row_def(self, height, star=False):
        """Create row definition"""
        from System.Windows.Controls import RowDefinition
        from System.Windows import GridLength, GridUnitType
        rd = RowDefinition()
        if star:
            rd.Height = GridLength(height, GridUnitType.Star)
        else:
            rd.Height = GridLength(height)
        return rd
    
    def header_panel(self, row):
        """Create header"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color(Colors.HEADER))
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 0, 0, 1)
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        
        title = TextBlock()
        title.Text = u"\U0001F50D Preview - Items to Purge"
        title.FontSize = 18
        title.FontWeight = FontWeights.Bold
        stack.Children.Add(title)
        
        count = TextBlock()
        count.Text = "Total: {} items".format(len(self.items))
        count.FontSize = 12
        count.Foreground = SolidColorBrush(self.parse_color("#FF666666"))
        count.Margin = Thickness(0, 5, 0, 0)
        stack.Children.Add(count)
        
        border.Child = stack
        return border
    
    def list_panel(self, row):
        """Create item list grouped by category"""
        border = Border()
        border.Padding = Thickness(10)
        Grid.SetRow(border, row)
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        
        main_stack = StackPanel()
        
        # Group items by category
        from collections import defaultdict
        grouped = defaultdict(list)
        for item in self.items:
            if isinstance(item, dict):
                category = item.get('category', 'Unknown')
            else:
                if hasattr(item, 'element') and item.element and hasattr(item.element, 'Category'):
                    category = item.element.Category.Name if item.element.Category else 'No Category'
                else:
                    category = 'Unknown'
            grouped[category].append(item)
        
        # Create section for each category
        for category_name in sorted(grouped.keys()):
            items_in_cat = grouped[category_name]
            
            # Category header
            header_border = Border()
            header_border.Background = SolidColorBrush(self.parse_color("#FFE3F2FD"))
            header_border.Padding = Thickness(10, 5, 10, 5)
            header_border.Margin = Thickness(0, 10, 0, 0)
            
            header_text = TextBlock()
            header_text.Text = u"{} ({} items)".format(category_name, len(items_in_cat))
            header_text.FontWeight = FontWeights.Bold
            header_text.FontSize = 12
            header_border.Child = header_text
            main_stack.Children.Add(header_border)
            
            # DataGrid for this category
            grid = DataGrid()
            grid.AutoGenerateColumns = False
            grid.IsReadOnly = True
            grid.HeadersVisibility = System.Windows.Controls.DataGridHeadersVisibility.Column
            grid.AlternatingRowBackground = SolidColorBrush(self.parse_color("#FFF9F9F9"))
            grid.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.Horizontal
            grid.MaxHeight = 300  # Limit height per category
            
            # Columns
            col_name = DataGridTextColumn()
            col_name.Header = "Name"
            col_name.Binding = System.Windows.Data.Binding("name")
            col_name.Width = System.Windows.Controls.DataGridLength(400)
            
            col_type = DataGridTextColumn()
            col_type.Header = "Type"
            col_type.Binding = System.Windows.Data.Binding("type")
            col_type.Width = System.Windows.Controls.DataGridLength(150)
            
            col_id = DataGridTextColumn()
            col_id.Header = "ID"
            col_id.Binding = System.Windows.Data.Binding("id")
            col_id.Width = System.Windows.Controls.DataGridLength(100)
            
            grid.Columns.Add(col_name)
            grid.Columns.Add(col_type)
            grid.Columns.Add(col_id)
            
            # Populate data for this category
            collection = ObservableCollection[object]()
            for item in items_in_cat:
                # Handle both dictionary and PurgeCategoryItem object formats
                if isinstance(item, dict):
                    name = item.get('name', 'Unknown')
                    item_type = item.get('type', 'Unknown')
                    item_id = str(item.get('id', ''))
                else:
                    name = getattr(item, 'name', 'Unknown')
                    item_type = getattr(item, 'item_type', 'Unknown')
                    item_id = str(getattr(item, 'element_id', ''))
                
                # Create row object for WPF binding
                class ItemRow:
                    def __init__(self, name, item_type, item_id):
                        self.name = name
                        self.type = item_type
                        self.id = item_id
                
                row = ItemRow(name, item_type, item_id)
                collection.Add(row)
            
            grid.ItemsSource = collection
            main_stack.Children.Add(grid)
        
        scroll.Content = main_stack
        border.Child = scroll
        return border
    
    def action_panel(self, row):
        """Create action buttons"""
        border = Border()
        border.Background = SolidColorBrush(self.parse_color("#FFF5F5F5"))
        border.BorderBrush = SolidColorBrush(self.parse_color(Colors.BORDER))
        border.BorderThickness = Thickness(0, 1, 0, 0)
        border.Padding = Thickness(15, 10, 15, 10)
        Grid.SetRow(border, row)
        
        stack = StackPanel()
        stack.Orientation = System.Windows.Controls.Orientation.Horizontal
        stack.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        
        # Info text
        info = TextBlock()
        info.Text = "Review items carefully before purging"
        info.VerticalAlignment = System.Windows.VerticalAlignment.Center
        info.Margin = Thickness(0, 0, 20, 0)
        info.Foreground = SolidColorBrush(self.parse_color("#FF666666"))
        stack.Children.Add(info)
        
        # Export button
        btn_export = Button()
        btn_export.Content = u"\U0001F4BE Export to Excel"
        btn_export.Width = 130
        btn_export.Height = 35
        btn_export.Margin = Thickness(0, 0, 5, 0)
        btn_export.Background = SolidColorBrush(self.parse_color("#FF4CAF50"))
        btn_export.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_export.Click += self.on_export_click
        stack.Children.Add(btn_export)
        
        # Close button
        btn_close = Button()
        btn_close.Content = u"\u2714 OK"
        btn_close.Width = 100
        btn_close.Height = 35
        btn_close.Background = SolidColorBrush(self.parse_color("#FF2196F3"))
        btn_close.Foreground = SolidColorBrush(self.parse_color(Colors.White))
        btn_close.Click += lambda s, e: self.Close()
        stack.Children.Add(btn_close)
        
        border.Child = stack
        return border
    
    def on_export_click(self, sender, e):
        """Export items to Excel CSV"""
        try:
            from System.Windows.Forms import SaveFileDialog, DialogResult
            
            # Show save dialog
            dialog = SaveFileDialog()
            dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            dialog.DefaultExt = "csv"
            dialog.FileName = "PurgePreview_{}.csv".format(
                System.DateTime.Now.ToString("yyyyMMdd_HHmmss")
            )
            
            if dialog.ShowDialog() == DialogResult.OK:
                import codecs
                
                # Open file with UTF-8 encoding
                with codecs.open(dialog.FileName, 'w', encoding='utf-8-sig') as f:
                    # Write header
                    f.write("Category,Name,Type,ID\n")
                    
                    # Group items by category
                    from collections import defaultdict
                    grouped = defaultdict(list)
                    for item in self.items:
                        if isinstance(item, dict):
                            category = item.get('category', 'Unknown')
                        else:
                            if hasattr(item, 'element') and item.element and hasattr(item.element, 'Category'):
                                category = item.element.Category.Name if item.element.Category else 'No Category'
                            else:
                                category = 'Unknown'
                        grouped[category].append(item)
                    
                    # Write data grouped by category
                    for category_name in sorted(grouped.keys()):
                        items_in_cat = grouped[category_name]
                        
                        for item in items_in_cat:
                            # Get values
                            if isinstance(item, dict):
                                name = item.get('name', 'Unknown')
                                item_type = item.get('type', 'Unknown')
                                item_id = str(item.get('id', ''))
                            else:
                                name = getattr(item, 'name', 'Unknown')
                                item_type = getattr(item, 'item_type', 'Unknown')
                                item_id = str(getattr(item, 'element_id', ''))
                            
                            # Escape quotes and commas
                            name = '"{}"'.format(name.replace('"', '""'))
                            item_type = '"{}"'.format(item_type.replace('"', '""'))
                            category_name_escaped = '"{}"'.format(category_name.replace('"', '""'))
                            
                            # Write row
                            f.write("{},{},{},{}\n".format(
                                category_name_escaped, name, item_type, item_id
                            ))
                
                MessageBox.Show(
                    "Export successful!\n\nFile: {}".format(dialog.FileName),
                    "Export Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                )
        
        except Exception as ex:
            MessageBox.Show(
                "Export failed:\n\n{}".format(str(ex)),
                "Export Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            )