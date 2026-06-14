# -*- coding: utf-8 -*-
"""
Sheet Manager - UI Components
Reusable UI component factories

Copyright © Dang Quoc Truong (DQT)
"""

import System
from System.Windows import Thickness
from System.Windows.Controls import Button, ComboBox, TextBox, DataGrid, DataGridTextColumn, DataGridCheckBoxColumn
from System.Windows.Controls import Grid, ColumnDefinition, Label
from System.Windows import GridLength, GridUnitType
from System.Windows.Media import SolidColorBrush, Color, Brushes
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength, DataGridLengthUnitType

from core.config import Config


class ButtonFactory(object):
    """Factory for creating buttons"""
    
    @staticmethod
    def create_toolbar_button(text, width=100):
        """Create a toolbar button"""
        btn = Button()
        btn.Content = text
        btn.Width = width
        btn.Height = 30
        btn.Margin = Thickness(0, 0, 5, 0)
        
        argb = Config.get_color_argb(Config.PRIMARY_COLOR)
        btn.Background = SolidColorBrush(Color.FromArgb(*argb))
        btn.Foreground = Brushes.White
        btn.FontSize = 12
        
        return btn


class ComboBoxFactory(object):
    """Factory for creating combo boxes"""
    
    @staticmethod
    def create_filter_combobox(width=150):
        """Create a filter combo box"""
        combo = ComboBox()
        combo.Width = width
        combo.Height = 30
        combo.Margin = Thickness(0, 0, 10, 0)
        combo.FontSize = 12
        return combo


class SearchBoxFactory(object):
    """Factory for creating search text boxes"""
    
    @staticmethod
    def create_search_textbox(width=200):
        """Create a search text box"""
        textbox = TextBox()
        textbox.Width = width
        textbox.Height = 30
        textbox.FontSize = 12
        textbox.Margin = Thickness(5, 0, 0, 0)
        return textbox


class DataGridHelper(object):
    """Helper for DataGrid operations"""
    
    @staticmethod
    def create_datagrid():
        """Create a standard data grid"""
        grid = DataGrid()
        grid.AutoGenerateColumns = False
        grid.CanUserAddRows = False
        grid.CanUserDeleteRows = False
        grid.CanUserReorderColumns = False
        grid.CanUserSortColumns = True
        grid.SelectionMode = System.Windows.Controls.DataGridSelectionMode.Extended
        grid.GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.All
        grid.HeadersVisibility = System.Windows.Controls.DataGridHeadersVisibility.Column
        grid.RowHeight = 28
        
        # Background
        argb = Config.get_color_argb(Config.BACKGROUND_COLOR)
        grid.Background = SolidColorBrush(Color.FromArgb(*argb))
        
        return grid
    
    @staticmethod
    def create_checkbox_column(header, binding_path, width=60):
        """Create a checkbox column"""
        column = DataGridCheckBoxColumn()
        column.Header = header
        column.Binding = Binding(binding_path)
        column.Width = DataGridLength(width)
        column.IsReadOnly = False
        return column
    
    @staticmethod
    def create_status_column(header, binding_path, width=40):
        """Create a status column"""
        column = DataGridTextColumn()
        column.Header = header
        column.Binding = Binding(binding_path)
        column.Width = DataGridLength(width)
        column.IsReadOnly = True
        return column
    
    @staticmethod
    def create_text_column(header, binding_path, width=200, is_readonly=True):
        """Create a text column"""
        from System.Windows.Data import BindingMode
        
        column = DataGridTextColumn()
        column.Header = header
        
        # Create binding with proper mode
        binding = Binding(binding_path)
        if not is_readonly:
            binding.Mode = BindingMode.TwoWay
            binding.UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
        
        column.Binding = binding
        column.Width = DataGridLength(width)
        column.IsReadOnly = is_readonly
        return column
    
    @staticmethod
    def create_editable_text_column(header, binding_path, width=200):
        """Create an editable text column"""
        return DataGridHelper.create_text_column(header, binding_path, width, is_readonly=False)


class ToolbarBuilder(object):
    """Builder for toolbar grids"""
    
    @staticmethod
    def create_toolbar_grid(rows=2):
        """Create a toolbar grid"""
        toolbar = Grid()
        toolbar.Margin = Thickness(10)
        
        # Create rows
        for i in range(rows):
            toolbar.RowDefinitions.Add(System.Windows.Controls.RowDefinition(
                Height=GridLength(40)))
        
        # Create columns (will be added as needed)
        for i in range(15):  # 15 columns for flexibility
            toolbar.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Auto)))
        
        return toolbar
    
    @staticmethod
    def add_button_to_toolbar(toolbar, button, column, row=0):
        """Add button to toolbar grid"""
        Grid.SetColumn(button, column)
        Grid.SetRow(button, row)
        toolbar.Children.Add(button)
    
    @staticmethod
    def add_combobox_to_toolbar(toolbar, combobox, column, row=0):
        """Add combobox to toolbar grid"""
        Grid.SetColumn(combobox, column)
        Grid.SetRow(combobox, row)
        toolbar.Children.Add(combobox)


class StatusBarBuilder(object):
    """Builder for status bars"""
    
    @staticmethod
    def create_status_text():
        """Create a status text block"""
        from System.Windows.Controls import TextBlock
        
        status = TextBlock()
        status.Text = "Ready"
        status.FontSize = 12
        status.Margin = Thickness(10, 5, 10, 5)
        status.VerticalAlignment = System.Windows.VerticalAlignment.Center
        
        return status