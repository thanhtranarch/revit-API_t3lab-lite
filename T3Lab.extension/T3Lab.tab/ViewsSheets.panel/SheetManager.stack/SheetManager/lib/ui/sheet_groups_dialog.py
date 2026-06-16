# -*- coding: utf-8 -*-
"""
Sheet Manager - Sheet Groups Dialog
Manage custom sheet groups

Copyright © Dang Quoc Truong (DQT)
"""

from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage, Thickness, GridLength
from System.Windows.Controls import (Grid, Label, Button, TextBox, ListBox, CheckBox,
                                      StackPanel, RowDefinition, ColumnDefinition)
import System


class SheetGroupsDialog(Window):
    """Dialog for managing custom sheet groups"""
    
    def __init__(self, groups_service, doc, all_sheets=None):
        self.groups_service = groups_service
        self.doc = doc
        self.current_group = None
        self.all_sheets = all_sheets or []
        
        self._build_ui()
        self._load_data()
    
    def _build_ui(self):
        """Build dialog UI"""
        self.Title = "Sheet Groups Management"
        self.Width = 800
        self.Height = 600
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        # Main grid
        main_grid = Grid()
        main_grid.Margin = Thickness(20)
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        main_grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        main_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(300)))
        main_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(20)))
        main_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, System.Windows.GridUnitType.Star)))
        
        # Left panel - Groups list
        left_panel = self._create_left_panel()
        Grid.SetRow(left_panel, 0)
        Grid.SetColumn(left_panel, 0)
        main_grid.Children.Add(left_panel)
        
        # Right panel - Sheets in group
        right_panel = self._create_right_panel()
        Grid.SetRow(right_panel, 0)
        Grid.SetColumn(right_panel, 2)
        main_grid.Children.Add(right_panel)
        
        # Close button
        close_btn = Button()
        close_btn.Content = "Close"
        close_btn.Width = 100
        close_btn.Height = 35
        close_btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        close_btn.Margin = Thickness(0, 10, 0, 0)
        close_btn.Click += lambda s, e: self.Close()
        Grid.SetRow(close_btn, 1)
        Grid.SetColumn(close_btn, 2)
        main_grid.Children.Add(close_btn)
        
        self.Content = main_grid
    
    def _create_left_panel(self):
        """Create left panel with groups list"""
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        
        # Label
        label = Label()
        label.Content = "Sheet Groups:"
        label.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(label, 0)
        grid.Children.Add(label)
        
        # ListBox
        self.groups_listbox = ListBox()
        self.groups_listbox.SelectionChanged += self.on_group_selected
        Grid.SetRow(self.groups_listbox, 1)
        grid.Children.Add(self.groups_listbox)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        Grid.SetRow(btn_panel, 2)
        
        new_btn = Button()
        new_btn.Content = "New"
        new_btn.Width = 70
        new_btn.Height = 30
        new_btn.Margin = Thickness(5, 0, 5, 0)
        new_btn.Click += self.on_new_group_click
        btn_panel.Children.Add(new_btn)
        
        rename_btn = Button()
        rename_btn.Content = "Rename"
        rename_btn.Width = 70
        rename_btn.Height = 30
        rename_btn.Margin = Thickness(5, 0, 5, 0)
        rename_btn.Click += self.on_rename_group_click
        btn_panel.Children.Add(rename_btn)
        
        delete_btn = Button()
        delete_btn.Content = "Delete"
        delete_btn.Width = 70
        delete_btn.Height = 30
        delete_btn.Margin = Thickness(5, 0, 5, 0)
        delete_btn.Click += self.on_delete_group_click
        btn_panel.Children.Add(delete_btn)
        
        grid.Children.Add(btn_panel)
        
        return grid
    
    def _create_right_panel(self):
        """Create right panel with sheets"""
        grid = Grid()
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(1, System.Windows.GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))
        
        # Label
        label = Label()
        label.Content = "Sheets in Group:"
        label.FontWeight = System.Windows.FontWeights.Bold
        Grid.SetRow(label, 0)
        grid.Children.Add(label)
        
        # ScrollViewer for checkboxes
        from System.Windows.Controls import ScrollViewer
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        Grid.SetRow(scroll, 1)
        
        self.sheets_panel = StackPanel()
        scroll.Content = self.sheets_panel
        grid.Children.Add(scroll)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        Grid.SetRow(btn_panel, 2)
        
        select_all_btn = Button()
        select_all_btn.Content = "Select All"
        select_all_btn.Width = 90
        select_all_btn.Height = 30
        select_all_btn.Margin = Thickness(5, 0, 5, 0)
        select_all_btn.Click += self.on_select_all_click
        btn_panel.Children.Add(select_all_btn)
        
        select_none_btn = Button()
        select_none_btn.Content = "Select None"
        select_none_btn.Width = 90
        select_none_btn.Height = 30
        select_none_btn.Margin = Thickness(5, 0, 5, 0)
        select_none_btn.Click += self.on_select_none_click
        btn_panel.Children.Add(select_none_btn)
        
        apply_btn = Button()
        apply_btn.Content = "Apply"
        apply_btn.Width = 90
        apply_btn.Height = 30
        apply_btn.Margin = Thickness(5, 0, 5, 0)
        apply_btn.Background = System.Windows.Media.Brushes.LightGreen
        apply_btn.Click += self.on_apply_sheets_click
        btn_panel.Children.Add(apply_btn)
        
        grid.Children.Add(btn_panel)
        
        return grid
    
    def _load_data(self):
        """Load groups"""
        self.refresh_groups()
    
    def refresh_groups(self):
        """Refresh groups list"""
        groups = self.groups_service.get_all_groups()
        
        self.groups_listbox.Items.Clear()
        for group_name in groups:
            self.groups_listbox.Items.Add(group_name)
    
    def on_group_selected(self, sender, args):
        """Group selected"""
        if self.groups_listbox.SelectedIndex < 0:
            return
        
        self.current_group = self.groups_listbox.SelectedItem
        self.load_sheets_in_group()
    
    def load_sheets_in_group(self):
        """Load sheets in current group"""
        if not self.current_group:
            return
        
        sheets_in_group = self.groups_service.get_sheets_in_group(self.current_group)
        
        # Clear and rebuild checkboxes
        self.sheets_panel.Children.Clear()
        
        for sheet_info in self.all_sheets:
            cb = CheckBox()
            cb.Content = "{} - {}".format(sheet_info['number'], sheet_info['name'])
            cb.Tag = sheet_info['id']
            cb.IsChecked = sheet_info['id'] in sheets_in_group
            cb.Margin = Thickness(5, 2, 5, 2)
            self.sheets_panel.Children.Add(cb)
    
    def on_new_group_click(self, sender, args):
        """Create new group"""
        group_name = "Group {}".format(self.groups_listbox.Items.Count + 1)
        
        success = self.groups_service.create_group(group_name)
        
        if success:
            self.refresh_groups()
            
            # Select new group
            for i in range(self.groups_listbox.Items.Count):
                if self.groups_listbox.Items[i] == group_name:
                    self.groups_listbox.SelectedIndex = i
                    break
            
            MessageBox.Show("Sheet group created: {}".format(group_name), "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        else:
            MessageBox.Show("Failed to create group", "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_rename_group_click(self, sender, args):
        """Rename selected group"""
        if not self.current_group:
            MessageBox.Show("Please select a group to rename", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Create input dialog
        input_window = Window()
        input_window.Title = "Rename Sheet Group"
        input_window.Width = 400
        input_window.Height = 150
        input_window.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        
        grid = Grid()
        grid.Margin = Thickness(20)
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        
        label = Label()
        label.Content = "New name:"
        Grid.SetRow(label, 0)
        grid.Children.Add(label)
        
        textbox = TextBox()
        textbox.Text = self.current_group
        textbox.Padding = Thickness(5)
        Grid.SetRow(textbox, 1)
        grid.Children.Add(textbox)
        
        btn_panel = StackPanel()
        btn_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        btn_panel.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        Grid.SetRow(btn_panel, 2)
        
        result_holder = [False]
        
        def on_ok(s, e):
            result_holder[0] = True
            input_window.Close()
        
        def on_cancel(s, e):
            result_holder[0] = False
            input_window.Close()
        
        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 80
        ok_btn.Height = 30
        ok_btn.Margin = Thickness(5, 0, 5, 0)
        ok_btn.Click += on_ok
        btn_panel.Children.Add(ok_btn)
        
        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 80
        cancel_btn.Height = 30
        cancel_btn.Click += on_cancel
        btn_panel.Children.Add(cancel_btn)
        
        grid.Children.Add(btn_panel)
        input_window.Content = grid
        
        input_window.ShowDialog()
        
        if not result_holder[0]:
            return
        
        new_name = textbox.Text.strip()
        if not new_name or new_name == self.current_group:
            return
        
        success = self.groups_service.rename_group(self.current_group, new_name)
        
        if success:
            old_name = self.current_group
            self.current_group = new_name
            self.refresh_groups()
            
            # Reselect
            for i in range(self.groups_listbox.Items.Count):
                if self.groups_listbox.Items[i] == new_name:
                    self.groups_listbox.SelectedIndex = i
                    break
            
            MessageBox.Show("Group renamed to: {}".format(new_name), "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        else:
            MessageBox.Show("Failed to rename group", "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_delete_group_click(self, sender, args):
        """Delete selected group"""
        if not self.current_group:
            return
        
        result = MessageBox.Show("Delete this group?\n\n{}".format(self.current_group), "Confirm",
                               MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != System.Windows.MessageBoxResult.Yes:
            return
        
        success = self.groups_service.delete_group(self.current_group)
        
        if success:
            self.current_group = None
            self.refresh_groups()
            self.sheets_panel.Children.Clear()
            
            MessageBox.Show("Group deleted", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
        else:
            MessageBox.Show("Failed to delete group", "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def on_select_all_click(self, sender, args):
        """Select all sheets"""
        for child in self.sheets_panel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = True
    
    def on_select_none_click(self, sender, args):
        """Deselect all sheets"""
        for child in self.sheets_panel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = False
    
    def on_apply_sheets_click(self, sender, args):
        """Apply sheet selection to group"""
        if not self.current_group:
            return
        
        # Get selected sheets
        selected_ids = []
        for child in self.sheets_panel.Children:
            if isinstance(child, CheckBox) and child.IsChecked:
                selected_ids.append(child.Tag)
        
        success = self.groups_service.set_sheets_in_group(self.current_group, selected_ids)
        
        if success:
            # Refresh to show updated state
            self.load_sheets_in_group()
            
            MessageBox.Show("Group updated: {} sheets".format(len(selected_ids)),
                          "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        else:
            MessageBox.Show("Failed to update group", "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
