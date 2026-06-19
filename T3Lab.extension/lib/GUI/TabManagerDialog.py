# -*- coding: utf-8 -*-
"""Tab Manager Dialog class."""

import os
from pyrevit import forms
from System.Windows import WindowState
from System.Collections.ObjectModel import ObservableCollection

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'TabManager.xaml')

class TabItemModel(object):
    def __init__(self, name, is_checked):
        self.Name = name
        self.IsChecked = is_checked

class TabManagerWindow(forms.WPFWindow):
    def __init__(self, current_lst):
        # WPFWindow.__init__ loads XAML
        forms.WPFWindow.__init__(self, _XAML)
        
        self.all_items = [TabItemModel(item.item, item.state) for item in current_lst]
        self.filtered_items = ObservableCollection[object]()
        
        self.BtnApply.Click += self._on_apply
        self.BtnClose.Click += self._on_close
        
        self.applied = False
        self.selected_names = []

        # Load initial items
        self._filter_list("")

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()

    def SearchBox_TextChanged(self, sender, e):
        search_text = self.SearchBox.Text.strip().lower()
        self._filter_list(search_text)

    def _filter_list(self, search_text):
        self.filtered_items.Clear()
        for item in self.all_items:
            if not search_text or search_text in item.Name.lower():
                self.filtered_items.Add(item)
        self.TabListBox.ItemsSource = self.filtered_items

    def _on_apply(self, sender, e):
        # Collect checked items
        self.selected_names = [item.Name for item in self.all_items if item.IsChecked]
        self.applied = True
        self.Close()

    def _on_close(self, sender, e):
        self.Close()

def show_tab_manager_dialog(current_lst):
    """Show the Tab Manager dialog."""
    dlg = TabManagerWindow(current_lst)
    dlg.ShowDialog()
    if dlg.applied:
        return dlg.selected_names
    return None
