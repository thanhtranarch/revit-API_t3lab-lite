# -*- coding: utf-8 -*-
"""
Select From Dict Dialog
GUI dialog for selecting items from a dictionary.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "Select From Dict Dialog"

import os

from pyrevit import forms

from GUI.forms import my_WPF

import clr
clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")

from System import Uri, UriKind
from System.Collections.Generic import List
from System.Windows import Visibility, WindowState
from System.Windows.Media.Imaging import BitmapImage
import wpf

PATH_SCRIPT = os.path.dirname(__file__)


class ListItem:
    def __init__(self, Name='Unnamed', element=None, checked=False):
        self.Name      = Name
        self.IsChecked = checked
        self.element   = element


class SelectFromDict(my_WPF):
    def __init__(self, items,
                 title='__title',
                 label="Select Elements:",
                 button_name='Select',
                 version='version= 1.0',
                 SelectMultiple=True):
        self.SelectMultiple = SelectMultiple
        self.given_dict_items = {k: v for k, v in items.items() if k}
        self.items = self.generate_list_items()
        self.selected_items = []

        path_xaml_file = os.path.join(PATH_SCRIPT, 'Tools', 'SelectFromDict.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        self.main_title.Text     = title
        self.text_label.Content  = label
        self.button_main.Content = button_name
        self.footer_version.Text = version

        if not SelectMultiple:
            self.UI_Buttons_all_none.Visibility = Visibility.Collapsed

        self.main_ListBox.ItemsSource = self.items
        self.ShowDialog()

    def generate_list_items(self):
        return [ListItem(Name=name, element=element)
                for name, element in self.given_dict_items.items()]

    def button_select_all(self, sender, e):
        for item in self.items:
            item.IsChecked = True
        self.main_ListBox.Items.Refresh()

    def button_select_none(self, sender, e):
        for item in self.items:
            item.IsChecked = False
        self.main_ListBox.Items.Refresh()

    def UIe_ItemChecked(self, sender, e):
        if self.SelectMultiple:
            return
        checked_item = sender.DataContext
        for item in self.items:
            if item is not checked_item:
                item.IsChecked = False
        self.main_ListBox.Items.Refresh()

    def text_filter_updated(self, sender, e):
        filter_text = self.textbox_filter.Text.strip().lower() if self.textbox_filter.Text else ''
        if filter_text:
            self.main_ListBox.ItemsSource = [
                item for item in self.items if filter_text in item.Name.lower()
            ]
        else:
            self.main_ListBox.ItemsSource = self.items

    def button_select(self, sender, e):
        self.selected_items = [item.element for item in self.items if item.IsChecked]
        self.Close()

def select_from_dict(elements_dict,
                     title='__title__',
                     label="Select Elements:",
                     button_name='Select',
                     version='Version: 1.0',
                     SelectMultiple=True):
    if isinstance(elements_dict, list):
        elements_dict = {i: i for i in elements_dict}

    GUI_select = SelectFromDict(
        items=elements_dict, title=title, label=label,
        button_name=button_name, version=version,
        SelectMultiple=SelectMultiple
    )
    return GUI_select.selected_items
