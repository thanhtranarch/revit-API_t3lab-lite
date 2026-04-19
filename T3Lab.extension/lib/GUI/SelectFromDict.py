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

        self.add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'SelectFromDict.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        self.main_title.Text     = title
        self.text_label.Content  = label
        self.button_main.Content = button_name
        self.footer_version.Text = version

        if not SelectMultiple:
            self.UI_Buttons_all_none.Visibility = Visibility.Collapsed

        self.main_ListBox.ItemsSource = self.items
        self._load_logo()
        self.ShowDialog()

    def _load_logo(self):
        try:
            logo_path = os.path.join(PATH_SCRIPT, 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.logo_image.Source = bitmap
        except Exception:
            pass

    def __iter__(self):
        return iter(self.selected_items)

    def generate_list_items(self):
        list_of_items = List[type(ListItem())]()
        for type_name, floor_type in sorted(self.given_dict_items.items()):
            list_of_items.Add(ListItem(type_name, floor_type))
        return list_of_items

    def add_wpf_resource(self):
        super(SelectFromDict, self).add_wpf_resource()

    # Event handlers
    def text_filter_updated(self, sender, e):
        filtered = List[type(ListItem())]()
        kw = self.textbox_filter.Text
        if not kw:
            self.main_ListBox.ItemsSource = self.items
            return
        for item in self.items:
            if kw.lower() in item.Name.lower():
                filtered.Add(item)
        self.main_ListBox.ItemsSource = filtered

    def UIe_ItemChecked(self, sender, e):
        if not self.SelectMultiple:
            filtered = List[type(ListItem())]()
            for item in self.main_ListBox.Items:
                item.IsChecked = item.Name == sender.Content.Text
                filtered.Add(item)
            self.main_ListBox.ItemsSource = filtered

    def select_mode(self, mode):
        items = List[type(ListItem())]()
        checked = mode == 'all'
        for item in self.main_ListBox.ItemsSource:
            item.IsChecked = checked
            items.Add(item)
        self.main_ListBox.ItemsSource = items

    def button_select_all(self, sender, e):
        self.select_mode('all')

    def button_select_none(self, sender, e):
        self.select_mode('none')

    def button_select(self, sender, e):
        self.textbox_filter.Text = ''
        self.Close()
        selected = []
        for item in self.main_ListBox.ItemsSource:
            if item.IsChecked:
                selected.append(item.element)
        self.selected_items = selected

    def button_close(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized


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
    return list(GUI_select)
