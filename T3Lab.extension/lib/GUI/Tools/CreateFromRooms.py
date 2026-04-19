# -*- coding: utf-8 -*-
"""
Create From Rooms

Create elements based on room boundaries and properties.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Create From Rooms"

import os, clr

from pyrevit import forms
from GUI.forms         import my_WPF
from Snippets._convert import convert_internal_units

clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")

from System import Uri, UriKind
from System.Collections.Generic import List
from System.Windows import WindowState
from System.Windows.Media.Imaging import BitmapImage
import wpf

PATH_SCRIPT = os.path.dirname(__file__)


class ListItem:
    def __init__(self, Name='Unnamed', element=None, checked=False):
        self.Name      = Name
        self.IsChecked = checked
        self.element   = element


class CreateFromRooms(my_WPF):
    selected_type = []
    offset        = 0

    def __init__(self, items, title='__title', label="Select Type:",
                 button_name='Create', version='version= 1.0'):
        self.items       = items
        self.title       = title
        self.label       = label
        self.button_name = button_name
        self.version     = version

        super(CreateFromRooms, self).add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'CreateFromRooms.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        self.update_UI()
        self._load_logo()
        self.ShowDialog()

    def _load_logo(self):
        try:
            logo_path = os.path.join(PATH_SCRIPT, '..', 'T3Lab_logo.png')
            logo_path = os.path.normpath(logo_path)
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.logo_image.Source = bitmap
        except Exception:
            pass

    def update_UI(self):
        self.main_title.Text          = self.title
        self.text_label.Content       = self.label
        self.button_main.Content      = self.button_name
        self.footer_version.Text      = self.version

        self.items                    = self.generate_list_items()
        self.main_ListBox.ItemsSource = self.items

    def generate_list_items(self):
        list_of_items = List[type(ListItem())]()
        first = True
        for type_name, typ in sorted(self.items.items()):
            checked = True if first else False
            first   = False
            list_of_items.Add(ListItem(type_name, typ, checked))
        return list_of_items

    def text_filter_updated(self, sender, e):
        filtered_list_of_items = List[type(ListItem())]()
        filter_keyword = self.textbox_filter.Text

        if not filter_keyword:
            self.main_ListBox.ItemsSource = self.items
            return

        for item in self.items:
            if filter_keyword.lower() in item.Name.lower():
                filtered_list_of_items.Add(item)
            else:
                if item.IsChecked:
                    item.IsChecked = False

        self.main_ListBox.ItemsSource = filtered_list_of_items

    def UIe_ItemChecked(self, sender, e):
        filtered_list_of_items = List[type(ListItem())]()
        for item in self.main_ListBox.Items:
            item.IsChecked = item.Name == sender.Content.Text
            filtered_list_of_items.Add(item)
        self.main_ListBox.ItemsSource = filtered_list_of_items

    def NumberValidationTextBox(self, sender, e):
        import re
        regex = re.compile('[^0-9.]+')
        e.Handled = regex.match(e.Text)

    def button_close(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def button_run(self, sender, e):
        self.textbox_filter.Text = ''
        self.Close()

        try:
            selected_items = [item.element for item in self.main_ListBox.ItemsSource if item.IsChecked]
            self.offset    = convert_internal_units(float(self.UI_offset.Text),
                                                    get_internal=True,
                                                    units='cm')
            self.selected_type = selected_items[0]
        except Exception:
            self.offset = 0
