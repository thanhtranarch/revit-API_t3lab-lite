# -*- coding: utf-8 -*-
"""
Find Replace Dialog
GUI dialog for find and replace operations on Revit elements.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "Find Replace Dialog"

import os
from pyrevit import forms
from Autodesk.Revit.DB import (Transaction, View, ViewPlan, ViewSection,
                                View3D, ViewSchedule, ViewDrafting)
from Autodesk.Revit.Exceptions import ArgumentException

import clr
clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Uri, UriKind
from System.Diagnostics.Process import Start
from System.Windows import WindowState
from System.Windows.Media.Imaging import BitmapImage
import wpf

from GUI.forms import my_WPF

PATH_SCRIPT = os.path.dirname(__file__)


class FindReplace(my_WPF):
    """GUI for [Views: Find and Replace]"""
    run = False

    def __init__(self, title, label="Find and Replace", button_name="Rename"):
        path_xaml_file = os.path.join(PATH_SCRIPT, 'FindReplace.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        self.UI_label.Content       = label
        self.UI_main_button.Content = button_name
        self.main_title.Text        = title

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

    def find_replace(self, name):
        return self.prefix + str(name).replace(self.find, self.replace) + self.suffix

    @property
    def find(self):
        return self.input_find.Text

    @property
    def replace(self):
        return self.input_replace.Text

    @property
    def prefix(self):
        return self.input_prefix.Text

    @property
    def suffix(self):
        return self.input_suffix.Text

    def button_close(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def Hyperlink_RequestNavigate(self, sender, e):
        Start(e.Uri.AbsoluteUri)

    def button_run(self, sender, e):
        self.Close()
