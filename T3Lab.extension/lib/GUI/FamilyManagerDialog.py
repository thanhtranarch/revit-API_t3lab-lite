# -*- coding: utf-8 -*-
"""Family Manager — event handling for the Family Manager launcher window."""

import os

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'FamilyManager.xaml')


class FamilyManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_family_management.Click += self._open_family_management
        self.btn_load_family.Click += self._open_load_family

        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome
        self.PreviewKeyDown += self._on_key_down

    def _open_family_management(self, sender, e):
        self.Close()
        try:
            from GUI.FamilyManagementDialog import show_family_management
            show_family_management()
        except Exception as ex:
            forms.alert("Error opening Family Management:\n{}".format(ex))

    def _open_load_family(self, sender, e):
        self.Close()
        try:
            from GUI.FamilyLoaderDialog import show_family_loader
            show_family_loader()
        except Exception as ex:
            forms.alert("Error opening Family Loader:\n{}".format(ex))

    def _minimize(self, sender, e):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def _maximize(self, sender, e):
        import System.Windows
        if self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
        else:
            self.WindowState = System.Windows.WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()

    def _on_key_down(self, sender, e):
        import System.Windows.Input as WI
        if e.Key == WI.Key.Escape:
            self.Close()


def show_family_manager(script_dir, revit):
    FamilyManagerWindow(script_dir, revit).ShowDialog()
