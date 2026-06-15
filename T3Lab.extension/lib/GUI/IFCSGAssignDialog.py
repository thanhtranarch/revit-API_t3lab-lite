# -*- coding: utf-8 -*-
"""IFC-SG Assign — event handling for the IFC-SG Assignment launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'IFCSGAssign.xaml')


class IFCSGAssignWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_auto_assign.Click += self._on_auto_assign
        self.btn_manual_assign.Click += self._on_manual_assign

        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome

    def _launch(self, rel_path):
        script_path = os.path.normpath(os.path.join(self._script_dir, rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': self._revit}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _on_auto_assign(self, sender, e):
        self._launch("../Auto Assign.pushbutton/script.py")

    def _on_manual_assign(self, sender, e):
        self._launch("../Manual Assign.pushbutton/script.py")

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


def show_ifcsg_assign(script_dir, revit):
    IFCSGAssignWindow(script_dir, revit).ShowDialog()
