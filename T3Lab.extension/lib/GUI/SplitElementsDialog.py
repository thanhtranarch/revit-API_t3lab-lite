# -*- coding: utf-8 -*-
"""Split Elements — event handling for the Split Elements launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'SplitElements.xaml')


class SplitElementsWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_split_walls.Click += self._on_split_walls
        self.btn_split_columns.Click += self._on_split_columns
        self.btn_split_floors.Click += self._on_split_floors

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

    def _on_split_walls(self, sender, e):
        self._launch("../Split.pulldown/Wall_Split.pushbutton/script.py")

    def _on_split_columns(self, sender, e):
        self._launch("../Split.pulldown/Column_Split.pushbutton/script.py")

    def _on_split_floors(self, sender, e):
        self._launch("../Split.pulldown/Floor_Split.pushbutton/script.py")

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


def show_split_elements(script_dir, revit):
    SplitElementsWindow(script_dir, revit).ShowDialog()
