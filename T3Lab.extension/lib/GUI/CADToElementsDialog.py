# -*- coding: utf-8 -*-
"""CAD to Elements — event handling for the CAD to Elements launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'CADToElements.xaml')


class CADToElementsWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_cad_to_wall.Click += self._on_cad_to_wall
        self.btn_cad_to_floor.Click += self._on_cad_to_floor
        self.btn_cad_to_beam.Click += self._on_cad_to_beam

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

    def _on_cad_to_wall(self, sender, e):
        self._launch("../CadtoWall.pushbutton/script.py")

    def _on_cad_to_floor(self, sender, e):
        self._launch("../CadtoFloor.pushbutton/script.py")

    def _on_cad_to_beam(self, sender, e):
        self._launch("../Beam.pushbutton/script.py")

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


def show_cad_to_elements(script_dir, revit):
    CADToElementsWindow(script_dir, revit).ShowDialog()
