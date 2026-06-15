# -*- coding: utf-8 -*-
"""Datum Manager — event handling for the Datum Manager launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'DatumManager.xaml')


class DatumManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_save_grids.Click += self._on_save_grids
        self.btn_restore_grids.Click += self._on_restore_grids
        self.btn_restore_all_grids.Click += self._on_restore_all_grids
        self.btn_align_gridlines.Click += self._on_align_gridlines
        self.btn_convert_grid.Click += self._on_convert_grid
        self.btn_align_levels.Click += self._on_align_levels
        self.btn_convert_level.Click += self._on_convert_level

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

    def _on_save_grids(self, sender, e):
        self._launch("../../../Annotation.panel/Graphic 2.stack/Grids.pulldown/Save Grids.pushbutton/script.py")

    def _on_restore_grids(self, sender, e):
        self._launch("../../../Annotation.panel/Graphic 2.stack/Grids.pulldown/Restore Grids.pushbutton/script.py")

    def _on_restore_all_grids(self, sender, e):
        self._launch("../../../Annotation.panel/Graphic 2.stack/Grids.pulldown/Restore All Grids.pushbutton/script.py")

    def _on_align_gridlines(self, sender, e):
        self._launch("../Datum.pulldown/Gridline.pulldown/Align Gridline.pushbutton/script.py")

    def _on_convert_grid(self, sender, e):
        self._launch("../Datum.pulldown/Gridline.pulldown/ConvertGridline.pushbutton/script.py")

    def _on_align_levels(self, sender, e):
        self._launch("../Datum.pulldown/Level.pulldown/Align Level.pushbutton/script.py")

    def _on_convert_level(self, sender, e):
        self._launch("../Datum.pulldown/Level.pulldown/ConvertLevel.pushbutton/script.py")

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


def show_datum_manager(script_dir, revit):
    DatumManagerWindow(script_dir, revit).ShowDialog()
