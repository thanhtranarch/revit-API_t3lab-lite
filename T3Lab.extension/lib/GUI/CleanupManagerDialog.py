# -*- coding: utf-8 -*-
"""Cleanup Manager — event handling for the Cleanup Manager launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'CleanupManager.xaml')


class CleanupManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_smart_purge.Click += self._on_smart_purge
        self.btn_advanced_purge.Click += self._on_advanced_purge
        self.btn_smart_delete.Click += self._on_smart_delete

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

    def _on_smart_purge(self, sender, e):
        self._launch("../SmartPurge.pushbutton/script.py")

    def _on_advanced_purge(self, sender, e):
        self._launch("../AdvancedPurge.pushbutton/script.py")

    def _on_smart_delete(self, sender, e):
        self._launch("../SmartDelete.pushbutton/script.py")

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


def show_cleanup_manager(script_dir, revit):
    CleanupManagerWindow(script_dir, revit).ShowDialog()
