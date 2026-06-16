# -*- coding: utf-8 -*-
"""View Hub — event handling for the View Hub launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ViewHub.xaml')


class ViewHubWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_view_manager.Click += self._on_view_manager
        self.btn_view_templates.Click += self._on_view_templates
        self.btn_room_plans.Click += self._on_room_plans

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

    def _on_view_manager(self, sender, e):
        self._launch("../ViewManager/script.py")

    def _on_view_templates(self, sender, e):
        self._launch("../ViewTemplate/script.py")

    def _on_room_plans(self, sender, e):
        self._launch("../Create Room Plan/script.py")

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


def show_view_hub(script_dir, revit):
    ViewHubWindow(script_dir, revit).ShowDialog()
