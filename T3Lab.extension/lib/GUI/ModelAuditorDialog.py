# -*- coding: utf-8 -*-
"""Model Auditor — event handling for the Model Auditor launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ModelAuditor.xaml')


class ModelAuditorWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_health_check.Click += self._on_health_check
        self.btn_model_checker.Click += self._on_model_checker
        self.btn_warning_manager.Click += self._on_warning_manager
        self.btn_inplace_model.Click += self._on_inplace_model
        self.btn_material_list.Click += self._on_material_list

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

    def _on_health_check(self, sender, e):
        self._launch("../HealthCheck/script.py")

    def _on_model_checker(self, sender, e):
        self._launch("../ModelChecker/script.py")

    def _on_warning_manager(self, sender, e):
        self._launch("../Warning/script.py")

    def _on_inplace_model(self, sender, e):
        self._launch("../InPlaceModel/script.py")

    def _on_material_list(self, sender, e):
        self._launch("../MaterialList/script.py")

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


def show_model_auditor(script_dir, revit):
    ModelAuditorWindow(script_dir, revit).ShowDialog()
