# -*- coding: utf-8 -*-
"""Sheet Hub — event handling for the Sheet Hub launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'SheetHub.xaml')


class SheetHubWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_sheet_manager.Click += self._on_sheet_manager
        self.btn_sheet_renumber.Click += self._on_sheet_renumber

    def _launch(self, rel_path):
        script_path = os.path.normpath(os.path.join(self._script_dir, rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': self._revit}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _on_sheet_manager(self, sender, e):
        self._launch("../SheetManager.pushbutton/script.py")

    def _on_sheet_renumber(self, sender, e):
        self._launch("../Sheet re-number.pushbutton/script.py")


def show_sheet_hub(script_dir, revit):
    SheetHubWindow(script_dir, revit).ShowDialog()
