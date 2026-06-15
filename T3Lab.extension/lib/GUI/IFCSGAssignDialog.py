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


def show_ifcsg_assign(script_dir, revit):
    IFCSGAssignWindow(script_dir, revit).ShowDialog()
