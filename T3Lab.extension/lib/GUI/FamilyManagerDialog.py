# -*- coding: utf-8 -*-
"""Family Manager — event handling for the Family Manager launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'FamilyManager.xaml')


class FamilyManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_family_management.Click += self._open_family_management
        self.btn_load_family.Click += self._open_load_family

    def _launch(self, rel_path):
        script_path = os.path.normpath(os.path.join(self._script_dir, rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': self._revit}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _open_family_management(self, sender, e):
        self._launch("../Family Work 2.stack/Family Management.pushbutton/script.py")

    def _open_load_family(self, sender, e):
        self._launch("../Family Work 2.stack/Load Family.pushbutton/script.py")


def show_family_manager(script_dir, revit):
    FamilyManagerWindow(script_dir, revit).ShowDialog()
