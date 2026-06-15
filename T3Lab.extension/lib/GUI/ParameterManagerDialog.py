# -*- coding: utf-8 -*-
"""Parameter Manager — event handling for the Parameter Manager launcher window."""

import os
import __builtin__

from pyrevit import forms

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ParameterManager.xaml')


class ParameterManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit

        self.btn_transfer_params.Click += self._open_transfer_params
        self.btn_text_to_element.Click += self._open_text_to_element
        self.btn_values_to_region.Click += self._open_values_to_region

    def _launch(self, rel_path):
        script_path = os.path.normpath(os.path.join(self._script_dir, rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': self._revit}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _open_transfer_params(self, sender, e):
        self._launch("../Transfer Para.pushbutton/script.py")

    def _open_text_to_element(self, sender, e):
        self._launch("../Text to element.pushbutton/script.py")

    def _open_values_to_region(self, sender, e):
        self._launch("../Values to Filled Region .pushbutton/script.py")


def show_parameter_manager(script_dir, revit):
    ParameterManagerWindow(script_dir, revit).ShowDialog()
