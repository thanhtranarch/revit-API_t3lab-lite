# -*- coding: utf-8 -*-
"""
Forms
=====
Utility functions and form helpers for GUI dialogs.
Universal compatibility for CPython 3 and IronPython.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Forms"

from GUI.WPF_Base import my_WPF, T3WPFWindow, WPFWindow
from GUI.FindReplace import FindReplace
from GUI.SelectFromDict import select_from_dict
from GUI.T3Dialog import T3Dialog, show_info, show_warning, show_error, confirm


class ListItem(object):
    """Helper Class for displaying selected sheets in my custom GUI."""
    def __init__(self, Name='Unnamed', element=None, checked=False):
        self.Name       = Name
        self.IsChecked  = checked
        self.element    = element