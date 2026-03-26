# -*- coding: utf-8 -*-
"""
Forms

Utility functions and form helpers for GUI dialogs.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Forms"

from WPF_Base import my_WPF
from FindReplace import FindReplace
from SelectFromDict import select_from_dict


class ListItem:
    """Helper Class for displaying selected sheets in my custom GUI."""
    def __init__(self,  Name='Unnamed', element = None, checked = False):
        self.Name       = Name
        self.IsChecked  = checked
        self.element    = element