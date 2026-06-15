# -*- coding: utf-8 -*-
__title__ = "Sheet\nHub"
__author__ = "Dang Quoc Truong & Tran Tien Thanh"
__doc__ = "Sheet Hub — Manage Sheets and Re-number Sheets in one dialog."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.SheetHubDialog import show_sheet_hub

if __name__ == '__main__':
    show_sheet_hub(os.path.dirname(__file__), __revit__)
