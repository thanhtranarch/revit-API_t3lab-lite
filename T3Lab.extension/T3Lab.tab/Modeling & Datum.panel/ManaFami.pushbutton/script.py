# -*- coding: utf-8 -*-
__title__ = "Family\nManager"
__author__ = "Tran Tien Thanh & Dang Quoc Truong"
__doc__ = "Family Manager — Manage families and load new families in one dialog."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.ManaFamiDialog import show_family_manager

if __name__ == '__main__':
    show_family_manager(os.path.dirname(__file__), __revit__)
