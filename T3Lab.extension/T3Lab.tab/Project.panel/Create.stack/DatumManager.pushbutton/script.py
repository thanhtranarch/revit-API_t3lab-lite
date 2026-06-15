# -*- coding: utf-8 -*-
__title__ = "Datum\nManager"
__author__ = "Tran Tien Thanh & Dang Quoc Truong"
__doc__ = "Datum Manager — Manage Grids and Levels in one unified tool. Save/restore grid positions, align extents, and convert 2D/3D."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.DatumManagerDialog import show_datum_manager

if __name__ == '__main__':
    show_datum_manager(os.path.dirname(__file__), __revit__)
