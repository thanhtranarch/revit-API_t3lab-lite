# -*- coding: utf-8 -*-
__title__ = "Smart\nAlign"
__author__ = "T3Lab"
__doc__ = "Smart Align — Graphical alignment and distribution hub."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.SmartAlignDialog import show_smart_align

if __name__ == '__main__':
    show_smart_align()
