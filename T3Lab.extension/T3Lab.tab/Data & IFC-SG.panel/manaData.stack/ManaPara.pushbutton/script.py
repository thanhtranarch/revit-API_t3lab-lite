# -*- coding: utf-8 -*-
__title__ = "Parameter\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Parameter Manager — Transfer, Text-to-Element, and Values-to-Region tools."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.ManaParaDialog import show_parameter_manager

if __name__ == '__main__':
    show_parameter_manager(os.path.dirname(__file__), __revit__)
