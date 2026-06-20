# -*- coding: utf-8 -*-
__title__ = "Split\nElements"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Split Elements — Split Walls, Columns, or Floors at selected levels."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.SplitElementsDialog import show_split_elements

if __name__ == '__main__':
    show_split_elements(os.path.dirname(__file__), __revit__)
