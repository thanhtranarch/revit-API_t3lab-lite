# -*- coding: utf-8 -*-
__title__ = "CAD to\nElements"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "CAD to Elements — Convert CAD linework into Walls, Floors, or Beams."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.CADToElementsDialog import show_cad_to_elements

if __name__ == '__main__':
    show_cad_to_elements(os.path.dirname(__file__), __revit__)
