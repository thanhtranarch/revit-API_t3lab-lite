# -*- coding: utf-8 -*-
__title__ = "IFC-SG\nAssign"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "IFC-SG Assign — Auto or Manual assignment of IFC export classes."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.IFCSGAssignDialog import show_ifcsg_assign

if __name__ == '__main__':
    show_ifcsg_assign(os.path.dirname(__file__), __revit__)
