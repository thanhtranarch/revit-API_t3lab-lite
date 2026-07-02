# -*- coding: utf-8 -*-
__title__ = "IFC-SG\nSuite"
__author__ = "T3Lab"
__doc__ = "Unified IFC-SG manager: Assign subtypes and verify parameter compliance."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.IFCSGDialog import show_ifcsg_suite

if __name__ == '__main__':
    show_ifcsg_suite(os.path.dirname(__file__), __revit__)
