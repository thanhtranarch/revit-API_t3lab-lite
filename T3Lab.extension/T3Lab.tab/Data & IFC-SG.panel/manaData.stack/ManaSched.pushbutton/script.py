# -*- coding: utf-8 -*-
__title__ = "Schedule\nManager"
__author__ = "Tran Tien Thanh"
__doc__ = "Schedule Manager — unified Export/Import Excel and Schedule Duplication tool."

import sys
import os

_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
_lib = os.path.join(_ext_dir, 'lib')
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.ManaSchedDialog import show_schedule_manager

show_schedule_manager(os.path.dirname(__file__), __revit__)
