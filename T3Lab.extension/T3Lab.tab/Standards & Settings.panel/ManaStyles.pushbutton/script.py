# -*- coding: utf-8 -*-
__title__ = "Visual &amp;\nStyles"
__author__ = "T3Lab"
__doc__ = "Visual & Style Manager — Manage fill patterns, line styles, line patterns, color splasher overrides, and coordinate locations."

import os
import sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.ManaStylesDialog import show_visual_settings

if __name__ == '__main__':
    show_visual_settings(os.path.dirname(__file__), __revit__)
