# -*- coding: utf-8 -*-
"""Text to Element — transfer text note content to element parameters via
bounding-box intersection in the active view.
"""

__title__ = "Text to\nElement"
__author__ = "Tran Tien Thanh"
__doc__ = (
    "Transfer text note content to element parameters via bounding-box "
    "intersection. Select a target category and parameter, then run "
    "'Find Intersections' to preview matches before writing."
)

import sys
import os

_lib = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
)
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.TextToElementDialog import show_text_to_element

show_text_to_element(__revit__)
