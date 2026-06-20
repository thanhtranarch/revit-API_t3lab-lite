# -*- coding: utf-8 -*-
"""ManaAlign — Unified alignment and dimensioning manager.

Consolidates:
  - Smart Align (Graphical align & distribute)
  - Auto Dimension (Automated dimension chains)
  - Snap Dimension (Snapping elements/dimensions to grids)
  - Reset Graphic Overrides (Title bar action)

Author: T3Lab
"""
__title__ = "Mana\nAlign"
__author__ = "T3Lab"
__persistentengine__ = True

import os
import sys

# Add lib directory to system path
# __file__ is T3Lab.extension/T3Lab.tab/Annotation.panel/ManaAlign.pushbutton/script.py
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Import and show the dialog
import GUI.ManaAlignDialog as ManaAlignDialog

if __name__ == '__main__':
    ManaAlignDialog.show_dialog()
