# -*- coding: utf-8 -*-
"""Auto Dimension — automatically create dimension chains for selected elements.

Author: T3Lab
"""
__title__ = "Auto\nDimension"
__author__ = "T3Lab"
__persistentengine__ = True

import os
import sys

# Add lib directory to system path
# __file__ is T3Lab.extension/T3Lab.tab/Annotation & Select.panel/Mana.stack/ManaAlign.pushbutton/script.py
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Import and show the dialog
import GUI.AutoDimensionDialog as AutoDimensionDialog

if __name__ == '__main__':
    AutoDimensionDialog.show_dialog()
