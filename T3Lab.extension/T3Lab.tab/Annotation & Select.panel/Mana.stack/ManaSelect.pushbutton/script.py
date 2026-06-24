# -*- coding: utf-8 -*-
"""ManaSelect — Unified smart selection manager.

Consolidates:
  - Quick Select (Query by parameters/text)
  - Select Similar (Match type/family/category)
  - Select on Sheets (Title blocks & CAD imports)
  - Sidebar Quick Filters (Linked, In-place, Category, Grouped, Material)

Author: T3Lab
"""
__title__ = "Mana\nSelect"
__author__ = "T3Lab"

import os
import sys

# Add lib directory to system path
# __file__ is T3Lab.extension/T3Lab.tab/Annotation & Select.panel/Mana.stack/ManaSelect.pushbutton/script.py
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Import and show the dialog
import GUI.ManaSelectDialog as ManaSelectDialog

if __name__ == '__main__':
    ManaSelectDialog.show_dialog()
