# -*- coding: utf-8 -*-
"""ManaAnno — Unified annotation and text note manager.

Consolidates:
  - Dimensions (Audit & manage dimension types/instances)
  - Text Notes (Audit & search text note contents)
  - Tag Checker (Search & delete orphan tags)
  - DimText (Manage dimension text overrides)
  - Utilities (Renumber along spline, Copy annotations, Upper all)

Author: T3Lab
"""
__title__ = "Mana\nAnno"
__author__ = "T3Lab"

import os
import sys

# Add lib directory to system path
# __file__ is T3Lab.extension/T3Lab.tab/Annotation.panel/ManaAnno.pushbutton/script.py
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Import and show the dialog
import GUI.AnnotationManagerDialog as AnnotationManagerDialog

if __name__ == '__main__':
    AnnotationManagerDialog.show_dialog()
