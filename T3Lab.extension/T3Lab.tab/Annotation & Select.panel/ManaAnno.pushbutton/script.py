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
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Evict stale cached modules so code changes take effect without a full pyRevit reload
_stale = [k for k in list(sys.modules.keys())
          if k in ('GUI.ManaAnnoDialog', 'GUI.DimTextDialog',
                   'GUI.TagCheckerDialog', 'ManaAnnoDialog',
                   'DimTextDialog', 'TagCheckerDialog')]
for _k in _stale:
    del sys.modules[_k]

import GUI.ManaAnnoDialog as ManaAnnoDialog

if __name__ == '__main__':
    ManaAnnoDialog.show_dialog()
