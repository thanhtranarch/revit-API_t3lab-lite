#! python3
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
# ─── CPython 3 & lib bootstrap ────────────────────────────────────────────────
for _env in ('APPDATA', 'PROGRAMDATA'):
    _base = os.environ.get(_env, '')
    if _base:
        for _clone in ('pyRevit-Master', 'pyRevit'):
            _ceng = os.path.join(_base, _clone, 'bin', 'cengines', 'CPY3123')
            if os.path.isdir(_ceng):
                for _d in (_ceng, os.path.join(_ceng, 'Lib')):
                    if hasattr(os, 'add_dll_directory'):
                        try:
                            os.add_dll_directory(_d)
                        except Exception:
                            pass
                for _p in (_ceng, os.path.join(_ceng, 'Lib'), os.path.join(_ceng, 'python312.zip')):
                    if os.path.exists(_p) and _p not in sys.path:
                        sys.path.insert(0, _p)

_cur = os.path.dirname(os.path.abspath(__file__))
while _cur and not os.path.exists(os.path.join(_cur, 'lib')):
    _parent = os.path.dirname(_cur)
    if _parent == _cur:
        break
    _cur = _parent
_lib_dir = os.path.join(_cur, 'lib')
if os.path.exists(_lib_dir) and _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass
# ──────────────────────────────────────────────────────────────────────────────

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
