#! python3
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

import os, sys
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
import os

_lib = os.path.normpath(
    os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
)
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.TextToElementDialog import show_text_to_element

show_text_to_element(__revit__)
