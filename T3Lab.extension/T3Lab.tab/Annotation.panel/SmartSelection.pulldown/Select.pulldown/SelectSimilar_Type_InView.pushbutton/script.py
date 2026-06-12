# -*- coding: utf-8 -*-
__title__   = "Select Similar\nType (View)"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select Similar: Type (in View)

Improved version of Revit's built-in 'Select All Instances: Visible in View'.

You can seed the selection with MULTIPLE elements at once and it also
handles unusual elements that the built-in tool ignores
(Lines, Grids, Levels, Rooms, Areas, Scope Boxes, Reference Planes...).
___________________________________________________________
How-to:
- Select one or more instances in the active view
- Run the tool
- All instances of the same Type (visible in the active view) get selected
___________________________________________________________
Dang Quoc Truong - DQT (c) 2026
"""

# ----------------------------------------------------------------- LIB PATH
import os as _os, sys as _sys
_here = _os.path.dirname(__file__)
for _up in range(6):
    _libp = _os.path.join(_here, 'lib')
    if _os.path.isdir(_os.path.join(_libp, 'dqt_select')) and _libp not in _sys.path:
        _sys.path.append(_libp)
        break
    _here = _os.path.dirname(_here)
# -----------------------------------------------------------------------------

from dqt_select.core import select_similar_type

if __name__ == '__main__':
    select_similar_type(mode='view')
