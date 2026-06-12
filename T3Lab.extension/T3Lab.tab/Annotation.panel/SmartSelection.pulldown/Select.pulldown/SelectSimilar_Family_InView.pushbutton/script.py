# -*- coding: utf-8 -*-
__title__   = "Select Similar\nFamily (View)"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select Similar: Family (in View)

Select all instances in the active view that belong to the same Family
as the currently selected element(s). Useful when a family has several
types and you want every type at once.
___________________________________________________________
How-to:
- Select one or more instances in the active view
- Run the tool
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

from dqt_select.core import select_similar_family

if __name__ == '__main__':
    select_similar_family(mode='view')
