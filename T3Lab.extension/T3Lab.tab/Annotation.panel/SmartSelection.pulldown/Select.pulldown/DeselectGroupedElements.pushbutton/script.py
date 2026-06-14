# -*- coding: utf-8 -*-
__title__   = "Deselect\nGrouped"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Deselect Grouped Elements

Remove every element that is part of a Group from the current selection.
Handy when you want to edit only the loose / ungrouped elements.
___________________________________________________________
How-to:
- Select a bunch of elements (some inside groups, some not)
- Run the tool
- Elements that belong to a group are removed from the selection
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

# ----------------------------------------------------------------- IMPORTS
from pyrevit import forms

from dqt_select.compat import to_element_id_list, eid_int, notify, INVALID_ELEMENT_ID_INT

# ----------------------------------------------------------------- VARIABLES
doc = __revit__.ActiveUIDocument.Document          # noqa: F821
uidoc = __revit__.ActiveUIDocument                 # noqa: F821

# ----------------------------------------------------------------- MAIN
if __name__ == '__main__':
    selected_ids = uidoc.Selection.GetElementIds()

    if not selected_ids or selected_ids.Count == 0:
        forms.alert('Nothing is selected. Select some elements first.',
                    title='DQT - Deselect Grouped', exitscript=True)

    kept = []
    removed = 0
    for eid in selected_ids:
        elem = doc.GetElement(eid)
        if elem is None:
            continue
        # GroupId equals InvalidElementId (-1) when the element is NOT in a group
        if eid_int(elem.GroupId) == INVALID_ELEMENT_ID_INT:
            kept.append(elem.Id)
        else:
            removed += 1

    uidoc.Selection.SetElementIds(to_element_id_list(kept))

    notify('Removed {} grouped element(s). {} kept.'.format(removed, len(kept)),
                title='DQT - Deselect Grouped')
