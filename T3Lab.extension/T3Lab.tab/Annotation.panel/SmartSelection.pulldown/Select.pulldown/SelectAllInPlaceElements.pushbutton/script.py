# -*- coding: utf-8 -*-
__title__   = "Select\nIn-Place"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select All In-Place Elements

Select every In-Place family instance visible in the active view.
Optionally extend the search to the whole model.
___________________________________________________________
How-to:
- Run the tool from any view
- All In-Place families in the active view are selected
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

from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance

from dqt_select.compat import to_element_id_list, notify

# ----------------------------------------------------------------- VARIABLES
doc = __revit__.ActiveUIDocument.Document          # noqa: F821
uidoc = __revit__.ActiveUIDocument                 # noqa: F821


# ----------------------------------------------------------------- FUNCTIONS
def get_in_place_instances(scope_view_id=None):
    """Return all In-Place FamilyInstances, optionally scoped to a view."""
    if scope_view_id is not None:
        collector = FilteredElementCollector(doc, scope_view_id)
    else:
        collector = FilteredElementCollector(doc)

    instances = collector.OfClass(FamilyInstance) \
                         .WhereElementIsNotElementType().ToElements()

    result = []
    for fi in instances:
        try:
            if fi.Symbol and fi.Symbol.Family and fi.Symbol.Family.IsInPlace:
                result.append(fi)
        except Exception:
            continue
    return result


# ----------------------------------------------------------------- MAIN
if __name__ == '__main__':
    scope = forms.CommandSwitchWindow.show(
        ['Active View', 'Whole Model'],
        message='Where do you want to select In-Place elements?',
    )
    if not scope:
        forms.alert('Cancelled.', title='DQT - Select In-Place', exitscript=True)

    view_id = doc.ActiveView.Id if scope == 'Active View' else None
    in_place = get_in_place_instances(view_id)

    if not in_place:
        forms.alert('No In-Place elements found in {}.'.format(scope.lower()),
                    title='DQT - Select In-Place', exitscript=True)

    ids = [e.Id for e in in_place]
    uidoc.Selection.SetElementIds(to_element_id_list(ids))

    notify('Selected {} In-Place element(s) in {}.'.format(len(ids), scope.lower()),
                title='DQT - Select In-Place')
