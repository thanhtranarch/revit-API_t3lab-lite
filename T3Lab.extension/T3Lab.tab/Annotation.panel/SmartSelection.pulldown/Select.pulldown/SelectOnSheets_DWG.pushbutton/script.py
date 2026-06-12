# -*- coding: utf-8 -*-
__title__   = "On Sheets:\nDWGs"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select on Sheets: DWGs

Select the imported / linked DWG (ImportInstance) elements that live on
the selected sheets.
___________________________________________________________
How-to:
- Select sheets in the Project Browser, OR
- Run with nothing selected and pick sheets from the list
- The DWGs placed on those sheets get selected
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

from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet,
                               ImportInstance)

from dqt_select.compat import to_element_id_list, eid_int, notify

# ----------------------------------------------------------------- VARIABLES
doc = __revit__.ActiveUIDocument.Document          # noqa: F821
uidoc = __revit__.ActiveUIDocument                 # noqa: F821


# ----------------------------------------------------------------- HELPERS
def get_target_sheets():
    """Sheets from current selection, or ask the user to pick some."""
    sel_ids = uidoc.Selection.GetElementIds()
    sheets = [doc.GetElement(i) for i in sel_ids
              if isinstance(doc.GetElement(i), ViewSheet)]
    if sheets:
        return sheets

    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    if not all_sheets:
        forms.alert('There are no sheets in this model.',
                    title='DQT - On Sheets: DWGs', exitscript=True)

    sheet_map = {'{} - {}'.format(s.SheetNumber, s.Name): s for s in all_sheets}
    chosen = forms.SelectFromList.show(
        sorted(sheet_map.keys()),
        title='DQT - Pick Sheets',
        button_name='Select DWGs',
        multiselect=True,
    )
    if not chosen:
        forms.alert('No sheet selected. Cancelled.',
                    title='DQT - On Sheets: DWGs', exitscript=True)
    return [sheet_map[c] for c in chosen]


# ----------------------------------------------------------------- MAIN
if __name__ == '__main__':
    sheets = get_target_sheets()
    sheet_ids = set(eid_int(s.Id) for s in sheets)

    all_imports = FilteredElementCollector(doc) \
        .OfClass(ImportInstance) \
        .WhereElementIsNotElementType().ToElements()

    dwg_ids = [imp.Id for imp in all_imports if eid_int(imp.OwnerViewId) in sheet_ids]

    if dwg_ids:
        uidoc.Selection.SetElementIds(to_element_id_list(dwg_ids))
        notify('Selected {} DWG(s) on {} sheet(s).'.format(
            len(dwg_ids), len(sheets)),
            title='DQT - On Sheets: DWGs')
    else:
        forms.alert('No DWGs found on the selected sheets.',
                    title='DQT - On Sheets: DWGs')
