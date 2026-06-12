# -*- coding: utf-8 -*-
__title__   = "On Sheets:\nTitleBlocks"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select on Sheets: Title Blocks

Select the Title Block elements that live on the selected sheets.
___________________________________________________________
How-to:
- Select sheets in the Project Browser, OR
- Run with nothing selected and pick sheets from the list
- The title blocks on those sheets get selected
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
                               BuiltInCategory)

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

    # Fall back to a picker
    all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    if not all_sheets:
        forms.alert('There are no sheets in this model.',
                    title='DQT - On Sheets: TitleBlocks', exitscript=True)

    sheet_map = {'{} - {}'.format(s.SheetNumber, s.Name): s for s in all_sheets}
    chosen = forms.SelectFromList.show(
        sorted(sheet_map.keys()),
        title='DQT - Pick Sheets',
        button_name='Select Title Blocks',
        multiselect=True,
    )
    if not chosen:
        forms.alert('No sheet selected. Cancelled.',
                    title='DQT - On Sheets: TitleBlocks', exitscript=True)
    return [sheet_map[c] for c in chosen]


# ----------------------------------------------------------------- MAIN
if __name__ == '__main__':
    sheets = get_target_sheets()
    sheet_ids = set(eid_int(s.Id) for s in sheets)

    all_tb = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_TitleBlocks) \
        .WhereElementIsNotElementType().ToElements()

    tb_ids = [tb.Id for tb in all_tb if eid_int(tb.OwnerViewId) in sheet_ids]

    if tb_ids:
        uidoc.Selection.SetElementIds(to_element_id_list(tb_ids))
        notify('Selected {} title block(s) on {} sheet(s).'.format(
            len(tb_ids), len(sheets)),
            title='DQT - On Sheets: TitleBlocks')
    else:
        forms.alert('No title blocks found on the selected sheets.',
                    title='DQT - On Sheets: TitleBlocks')
