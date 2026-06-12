# -*- coding: utf-8 -*-
__title__   = "Select By\nCategory"
__author__  = "Dang Quoc Truong - DQT"
__doc__     = """Select By Category

Pick elements in the view while Revit only allows the categories you
chose. Great for selecting a messy area where you only care about a few
categories (e.g. just Walls + Doors, ignoring everything else).
___________________________________________________________
How-to:
- Run the tool
- Tick the categories you want to be selectable
- Drag a window / pick elements - only the allowed categories respond
- Press [Finish] (or Esc) to apply the selection
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

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from dqt_select.compat import to_element_id_list, eid_int, notify

# ----------------------------------------------------------------- VARIABLES
doc = __revit__.ActiveUIDocument.Document          # noqa: F821
uidoc = __revit__.ActiveUIDocument                 # noqa: F821
app = __revit__.Application                         # noqa: F821
rvt_year = int(app.VersionNumber)
selection = uidoc.Selection


# ----------------------------------------------------------------- HELPERS
def is_valid_category(cat):
    """Filter out internal / invalid categories."""
    try:
        if cat is None:
            return False
        # Must allow bounding box visibility -> roughly = "real" category
        if not cat.CanAddSubcategory and cat.CategoryType == DB.CategoryType.Internal:
            return False
        if rvt_year > 2022:
            try:
                if cat.BuiltInCategory == BuiltInCategory.INVALID:
                    return False
            except AttributeError:
                pass
        return True
    except Exception:
        return False


# ----------------------------------------------------------------- ISelectionFilter
class CategorySelectionFilter(ISelectionFilter):
    """Only allow elements whose category is in the allowed set."""

    def __init__(self, allowed_cat_int_ids):
        self.allowed = set(allowed_cat_int_ids)

    def AllowElement(self, element):
        try:
            if element.Category is None:
                return False
            return eid_int(element.Category.Id) in self.allowed
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


# ----------------------------------------------------------------- MAIN
if __name__ == '__main__':
    # Build category list
    all_cats = [c for c in doc.Settings.Categories if is_valid_category(c)]

    # Manually add a few non-model categories that are still pickable
    for extra_bic in (BuiltInCategory.OST_Grids,
                      BuiltInCategory.OST_Levels,
                      BuiltInCategory.OST_Viewports):
        try:
            extra = doc.Settings.Categories.get_Item(extra_bic)
            if extra is not None and extra not in all_cats:
                all_cats.append(extra)
        except Exception:
            pass

    cat_map = {c.Name: c for c in all_cats}
    names = sorted(cat_map.keys())

    chosen = forms.SelectFromList.show(
        names,
        title='DQT - Select By Category',
        button_name='Select',
        multiselect=True,
    )

    if not chosen:
        forms.alert('No category selected. Cancelled.',
                    title='DQT - Select By Category', exitscript=True)

    allowed_ids = [eid_int(cat_map[n].Id) for n in chosen]
    sel_filter = CategorySelectionFilter(allowed_ids)

    try:
        refs = selection.PickObjects(
            ObjectType.Element, sel_filter,
            'Pick elements (only chosen categories) then click Finish')
    except OperationCanceledException:
        forms.alert('Picking cancelled.',
                    title='DQT - Select By Category', exitscript=True)
        refs = []

    picked_ids = [doc.GetElement(r).Id for r in refs if doc.GetElement(r)]

    if picked_ids:
        uidoc.Selection.SetElementIds(to_element_id_list(picked_ids))
        notify('Selected {} element(s).'.format(len(picked_ids)),
                    title='DQT - Select By Category')
    else:
        forms.alert('Nothing picked.', title='DQT - Select By Category')
