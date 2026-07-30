# -*- coding: utf-8 -*-
"""pyDQT - Select Tools | Core selection engine.

Self contained selection logic for the pyDQT Select pulldown.
Provides "Select Similar" by Type / Family / Category, scoped to either
the active View or the whole Model, with smart per-element filtering rules
so that picking one Line / Grid / Room etc. does not pull in unrelated
elements.

Dang Quoc Truong - DQT (c) 2026
"""

from pyrevit import forms

from Autodesk.Revit.DB import FilteredElementCollector

try:
    from dqt_select.compat import eid_int, to_element_id_list, notify
except ImportError:
    from compat import eid_int, to_element_id_list, notify


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _get_uidoc():
    return __revit__.ActiveUIDocument  # noqa: F821


def _get_selected_elements(uidoc):
    """Return currently selected (instance) elements as a python list."""
    doc = uidoc.Document
    ids = uidoc.Selection.GetElementIds()
    return [doc.GetElement(i) for i in ids if doc.GetElement(i) is not None]


def _collector(doc, mode):
    """Return a FilteredElementCollector scoped to view or model.

    :param mode: 'view' -> active view only, 'model' -> whole document
    """
    if mode == 'view':
        return FilteredElementCollector(doc, doc.ActiveView.Id)
    return FilteredElementCollector(doc)


def _set_selection(uidoc, ids):
    """Apply selection and report the count to the user."""
    unique_ids = list({eid_int(i): i for i in ids}.values())
    uidoc.Selection.SetElementIds(to_element_id_list(unique_ids))
    return len(unique_ids)


# ----------------------------------------------------------------------------
# Type signature rules
# ----------------------------------------------------------------------------
# Some elements have no meaningful "Type", so "Select Similar Type" should
# fall back to matching by Category. These categories are matched by
# BuiltInCategory rather than by GetTypeId().
#
# The rules themselves live in Snippets/_similar.py so the MCP server's
# operate_element(select_similar) shares them — this module can't be imported
# from the server (it calls __revit__ at module scope and pulls pyrevit.forms).
from Snippets._similar import (                                  # noqa: E402
    is_category_only as _is_category_only,
    family_id_of as _shared_family_id_of,
)


# ----------------------------------------------------------------------------
# SELECT SIMILAR : CATEGORY
# ----------------------------------------------------------------------------
def select_similar_category(mode='view'):
    """Select every element sharing the category of the current selection."""
    uidoc = _get_uidoc()
    doc = uidoc.Document
    selected = _get_selected_elements(uidoc)

    if not selected:
        forms.alert('Please select at least one element first.',
                    title='DQT - Select Similar: Category')
        return

    cat_ids = set()
    for e in selected:
        if e.Category is not None:
            cat_ids.add(eid_int(e.Category.Id))

    if not cat_ids:
        forms.alert('Selected elements have no category to match.',
                    title='DQT - Select Similar: Category')
        return

    collector = _collector(doc, mode).WhereElementIsNotElementType().ToElements()

    result_ids = []
    for e in collector:
        if e.Category is not None and eid_int(e.Category.Id) in cat_ids:
            result_ids.append(e.Id)

    count = _set_selection(uidoc, result_ids)
    scope = 'view' if mode == 'view' else 'model'
    notify('Selected {} element(s) by category in {}.'.format(count, scope),
                title='DQT - Select Similar: Category')


# ----------------------------------------------------------------------------
# SELECT SIMILAR : FAMILY
# ----------------------------------------------------------------------------
def select_similar_family(mode='view'):
    """Select every instance sharing the family of the current selection."""
    uidoc = _get_uidoc()
    doc = uidoc.Document
    selected = _get_selected_elements(uidoc)

    if not selected:
        forms.alert('Please select at least one element first.',
                    title='DQT - Select Similar: Family')
        return

    # Collect target family ids (only FamilyInstance-based elements have one)
    family_ids = set()
    category_fallback_ids = set()    # for elements without a Family
    for e in selected:
        fam_id = _family_id_of(e, doc)
        if fam_id is not None:
            family_ids.add(eid_int(fam_id))
        elif e.Category is not None:
            category_fallback_ids.add(eid_int(e.Category.Id))

    if not family_ids and not category_fallback_ids:
        forms.alert('Selected elements have no family to match.',
                    title='DQT - Select Similar: Family')
        return

    collector = _collector(doc, mode).WhereElementIsNotElementType().ToElements()

    result_ids = []
    for e in collector:
        fam_id = _family_id_of(e, doc)
        if fam_id is not None and eid_int(fam_id) in family_ids:
            result_ids.append(e.Id)
        elif fam_id is None and e.Category is not None \
                and eid_int(e.Category.Id) in category_fallback_ids:
            result_ids.append(e.Id)

    count = _set_selection(uidoc, result_ids)
    scope = 'view' if mode == 'view' else 'model'
    notify('Selected {} element(s) by family in {}.'.format(count, scope),
                title='DQT - Select Similar: Family')


def _family_id_of(element, doc):
    """Return the Family ElementId of an element, or None.

    Works for FamilyInstance (loadable & in-place), and also for annotation
    types that expose a Family. System family hosts (Walls, Floors etc.) have
    no reliable Family, so those return None and the caller falls back to
    category matching. Implementation shared with the MCP server.
    """
    return _shared_family_id_of(element, doc)


# ----------------------------------------------------------------------------
# SELECT SIMILAR : TYPE  (Super Select)
# ----------------------------------------------------------------------------
def select_similar_type(mode='view'):
    """Improved 'Select All Instances' that supports multiple seeds and
    handles unusual elements (lines, grids, rooms, scope boxes, ...).
    """
    uidoc = _get_uidoc()
    doc = uidoc.Document
    selected = _get_selected_elements(uidoc)

    if not selected:
        forms.alert('Please select at least one element first.',
                    title='DQT - Select Similar: Type')
        return

    type_ids = set()          # match these GetTypeId()
    category_only_ids = set()  # match these by category (no usable type)

    for e in selected:
        if _is_category_only(e):
            if e.Category is not None:
                category_only_ids.add(eid_int(e.Category.Id))
            else:
                # No category at all -> match by python class name later
                category_only_ids.add(('cls', e.__class__.__name__))
        else:
            try:
                tid = e.GetTypeId()
                if tid is not None and eid_int(tid) != -1:
                    type_ids.add(eid_int(tid))
                elif e.Category is not None:
                    category_only_ids.add(eid_int(e.Category.Id))
            except Exception:
                if e.Category is not None:
                    category_only_ids.add(eid_int(e.Category.Id))

    collector = _collector(doc, mode).WhereElementIsNotElementType().ToElements()

    result_ids = []
    for e in collector:
        matched = False

        # 1) Match by type id
        if type_ids:
            try:
                tid = e.GetTypeId()
                if tid is not None and eid_int(tid) in type_ids:
                    matched = True
            except Exception:
                pass

        # 2) Match by category (lines, grids, rooms, scope boxes, ...)
        if not matched and category_only_ids:
            if e.Category is not None and eid_int(e.Category.Id) in category_only_ids:
                matched = True
            elif ('cls', e.__class__.__name__) in category_only_ids:
                matched = True

        if matched:
            result_ids.append(e.Id)

    count = _set_selection(uidoc, result_ids)
    scope = 'view' if mode == 'view' else 'model'
    notify('Selected {} element(s) by type in {}.'.format(count, scope),
                title='DQT - Select Similar: Type')
