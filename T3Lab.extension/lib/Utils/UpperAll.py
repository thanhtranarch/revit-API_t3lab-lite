# -*- coding: utf-8 -*-
"""
Upper All Text — selection-aware uppercase tool.

Behavior:
- Selection contains Dimensions  -> uppercase overrides on those dims only.
- Selection contains TextNotes   -> uppercase text on those notes only.
- Selection contains both        -> process both kinds, ignore other elements.
- Selection contains other only  -> do nothing (avoid accidental bulk run).
- Selection is empty             -> uppercase project-wide:
    * View names (all non-template views)
    * Sheet names
    * Title block instance text params
    * All TextNotes in the document
    * All Dimension overrides in the document
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, View, ViewSheet, Dimension, SpotDimension,
    TextNote, StorageType, Group, GroupType, Family, FamilySymbol,
    FamilyInstance, SpatialElement, Level, Grid, ElementType,
    IFailuresPreprocessor, FailureProcessingResult
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit

from Snippets._compat import eid_value

_SKIP_PARAM_TYPE_NAMES = frozenset(("URL", "Image"))

def _is_skippable_string_param(param):
    try:
        return param.Definition.ParameterType.ToString() in _SKIP_PARAM_TYPE_NAMES
    except Exception:
        return False


class WarningSwallower(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        try:
            failures = failuresAccessor.GetFailureMessages()
            for f in failures:
                failuresAccessor.DeleteWarning(f)
        except Exception:
            pass
        return FailureProcessingResult.Continue


def needs_upper(text):
    return bool(text) and any(c.islower() for c in text)

def set_string_param(param, new_val):
    if param is None or param.IsReadOnly or param.StorageType != StorageType.String:
        return False
    try:
        param.Set(new_val)
        return True
    except Exception:
        return False

def upper_dimension(dim):
    if isinstance(dim, SpotDimension):
        return False
    def upd(target):
        if target.Above:         target.Above         = target.Above.upper()
        if target.Below:         target.Below         = target.Below.upper()
        if target.Prefix:        target.Prefix        = target.Prefix.upper()
        if target.Suffix:        target.Suffix        = target.Suffix.upper()
        if target.ValueOverride: target.ValueOverride = target.ValueOverride.upper()
    try:
        if dim.HasOneSegment():
            upd(dim)
        else:
            for seg in dim.Segments:
                upd(seg)
        return True
    except Exception:
        return False

def upper_text_note(note):
    try:
        txt = note.Text
        if not needs_upper(txt):
            return False
        note.Text = txt.upper()
        return True
    except Exception:
        return False

def upper_element_string_params(elem, doc):
    count = 0
    for p in elem.Parameters:
        try:
            if p.IsReadOnly or p.StorageType != StorageType.String:
                continue
            if _is_skippable_string_param(p):
                continue
            val = p.AsString()
            if not needs_upper(val):
                continue
            if set_string_param(p, val.upper()):
                count += 1
        except Exception:
            continue
    return count

def rename_safely(elem, new_name):
    try:
        elem.Name = new_name
        return True
    except Exception:
        return False

def rename_element_name(el):
    try:
        if needs_upper(el.Name):
            return rename_safely(el, el.Name.upper())
    except Exception:
        pass
    return False

def rename_spatial_element(se):
    try:
        p = se.get_Parameter(BuiltInParameter.ROOM_NAME)
        if p and p.StorageType == StorageType.String and not p.IsReadOnly:
            val = p.AsString()
            if needs_upper(val):
                return set_string_param(p, val.upper())

        p = se.LookupParameter("Name")
        if p and p.StorageType == StorageType.String and not p.IsReadOnly:
            val = p.AsString()
            if needs_upper(val):
                return set_string_param(p, val.upper())
    except Exception:
        pass
    return False


def process_selection(elements, doc):
    s = {"views": 0, "sheets": 0, "groups": 0, "grouptypes": 0, "families": 0,
         "familytypes": 0, "spatial": 0, "levels": 0, "grids": 0, "titleblocks": 0,
         "notes": 0, "dims": 0, "skipped": 0}
    for el in elements:
        try:
            if isinstance(el, Dimension):
                if upper_dimension(el):
                    s["dims"] += 1
            elif isinstance(el, TextNote):
                if upper_text_note(el):
                    s["notes"] += 1
            elif isinstance(el, View):
                if el.IsTemplate:
                    continue
                if rename_element_name(el):
                    s["views"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
                tos = el.LookupParameter("Title on Sheet")
                if tos:
                    cur = tos.AsString()
                    if needs_upper(cur):
                        set_string_param(tos, cur.upper())
            elif isinstance(el, ViewSheet):
                if rename_element_name(el):
                    s["sheets"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
            elif isinstance(el, Group):
                if rename_element_name(el):
                    s["groups"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
                try:
                    gt = el.GroupType
                    if gt and rename_element_name(gt):
                        s["grouptypes"] += 1
                except Exception:
                    pass
            elif isinstance(el, GroupType):
                if rename_element_name(el):
                    s["grouptypes"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
            elif isinstance(el, Family):
                if rename_element_name(el):
                    s["families"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
            elif isinstance(el, FamilySymbol):
                if rename_element_name(el):
                    s["familytypes"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
            elif isinstance(el, Level):
                if rename_element_name(el):
                    s["levels"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
            elif isinstance(el, Grid):
                if rename_element_name(el):
                    s["grids"] += 1
                else:
                    if needs_upper(el.Name):
                        s["skipped"] += 1
            elif isinstance(el, SpatialElement):
                if rename_spatial_element(el):
                    s["spatial"] += 1
            elif isinstance(el, FamilyInstance):
                try:
                    sym = el.Symbol
                    if sym and rename_element_name(sym):
                        s["familytypes"] += 1
                    fam = sym.Family
                    if fam and rename_element_name(fam):
                        s["families"] += 1
                except Exception:
                    pass
            elif el.Category and eid_value(el.Category.Id) == int(BuiltInCategory.OST_TitleBlocks):
                if upper_element_string_params(el, doc) > 0:
                    s["titleblocks"] += 1
            
            try:
                type_id = el.GetTypeId()
                if type_id and eid_value(type_id) != -1:
                    elem_type = doc.GetElement(type_id)
                    if elem_type and isinstance(elem_type, ElementType):
                        if rename_element_name(elem_type):
                            s["familytypes"] += 1
            except Exception:
                pass
        except Exception:
            pass
    return s


def process_all_text(doc):
    s = {"views": 0, "sheets": 0, "groups": 0, "grouptypes": 0, "families": 0,
         "familytypes": 0, "spatial": 0, "levels": 0, "grids": 0, "titleblocks": 0,
         "notes": 0, "dims": 0, "skipped": 0}
    
    for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType():
        try:
            if v.IsTemplate:
                continue
            if rename_element_name(v):
                s["views"] += 1
            else:
                if needs_upper(v.Name):
                    s["skipped"] += 1
            tos = v.LookupParameter("Title on Sheet")
            if tos:
                cur = tos.AsString()
                if needs_upper(cur):
                    set_string_param(tos, cur.upper())
        except Exception:
            pass

    for sh in FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType():
        try:
            if rename_element_name(sh):
                s["sheets"] += 1
            else:
                if needs_upper(sh.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for gt in FilteredElementCollector(doc).OfClass(GroupType):
        try:
            if rename_element_name(gt):
                s["grouptypes"] += 1
            else:
                if needs_upper(gt.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for g in FilteredElementCollector(doc).OfClass(Group).WhereElementIsNotElementType():
        try:
            if rename_element_name(g):
                s["groups"] += 1
            else:
                if needs_upper(g.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for fam in FilteredElementCollector(doc).OfClass(Family):
        try:
            if rename_element_name(fam):
                s["families"] += 1
            else:
                if needs_upper(fam.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for fs in FilteredElementCollector(doc).OfClass(FamilySymbol):
        try:
            if rename_element_name(fs):
                s["familytypes"] += 1
            else:
                if needs_upper(fs.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for se in FilteredElementCollector(doc).OfClass(SpatialElement).WhereElementIsNotElementType():
        try:
            if rename_spatial_element(se):
                s["spatial"] += 1
        except Exception:
            pass

    for lvl in FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType():
        try:
            if rename_element_name(lvl):
                s["levels"] += 1
            else:
                if needs_upper(lvl.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for grd in FilteredElementCollector(doc).OfClass(Grid).WhereElementIsNotElementType():
        try:
            if rename_element_name(grd):
                s["grids"] += 1
            else:
                if needs_upper(grd.Name):
                    s["skipped"] += 1
        except Exception:
            pass

    for tb in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType():
        try:
            if upper_element_string_params(tb, doc) > 0:
                s["titleblocks"] += 1
        except Exception:
            pass

    for tn in FilteredElementCollector(doc).OfClass(TextNote).WhereElementIsNotElementType():
        try:
            if upper_text_note(tn):
                s["notes"] += 1
        except Exception:
            pass

    for d in FilteredElementCollector(doc).OfClass(Dimension).WhereElementIsNotElementType():
        try:
            if upper_dimension(d):
                s["dims"] += 1
        except Exception:
            pass

    return s


def run():
    uidoc = revit.uidoc
    doc = revit.doc
    selected = [doc.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]

    t = Transaction(doc, "Upper All Text")
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(WarningSwallower())
    t.SetFailureHandlingOptions(options)
    
    t.Start()
    try:
        if selected:
            s = process_selection(selected, doc)
            t.Commit()
            
            total_processed = (
                s["views"] + s["sheets"] + s["groups"] + s["grouptypes"] +
                s["families"] + s["familytypes"] + s["spatial"] + s["levels"] +
                s["grids"] + s["titleblocks"] + s["notes"] + s["dims"]
            )
            
            if total_processed == 0 and s["skipped"] == 0:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                TaskDialog.Show(
                    "Upper All Text",
                    "No valid elements were processed in the selection.\n\n"
                    "Supported elements: Dimensions, Text Notes, Views, Sheets, Groups, Families/Types, Rooms/Spaces/Areas, Levels, Grids, and Title Blocks.\n\n"
                    "Tip: Clear the selection to process the entire project.")
                return
            
            msg = ("Uppercase applied to selection:\n"
                   "  - Views:        {}\n"
                   "  - Sheets:       {}\n"
                   "  - Groups:       {}\n"
                   "  - Group Types:  {}\n"
                   "  - Families:     {}\n"
                   "  - Family Types: {}\n"
                   "  - Rooms/Spaces: {}\n"
                   "  - Levels:       {}\n"
                   "  - Grids:        {}\n"
                   "  - Title blocks: {}\n"
                   "  - TextNotes:    {}\n"
                   "  - Dimensions:   {}\n"
                   "  - Skipped (locked/duplicate): {}").format(
                       s["views"], s["sheets"], s["groups"], s["grouptypes"],
                       s["families"], s["familytypes"], s["spatial"], s["levels"],
                       s["grids"], s["titleblocks"], s["notes"], s["dims"],
                       s["skipped"])
        else:
            s = process_all_text(doc)
            t.Commit()
            msg = ("Uppercase applied across project:\n"
                   "  - Views:        {}\n"
                   "  - Sheets:       {}\n"
                   "  - Groups:       {}\n"
                   "  - Group Types:  {}\n"
                   "  - Families:     {}\n"
                   "  - Family Types: {}\n"
                   "  - Rooms/Spaces: {}\n"
                   "  - Levels:       {}\n"
                   "  - Grids:        {}\n"
                   "  - Title blocks: {}\n"
                   "  - TextNotes:    {}\n"
                   "  - Dimensions:   {}\n"
                   "  - Skipped (locked/duplicate): {}").format(
                       s["views"], s["sheets"], s["groups"], s["grouptypes"],
                       s["families"], s["familytypes"], s["spatial"], s["levels"],
                       s["grids"], s["titleblocks"], s["notes"], s["dims"],
                       s["skipped"])
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        TaskDialog.Show("Upper All Text", "Error: {}".format(ex))
        return

    TaskDialog.Show("Upper All Text", msg)
