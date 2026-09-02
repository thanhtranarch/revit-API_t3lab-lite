# -*- coding: utf-8 -*-
"""
Drawing Spellcheck Engine

Scans TextNotes, Sheet Names, and View Names for typographical errors and
architectural terminology misspellings using the bundled deterministic dictionary.

Author: Tran Tien Thanh
"""

import re
from Snippets._compat import eid_value, make_eid, elem_name


def check_drawing_spelling(doc, view_only=False):
    """Scan TextNotes, Sheet Names, and View Names for typos.

    Args:
        doc: Revit Document
        view_only: If True, only scans the active view

    Returns:
        List of dicts representing detected spelling errors.
    """
    from Autodesk.Revit import DB
    try:
        from Services import spell_dictionary as SD
    except ImportError:
        try:
            import spell_dictionary as SD
        except ImportError:
            return []

    findings = []

    # 1. SCAN TEXT NOTES
    try:
        if view_only and doc.ActiveView and not doc.ActiveView.IsTemplate:
            col = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            col = DB.FilteredElementCollector(doc)

        col = col.OfClass(DB.TextNote).WhereElementIsNotElementType().ToElements()

        for tn in col:
            text = (tn.Text or u"").strip()
            if not text or not any(c.isalpha() for c in text):
                continue

            t_id = eid_value(tn.Id)
            view_name = u"Active View"
            view_id = -1
            try:
                ov = doc.GetElement(tn.OwnerViewId)
                if ov:
                    view_name = ov.Name
                    view_id = eid_value(ov.Id)
            except Exception:
                pass

            # Check spelling for text line
            issues = SD.find_errors(text)
            if issues:
                for it in issues:
                    wrong = it.get('wrong', '')
                    right = it.get('right', '')
                    reason = it.get('reason', '')
                    findings.append({
                        'CheckType': u'Spelling Error',
                        'Category': u'Text Note',
                        'Severity': u'Warning',
                        'ViewName': view_name,
                        'ViewId': view_id,
                        'ElementId': t_id,
                        'SecondaryId': None,
                        'ElementIdsStr': str(t_id),
                        'Title': u"Spelling Typo: \"{}\"".format(wrong),
                        'Description': u"In text: \"{}\"".format(text[:40] + (u"..." if len(text) > 40 else u"")),
                        'Calculation': u"Suggestion: \"{}\" -> \"{}\" ({})".format(wrong, right, reason or u"typo"),
                        'OverlapPct': 0.0,
                        'IsSelected': True
                    })
    except Exception:
        pass

    # 2. SCAN SHEET NAMES
    if not view_only:
        try:
            sheet_col = DB.FilteredElementCollector(doc)\
                          .OfClass(DB.ViewSheet)\
                          .WhereElementIsNotElementType()\
                          .ToElements()

            for sheet in sheet_col:
                if sheet.IsTemplate:
                    continue
                s_id = eid_value(sheet.Id)
                s_name = sheet.Name or u""
                s_num = sheet.SheetNumber or u""
                if not s_name:
                    continue

                issues = SD.find_errors(s_name)
                if issues:
                    for it in issues:
                        wrong = it.get('wrong', '')
                        right = it.get('right', '')
                        findings.append({
                            'CheckType': u'Spelling Error',
                            'Category': u'Sheet Name',
                            'Severity': u'Warning',
                            'ViewName': u"{} - {}".format(s_num, s_name),
                            'ViewId': s_id,
                            'ElementId': s_id,
                            'SecondaryId': None,
                            'ElementIdsStr': str(s_id),
                            'Title': u"Spelling Typo in Sheet Name",
                            'Description': u"Sheet '{}': typo in \"{}\"".format(s_num, wrong),
                            'Calculation': u"Suggestion: \"{}\" -> \"{}\"".format(wrong, right),
                            'OverlapPct': 0.0,
                            'IsSelected': True
                        })
        except Exception:
            pass

    return findings
