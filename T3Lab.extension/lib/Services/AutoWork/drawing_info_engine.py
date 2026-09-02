# -*- coding: utf-8 -*-
"""
Drawing & Information QA/QC Engine

Audits sheets, placed views, rooms, and elements for missing parameters,
empty titleblock metadata, missing view templates, unplaced rooms, and duplicate values.

Author: Tran Tien Thanh
"""

from collections import defaultdict
from Snippets._compat import eid_value, make_eid, elem_name


def check_drawing_and_sheet_info(doc):
    """Audit drawings, sheets, views, and spatial elements for information errors.

    Returns:
        List of dicts representing detected information defects.
    """
    from Autodesk.Revit import DB

    issues = []

    # 1. AUDIT SHEETS & TITLEBLOCK METADATA
    try:
        sheet_collector = DB.FilteredElementCollector(doc)\
                            .OfClass(DB.ViewSheet)\
                            .WhereElementIsNotElementType()\
                            .ToElements()

        sheet_numbers = defaultdict(list)

        for sheet in sheet_collector:
            if sheet.IsTemplate:
                continue

            s_id = eid_value(sheet.Id)
            s_num = sheet.SheetNumber or u""
            s_name = sheet.Name or u""
            sheet_numbers[s_num].append(sheet)

            # Check Viewports on Sheet
            viewports = sheet.GetAllViewports()
            if len(viewports) == 0:
                issues.append({
                    'CheckType': u'Sheet Completeness',
                    'Category': u'Empty Sheet',
                    'Severity': u'Warning',
                    'ViewName': u"{} - {}".format(s_num, s_name),
                    'ViewId': s_id,
                    'ElementId': s_id,
                    'SecondaryId': None,
                    'ElementIdsStr': str(s_id),
                    'Title': u"Sheet has no placed views",
                    'Description': u"Sheet '{}' contains 0 viewports or schedules.".format(s_num),
                    'Calculation': u"Viewports: 0 (Empty sheet)",
                    'OverlapPct': 0.0,
                    'IsSelected': True
                })

            # Check Sheet Metadata (Drawn By, Checked By, Issue Date)
            params_to_check = [
                (DB.BuiltInParameter.SHEET_DRAWN_BY, u"Drawn By", u"Info"),
                (DB.BuiltInParameter.SHEET_CHECKED_BY, u"Checked By", u"Warning"),
                (DB.BuiltInParameter.SHEET_ISSUE_DATE, u"Issue Date", u"Warning"),
            ]
            missing_meta = []
            for bip, lbl, sev in params_to_check:
                p = sheet.get_Parameter(bip)
                if not p or not p.AsString() or not p.AsString().strip():
                    missing_meta.append(lbl)

            if missing_meta:
                issues.append({
                    'CheckType': u'Information Error',
                    'Category': u'Sheet Metadata',
                    'Severity': u'Warning' if u'Checked By' in missing_meta else u'Info',
                    'ViewName': u"{} - {}".format(s_num, s_name),
                    'ViewId': s_id,
                    'ElementId': s_id,
                    'SecondaryId': None,
                    'ElementIdsStr': str(s_id),
                    'Title': u"Missing Titleblock Metadata",
                    'Description': u"Empty parameter(s): {}".format(u", ".join(missing_meta)),
                    'Calculation': u"Missing: {}/3 metadata fields".format(len(missing_meta)),
                    'OverlapPct': 0.0,
                    'IsSelected': True
                })

            # Check Placed Views on this Sheet for View Templates
            for vp_id in viewports:
                vp = doc.GetElement(vp_id)
                if not vp:
                    continue
                v = doc.GetElement(vp.ViewId)
                if v and not v.IsTemplate:
                    v_id = eid_value(v.Id)
                    # Check View Template
                    vt_id = eid_value(v.ViewTemplateId)
                    if vt_id == -1:
                        issues.append({
                            'CheckType': u'Drawing Standard',
                            'Category': u'View Template',
                            'Severity': u'Warning',
                            'ViewName': u"{} - {}".format(s_num, v.Name),
                            'ViewId': v_id,
                            'ElementId': v_id,
                            'SecondaryId': None,
                            'ElementIdsStr': str(v_id),
                            'Title': u"No View Template Assigned",
                            'Description': u"View '{}' on sheet '{}' has no View Template applied.".format(v.Name, s_num),
                            'Calculation': u"Template ID: None (Manual overrides active)",
                            'OverlapPct': 0.0,
                            'IsSelected': True
                        })

        # Check Duplicate Sheet Numbers
        for s_num, s_list in sheet_numbers.items():
            if len(s_list) > 1:
                for s in s_list:
                    s_id = eid_value(s.Id)
                    issues.append({
                        'CheckType': u'Information Error',
                        'Category': u'Duplicate Number',
                        'Severity': u'Critical',
                        'ViewName': u"{} - {}".format(s_num, s.Name),
                        'ViewId': s_id,
                        'ElementId': s_id,
                        'SecondaryId': None,
                        'ElementIdsStr': str(s_id),
                        'Title': u"Duplicate Sheet Number",
                        'Description': u"Sheet Number '{}' is duplicated across {} sheets.".format(s_num, len(s_list)),
                        'Calculation': u"Count: {} duplicates".format(len(s_list)),
                        'OverlapPct': 0.0,
                        'IsSelected': True
                    })

    except Exception:
        pass

    # 2. AUDIT ROOMS & SPATIAL INFORMATION
    try:
        room_collector = DB.FilteredElementCollector(doc)\
                           .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                           .WhereElementIsNotElementType()\
                           .ToElements()

        room_numbers = defaultdict(list)

        for room in room_collector:
            r_id = eid_value(room.Id)
            r_num = room.Number or u""
            r_name = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
            r_name_str = r_name.AsString() if r_name else u""

            # Check Unplaced Rooms
            if room.Location is None:
                issues.append({
                    'CheckType': u'Spatial QA/QC',
                    'Category': u'Unplaced Room',
                    'Severity': u'Warning',
                    'ViewName': u"Model-wide",
                    'ViewId': -1,
                    'ElementId': r_id,
                    'SecondaryId': None,
                    'ElementIdsStr': str(r_id),
                    'Title': u"Unplaced Room in Project",
                    'Description': u"Room #{} '{}' is unplaced (0 Area).".format(r_num, r_name_str),
                    'Calculation': u"Area: 0.00 m² (Location is None)",
                    'OverlapPct': 0.0,
                    'IsSelected': True
                })
                continue

            # Check Not Enclosed / Redundant Rooms
            try:
                area = room.Area
                if area <= 0.0001:
                    issues.append({
                        'CheckType': u'Spatial QA/QC',
                        'Category': u'Not Enclosed Room',
                        'Severity': u'Critical',
                        'ViewName': u"Model-wide",
                        'ViewId': -1,
                        'ElementId': r_id,
                        'SecondaryId': None,
                        'ElementIdsStr': str(r_id),
                        'Title': u"Not Enclosed or Redundant Room",
                        'Description': u"Room #{} '{}' has bounding issues (Area = 0).".format(r_num, r_name_str),
                        'Calculation': u"Bounding loop open / Redundant boundary",
                        'OverlapPct': 0.0,
                        'IsSelected': True
                    })
            except Exception:
                pass

            # Track room numbers for duplicates
            if r_num:
                room_numbers[r_num].append(room)

        # Check Duplicate Room Numbers
        for r_num, r_list in room_numbers.items():
            if len(r_list) > 1:
                for r in r_list:
                    r_id = eid_value(r.Id)
                    issues.append({
                        'CheckType': u'Information Error',
                        'Category': u'Duplicate Room #',
                        'Severity': u'Critical',
                        'ViewName': u"Model-wide",
                        'ViewId': -1,
                        'ElementId': r_id,
                        'SecondaryId': None,
                        'ElementIdsStr': str(r_id),
                        'Title': u"Duplicate Room Number",
                        'Description': u"Room Number '{}' is assigned to {} distinct rooms.".format(r_num, len(r_list)),
                        'Calculation': u"Count: {} occurrences".format(len(r_list)),
                        'OverlapPct': 0.0,
                        'IsSelected': True
                    })

    except Exception:
        pass

    return issues
