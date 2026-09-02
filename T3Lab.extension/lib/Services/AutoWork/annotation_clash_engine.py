# -*- coding: utf-8 -*-
"""
Annotation Clash Engine

Detects overlapping and colliding annotations (TextNotes, Tags, Dimensions)
in 2D views, calculates collision geometry, overlap percentages, and severity.

Author: Tran Tien Thanh
"""

import math
from Snippets._compat import eid_value, make_eid, elem_name


def _get_element_text_preview(elem):
    """Extract a representative text or name preview from an annotation element."""
    try:
        from Autodesk.Revit import DB
        if isinstance(elem, DB.TextNote):
            t = (elem.Text or u"").strip()
            return (t[:30] + u"...") if len(t) > 30 else t
        elif isinstance(elem, DB.IndependentTag):
            try:
                tagged_text = elem.TagText
                if tagged_text:
                    return u"[Tag] {}".format(tagged_text)
            except Exception:
                pass
            return u"[Tag] {}".format(elem_name(elem))
        elif isinstance(elem, DB.Dimension):
            try:
                val = elem.ValueString
                if val:
                    return u"[Dim] {}".format(val)
            except Exception:
                pass
            return u"[Dim] {}".format(elem_name(elem))
        return elem_name(elem)
    except Exception:
        return u"Annotation"


def _get_category_label(elem):
    """Return friendly category name."""
    try:
        from Autodesk.Revit import DB
        if isinstance(elem, DB.TextNote):
            return u"Text Note"
        elif isinstance(elem, DB.IndependentTag):
            return u"Tag"
        elif isinstance(elem, DB.Dimension):
            return u"Dimension"
        elif elem.Category:
            return elem.Category.Name
        return u"Annotation"
    except Exception:
        return u"Annotation"


def check_annotation_clashes(doc, views=None, min_overlap_pct=5.0):
    """Scan views for overlapping TextNotes, IndependentTags, and Dimensions.

    Args:
        doc: Revit Document
        views: List of View elements to inspect (default: active view or all views on sheets)
        min_overlap_pct: Minimum overlap ratio (%) to be flagged as a clash

    Returns:
        List of dicts representing detected annotation clashes.
    """
    from Autodesk.Revit import DB

    if views is None:
        if doc.ActiveView and not doc.ActiveView.IsTemplate:
            views = [doc.ActiveView]
        else:
            # Collect all views placed on sheets
            views = []
            collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet)
            for sheet in collector:
                for vp_id in sheet.GetAllViewports():
                    vp = doc.GetElement(vp_id)
                    if vp:
                        v = doc.GetElement(vp.ViewId)
                        if v and not v.IsTemplate and v not in views:
                            views.append(v)

    clashes = []

    for view in views:
        if view.IsTemplate:
            continue

        view_name = view.Name
        view_id = eid_value(view.Id)

        # Collect annotation elements in this view
        annos = []
        try:
            col_text = DB.FilteredElementCollector(doc, view.Id)\
                         .OfClass(DB.TextNote)\
                         .WhereElementIsNotElementType()\
                         .ToElements()
            annos.extend(col_text)
        except Exception:
            pass

        try:
            col_tags = DB.FilteredElementCollector(doc, view.Id)\
                         .OfClass(DB.IndependentTag)\
                         .WhereElementIsNotElementType()\
                         .ToElements()
            annos.extend(col_tags)
        except Exception:
            pass

        try:
            col_dims = DB.FilteredElementCollector(doc, view.Id)\
                         .OfClass(DB.Dimension)\
                         .WhereElementIsNotElementType()\
                         .ToElements()
            annos.extend(col_dims)
        except Exception:
            pass

        # Extract 2D bounding boxes in view coordinate space
        boxes = []
        for elem in annos:
            try:
                bbox = elem.get_BoundingBox(view)
                if not bbox:
                    continue
                min_pt = bbox.Min
                max_pt = bbox.Max
                width = abs(max_pt.X - min_pt.X)
                height = abs(max_pt.Y - min_pt.Y)
                area = width * height

                # Skip degenerate 0-size bounding boxes
                if width < 0.001 or height < 0.001 or area < 0.000001:
                    continue

                boxes.append({
                    'elem': elem,
                    'id': eid_value(elem.Id),
                    'min_x': min(min_pt.X, max_pt.X),
                    'max_x': max(min_pt.X, max_pt.X),
                    'min_y': min(min_pt.Y, max_pt.Y),
                    'max_y': max(min_pt.Y, max_pt.Y),
                    'area': area,
                    'preview': _get_element_text_preview(elem),
                    'cat': _get_category_label(elem)
                })
            except Exception:
                continue

        count = len(boxes)
        # Pairwise comparison
        for i in xrange(count):
            a = boxes[i]
            for j in xrange(i + 1, count):
                b = boxes[j]

                # Quick axis separation check
                if a['max_x'] <= b['min_x'] or a['min_x'] >= b['max_x']:
                    continue
                if a['max_y'] <= b['min_y'] or a['min_y'] >= b['max_y']:
                    continue

                # Calculate overlap box
                overlap_w = min(a['max_x'], b['max_x']) - max(a['min_x'], b['min_x'])
                overlap_h = min(a['max_y'], b['max_y']) - max(a['min_y'], b['min_y'])
                if overlap_w <= 0.0001 or overlap_h <= 0.0001:
                    continue

                overlap_area = overlap_w * overlap_h
                min_area = min(a['area'], b['area'])
                if min_area <= 0.000001:
                    continue

                overlap_pct = (overlap_area / min_area) * 100.0

                if overlap_pct >= min_overlap_pct:
                    # Determine severity based on overlap magnitude
                    if overlap_pct >= 50.0:
                        severity = u"Critical"
                        sev_badge = u"Critical"
                    elif overlap_pct >= 20.0:
                        severity = u"Warning"
                        sev_badge = u"Warning"
                    else:
                        severity = u"Info"
                        sev_badge = u"Info"

                    # Convert internal feet to mm (1 ft = 304.8 mm)
                    overlap_area_mm2 = overlap_area * 92903.04

                    clashes.append({
                        'CheckType': u'Annotation Clash',
                        'Category': u'{} ∩ {}'.format(a['cat'], b['cat']),
                        'Severity': severity,
                        'ViewName': view_name,
                        'ViewId': view_id,
                        'ElementId': a['id'],
                        'SecondaryId': b['id'],
                        'ElementIdsStr': u"{}, {}".format(a['id'], b['id']),
                        'Title': u"Overlapping Annotations",
                        'Description': u"\"{}\" clashes with \"{}\"".format(a['preview'], b['preview']),
                        'Calculation': u"Overlap: {:.1f}% ({:.0f} mm²)".format(overlap_pct, overlap_area_mm2),
                        'OverlapPct': overlap_pct,
                        'IsSelected': True
                    })

    return clashes
