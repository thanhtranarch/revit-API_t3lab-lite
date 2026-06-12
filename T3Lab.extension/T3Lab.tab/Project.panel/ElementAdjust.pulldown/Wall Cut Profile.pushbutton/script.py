# -*- coding: utf-8 -*-
"""
Wall Profile Cut from Linked Elements v4
Creates openings in walls based on intersecting elements from linked models.
Uses pyrevit.forms UI only (no custom WPF - proven stable).
Copyright (c) 2026 Dang Quoc Truong (DQT)
"""

__title__ = "Wall Profile\nCut from Link"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Cut wall profiles or create openings based on linked element intersections"

import clr
import os
import traceback

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, TransactionGroup, XYZ, Line,
    ElementId, RevitLinkInstance, Wall,
    SketchEditScope, FailureProcessingResult,
    FamilySymbol, FamilyInstance,
    HostObjectUtils, ShellLayerType, Level
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from pyrevit import forms, revit, script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

LOG_PATH = r"C:\Temp\DQT_WallProfileCut_log.txt"


# ==============================================================================
# HELPERS
# ==============================================================================

def _eid_int(eid):
    try:
        return eid.Value
    except:
        return eid.IntegerValue


def ft_to_mm(ft):
    return ft * 304.8


def mm_to_ft(mm):
    return mm / 304.8


def log(msg):
    try:
        folder = os.path.dirname(LOG_PATH)
        if not os.path.exists(folder):
            os.makedirs(folder)
        with open(LOG_PATH, "a") as f:
            f.write(str(msg) + "\n")
    except:
        pass


# ==============================================================================
# CATEGORIES
# ==============================================================================

CATEGORIES = {
    "Structural Framing": BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": BuiltInCategory.OST_StructuralColumns,
    "Columns": BuiltInCategory.OST_Columns,
    "Ducts": BuiltInCategory.OST_DuctCurves,
    "Pipes": BuiltInCategory.OST_PipeCurves,
    "Cable Trays": BuiltInCategory.OST_CableTray,
    "Conduits": BuiltInCategory.OST_Conduit,
    "Floors": BuiltInCategory.OST_Floors,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
}


# ==============================================================================
# SELECTION FILTERS
# ==============================================================================

class LinkFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, RevitLinkInstance)
        except:
            return False

    def AllowReference(self, ref, point):
        return False


class WallFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if not isinstance(elem, Wall):
                return False
            kind = elem.WallType.Kind
            if kind == DB.WallKind.Curtain or kind == DB.WallKind.Stacked:
                return False
            loc = elem.Location
            if loc is None or not hasattr(loc, 'Curve') or loc.Curve is None:
                return False
            return True
        except:
            return False

    def AllowReference(self, ref, point):
        return False


# ==============================================================================
# FAILURES PREPROCESSOR
# ==============================================================================

class WarningSwallower(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        for f in fa.GetFailureMessages():
            sev = f.GetSeverity()
            if sev == DB.FailureSeverity.Warning:
                fa.DeleteWarning(f)
            elif sev == DB.FailureSeverity.Error:
                # Try to resolve (e.g. unjoin elements, delete instances)
                if f.HasResolutions():
                    try:
                        # Try default resolution first
                        fa.ResolveFailure(f)
                    except:
                        try:
                            # Try deleting the problematic element
                            ids = f.GetFailingElementIds()
                            if ids and ids.Count > 0:
                                fa.DeleteElements(ids)
                        except:
                            pass
        return FailureProcessingResult.Continue


# ==============================================================================
# WALL VALIDATION
# ==============================================================================

def is_valid_wall(wall):
    try:
        if not isinstance(wall, Wall):
            return False
        kind = wall.WallType.Kind
        if kind == DB.WallKind.Curtain or kind == DB.WallKind.Stacked:
            return False
        loc = wall.Location
        if loc is None or not hasattr(loc, 'Curve') or loc.Curve is None:
            return False
        bb = wall.get_BoundingBox(None)
        if bb is None:
            return False
        return True
    except:
        return False


# ==============================================================================
# STEP 1: PICK LINK
# ==============================================================================

def pick_link():
    log("STEP 1: Pick link")
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, LinkFilter(),
            "Select a Revit Link instance"
        )
        elem = doc.GetElement(ref.ElementId)
        if not isinstance(elem, RevitLinkInstance):
            return None, None
        link_doc = elem.GetLinkDocument()
        if link_doc is None:
            forms.alert("Link is not loaded.", exitscript=True)
            return None, None
        log("  Link: " + link_doc.Title)
        return elem, link_doc
    except:
        return None, None


# ==============================================================================
# STEP 2: GET WALLS
# ==============================================================================

def get_walls_from_view():
    log("STEP 2: Get walls from view")
    view = doc.ActiveView
    coll = FilteredElementCollector(doc, view.Id) \
        .OfCategory(BuiltInCategory.OST_Walls) \
        .WhereElementIsNotElementType()
    walls = []
    for eid in list(coll.ToElementIds()):
        try:
            w = doc.GetElement(eid)
            if is_valid_wall(w):
                walls.append(w)
        except:
            pass
    log("  Valid walls: {}".format(len(walls)))
    return walls


def pick_walls():
    log("STEP 2: Pick walls")
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, WallFilter(),
            "Select walls, then click Finish"
        )
        walls = []
        if refs:
            for r in refs:
                try:
                    w = doc.GetElement(r.ElementId)
                    if is_valid_wall(w):
                        walls.append(w)
                except:
                    pass
        log("  Picked {} walls".format(len(walls)))
        return walls
    except:
        return []


# ==============================================================================
# STEP 3: COLLECT LINK ELEMENT BBOXES (safe - no geometry)
# ==============================================================================

def collect_link_bboxes(link_doc, link_inst, categories):
    """Collect element data from link.
    For line-based elements (beams, pipes, ducts): collect LocationCurve endpoints
    + cross-section dimensions for precise line-plane intersection.
    For all elements: collect bbox as fallback."""
    log("STEP 3: Collect link element data")
    link_transform = link_inst.GetTotalTransform()
    all_data = []

    # Categories that typically have LocationCurve (line-based)
    LINE_BASED_CATS = {
        "Structural Framing",
        "Ducts", "Pipes", "Cable Trays", "Conduits",
    }

    for cat_name, bic in categories:
        try:
            coll = FilteredElementCollector(link_doc) \
                .OfCategory(bic) \
                .WhereElementIsNotElementType()
            eids = list(coll.ToElementIds())
            log("  {} : {} elements".format(cat_name, len(eids)))
            is_line_cat = cat_name in LINE_BASED_CATS

            for eid in eids:
                try:
                    elem = link_doc.GetElement(eid)
                    if elem is None:
                        continue

                    bb = elem.get_BoundingBox(None)
                    if bb is None:
                        continue

                    mn_x, mn_y, mn_z = bb.Min.X, bb.Min.Y, bb.Min.Z
                    mx_x, mx_y, mx_z = bb.Max.X, bb.Max.Y, bb.Max.Z
                    dx = mx_x - mn_x
                    dy = mx_y - mn_y
                    dz = mx_z - mn_z
                    if dx < 0.001 and dy < 0.001 and dz < 0.001:
                        continue

                    # Transform bbox corners to host
                    cxs, cys, czs = [], [], []
                    for xf in [0, 1]:
                        for yf in [0, 1]:
                            for zf in [0, 1]:
                                gpt = link_transform.OfPoint(XYZ(
                                    mn_x + xf * dx,
                                    mn_y + yf * dy,
                                    mn_z + zf * dz
                                ))
                                cxs.append(gpt.X)
                                cys.append(gpt.Y)
                                czs.append(gpt.Z)

                    fam = ""
                    try:
                        p = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
                        if p:
                            fam = p.AsValueString() or ""
                    except:
                        pass

                    entry = {
                        'eid': _eid_int(eid),
                        'cat': cat_name,
                        'fam': fam,
                        'xmin': min(cxs), 'xmax': max(cxs),
                        'ymin': min(cys), 'ymax': max(cys),
                        'zmin': min(czs), 'zmax': max(czs),
                        'has_line': False,
                    }

                    # ── Try to get LocationCurve for line-based elements ──
                    if is_line_cat:
                        try:
                            loc = elem.Location
                            if loc and hasattr(loc, 'Curve') and loc.Curve is not None:
                                curve = loc.Curve
                                lp0 = curve.GetEndPoint(0)
                                lp1 = curve.GetEndPoint(1)
                                gp0 = link_transform.OfPoint(lp0)
                                gp1 = link_transform.OfPoint(lp1)

                                entry['has_line'] = True
                                entry['lp0_x'] = gp0.X
                                entry['lp0_y'] = gp0.Y
                                entry['lp0_z'] = gp0.Z
                                entry['lp1_x'] = gp1.X
                                entry['lp1_y'] = gp1.Y
                                entry['lp1_z'] = gp1.Z

                                # ── Cross-section from bbox + beam direction ──
                                bb_dx = max(cxs) - min(cxs)
                                bb_dy = max(cys) - min(cys)
                                bb_dz = max(czs) - min(czs)

                                beam_dx = gp1.X - gp0.X
                                beam_dy = gp1.Y - gp0.Y
                                beam_dz = gp1.Z - gp0.Z
                                beam_xy_len = (beam_dx**2 + beam_dy**2) ** 0.5

                                # Width: solve from bbox using beam direction
                                # bbox_x = L_xy * |cos(a)| + sec_w * |sin(a)|
                                # bbox_y = L_xy * |sin(a)| + sec_w * |cos(a)|
                                # where a = beam angle in XY plane
                                sec_w = 0
                                sec_h = 0

                                if beam_xy_len > 0.1:
                                    bd_nx = abs(beam_dx / beam_xy_len)
                                    bd_ny = abs(beam_dy / beam_xy_len)

                                    # Use the axis where beam has less component
                                    # for better numerical stability
                                    if bd_ny > 0.05:
                                        sec_w = (bb_dx - beam_xy_len * bd_nx) / bd_ny
                                    if bd_nx > 0.05:
                                        sec_w2 = (bb_dy - beam_xy_len * bd_ny) / bd_nx
                                        if sec_w <= 0:
                                            sec_w = sec_w2
                                        else:
                                            # Average of both estimates
                                            sec_w = (sec_w + sec_w2) / 2.0

                                    # Height: bbox Z minus beam slope
                                    sec_h = bb_dz - abs(beam_dz)
                                    if sec_h < 0.01:
                                        sec_h = bb_dz
                                else:
                                    # Nearly vertical beam
                                    sec_w = max(bb_dx, bb_dy)
                                    sec_h = max(bb_dx, bb_dy)

                                # Sanity check
                                if sec_w < 0.01:
                                    sec_w = min(bb_dx, bb_dy)
                                if sec_h < 0.01:
                                    sec_h = bb_dz

                                entry['sec_w'] = sec_w
                                entry['sec_h'] = sec_h
                        except:
                            pass

                    all_data.append(entry)
                except:
                    continue
        except:
            continue

    line_count = sum(1 for d in all_data if d['has_line'])
    log("  Total: {} elements ({} with LocationCurve)".format(
        len(all_data), line_count))
    return all_data


# ==============================================================================
# STEP 4: INTERSECTION CHECK (pure math - no Revit API)
# ==============================================================================

def find_intersections(walls, elem_data):
    """Hybrid intersection detection:
    - Line-based elements (has_line=True): precise line-plane intersection
      + cross-section projection
    - Others: bbox overlap on wall plane (original method)
    """
    log("STEP 4: Check intersections (hybrid)")
    results = []

    for wall in walls:
        try:
            curve = wall.Location.Curve
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            ddx = p1.X - p0.X
            ddy = p1.Y - p0.Y
            w_len = (ddx * ddx + ddy * ddy) ** 0.5
            if w_len < 0.001:
                continue
            wd_x = ddx / w_len
            wd_y = ddy / w_len
            wn_x = wall.Orientation.X
            wn_y = wall.Orientation.Y
            bb = wall.get_BoundingBox(None)
            base_z = bb.Min.Z
            top_z = bb.Max.Z
            try:
                w_width = wall.Width
            except:
                w_width = 0.5

            wid = _eid_int(wall.Id)
            ox, oy = p0.X, p0.Y

            # Quick reject bounds
            qr_xmin = bb.Min.X - 2.0
            qr_xmax = bb.Max.X + 2.0
            qr_ymin = bb.Min.Y - 2.0
            qr_ymax = bb.Max.Y + 2.0
            qr_zmin = bb.Min.Z - 2.0
            qr_zmax = bb.Max.Z + 2.0

        except:
            continue

        for ed in elem_data:
            # Quick AABB rejection
            if ed['xmax'] < qr_xmin or ed['xmin'] > qr_xmax:
                continue
            if ed['ymax'] < qr_ymin or ed['ymin'] > qr_ymax:
                continue
            if ed['zmax'] < qr_zmin or ed['zmin'] > qr_zmax:
                continue

            result = None
            method_used = "none"

            # ── Method A: Line-plane intersection (precise) ──
            if ed.get('has_line') and ed.get('sec_w', 0) > 0 and ed.get('sec_h', 0) > 0:
                result = _line_plane_intersection(
                    ed, ox, oy, wd_x, wd_y, wn_x, wn_y,
                    w_len, w_width, base_z, top_z
                )
                if result:
                    method_used = "line"

            # ── Method B: BBox overlap (fallback) ──
            if result is None:
                result = _bbox_intersection(
                    ed, ox, oy, wd_x, wd_y, wn_x, wn_y,
                    w_len, w_width, base_z, top_z
                )
                if result:
                    method_used = "bbox"

            if result is None:
                continue

            ov_a0, ov_a1, ov_z0, ov_z1 = result

            # Skip tiny
            if ft_to_mm(ov_a1 - ov_a0) < 10:
                continue
            if ft_to_mm(ov_z1 - ov_z0) < 10:
                continue

            results.append({
                'wall': wall,
                'wid': wid,
                'eid': ed['eid'],
                'cat': ed['cat'],
                'fam': ed['fam'],
                'a0': ov_a0, 'a1': ov_a1,
                'z0': ov_z0, 'z1': ov_z1,
                'w_dir_x': wd_x, 'w_dir_y': wd_y,
                'w_ox': ox, 'w_oy': oy,
                'w_len': w_len,
                'w_base_z': base_z,
                'w_top_z': top_z,
            })
            log("    eid={} [{}] {}x{}mm method={}".format(
                ed['eid'], ed['cat'],
                str(int(round(ft_to_mm(ov_a1 - ov_a0)))),
                str(int(round(ft_to_mm(ov_z1 - ov_z0)))),
                method_used
            ))
            if ed.get('has_line'):
                log("      sec_w={}mm sec_h={}mm".format(
                    str(int(round(ft_to_mm(ed.get('sec_w', 0))))),
                    str(int(round(ft_to_mm(ed.get('sec_h', 0)))))
                ))

    log("  Found {} intersections".format(len(results)))
    return results


def _line_plane_intersection(ed, ox, oy, wd_x, wd_y, wn_x, wn_y,
                              w_len, w_width, base_z, top_z):
    """
    Precise intersection using:
    - LocationCurve direction for ANGLE (opening size calculation)
    - BBox center for POSITION (accounts for beam offsets/justification)
    Returns (a0, a1, z0, z1) or None.
    """
    # Beam endpoints (in host coords) - for direction only
    bx0, by0, bz0 = ed['lp0_x'], ed['lp0_y'], ed['lp0_z']
    bx1, by1, bz1 = ed['lp1_x'], ed['lp1_y'], ed['lp1_z']

    # Beam direction (from LocationCurve)
    bdx = bx1 - bx0
    bdy = by1 - by0
    bdz = bz1 - bz0
    beam_len = (bdx**2 + bdy**2 + bdz**2) ** 0.5
    if beam_len < 0.01:
        return None

    # Check beam crosses wall plane (not parallel)
    denom = bdx * wn_x + bdy * wn_y
    if abs(denom) < 0.0001:
        return None  # parallel to wall

    # ── Use BBox CENTER projected along BEAM DIRECTION to wall plane ──
    # (accounts for y/z offsets, justification, AND angled beam position)
    bbox_cx = (ed['xmin'] + ed['xmax']) / 2.0
    bbox_cy = (ed['ymin'] + ed['ymax']) / 2.0
    bbox_cz = (ed['zmin'] + ed['zmax']) / 2.0

    # Distance from bbox center to wall plane (along wall normal)
    dist_to_wall = (bbox_cx - ox) * wn_x + (bbox_cy - oy) * wn_y

    # Check bbox is near wall
    max_reach = w_width / 2.0 + beam_len * 0.6
    if abs(dist_to_wall) > max_reach:
        return None

    # Find where line through bbox center IN BEAM DIRECTION crosses wall plane
    # Line: P = bbox_center + t * beam_dir
    # Wall plane: dot(P - wall_origin, wall_normal) = 0
    # t = -dist_to_wall / denom
    t = -dist_to_wall / denom
    ix = bbox_cx + t * bdx
    iy = bbox_cy + t * bdy
    iz = bbox_cz + t * bdz  # Z also adjusts for sloped beams

    # Project intersection point onto wall axes
    vx = ix - ox
    vy = iy - oy
    along_pos = vx * wd_x + vy * wd_y

    # Check within wall length
    if along_pos < -0.5 or along_pos > w_len + 0.5:
        return None

    # Check Z within wall height (use bbox center Z)
    if iz < base_z - 1.0 or iz > top_z + 1.0:
        return None

    # ── Calculate opening size from cross-section + angle ──
    sec_w = ed['sec_w']
    sec_h = ed['sec_h']

    beam_xy_len = (bdx**2 + bdy**2) ** 0.5
    if beam_xy_len < 0.001:
        half_a = sec_w / 2.0
        half_z = sec_h / 2.0
    else:
        beam_xy_dx = bdx / beam_xy_len
        beam_xy_dy = bdy / beam_xy_len
        cos_theta = abs(beam_xy_dx * wn_x + beam_xy_dy * wn_y)
        sin_theta = abs(beam_xy_dx * wd_x + beam_xy_dy * wd_y)
        if cos_theta < 0.01:
            cos_theta = 0.01

        # Opening width = beam_width / cos(theta) + wall_thickness * tan(theta)
        projected_w = sec_w / cos_theta + w_width * sin_theta / cos_theta

        if projected_w > w_len * 0.5:
            return None  # too wide, fallback to bbox

        half_a = projected_w / 2.0

        # Vertical: beam height + slope contribution
        sin_slope = abs(bdz) / beam_len
        slope_extra = w_width * sin_slope / cos_theta
        half_z = sec_h / 2.0 + slope_extra / 2.0

    # Opening rectangle centered at bbox center projected onto wall
    ov_a0 = along_pos - half_a
    ov_a1 = along_pos + half_a
    ov_z0 = iz - half_z
    ov_z1 = iz + half_z

    # Clamp to wall bounds
    ov_a0 = max(ov_a0, 0.0)
    ov_a1 = min(ov_a1, w_len)
    ov_z0 = max(ov_z0, base_z)
    ov_z1 = min(ov_z1, top_z)

    if ov_a1 - ov_a0 < 0.003 or ov_z1 - ov_z0 < 0.003:
        return None

    return (ov_a0, ov_a1, ov_z0, ov_z1)


def _bbox_intersection(ed, ox, oy, wd_x, wd_y, wn_x, wn_y,
                        w_len, w_width, base_z, top_z):
    """
    Original bbox overlap method (fallback for elements without LocationCurve).
    Returns (a0, a1, z0, z1) or None.
    """
    along_vals = []
    normal_vals = []
    for xf in [0, 1]:
        for yf in [0, 1]:
            px = ed['xmin'] + xf * (ed['xmax'] - ed['xmin'])
            py = ed['ymin'] + yf * (ed['ymax'] - ed['ymin'])
            vx = px - ox
            vy = py - oy
            along_vals.append(vx * wd_x + vy * wd_y)
            normal_vals.append(vx * wn_x + vy * wn_y)

    half_w = w_width / 2.0 + 0.05
    if min(normal_vals) > half_w or max(normal_vals) < -half_w:
        return None

    ov_a0 = max(min(along_vals), 0.0)
    ov_a1 = min(max(along_vals), w_len)
    if ov_a1 - ov_a0 < 0.033:
        return None

    ov_z0 = max(ed['zmin'], base_z)
    ov_z1 = min(ed['zmax'], top_z)
    if ov_z1 - ov_z0 < 0.033:
        return None

    return (ov_a0, ov_a1, ov_z0, ov_z1)


# ==============================================================================
# STEP 5A: CREATE WALL OPENINGS (NewOpening)
# ==============================================================================

def apply_wall_openings(results, offset_ft):
    log("STEP 5A: Creating {} wall openings".format(len(results)))
    ok, fail = 0, 0

    tg = TransactionGroup(doc, "DQT - Wall Profile Cut")
    tg.Start()

    for r in results:
        t = Transaction(doc, "DQT - Wall Opening")
        fo = t.GetFailureHandlingOptions()
        fo.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(fo)
        t.Start()
        try:
            dx, dy = r['w_dir_x'], r['w_dir_y']
            pt_min = XYZ(
                r['w_ox'] + dx * r['a0'] - dx * offset_ft,
                r['w_oy'] + dy * r['a0'] - dy * offset_ft,
                r['z0'] - offset_ft
            )
            pt_max = XYZ(
                r['w_ox'] + dx * r['a1'] + dx * offset_ft,
                r['w_oy'] + dy * r['a1'] + dy * offset_ft,
                r['z1'] + offset_ft
            )
            opening = doc.Create.NewOpening(r['wall'], pt_min, pt_max)
            if opening:
                try:
                    cp = opening.get_Parameter(
                        BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                    if cp and not cp.IsReadOnly:
                        cp.Set("DQT_WallProfileCut")
                except:
                    pass
                ok += 1
            else:
                fail += 1
            t.Commit()
        except Exception as ex:
            fail += 1
            log("  Opening FAIL: " + str(ex))
            if t.HasStarted():
                t.RollBack()

    tg.Assimilate()
    log("  Result: {} OK, {} failed".format(ok, fail))
    return ok, fail


# ==============================================================================
# STEP 5B: EDIT WALL PROFILE (SketchEditScope)
# ==============================================================================

def apply_edit_profile(results, offset_ft):
    """Edit wall profiles. Returns (ok_count, fail_count, failed_results).
    failed_results contains the original result dicts that failed,
    so caller can retry with Wall Opening method."""
    log("STEP 5B: Edit wall profiles")

    # Group by wall, keeping reference to original results
    wall_map = {}
    for r in results:
        wid = r['wid']
        if wid not in wall_map:
            wall_map[wid] = {
                'wall': r['wall'],
                'ops': [],
                'results': [],  # keep original results for fallback
                'w_dir_x': r['w_dir_x'],
                'w_dir_y': r['w_dir_y'],
                'w_ox': r['w_ox'],
                'w_oy': r['w_oy'],
                'w_len': r['w_len'],
                'w_base_z': r['w_base_z'],
                'w_top_z': r['w_top_z'],
            }
        wall_map[wid]['ops'].append((r['a0'], r['a1'], r['z0'], r['z1']))
        wall_map[wid]['results'].append(r)

    ok, fail = 0, 0
    failed_results = []  # results that failed Edit Profile

    for wid, data in wall_map.items():
        wall = data['wall']
        w_len = data['w_len']
        base_z = data['w_base_z']
        top_z = data['w_top_z']

        # ── Check if wall is in a Group → cannot edit profile ──
        try:
            group_id = wall.GroupId
            in_group = _eid_int(group_id) > 0
        except:
            in_group = False

        if in_group:
            log("  Wall {}: in Group -> fallback to Wall Opening".format(wid))
            fail += len(data['ops'])
            failed_results.extend(data['results'])
            continue

        # Minimum distance from wall edge for sketch inner loop
        sketch_edge = mm_to_ft(0.5)  # 0.5mm

        # ── Apply offset and clamp ALL openings ──
        clamped = []
        for idx_op, (a0, a1, z0, z1) in enumerate(data['ops']):
            a0 -= offset_ft
            a1 += offset_ft
            z0 -= offset_ft
            z1 += offset_ft

            # Clamp to wall bounds with minimal edge
            a0 = max(a0, sketch_edge)
            a1 = min(a1, w_len - sketch_edge)
            z0 = max(z0, base_z + sketch_edge)
            z1 = min(z1, top_z - sketch_edge)

            if a1 - a0 > mm_to_ft(10) and z1 - z0 > mm_to_ft(10):
                clamped.append([a0, a1, z0, z1])

        if not clamped:
            log("  Wall {}: no valid openings after clamping".format(wid))
            fail += len(data['ops'])
            failed_results.extend(data['results'])
            continue

        # ── Merge overlapping rectangles ──
        merged = _merge_rects(clamped)
        log("  Wall {}: {} raw -> {} merged openings".format(
            wid, len(data['ops']), len(merged)))

        try:
            if not wall.CanHaveProfileSketch():
                log("    Cannot have profile sketch")
                fail += len(data['ops'])
                failed_results.extend(data['results'])
                continue
        except:
            fail += len(data['ops'])
            failed_results.extend(data['results'])
            continue

        # Create sketch if needed
        # ── Unjoin ALL geometry to prevent "Can't keep elements joined" ──
        joined_pairs = []
        t_unjoin = Transaction(doc, "DQT - Unjoin Wall")
        fo_uj = t_unjoin.GetFailureHandlingOptions()
        fo_uj.SetFailuresPreprocessor(WarningSwallower())
        t_unjoin.SetFailureHandlingOptions(fo_uj)
        t_unjoin.Start()
        try:
            # Unjoin wall ends (only if wall itself is not in group)
            try:
                DB.WallUtils.DisallowWallJoinAtEnd(wall, 0)
                DB.WallUtils.DisallowWallJoinAtEnd(wall, 1)
            except:
                pass

            # Unjoin ALL geometry joined to this wall
            # BUT skip elements that are in groups (can't modify group members)
            try:
                joined_ids = DB.JoinGeometryUtils.GetJoinedElements(doc, wall)
                if joined_ids:
                    for jid in joined_ids:
                        try:
                            other = doc.GetElement(jid)
                            if other is None:
                                continue
                            # Check if the OTHER element is in a group
                            other_in_group = False
                            try:
                                gid = other.GroupId
                                other_in_group = _eid_int(gid) > 0
                            except:
                                pass
                            if other_in_group:
                                log("    Skip unjoin with id={} (in group)".format(
                                    _eid_int(jid)))
                                continue
                            DB.JoinGeometryUtils.UnjoinGeometry(doc, wall, other)
                            joined_pairs.append(jid)
                        except:
                            pass
            except:
                pass

            t_unjoin.Commit()
            if joined_pairs:
                log("    Unjoined {} elements from wall".format(len(joined_pairs)))
        except:
            if t_unjoin.HasStarted():
                t_unjoin.RollBack()

        has_sketch = False
        try:
            has_sketch = _eid_int(wall.SketchId) > 0
        except:
            pass

        if not has_sketch:
            t1 = Transaction(doc, "DQT - Create Sketch")
            fo1 = t1.GetFailureHandlingOptions()
            fo1.SetFailuresPreprocessor(WarningSwallower())
            t1.SetFailureHandlingOptions(fo1)
            t1.Start()
            try:
                wall.CreateProfileSketch()
                doc.Regenerate()
                t1.Commit()
            except Exception as ex:
                log("    CreateProfileSketch FAIL: " + str(ex))
                if t1.HasStarted():
                    t1.RollBack()
                fail += len(data['ops'])
                failed_results.extend(data['results'])
                continue

        try:
            sketch = doc.GetElement(wall.SketchId)
        except:
            fail += len(data['ops'])
            failed_results.extend(data['results'])
            continue
        if sketch is None:
            fail += len(data['ops'])
            failed_results.extend(data['results'])
            continue

        ses = SketchEditScope(doc, "DQT - Edit Profile")
        try:
            ses.Start(sketch.Id)
        except Exception as ex:
            log("    SketchEditScope.Start FAIL: " + str(ex))
            fail += len(data['ops'])
            failed_results.extend(data['results'])
            continue

        t2 = Transaction(doc, "DQT - Add Openings")
        fo2 = t2.GetFailureHandlingOptions()
        fo2.SetFailuresPreprocessor(WarningSwallower())
        t2.SetFailureHandlingOptions(fo2)
        t2.Start()

        created = 0
        try:
            sp = sketch.SketchPlane
            sp_plane = sp.GetPlane()
            sp_o = sp_plane.Origin
            sp_n = sp_plane.Normal
            sn_x, sn_y, sn_z = sp_n.X, sp_n.Y, sp_n.Z
            so_x, so_y, so_z = sp_o.X, sp_o.Y, sp_o.Z

            dx = data['w_dir_x']
            dy = data['w_dir_y']
            
            # IMPORTANT: Use sketch plane origin projected onto wall direction
            # as the reference point for along-wall positions.
            # The intersection detection uses wall p0 (centerline endpoint 0)
            # but sketch plane may be offset. We need to compute the offset
            # between wall p0 and sketch plane origin along the wall direction.
            w_ox = data['w_ox']
            w_oy = data['w_oy']
            
            # Vector from wall p0 to sketch plane origin
            sp_offset_x = so_x - w_ox
            sp_offset_y = so_y - w_oy
            # Along-wall component of this offset
            along_offset = sp_offset_x * dx + sp_offset_y * dy
            
            # Debug: log coordinate system comparison
            log("    Wall p0=({},{}) dir=({},{})".format(
                str(round(w_ox, 3)), str(round(w_oy, 3)),
                str(round(dx, 4)), str(round(dy, 4))))
            log("    Sketch origin=({},{},{}) normal=({},{},{})".format(
                str(round(so_x, 3)), str(round(so_y, 3)), str(round(so_z, 3)),
                str(round(sn_x, 4)), str(round(sn_y, 4)), str(round(sn_z, 4))))
            log("    Along offset from p0 to sketch: {}mm".format(
                str(int(round(along_offset * 304.8)))))
            
            for a0, a1, z0, z1 in merged:
                # 4 corners on the SKETCH PLANE
                # a0, a1 are measured from wall p0 along wall direction
                # We generate points along wall direction from wall p0,
                # then project onto sketch plane
                
                # Points along wall centerline at correct along-wall positions
                p_a0_x = w_ox + dx * a0
                p_a0_y = w_oy + dy * a0
                p_a1_x = w_ox + dx * a1
                p_a1_y = w_oy + dy * a1
                
                # Create points and project to sketch plane
                raw_pts = [
                    XYZ(p_a0_x, p_a0_y, z0),
                    XYZ(p_a1_x, p_a1_y, z0),
                    XYZ(p_a1_x, p_a1_y, z1),
                    XYZ(p_a0_x, p_a0_y, z1),
                ]

                # Project onto sketch plane
                proj_pts = []
                for pt in raw_pts:
                    vx = pt.X - so_x
                    vy = pt.Y - so_y
                    vz = pt.Z - so_z
                    d = vx * sn_x + vy * sn_y + vz * sn_z
                    proj_pts.append(XYZ(
                        pt.X - sn_x * d,
                        pt.Y - sn_y * d,
                        pt.Z - sn_z * d
                    ))

                try:
                    doc.Create.NewModelCurve(
                        Line.CreateBound(proj_pts[0], proj_pts[1]), sp)
                    doc.Create.NewModelCurve(
                        Line.CreateBound(proj_pts[1], proj_pts[2]), sp)
                    doc.Create.NewModelCurve(
                        Line.CreateBound(proj_pts[2], proj_pts[3]), sp)
                    doc.Create.NewModelCurve(
                        Line.CreateBound(proj_pts[3], proj_pts[0]), sp)
                    created += 1
                    log("    Opening {}x{}mm at a={}-{}, z={}-{}".format(
                        str(int(round(ft_to_mm(a1 - a0)))),
                        str(int(round(ft_to_mm(z1 - z0)))),
                        str(int(round(ft_to_mm(a0)))),
                        str(int(round(ft_to_mm(a1)))),
                        str(int(round(ft_to_mm(z0 - base_z)))),
                        str(int(round(ft_to_mm(z1 - base_z))))
                    ))
                except Exception as ex:
                    log("    ModelCurve FAIL: " + str(ex))

            if created > 0:
                t2.Commit()
                ses.Commit(WarningSwallower())
                ok += created
                log("    {} openings created".format(created))
            else:
                t2.RollBack()
                ses.Dispose()
                fail += len(data['ops'])
                failed_results.extend(data['results'])

        except Exception as ex:
            log("    Edit FAIL: " + str(ex))
            if t2.HasStarted():
                t2.RollBack()
            try:
                ses.Dispose()
            except:
                pass
            fail += len(data['ops'])
            failed_results.extend(data['results'])

    log("  Result: {} OK, {} failed, {} for fallback".format(ok, fail, len(failed_results)))
    return ok, fail, failed_results


def _merge_rects(rects):
    """
    Merge TRULY overlapping 2D rectangles only.
    Each rect = [a0, a1, z0, z1].
    Only merge when rects actually overlap by more than a small threshold
    in BOTH dimensions. This prevents chain-merging beams that are merely
    adjacent into one giant opening.
    """
    if len(rects) <= 1:
        return rects

    # Minimum overlap required to merge (in feet) ~10mm
    MIN_OVERLAP = 0.033

    def truly_overlaps(r1, r2):
        """Check if two rectangles have significant overlap in BOTH axes"""
        # Overlap amount in along-wall direction
        a_overlap = min(r1[1], r2[1]) - max(r1[0], r2[0])
        if a_overlap < MIN_OVERLAP:
            return False
        # Overlap amount in vertical direction
        z_overlap = min(r1[3], r2[3]) - max(r1[2], r2[2])
        if z_overlap < MIN_OVERLAP:
            return False
        return True

    def union_rect(r1, r2):
        return [
            min(r1[0], r2[0]),
            max(r1[1], r2[1]),
            min(r1[2], r2[2]),
            max(r1[3], r2[3]),
        ]

    changed = True
    current = [list(r) for r in rects]

    # Limit iterations to prevent infinite loop
    max_iter = 20
    iteration = 0

    while changed and iteration < max_iter:
        changed = False
        iteration += 1
        new_list = []
        used = [False] * len(current)

        for i in range(len(current)):
            if used[i]:
                continue
            merged = list(current[i])
            for j in range(i + 1, len(current)):
                if used[j]:
                    continue
                if truly_overlaps(merged, current[j]):
                    merged = union_rect(merged, current[j])
                    used[j] = True
                    changed = True
            new_list.append(merged)
            used[i] = True

        current = new_list

    return current


# ==============================================================================
# STEP 5C: PLACE OPENING FAMILY INSTANCE
# ==============================================================================

class OpeningFamilyItem(object):
    """Item for family type selection list"""
    def __init__(self, symbol):
        self.symbol = symbol
        self.symbol_id = symbol.Id
        try:
            fam_name = symbol.Family.Name
        except:
            fam_name = "Unknown"
        try:
            type_name = symbol.get_Parameter(
                BuiltInParameter.ALL_MODEL_TYPE_NAME
            ).AsString() or symbol.Name
        except:
            type_name = symbol.Name
        self.name = "{} : {}".format(fam_name, type_name)

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.name


def get_opening_families():
    """Collect all wall-hosted opening/void family types in the project"""
    results = []

    # Collect from multiple categories that typically contain opening families
    cats = [
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Doors,
    ]

    for bic in cats:
        try:
            coll = FilteredElementCollector(doc) \
                .OfCategory(bic) \
                .OfClass(FamilySymbol)
            for sym in coll:
                try:
                    fam = sym.Family
                    if fam is None:
                        continue
                    # Check if wall-hosted
                    host_param = fam.get_Parameter(
                        BuiltInParameter.FAMILY_HOSTING_BEHAVIOR
                    )
                    if host_param:
                        host_val = host_param.AsInteger()
                        # 1 = Wall-hosted
                        if host_val == 1:
                            results.append(OpeningFamilyItem(sym))
                            continue

                    # Also check family placement type
                    try:
                        fp = fam.FamilyPlacementType
                        fp_str = str(fp)
                        if "Wall" in fp_str or "OneLevelBasedHost" in fp_str:
                            results.append(OpeningFamilyItem(sym))
                            continue
                    except:
                        pass

                    # Fallback: check family name for opening keywords
                    fam_name_lower = fam.Name.lower()
                    if any(kw in fam_name_lower for kw in [
                        "opening", "void", "cutout", "penetration",
                        "sleeve", "hole"
                    ]):
                        results.append(OpeningFamilyItem(sym))
                except:
                    continue
        except:
            continue

    # Deduplicate by symbol id
    seen = set()
    unique = []
    for item in results:
        sid = _eid_int(item.symbol_id)
        if sid not in seen:
            seen.add(sid)
            unique.append(item)

    return sorted(unique, key=lambda x: x.name)


def apply_place_family(results, offset_ft, family_symbol):
    """Place a wall-hosted opening family at each intersection location.
    Validates that opening fits within wall bounds before placing.
    Uses Sill Height + MoveElement for correct vertical positioning."""
    log("STEP 5C: Place opening family instances")
    ok, fail, skip = 0, 0, 0

    # Minimum margin from wall edges (in feet) ~50mm
    EDGE_MARGIN = mm_to_ft(50)
    # Minimum opening size (in feet) ~50mm
    MIN_SIZE = mm_to_ft(50)

    # Activate symbol if needed
    t_act = Transaction(doc, "DQT - Activate Symbol")
    t_act.Start()
    try:
        if not family_symbol.IsActive:
            family_symbol.Activate()
            doc.Regenerate()
        t_act.Commit()
    except:
        if t_act.HasStarted():
            t_act.RollBack()

    tg = TransactionGroup(doc, "DQT - Place Opening Families")
    tg.Start()

    for r in results:
        wall = r['wall']
        dx, dy = r['w_dir_x'], r['w_dir_y']
        w_len = r['w_len']
        base_z = r['w_base_z']
        top_z = r['w_top_z']

        # ── Clamp opening within wall bounds ──
        a0 = r['a0'] - offset_ft
        a1 = r['a1'] + offset_ft
        z0 = r['z0'] - offset_ft
        z1 = r['z1'] + offset_ft

        # Clamp to wall extents with margin
        a0 = max(a0, EDGE_MARGIN)
        a1 = min(a1, w_len - EDGE_MARGIN)
        z0 = max(z0, base_z + EDGE_MARGIN)
        z1 = min(z1, top_z - EDGE_MARGIN)

        # Check minimum size after clamping
        width_ft = a1 - a0
        height_ft = z1 - z0

        if width_ft < MIN_SIZE or height_ft < MIN_SIZE:
            skip += 1
            log("  SKIP wall {}: opening too small or at edge ({}x{}mm)".format(
                r['wid'],
                str(int(round(ft_to_mm(width_ft)))),
                str(int(round(ft_to_mm(height_ft))))
            ))
            continue

        # Along-wall center
        a_center = (a0 + a1) / 2.0
        # Vertical center
        z_center = (z0 + z1) / 2.0

        # Target center point on the wall face
        target_pt = XYZ(
            r['w_ox'] + dx * a_center,
            r['w_oy'] + dy * a_center,
            z_center
        )

        log("  Wall {} elem {}: target=({},{},{}), W={}mm H={}mm".format(
            r['wid'], r['eid'],
            str(round(target_pt.X, 2)),
            str(round(target_pt.Y, 2)),
            str(round(target_pt.Z, 2)),
            str(int(round(ft_to_mm(width_ft)))),
            str(int(round(ft_to_mm(height_ft))))
        ))

        t = Transaction(doc, "DQT - Place Opening")
        fo = t.GetFailureHandlingOptions()
        fo.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(fo)
        t.Start()
        try:
            inst = None

            # ── Method A: Place on wall face Reference (best for position) ──
            try:
                face_refs = HostObjectUtils.GetSideFaces(
                    wall, ShellLayerType.Exterior
                )
                if face_refs and face_refs.Count > 0:
                    face_ref = face_refs[0]
                    ref_dir = XYZ(0, 0, 0)
                    inst = doc.Create.NewFamilyInstance(
                        face_ref, target_pt, ref_dir, family_symbol
                    )
                    if inst:
                        log("    Placed via face Reference (Method A)")
            except Exception as ex:
                log("    Method A failed: " + str(ex))
                inst = None

            # ── Method B: Place with Level (wall-hosted families) ──
            if inst is None:
                try:
                    wall_level_id = wall.LevelId
                    wall_level = doc.GetElement(wall_level_id)
                    if wall_level and isinstance(wall_level, Level):
                        inst = doc.Create.NewFamilyInstance(
                            target_pt, family_symbol, wall,
                            wall_level,
                            DB.Structure.StructuralType.NonStructural
                        )
                        if inst:
                            log("    Placed via Level overload (Method B)")
                except Exception as ex:
                    log("    Method B failed: " + str(ex))
                    inst = None

            # ── Method C: Basic overload + MoveElement ──
            if inst is None:
                try:
                    inst = doc.Create.NewFamilyInstance(
                        target_pt, family_symbol, wall,
                        DB.Structure.StructuralType.NonStructural
                    )
                    if inst:
                        log("    Placed via basic overload (Method C)")
                except Exception as ex:
                    log("    Method C failed: " + str(ex))
                    inst = None

            if inst is None:
                fail += 1
                log("    FAIL: all placement methods failed")
                t.RollBack()
                continue

            # ── Set Width ──
            _try_set_dim(inst, width_ft, [
                "Width", "Opening Width", "w", "W",
                "width", "opening width",
                BuiltInParameter.FAMILY_WIDTH_PARAM,
                BuiltInParameter.GENERIC_WIDTH,
                BuiltInParameter.FURNITURE_WIDTH,
                BuiltInParameter.DOOR_WIDTH,
                BuiltInParameter.WINDOW_WIDTH,
                BuiltInParameter.CASEWORK_WIDTH,
            ])

            # ── Set Height ──
            _try_set_dim(inst, height_ft, [
                "Height", "Opening Height", "h", "H",
                "height", "opening height",
                BuiltInParameter.FAMILY_HEIGHT_PARAM,
                BuiltInParameter.GENERIC_HEIGHT,
                BuiltInParameter.DOOR_HEIGHT,
                BuiltInParameter.WINDOW_HEIGHT,
                BuiltInParameter.CASEWORK_HEIGHT,
            ])

            # ── Verify & correct position ──
            doc.Regenerate()
            try:
                inst_bb = inst.get_BoundingBox(None)
                if inst_bb:
                    act_cx = (inst_bb.Min.X + inst_bb.Max.X) / 2.0
                    act_cy = (inst_bb.Min.Y + inst_bb.Max.Y) / 2.0
                    act_cz = (inst_bb.Min.Z + inst_bb.Max.Z) / 2.0
                    mdx = target_pt.X - act_cx
                    mdy = target_pt.Y - act_cy
                    mdz = target_pt.Z - act_cz
                    if abs(mdx) > 0.01 or abs(mdy) > 0.01 or abs(mdz) > 0.01:
                        DB.ElementTransformUtils.MoveElement(
                            doc, inst.Id, XYZ(mdx, mdy, mdz)
                        )
                        log("    Corrected: dX={}mm dY={}mm dZ={}mm".format(
                            str(int(round(ft_to_mm(mdx)))),
                            str(int(round(ft_to_mm(mdy)))),
                            str(int(round(ft_to_mm(mdz))))
                        ))
            except Exception as ex:
                log("    Position verify error: " + str(ex))

            # ── Tag with DQT comment ──
            try:
                cp = inst.get_Parameter(
                    BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                if cp and not cp.IsReadOnly:
                    cp.Set("DQT_WallProfileCut")
            except:
                pass

            t.Commit()

            if t.GetStatus() == DB.TransactionStatus.Committed:
                ok += 1
                log("    OK")
            else:
                fail += 1
                log("    FAIL: transaction not committed")

        except Exception as ex:
            fail += 1
            log("    FAIL: " + str(ex))
            if t.HasStarted():
                t.RollBack()

    tg.Assimilate()
    log("  Result: {} OK, {} failed, {} skipped".format(ok, fail, skip))
    return ok, fail + skip


def _try_set_dim(inst, value_ft, param_names):
    """Try to set a dimension parameter by various names"""
    for pn in param_names:
        try:
            if isinstance(pn, str):
                p = inst.LookupParameter(pn)
            else:
                # BuiltInParameter
                p = inst.get_Parameter(pn)
            if p and not p.IsReadOnly:
                p.Set(value_ft)
                return True
        except:
            continue
    return False


# ==============================================================================
# DISPLAY ITEM for forms.SelectFromList
# ==============================================================================

class IntersectionItem(object):
    """Item for SelectFromList display"""
    def __init__(self, idx, result):
        self.idx = idx
        self.result = result
        w_mm = int(round(ft_to_mm(result['a1'] - result['a0'])))
        h_mm = int(round(ft_to_mm(result['z1'] - result['z0'])))
        self.name = "Wall {} | {} [{}] | {}x{}mm".format(
            result['wid'],
            result['cat'],
            result['eid'],
            w_mm, h_mm
        )
        if result['fam']:
            self.name += " | " + result['fam']

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.name


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    log("")
    log("=" * 60)
    log("Wall Profile Cut v4 - START")
    log("=" * 60)

    # ── Step 1: Pick link ──
    link_inst, link_doc = pick_link()
    if link_inst is None:
        return

    # ── Step 2: Select categories ──
    sel_cats = forms.SelectFromList.show(
        sorted(CATEGORIES.keys()),
        title="DQT - Select Categories from Link",
        multiselect=True,
        button_name="Next"
    )
    if not sel_cats:
        return

    cats = [(name, CATEGORIES[name]) for name in sel_cats]

    # ── Step 3: Select walls ──
    wall_mode = forms.CommandSwitchWindow.show(
        ["All Walls in Active View", "Pick Walls Manually"],
        message="How to select walls?"
    )
    if wall_mode is None:
        return

    if wall_mode == "Pick Walls Manually":
        walls = pick_walls()
    else:
        walls = get_walls_from_view()

    if not walls:
        forms.alert("No valid walls found.", exitscript=True)
        return

    # ── Step 4: Collect link elements & find intersections ──
    with forms.ProgressBar(title="Scanning link elements...") as pb:
        pb.update_progress(10, 100)
        elem_data = collect_link_bboxes(link_doc, link_inst, cats)
        pb.update_progress(50, 100)

        if not elem_data:
            forms.alert("No elements found in selected categories.")
            return

        results = find_intersections(walls, elem_data)
        pb.update_progress(100, 100)

    if not results:
        forms.alert(
            "No intersections found.\n\n"
            "Checked {} elements against {} walls.".format(
                len(elem_data), len(walls))
        )
        return

    # ── Step 5: Show results for selection ──
    items = [IntersectionItem(i, r) for i, r in enumerate(results)]

    selected = forms.SelectFromList.show(
        items,
        title="DQT - Found {} Intersections on {} Walls".format(
            len(results),
            len(set(r['wid'] for r in results))
        ),
        multiselect=True,
        name_attr='name',
        button_name="Create Openings ({})".format(len(items))
    )

    if not selected:
        return

    sel_results = [s.result for s in selected]
    log("Selected {} intersections to create".format(len(sel_results)))

    # ── Step 6: Choose method & offset ──
    method = forms.CommandSwitchWindow.show(
        ["Wall Opening (Recommended)",
         "Edit Wall Profile",
         "Place Opening Family"],
        message="Opening method:"
    )
    if method is None:
        return

    # If Place Opening Family, select the family type first
    sel_symbol = None
    if method == "Place Opening Family":
        fam_items = get_opening_families()
        if not fam_items:
            forms.alert(
                "No wall-hosted opening families found in project.\n\n"
                "Load an opening/void family first,\n"
                "then try again.",
                exitscript=True
            )
            return

        sel_fam = forms.SelectFromList.show(
            fam_items,
            title="DQT - Select Opening Family Type",
            multiselect=False,
            name_attr='name',
            button_name="Use This Family"
        )
        if not sel_fam:
            return
        sel_symbol = sel_fam.symbol

    offset_str = forms.ask_for_string(
        prompt="Offset (mm) - buffer around each element:",
        default="25",
        title="DQT - Offset"
    )
    if offset_str is None:
        return
    try:
        offset_mm = float(offset_str)
    except:
        offset_mm = 25.0
    offset_ft = mm_to_ft(offset_mm)

    # ── Step 7: Apply ──
    if method == "Edit Wall Profile":
        ok, fail, failed_results = apply_edit_profile(sel_results, offset_ft)
        # Auto fallback: retry failed with Wall Opening
        if failed_results:
            log("Fallback: {} results failed Edit Profile, trying Wall Opening".format(
                len(failed_results)))
            ok2, fail2 = apply_wall_openings(failed_results, offset_ft)
            ok += ok2
            fail = fail - len(failed_results) + fail2
            if ok2 > 0:
                method = "Edit Wall Profile + Wall Opening (fallback)"
    elif method == "Place Opening Family":
        ok, fail = apply_place_family(sel_results, offset_ft, sel_symbol)
    else:
        ok, fail = apply_wall_openings(sel_results, offset_ft)

    # ── Report ──
    msg = "Created {} opening(s) successfully.".format(ok)
    if fail > 0:
        msg += "\n{} failed.".format(fail)

    output.print_md("## DQT - Wall Profile Cut Results")
    output.print_md("- **Link:** {}".format(link_doc.Title))
    output.print_md("- **Categories:** {}".format(", ".join(sel_cats)))
    output.print_md("- **Walls scanned:** {}".format(len(walls)))
    output.print_md("- **Intersections found:** {}".format(len(results)))
    output.print_md("- **Openings created:** {}".format(ok))
    if fail > 0:
        output.print_md("- **Failed:** {}".format(fail))
    output.print_md("- **Method:** {}".format(method))
    output.print_md("- **Offset:** {} mm".format(offset_mm))

    forms.alert(msg)
    log("=== DONE: {} OK, {} fail ===".format(ok, fail))


# ==============================================================================

if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log("FATAL: " + traceback.format_exc())
        forms.alert("Error: {}\n\nLog: {}".format(str(ex), LOG_PATH))