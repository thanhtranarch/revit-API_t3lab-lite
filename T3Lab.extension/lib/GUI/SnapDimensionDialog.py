from System.Windows import WindowState
# -*- coding: utf-8 -*-
"""Snap to Grid v12 - Round wall/column/beam distances to nearest gridline.

Added: Scope option - scan all elements or only current selection.

Copyright (c) 2026 by Dang Quoc Truong (DQT)
"""

__title__ = "Snap to\nGrid"
__author__ = "DQT"
__doc__ = "Round wall/column/beam offset from grid to whole millimeters"

import clr
import math

clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")
clr.AddReference("System.Data")

from System.IO import MemoryStream
from System.Text import Encoding
from System.Windows.Markup import XamlReader
from System.Data import DataTable
from System.Collections.Generic import List
from collections import OrderedDict

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction,
    XYZ, ElementId, Grid, Wall, FamilyInstance,
    LocationCurve, LocationPoint, ElementTransformUtils,
)
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

FEET_TO_MM = 304.8
SNAP_TOL = 0.0001


def get_id_value(eid):
    try:
        return eid.IntegerValue
    except:
        return eid.Value


def round_to(v, p):
    if p <= 0:
        return v
    return round(v / p) * p


def line_dir_2d(p0, p1):
    dx, dy = p1.X - p0.X, p1.Y - p0.Y
    ln = math.sqrt(dx * dx + dy * dy)
    if ln < 1e-12:
        return None
    return XYZ(dx / ln, dy / ln, 0)


def is_parallel(d1, d2):
    if d1 is None or d2 is None:
        return False
    return abs(d1.X * d2.X + d1.Y * d2.Y) >= math.cos(math.radians(5.0))


def is_perpendicular(d1, d2):
    if d1 is None or d2 is None:
        return False
    return abs(d1.X * d2.X + d1.Y * d2.Y) <= math.sin(math.radians(5.0))


def signed_perp(origin, direction, point):
    vx = point.X - origin.X
    vy = point.Y - origin.Y
    return vx * direction.Y - vy * direction.X


def group_grids_by_direction(grids_info):
    groups = []
    for g, go, gd in grids_info:
        placed = False
        for group in groups:
            if is_parallel(group[0][2], gd):
                group.append((g, go, gd))
                placed = True
                break
        if not placed:
            groups.append([(g, go, gd)])
    return groups


def get_line_elem_half_width(elem):
    if isinstance(elem, Wall):
        try:
            return doc.GetElement(elem.GetTypeId()).Width / 2.0
        except:
            return 0.0
    try:
        etype = doc.GetElement(elem.GetTypeId())
        if etype:
            for bip in [DB.BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH,
                        DB.BuiltInParameter.FAMILY_WIDTH_PARAM]:
                try:
                    p = etype.get_Parameter(bip)
                    if p and p.HasValue and p.AsDouble() > 0:
                        return p.AsDouble() / 2.0
                except:
                    pass
            for bip in [DB.BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH]:
                try:
                    p = elem.get_Parameter(bip)
                    if p and p.HasValue and p.AsDouble() > 0:
                        return p.AsDouble() / 2.0
                except:
                    pass
            for name in ["b", "B", "Width", "W", "bf", "Bf"]:
                try:
                    p = etype.LookupParameter(name)
                    if p and p.HasValue and p.AsDouble() > 0:
                        return p.AsDouble() / 2.0
                except:
                    pass
    except:
        pass
    return 0.0


def calc_snap(sc, half_w, grid_dir, precision):
    perp = XYZ(grid_dir.Y, -grid_dir.X, 0)
    s = 1.0 if sc >= 0 else -1.0
    center_mm = abs(sc) * FEET_TO_MM

    if half_w > 0:
        nf_signed = sc - s * half_w
        ff_signed = sc + s * half_w
        nf_mm = abs(nf_signed) * FEET_TO_MM
        ff_mm = abs(ff_signed) * FEET_TO_MM

        candidates = []
        for ref_signed, ref_mm in [(nf_signed, nf_mm),
                                    (ff_signed, ff_mm),
                                    (sc, center_mm)]:
            target_mm = round_to(ref_mm, precision)
            delta = target_mm - ref_mm
            if abs(delta) >= SNAP_TOL:
                candidates.append((ref_mm, target_mm, delta, ref_signed))

        if not candidates:
            return None

        candidates.sort(key=lambda c: abs(c[2]))
        shown_mm, snapped_mm, delta_mm, ref_signed = candidates[0]

        s_ref = 1.0 if ref_signed >= 0 else -1.0
        target_ref = s_ref * (snapped_mm / FEET_TO_MM)
        offset = ref_signed - sc
        target_sc = target_ref - offset
        move_ft = target_sc - sc
    else:
        target_mm = round_to(center_mm, precision)
        delta_mm = target_mm - center_mm
        if abs(delta_mm) < SNAP_TOL:
            return None
        shown_mm = center_mm
        snapped_mm = target_mm
        target_sc = s * (target_mm / FEET_TO_MM)
        move_ft = target_sc - sc

    move_vec = XYZ(perp.X * move_ft, perp.Y * move_ft, 0)
    return (shown_mm, snapped_mm, delta_mm, move_vec)


def calc_snap_endpoint(endpoint, grid_origin, grid_dir, precision):
    sd = signed_perp(grid_origin, grid_dir, endpoint)
    dist_mm = abs(sd) * FEET_TO_MM
    target_mm = round_to(dist_mm, precision)
    delta_mm = target_mm - dist_mm

    if abs(delta_mm) < SNAP_TOL:
        return None

    perp = XYZ(grid_dir.Y, -grid_dir.X, 0)
    s = 1.0 if sd >= 0 else -1.0
    target_sd = s * (target_mm / FEET_TO_MM)
    move_ft = target_sd - sd
    move_vec = XYZ(perp.X * move_ft, perp.Y * move_ft, 0)
    return (dist_mm, target_mm, delta_mm, move_vec)


# ==============================================================================
# ANALYSIS
# ==============================================================================
def analyze_wall(elem, grids_info, precision):
    loc = elem.Location
    if not isinstance(loc, LocationCurve):
        return []
    curve = loc.Curve
    mid = curve.Evaluate(0.5, True)
    elem_dir = line_dir_2d(curve.GetEndPoint(0), curve.GetEndPoint(1))
    if elem_dir is None:
        return []
    half_w = get_line_elem_half_width(elem)

    best_sd = None
    best_gdir = None
    best_gname = ""
    best_abs = float("inf")
    for g, go, gd in grids_info:
        if not is_parallel(elem_dir, gd):
            continue
        sd = signed_perp(go, gd, mid)
        if abs(sd) < best_abs:
            best_abs = abs(sd)
            best_sd = sd
            best_gdir = gd
            try:
                best_gname = DB.Element.Name.GetValue(g)
            except:
                best_gname = "?"
    if best_gdir is None:
        return []
    r = calc_snap(best_sd, half_w, best_gdir, precision)
    if not r:
        return []
    shown, snapped, delta, mv = r
    return [(best_gname, shown, snapped, delta, mv)]


def analyze_beam(elem, grids_info, precision):
    loc = elem.Location
    if not isinstance(loc, LocationCurve):
        return []
    curve = loc.Curve
    mid = curve.Evaluate(0.5, True)
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    elem_dir = line_dir_2d(p0, p1)
    if elem_dir is None:
        return []
    half_w = get_line_elem_half_width(elem)
    results = []

    # 1) Nearest parallel grid
    best_sd = None
    best_gdir = None
    best_gname = ""
    best_abs = float("inf")
    for g, go, gd in grids_info:
        if not is_parallel(elem_dir, gd):
            continue
        sd = signed_perp(go, gd, mid)
        if abs(sd) < best_abs:
            best_abs = abs(sd)
            best_sd = sd
            best_gdir = gd
            try:
                best_gname = DB.Element.Name.GetValue(g)
            except:
                best_gname = "?"
    if best_gdir is not None and best_sd is not None:
        r = calc_snap(best_sd, half_w, best_gdir, precision)
        if r:
            shown, snapped, delta, mv = r
            results.append((best_gname, shown, snapped, delta, mv))

    # 2) Perpendicular grids - nearest endpoint
    perp_grids = [(g, go, gd) for g, go, gd in grids_info
                  if is_perpendicular(elem_dir, gd)]
    if perp_grids:
        perp_groups = group_grids_by_direction(perp_grids)
        for group in perp_groups:
            best_ep = None
            best_ep_gdir = None
            best_ep_gname = ""
            best_ep_go = None
            best_ep_abs = float("inf")
            for g, go, gd in group:
                for ep in [p0, p1]:
                    sd = signed_perp(go, gd, ep)
                    if abs(sd) < best_ep_abs:
                        best_ep_abs = abs(sd)
                        best_ep = ep
                        best_ep_gdir = gd
                        best_ep_go = go
                        try:
                            best_ep_gname = DB.Element.Name.GetValue(g)
                        except:
                            best_ep_gname = "?"
            if best_ep is not None and best_ep_gdir is not None:
                r = calc_snap_endpoint(best_ep, best_ep_go,
                                       best_ep_gdir, precision)
                if r:
                    shown, snapped, delta, mv = r
                    results.append((best_ep_gname, shown, snapped, delta, mv))
    return results


def analyze_column(elem, grid_groups, precision):
    loc = elem.Location
    if isinstance(loc, LocationPoint):
        pt = loc.Point
    elif isinstance(loc, LocationCurve):
        pt = loc.Curve.Evaluate(0.5, True)
    else:
        return []
    results = []
    for group in grid_groups:
        best_sd = None
        best_gdir = None
        best_gname = ""
        best_abs = float("inf")
        for g, go, gd in group:
            sd = signed_perp(go, gd, pt)
            if abs(sd) < best_abs:
                best_abs = abs(sd)
                best_sd = sd
                best_gdir = gd
                try:
                    best_gname = DB.Element.Name.GetValue(g)
                except:
                    best_gname = "?"
        if best_gdir is None:
            continue
        r = calc_snap(best_sd, 0.0, best_gdir, precision)
        if r:
            shown, snapped, delta, mv = r
            results.append((best_gname, shown, snapped, delta, mv))
    return results


# ==============================================================================
# CATEGORY BIC SETS
# ==============================================================================
WALL_BICS = [BuiltInCategory.OST_Walls]
BEAM_BICS = [BuiltInCategory.OST_StructuralFraming]
COL_BICS = [BuiltInCategory.OST_Columns, BuiltInCategory.OST_StructuralColumns]

ALL_BICS = WALL_BICS + BEAM_BICS + COL_BICS


# ==============================================================================
# XAML LOADING
# ==============================================================================
import os
import io

xaml_path = os.path.join(os.path.dirname(__file__), "Tools", "SnapDimension.xaml")


class SI(object):
    def __init__(self, eid, cat, ft, gn, dist, snap, delta, lv, mv):
        self.sel = True
        self.eid = eid
        self.cat = cat
        self.ft = ft
        self.gn = gn
        self.dist = dist
        self.snap = snap
        self.delta = delta
        self.lv = lv
        self.mv = mv


class MainWin(object):
    CAT_MAP = {
        0: {"walls": True,  "columns": True,  "beams": True},
        1: {"walls": True,  "columns": False, "beams": False},
        2: {"walls": False, "columns": True,  "beams": False},
        3: {"walls": False, "columns": False, "beams": True},
        4: {"walls": True,  "columns": True,  "beams": False},
        5: {"walls": True,  "columns": False, "beams": True},
        6: {"walls": False, "columns": True,  "beams": True},
    }

    def __init__(self):
        self.items = []
        with io.open(xaml_path, 'r', encoding='utf-8') as f:
            xaml_content = f.read()
        self.w = XamlReader.Load(MemoryStream(Encoding.UTF8.GetBytes(xaml_content)))
        self.cmbScope = self.w.FindName("cmbScope")
        self.cmbCat = self.w.FindName("cmbCat")
        self.cmbPrec = self.w.FindName("cmbPrec")
        self.cmbMax = self.w.FindName("cmbMax")
        self.dg = self.w.FindName("dg")
        self.txtSt = self.w.FindName("txtSt")
        self.w.FindName("btnScan").Click += self._scan
        self.w.FindName("btnHL").Click += self._hl
        self.w.FindName("btnApply").Click += self._apply
        self.w.FindName("btnClose").Click += lambda s, e: self.w.Close()
        self.w.FindName("chkAll").Checked += lambda s, e: self._tog(True)
        self.w.FindName("chkAll").Unchecked += lambda s, e: self._tog(False)

        # Auto-detect: if elements are pre-selected, switch to Current Selection
        sel_ids = uidoc.Selection.GetElementIds()
        if sel_ids.Count > 0:
            self.cmbScope.SelectedIndex = 1

    def _prec(self):
        return [1.0, 0.5, 5.0, 10.0][self.cmbPrec.SelectedIndex]

    def _mm(self):
        return [0.1, 0.5, 1.0, 2.0, 5.0, float('inf')][self.cmbMax.SelectedIndex]

    def _get_cats(self):
        return self.CAT_MAP.get(self.cmbCat.SelectedIndex,
                                 {"walls": True, "columns": True, "beams": True})

    def _lv(self, e):
        for b in [DB.BuiltInParameter.WALL_BASE_CONSTRAINT,
                   DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                   DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                   DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM]:
            try:
                p = e.get_Parameter(b)
                if p and p.HasValue:
                    l = doc.GetElement(p.AsElementId())
                    if l:
                        return DB.Element.Name.GetValue(l)
            except:
                pass
        return "-"

    def _ft(self, e):
        try:
            t = doc.GetElement(e.GetTypeId())
            if t:
                fn = ""
                try:
                    fn = t.FamilyName
                except:
                    pass
                tn = DB.Element.Name.GetValue(t)
                return (fn + " : " + tn) if fn else tn
        except:
            pass
        return "-"

    def _collect_all(self, bic):
        return list(FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements())

    def _is_wall_bic(self, elem):
        try:
            cat_id = get_id_value(elem.Category.Id)
            return cat_id == get_id_value(
                DB.Category.GetCategory(doc, BuiltInCategory.OST_Walls).Id)
        except:
            return isinstance(elem, Wall)

    def _is_beam_bic(self, elem):
        try:
            cat_id = get_id_value(elem.Category.Id)
            return cat_id == get_id_value(
                DB.Category.GetCategory(doc, BuiltInCategory.OST_StructuralFraming).Id)
        except:
            return False

    def _is_column_bic(self, elem):
        try:
            cat_id = get_id_value(elem.Category.Id)
            col_id = get_id_value(
                DB.Category.GetCategory(doc, BuiltInCategory.OST_Columns).Id)
            scol_id = get_id_value(
                DB.Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns).Id)
            return cat_id == col_id or cat_id == scol_id
        except:
            return False

    def _get_elements(self, cats):
        """Get elements based on scope and category filter."""
        use_selection = (self.cmbScope.SelectedIndex == 1)

        walls = []
        beams = []
        columns = []

        if use_selection:
            sel_ids = uidoc.Selection.GetElementIds()
            if sel_ids.Count == 0:
                return walls, beams, columns

            for eid in sel_ids:
                elem = doc.GetElement(eid)
                if elem is None or elem.Location is None:
                    continue
                if cats["walls"] and self._is_wall_bic(elem):
                    walls.append(elem)
                elif cats["beams"] and self._is_beam_bic(elem):
                    beams.append(elem)
                elif cats["columns"] and self._is_column_bic(elem):
                    columns.append(elem)
        else:
            if cats["walls"]:
                walls = [e for e in self._collect_all(BuiltInCategory.OST_Walls)
                         if e.Location]
            if cats["beams"]:
                beams = [e for e in self._collect_all(BuiltInCategory.OST_StructuralFraming)
                         if e.Location]
            if cats["columns"]:
                for bic in [BuiltInCategory.OST_Columns,
                            BuiltInCategory.OST_StructuralColumns]:
                    columns.extend([e for e in self._collect_all(bic) if e.Location])

        return walls, beams, columns

    def _scan(self, s, a):
        self.items = []
        pr = self._prec()
        mx = self._mm()
        cats = self._get_cats()

        grids_info = []
        for g in FilteredElementCollector(doc).OfClass(Grid).WhereElementIsNotElementType().ToElements():
            try:
                c = g.Curve
                d = line_dir_2d(c.GetEndPoint(0), c.GetEndPoint(1))
                if d:
                    grids_info.append((g, c.GetEndPoint(0), d))
            except:
                pass

        if not grids_info:
            self.txtSt.Text = "No grids found."
            return

        grid_groups = group_grids_by_direction(grids_info)

        walls, beams, columns = self._get_elements(cats)
        ns = len(walls) + len(beams) + len(columns)

        if ns == 0:
            scope_name = "selection" if self.cmbScope.SelectedIndex == 1 else "project"
            self.txtSt.Text = "No matching elements in {}.".format(scope_name)
            return

        ni = 0

        for e in walls:
            for gn, dist, snap, delta, mv in analyze_wall(e, grids_info, pr):
                if abs(delta) > mx:
                    continue
                ni += 1
                cn = "Walls"
                try:
                    cn = e.Category.Name
                except:
                    pass
                self.items.append(SI(get_id_value(e.Id), cn, self._ft(e),
                                     gn, dist, snap, delta, self._lv(e), mv))

        for e in beams:
            for gn, dist, snap, delta, mv in analyze_beam(e, grids_info, pr):
                if abs(delta) > mx:
                    continue
                ni += 1
                cn = "Structural Framing"
                try:
                    cn = e.Category.Name
                except:
                    pass
                self.items.append(SI(get_id_value(e.Id), cn, self._ft(e),
                                     gn, dist, snap, delta, self._lv(e), mv))

        for e in columns:
            for gn, dist, snap, delta, mv in analyze_column(e, grid_groups, pr):
                if abs(delta) > mx:
                    continue
                ni += 1
                cn = "Columns"
                try:
                    cn = e.Category.Name
                except:
                    pass
                self.items.append(SI(get_id_value(e.Id), cn, self._ft(e),
                                     gn, dist, snap, delta, self._lv(e), mv))

        self._ref()
        scope_name = "selected" if self.cmbScope.SelectedIndex == 1 else "total"
        self.txtSt.Text = "Scanned {} ({}). Found {} fractional.".format(
            ns, scope_name, ni)

    def _ref(self):
        dt = DataTable()
        for c in ["Sel", "ID", "Category", "Family : Type", "Grid",
                   "Dist (mm)", "Snap (mm)", "Move (mm)", "Level"]:
            dt.Columns.Add(c)
        for i in self.items:
            r = dt.NewRow()
            r["Sel"] = "V" if i.sel else ""
            r["ID"] = str(i.eid)
            r["Category"] = i.cat
            r["Family : Type"] = i.ft
            r["Grid"] = i.gn
            r["Dist (mm)"] = str(round(i.dist, 3))
            r["Snap (mm)"] = str(round(i.snap, 3))
            r["Move (mm)"] = str(round(i.delta, 4))
            r["Level"] = i.lv
            dt.Rows.Add(r)
        self.dg.ItemsSource = dt.DefaultView

    def _hl(self, s, a):
        ids = List[ElementId]()
        for i in self.items:
            if i.sel:
                try:
                    ids.Add(ElementId(i.eid))
                except:
                    pass
        if ids.Count > 0:
            uidoc.Selection.SetElementIds(ids)
            self.txtSt.Text = "Highlighted {}.".format(ids.Count)

    def _apply(self, s, a):
        todo = [i for i in self.items if i.sel]
        if not todo:
            self.txtSt.Text = "Nothing selected."
            return

        moves = OrderedDict()
        for i in todo:
            if i.eid not in moves:
                moves[i.eid] = XYZ(0, 0, 0)
            moves[i.eid] = XYZ(
                moves[i.eid].X + i.mv.X,
                moves[i.eid].Y + i.mv.Y,
                moves[i.eid].Z + i.mv.Z,
            )

        ok = fail = 0
        t = Transaction(doc, "DQT - Snap to Grid")
        try:
            t.Start()
            for eid_int, mv in moves.items():
                try:
                    eid = ElementId(eid_int)
                    if doc.GetElement(eid):
                        ElementTransformUtils.MoveElement(doc, eid, mv)
                        ok += 1
                    else:
                        fail += 1
                except:
                    fail += 1
            t.Commit()
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            self.txtSt.Text = "Err: " + str(ex)
            return

        self.items = [i for i in self.items if not i.sel]
        self._ref()
        self.txtSt.Text = "Snapped {}.".format(ok) + (" {} failed.".format(fail) if fail else "")

    def _tog(self, v):
        for i in self.items:
            i.sel = v
        self._ref()

    def show(self):
        self.w.ShowDialog()


def show_dialog():
    try:
        MainWin().show()
    except Exception as e:
        TaskDialog.Show("DQT - Snap to Grid", str(e))

if __name__ == "__main__":
    show_dialog()