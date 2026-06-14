# -*- coding: utf-8 -*-
"""
CAD to Wall v3.0 - DQT
Reads lines from CAD, detects parallel pairs, computes centerlines,
and auto-creates Wall Types matching detected thickness.

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "CAD to\nWall"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Convert CAD lines to Revit Walls. Auto-creates wall types by detected thickness."

import clr
import sys
import os
import math

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")

import System
from System.Collections.Generic import List
from System.Windows import (
    Window, WindowStartupLocation,
    Thickness, HorizontalAlignment, VerticalAlignment,
    TextWrapping, Visibility,
    MessageBox, MessageBoxButton,
    MessageBoxResult, MessageBoxImage
)
from System.Windows.Controls import (
    StackPanel, TextBlock, Border, Grid, RowDefinition, ColumnDefinition,
    Button, ComboBox, ComboBoxItem, CheckBox, TextBox, ScrollViewer,
    Orientation, ScrollBarVisibility
)
from System.Windows.Media import SolidColorBrush, Color

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory,
    ElementId, XYZ, Line, Wall, WallType, Level,
    ImportInstance, CompoundStructure, CompoundStructureLayer,
    MaterialFunctionAssignment
)

from pyrevit import revit, forms

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application


def _eid_int(eid):
    """Get integer value from ElementId - Revit 2024 uses .IntegerValue, 2025+ uses .Value"""
    try:
        return eid.IntegerValue
    except:
        try:
            return eid.Value
        except:
            return int(str(eid))


# ============================================================
# CONSTANTS
# ============================================================
TOLERANCE = 0.01
MERGE_TOL = 0.15        # feet (~45mm) - merge collinear segments with small gaps
PARALLEL_TOL = 0.998    # slightly relaxed for near-parallel lines
MAX_WALL_THICKNESS = 2.0  # feet (~600mm)
THICKNESS_ROUND_MM = 1     # round thickness to nearest 1mm


# ============================================================
# CAD GEOMETRY EXTRACTION
# ============================================================
def get_cad_instances():
    """Get all CAD imports and links - compatible with Revit 2024-2026"""
    cad_list = []
    
    # Method 1: ImportInstance class (works in most versions)
    try:
        collector = FilteredElementCollector(doc).OfClass(ImportInstance)
        for inst in collector:
            try:
                _add_cad_to_list(inst, cad_list)
            except:
                pass
    except:
        pass
    
    # Method 2: If no results, try BuiltInCategory for imported/linked CAD
    if not cad_list:
        try:
            # Try linked CAD category
            collector2 = FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_ImportObjectStyles).WhereElementIsNotElementType()
            for elem in collector2:
                if isinstance(elem, ImportInstance):
                    try:
                        _add_cad_to_list(elem, cad_list)
                    except:
                        pass
        except:
            pass
    
    # Method 3: Try to find CAD links via RevitLinkType approach
    if not cad_list:
        try:
            collector3 = FilteredElementCollector(doc).OfClass(DB.CADLinkType)
            for cad_type in collector3:
                try:
                    # Find instances of this type
                    inst_collector = FilteredElementCollector(doc).OfClass(ImportInstance)
                    for inst in inst_collector:
                        try:
                            if _eid_int(inst.GetTypeId()) == _eid_int(cad_type.Id):
                                _add_cad_to_list(inst, cad_list)
                        except:
                            pass
                except:
                    pass
        except:
            pass
    
    # Method 4: Broadest search - all elements in RasterImages + ImportInstances
    if not cad_list:
        try:
            all_cats = [
                BuiltInCategory.OST_RasterImages,
            ]
            for cat in all_cats:
                try:
                    collector4 = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
                    for elem in collector4:
                        if isinstance(elem, ImportInstance):
                            _add_cad_to_list(elem, cad_list)
                except:
                    pass
        except:
            pass
    
    # Deduplicate by element id
    seen_ids = set()
    unique_list = []
    for cad in cad_list:
        if cad["id"] not in seen_ids:
            seen_ids.add(cad["id"])
            unique_list.append(cad)
    
    return unique_list


def _add_cad_to_list(inst, cad_list):
    """Helper to safely add a CAD instance to the list"""
    name = "Unknown CAD"
    try:
        cad_type = doc.GetElement(inst.GetTypeId())
        if cad_type:
            try:
                name = DB.Element.Name.GetValue(cad_type)
            except:
                try:
                    p = cad_type.LookupParameter("Name")
                    if p:
                        name = p.AsString()
                except:
                    pass
            if not name or name == "Unknown CAD":
                try:
                    name = str(_eid_int(cad_type.Id))
                except:
                    pass
    except:
        pass

    is_linked = False
    try:
        is_linked = inst.IsLinked
    except:
        # Revit 2026 may not have IsLinked - check via ExternalFileReference
        try:
            cad_type = doc.GetElement(inst.GetTypeId())
            if cad_type:
                efr = cad_type.GetExternalFileReference()
                if efr:
                    is_linked = True
        except:
            pass

    label = "{} [{}]".format(name, "Linked" if is_linked else "Imported")
    
    eid = 0
    try:
        eid = _eid_int(inst.Id)
    except:
        try:
            eid = inst.Id.Value
        except:
            pass
    
    cad_list.append({
        "element": inst,
        "name": label,
        "id": eid,
        "is_linked": is_linked
    })


def get_cad_layers(cad_instance):
    layers = set()
    try:
        geo_elem = cad_instance.get_Geometry(DB.Options())
        if geo_elem is None:
            return sorted(list(layers))
        for geo_obj in geo_elem:
            if isinstance(geo_obj, DB.GeometryInstance):
                sub_geo = geo_obj.GetInstanceGeometry()
                if sub_geo:
                    for sub_obj in sub_geo:
                        try:
                            gstyle = doc.GetElement(sub_obj.GraphicsStyleId)
                            if gstyle:
                                cat = gstyle.GraphicsStyleCategory
                                if cat:
                                    layers.add(cat.Name)
                        except:
                            pass
    except:
        pass
    return sorted(list(layers))


def extract_lines_from_cad(cad_instance, selected_layers):
    lines = []
    selected_set = set(selected_layers)
    try:
        geo_elem = cad_instance.get_Geometry(DB.Options())
        if geo_elem is None:
            return lines
        for geo_obj in geo_elem:
            if isinstance(geo_obj, DB.GeometryInstance):
                sub_geo = geo_obj.GetInstanceGeometry()
                if sub_geo:
                    for sub_obj in sub_geo:
                        try:
                            layer_name = ""
                            gstyle = doc.GetElement(sub_obj.GraphicsStyleId)
                            if gstyle:
                                cat = gstyle.GraphicsStyleCategory
                                if cat:
                                    layer_name = cat.Name
                            if layer_name not in selected_set:
                                continue
                            if isinstance(sub_obj, DB.Line):
                                p0 = sub_obj.GetEndPoint(0)
                                p1 = sub_obj.GetEndPoint(1)
                                if p0.DistanceTo(p1) > TOLERANCE:
                                    lines.append({"start": p0, "end": p1, "layer": layer_name})
                            elif isinstance(sub_obj, DB.PolyLine):
                                coords = sub_obj.GetCoordinates()
                                for i in range(len(coords) - 1):
                                    p0 = coords[i]
                                    p1 = coords[i + 1]
                                    if p0.DistanceTo(p1) > TOLERANCE:
                                        lines.append({"start": p0, "end": p1, "layer": layer_name})
                        except:
                            pass
    except Exception as e:
        print("Error extracting CAD: {}".format(str(e)))
    return lines


# ============================================================
# MERGE COLLINEAR
# ============================================================
def merge_collinear_lines(lines):
    if not lines:
        return lines
    merged = True
    result = list(lines)
    while merged:
        merged = False
        new_result = []
        used = [False] * len(result)
        for i in range(len(result)):
            if used[i]:
                continue
            cur = result[i]
            cs = cur["start"]
            ce = cur["end"]
            dx = ce.X - cs.X
            dy = ce.Y - cs.Y
            clen = math.sqrt(dx * dx + dy * dy)
            if clen < TOLERANCE:
                used[i] = True
                continue
            cdx = dx / clen
            cdy = dy / clen
            for j in range(i + 1, len(result)):
                if used[j]:
                    continue
                other = result[j]
                os_ = other["start"]
                oe = other["end"]
                odx = oe.X - os_.X
                ody = oe.Y - os_.Y
                olen = math.sqrt(odx * odx + ody * ody)
                if olen < TOLERANCE:
                    used[j] = True
                    continue
                dot = abs(cdx * (odx / olen) + cdy * (ody / olen))
                if dot < PARALLEL_TOL:
                    continue
                # Collinear check
                vx = os_.X - cs.X
                vy = os_.Y - cs.Y
                cross = abs(vx * cdy - vy * cdx)
                if cross > MERGE_TOL:
                    continue
                ns = ne = None
                if ce.DistanceTo(os_) < MERGE_TOL:
                    ns, ne = cs, oe
                elif ce.DistanceTo(oe) < MERGE_TOL:
                    ns, ne = cs, os_
                elif cs.DistanceTo(os_) < MERGE_TOL:
                    ns, ne = ce, oe
                elif cs.DistanceTo(oe) < MERGE_TOL:
                    ns, ne = ce, os_
                if ns and ne and ns.DistanceTo(ne) > TOLERANCE:
                    cur = {"start": ns, "end": ne, "layer": cur["layer"]}
                    cs, ce = ns, ne
                    dx = ce.X - cs.X
                    dy = ce.Y - cs.Y
                    clen = math.sqrt(dx * dx + dy * dy)
                    if clen > TOLERANCE:
                        cdx = dx / clen
                        cdy = dy / clen
                    used[j] = True
                    merged = True
            new_result.append(cur)
            used[i] = True
        result = new_result
    return result


# ============================================================
# PARALLEL PAIR DETECTION -> CENTERLINES
# ============================================================
def project_point_on_line_2d(px, py, ax, ay, dx, dy):
    vx = px - ax
    vy = py - ay
    t = vx * dx + vy * dy
    fx = ax + t * dx
    fy = ay + t * dy
    dist = math.sqrt((px - fx) ** 2 + (py - fy) ** 2)
    return t, dist


def find_parallel_pairs(lines):
    """Detect parallel line pairs and compute centerlines using UNION extent.
    One line can pair with multiple parallel lines on the other side."""
    n = len(lines)
    paired = [False] * n
    centerlines = []

    dirs = []
    for line in lines:
        dx = line["end"].X - line["start"].X
        dy = line["end"].Y - line["start"].Y
        length = math.sqrt(dx * dx + dy * dy)
        if length > TOLERANCE:
            dirs.append({"dx": dx / length, "dy": dy / length, "len": length})
        else:
            dirs.append({"dx": 0, "dy": 0, "len": 0})

    for i in range(n):
        if paired[i] or dirs[i]["len"] == 0:
            continue
        di = dirs[i]
        si = lines[i]["start"]

        # Collect ALL parallel lines within thickness range
        candidates = []
        for j in range(n):
            if j == i or paired[j] or dirs[j]["len"] == 0:
                continue
            dj = dirs[j]
            dot = abs(di["dx"] * dj["dx"] + di["dy"] * dj["dy"])
            if dot < PARALLEL_TOL:
                continue
            sj = lines[j]["start"]
            ej = lines[j]["end"]
            _, dist_s = project_point_on_line_2d(sj.X, sj.Y, si.X, si.Y, di["dx"], di["dy"])
            _, dist_e = project_point_on_line_2d(ej.X, ej.Y, si.X, si.Y, di["dx"], di["dy"])
            avg_dist = (dist_s + dist_e) / 2.0
            if avg_dist > MAX_WALL_THICKNESS or avg_dist < TOLERANCE:
                continue
            # Check minimal overlap (at least 20% of shorter line)
            t_js, _ = project_point_on_line_2d(sj.X, sj.Y, si.X, si.Y, di["dx"], di["dy"])
            t_je, _ = project_point_on_line_2d(ej.X, ej.Y, si.X, si.Y, di["dx"], di["dy"])
            overlap = min(di["len"], max(t_js, t_je)) - max(0, min(t_js, t_je))
            shorter = min(di["len"], dj["len"])
            if overlap < shorter * 0.2:
                continue
            candidates.append({"idx": j, "dist": avg_dist, "t_s": t_js, "t_e": t_je})

        if not candidates:
            continue

        # Group candidates by similar distance (same wall thickness)
        candidates.sort(key=lambda c: c["dist"])
        best_dist = candidates[0]["dist"]
        same_side = [c for c in candidates if abs(c["dist"] - best_dist) < mm_to_ft(20)]

        # Use UNION of all extents (line i + all matched lines on same side)
        all_t_values = [0.0, di["len"]]  # line i range
        for c in same_side:
            all_t_values.append(c["t_s"])
            all_t_values.append(c["t_e"])

        t_union_start = min(all_t_values)
        t_union_end = max(all_t_values)

        if t_union_end - t_union_start < TOLERANCE:
            continue

        # Centerline at midpoint between the two parallel sides
        # Points on line i side
        pi_s = XYZ(si.X + di["dx"] * t_union_start, si.Y + di["dy"] * t_union_start, 0)
        pi_e = XYZ(si.X + di["dx"] * t_union_end, si.Y + di["dy"] * t_union_end, 0)

        # Offset to center: perpendicular direction * half thickness
        perp_dx = -di["dy"]
        perp_dy = di["dx"]

        # Determine which side the parallel lines are on
        mid_c = same_side[0]
        sj = lines[mid_c["idx"]]["start"]
        vx = sj.X - si.X
        vy = sj.Y - si.Y
        side = vx * perp_dx + vy * perp_dy
        half_t = best_dist / 2.0
        if side > 0:
            offset_x = perp_dx * half_t
            offset_y = perp_dy * half_t
        else:
            offset_x = -perp_dx * half_t
            offset_y = -perp_dy * half_t

        cs = XYZ(pi_s.X + offset_x, pi_s.Y + offset_y, 0)
        ce = XYZ(pi_e.X + offset_x, pi_e.Y + offset_y, 0)

        if cs.DistanceTo(ce) > TOLERANCE:
            centerlines.append({
                "start": cs, "end": ce,
                "thickness": best_dist,
                "layer": lines[i]["layer"]
            })

        paired[i] = True
        for c in same_side:
            paired[c["idx"]] = True

    unpaired = [lines[i] for i in range(n) if not paired[i]]
    return centerlines, unpaired


# ============================================================
# WALL TYPE CREATION BY THICKNESS
# ============================================================
def round_thickness_mm(thickness_ft):
    """Round thickness in feet to nearest mm integer"""
    mm = thickness_ft * 304.8
    return int(round(mm / THICKNESS_ROUND_MM) * THICKNESS_ROUND_MM)


def group_by_thickness(centerlines):
    """Group centerlines by rounded thickness (mm).
    Returns dict: {thickness_mm: [centerline_list]}"""
    groups = {}
    for cl in centerlines:
        t_mm = round_thickness_mm(cl["thickness"])
        if t_mm not in groups:
            groups[t_mm] = []
        groups[t_mm].append(cl)
    return groups


def find_base_wall_type():
    """Find a basic single-layer wall type to use as template for duplication.
    Prefers 'Generic' types."""
    collector = FilteredElementCollector(doc).OfClass(WallType)
    generic_type = None
    any_basic = None

    for wt in collector:
        try:
            kind = wt.Kind
            if kind != DB.WallKind.Basic:
                continue

            name = DB.Element.Name.GetValue(wt)
            any_basic = wt

            name_lower = name.lower()
            if "generic" in name_lower:
                generic_type = wt
                break
        except:
            pass

    return generic_type or any_basic


def get_or_create_wall_type(thickness_mm, base_type):
    """Find existing or create new wall type with name 'Generic - XXXmm'.
    Returns WallType element."""

    target_name = "Generic - {}mm".format(thickness_mm)
    thickness_ft = thickness_mm / 304.8

    # Check if type already exists
    collector = FilteredElementCollector(doc).OfClass(WallType)
    for wt in collector:
        try:
            name = DB.Element.Name.GetValue(wt)
            if name == target_name:
                return wt
        except:
            pass

    # Duplicate base type
    try:
        new_type = base_type.Duplicate(target_name)
    except Exception as e:
        print("Error duplicating wall type: {}".format(str(e)))
        return base_type

    # Set thickness by modifying compound structure
    try:
        cs = new_type.GetCompoundStructure()
        if cs:
            layers = cs.GetLayers()
            if layers.Count == 1:
                # Single layer - just set width
                cs.SetLayerWidth(0, thickness_ft)
            else:
                # Multiple layers - set first structural layer
                found = False
                for idx in range(layers.Count):
                    layer = layers[idx]
                    if layer.Function == MaterialFunctionAssignment.Structure:
                        cs.SetLayerWidth(idx, thickness_ft)
                        found = True
                        break
                if not found:
                    # No structural layer found, set first layer
                    cs.SetLayerWidth(0, thickness_ft)

            new_type.SetCompoundStructure(cs)
        print("Created wall type: {} ({}mm = {} ft)".format(target_name, thickness_mm, str(round(thickness_ft, 4))))
    except Exception as e:
        print("Error setting wall thickness: {}".format(str(e)))

    return new_type


# ============================================================
# WALL CREATION (auto wall type by thickness)
# ============================================================
def create_walls_auto(centerlines, unpaired, level_id, height, use_unpaired, default_thickness_mm, structural=False):
    """Create walls with auto-generated wall types based on detected thickness"""
    created = 0
    failed = 0
    skipped = 0
    types_created = []

    level = doc.GetElement(level_id)
    level_elev = level.Elevation

    base_type = find_base_wall_type()
    if not base_type:
        print("ERROR: No basic wall type found in model!")
        return 0, 0, 0, []

    # Group centerlines by thickness
    groups = group_by_thickness(centerlines)

    t = Transaction(doc, "DQT - CAD to Wall")
    t.Start()

    try:
        # Create walls from paired centerlines (with auto thickness)
        wall_type_cache = {}  # thickness_mm -> WallType

        for t_mm, cl_list in groups.items():
            if t_mm not in wall_type_cache:
                wt = get_or_create_wall_type(t_mm, base_type)
                wall_type_cache[t_mm] = wt
                types_created.append("Generic - {}mm".format(t_mm))

            wt = wall_type_cache[t_mm]

            for cl in cl_list:
                try:
                    s = cl["start"]
                    e = cl["end"]
                    start = XYZ(s.X, s.Y, level_elev)
                    end = XYZ(e.X, e.Y, level_elev)

                    if start.DistanceTo(end) < TOLERANCE:
                        skipped += 1
                        continue

                    dx = abs(end.X - start.X)
                    dy = abs(end.Y - start.Y)
                    if dx < TOLERANCE and dy < TOLERANCE:
                        skipped += 1
                        continue

                    wall_line = Line.CreateBound(start, end)
                    new_wall = Wall.Create(
                        doc, wall_line, wt.Id, level_id,
                        height, 0.0, False, structural
                    )

                    if new_wall:
                        created += 1
                    else:
                        failed += 1
                except:
                    failed += 1

        # Create walls from unpaired lines (use default thickness)
        if use_unpaired and unpaired:
            default_wt_key = default_thickness_mm
            if default_wt_key not in wall_type_cache:
                wt = get_or_create_wall_type(default_wt_key, base_type)
                wall_type_cache[default_wt_key] = wt
                types_created.append("Generic - {}mm".format(default_wt_key))

            wt = wall_type_cache[default_wt_key]

            for ln in unpaired:
                try:
                    s = ln["start"]
                    e = ln["end"]
                    start = XYZ(s.X, s.Y, level_elev)
                    end = XYZ(e.X, e.Y, level_elev)

                    if start.DistanceTo(end) < TOLERANCE:
                        skipped += 1
                        continue

                    dx = abs(end.X - start.X)
                    dy = abs(end.Y - start.Y)
                    if dx < TOLERANCE and dy < TOLERANCE:
                        skipped += 1
                        continue

                    wall_line = Line.CreateBound(start, end)
                    new_wall = Wall.Create(
                        doc, wall_line, wt.Id, level_id,
                        height, 0.0, False, structural
                    )

                    if new_wall:
                        created += 1
                    else:
                        failed += 1
                except:
                    failed += 1

        t.Commit()
    except Exception as e:
        t.RollBack()
        print("Transaction error: {}".format(str(e)))

    return created, failed, skipped, types_created


# ============================================================
# HELPERS
# ============================================================
def get_levels():
    collector = FilteredElementCollector(doc).OfClass(Level)
    lvs = []
    for lv in collector:
        try:
            name = DB.Element.Name.GetValue(lv)
            lvs.append({"name": name, "id": lv.Id, "elevation": lv.Elevation})
        except:
            pass
    lvs.sort(key=lambda x: x["elevation"])
    return lvs


def ft_to_mm(feet):
    return str(int(round(feet * 304.8)))


def mm_to_ft(mm):
    return mm / 304.8


# ============================================================
# WPF UI
# ============================================================



class CADtoWallWindow(Window):
    def _get_xaml_content(self):
        import codecs
        current_dir = os.path.dirname(__file__)
        while current_dir and not current_dir.endswith('T3Lab.extension'):
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                break
            current_dir = parent
        xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "CadtoWall.xaml")
        with codecs.open(xaml_path, "r", "utf-8") as f:
            return f.read()

    def __init__(self):
        from System.IO import MemoryStream
        from System.Text import Encoding
        from System.Windows.Markup import XamlReader

        byte_array = Encoding.UTF8.GetBytes(self._get_xaml_content())
        stream = MemoryStream(byte_array)
        xr = XamlReader.Load(stream)

        self.Title = xr.Title
        self.Width = xr.Width
        self.Height = xr.Height
        self.MinWidth = xr.MinWidth
        self.MinHeight = xr.MinHeight
        self.WindowStartupLocation = xr.WindowStartupLocation
        self.Background = xr.Background
        self.Content = xr.Content
        self._xr = xr

        self.cmbCAD = self._xr.FindName("cmbCAD")
        self.cmbLevel = self._xr.FindName("cmbLevel")
        self.txtHeight = self._xr.FindName("txtHeight")
        self.txtDefaultThk = self._xr.FindName("txtDefaultThk")
        self.chkStructural = self._xr.FindName("chkStructural")
        self.chkMerge = self._xr.FindName("chkMerge")
        self.chkUnpaired = self._xr.FindName("chkUnpaired")
        self.chkSelectAll = self._xr.FindName("chkSelectAll")
        self.txtSummary = self._xr.FindName("txtSummary")
        self.txtSearch = self._xr.FindName("txtSearch")
        self.layerPanel = self._xr.FindName("layerPanel")
        self.btnRefresh = self._xr.FindName("btnRefresh")
        self.btnPreview = self._xr.FindName("btnPreview")
        self.btnCreate = self._xr.FindName("btnCreate")
        self.btnClose = self._xr.FindName("btnClose")
        self.txtStatus = self._xr.FindName("txtStatus")

        self.cad_list = []
        self.levels = []
        self.layer_checkboxes = {}

        self.cmbCAD.SelectionChanged += self.on_cad_changed
        self.chkSelectAll.Checked += self.on_select_all_checked
        self.chkSelectAll.Unchecked += self.on_select_all_unchecked
        self.txtSearch.TextChanged += self.on_search_changed
        self.btnRefresh.Click += self.on_refresh
        self.btnPreview.Click += self.on_preview
        self.btnCreate.Click += self.on_create
        self.btnClose.Click += self.on_close

        self._load_data()

    def _load_data(self):
        self.cad_list = get_cad_instances()
        self.cmbCAD.Items.Clear()
        if not self.cad_list:
            item = ComboBoxItem()
            item.Content = "No CAD found in model"
            item.IsEnabled = False
            self.cmbCAD.Items.Add(item)
        else:
            for cad in self.cad_list:
                item = ComboBoxItem()
                item.Content = cad["name"]
                self.cmbCAD.Items.Add(item)
            self.cmbCAD.SelectedIndex = 0

        self.levels = get_levels()
        self.cmbLevel.Items.Clear()
        for lv in self.levels:
            item = ComboBoxItem()
            item.Content = "{} (Elev: {} mm)".format(lv["name"], ft_to_mm(lv["elevation"]))
            self.cmbLevel.Items.Add(item)
        if self.levels:
            try:
                av = doc.ActiveView
                alid = av.GenLevel.Id if hasattr(av, 'GenLevel') and av.GenLevel else None
                if alid:
                    for i, lv in enumerate(self.levels):
                        if _eid_int(lv["id"]) == _eid_int(alid):
                            self.cmbLevel.SelectedIndex = i
                            break
                    else:
                        self.cmbLevel.SelectedIndex = 0
                else:
                    self.cmbLevel.SelectedIndex = 0
            except:
                self.cmbLevel.SelectedIndex = 0

    def _load_layers(self, cad_data):
        self.layerPanel.Children.Clear()
        self.layer_checkboxes = {}
        if not cad_data:
            return
        layers = get_cad_layers(cad_data["element"])
        if not layers:
            tb = TextBlock()
            tb.Text = "No layers found"
            tb.FontSize = 11
            tb.Margin = Thickness(8, 8, 8, 8)
            self.layerPanel.Children.Add(tb)
            return
        for layer_name in layers:
            border = Border()
            border.Padding = Thickness(8, 4, 8, 4)
            border.Margin = Thickness(0, 0, 0, 1)
            border.Tag = layer_name
            sp = StackPanel()
            sp.Orientation = Orientation.Horizontal
            cb = CheckBox()
            cb.VerticalContentAlignment = VerticalAlignment.Center
            cb.Margin = Thickness(0, 0, 8, 0)
            cb.IsChecked = System.Nullable[System.Boolean](False)
            cb.Tag = layer_name
            tb = TextBlock()
            tb.Text = layer_name
            tb.FontSize = 11
            tb.Foreground = SolidColorBrush(Color.FromRgb(51, 51, 51))
            tb.VerticalAlignment = VerticalAlignment.Center
            sp.Children.Add(cb)
            sp.Children.Add(tb)
            border.Child = sp
            self.layerPanel.Children.Add(border)
            self.layer_checkboxes[layer_name] = cb
        self.txtSummary.Text = "{} layers found. Select layers with wall lines.".format(len(layers))

    def _get_selected_layers(self):
        selected = []
        for name, cb in self.layer_checkboxes.items():
            try:
                if cb.IsChecked == True:
                    selected.append(name)
            except:
                pass
        return selected

    def _get_cad(self):
        idx = self.cmbCAD.SelectedIndex
        if idx < 0 or idx >= len(self.cad_list):
            return None
        return self.cad_list[idx]

    def _get_lv(self):
        idx = self.cmbLevel.SelectedIndex
        if idx < 0 or idx >= len(self.levels):
            return None
        return self.levels[idx]

    def _get_height(self):
        try:
            return mm_to_ft(float(self.txtHeight.Text.strip()))
        except:
            return mm_to_ft(3000)

    def _get_default_thk(self):
        try:
            return int(float(self.txtDefaultThk.Text.strip()))
        except:
            return 200

    def _chk(self, c):
        try:
            return c.IsChecked == True
        except:
            return False

    def _process(self):
        cad = self._get_cad()
        sel = self._get_selected_layers()
        lines = extract_lines_from_cad(cad["element"], sel)
        raw = len(lines)
        if self._chk(self.chkMerge):
            lines = merge_collinear_lines(lines)
        cl, up = find_parallel_pairs(lines)
        return cl, up, raw, len(lines)

    def on_cad_changed(self, sender, args):
        cad = self._get_cad()
        if cad:
            self._load_layers(cad)

    def on_select_all_checked(self, sender, args):
        for cb in self.layer_checkboxes.values():
            cb.IsChecked = System.Nullable[System.Boolean](True)

    def on_select_all_unchecked(self, sender, args):
        for cb in self.layer_checkboxes.values():
            cb.IsChecked = System.Nullable[System.Boolean](False)

    def on_search_changed(self, sender, args):
        txt = self.txtSearch.Text.strip().lower()
        for i in range(self.layerPanel.Children.Count):
            child = self.layerPanel.Children[i]
            if isinstance(child, Border) and child.Tag:
                name = str(child.Tag).lower()
                child.Visibility = Visibility.Visible if (not txt or txt in name) else Visibility.Collapsed

    def on_refresh(self, sender, args):
        cad = self._get_cad()
        if cad:
            self._load_layers(cad)

    def on_preview(self, sender, args):
        cad = self._get_cad()
        if not cad:
            MessageBox.Show("Select a CAD instance.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        sel = self._get_selected_layers()
        if not sel:
            MessageBox.Show("Select at least one layer.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        cl, up, raw, merged = self._process()
        groups = group_by_thickness(cl)

        msg = "Raw lines: {} | After merge: {}\n".format(raw, merged)
        msg += "Parallel pairs: {} centerline walls\n".format(len(cl))
        msg += "Unpaired lines: {}\n\n".format(len(up))

        if groups:
            msg += "Wall types to create:\n"
            for t_mm in sorted(groups.keys()):
                count = len(groups[t_mm])
                msg += "  Generic - {}mm : {} walls\n".format(t_mm, count)

        self.txtSummary.Text = msg
        self.txtStatus.Text = "Preview: {} wall types, {} total walls".format(len(groups), len(cl))

    def on_create(self, sender, args):
        cad = self._get_cad()
        if not cad:
            MessageBox.Show("Select a CAD instance.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        sel = self._get_selected_layers()
        if not sel:
            MessageBox.Show("Select at least one layer.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        lv = self._get_lv()
        if not lv:
            MessageBox.Show("Select a Level.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        height = self._get_height()
        if height <= 0:
            MessageBox.Show("Enter a valid wall height.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        cl, up, raw, merged = self._process()
        use_up = self._chk(self.chkUnpaired)
        total = len(cl) + (len(up) if use_up else 0)
        if total == 0:
            MessageBox.Show("No lines found.", "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Information)
            return

        groups = group_by_thickness(cl)
        msg = "Create walls?\n\n"
        for t_mm in sorted(groups.keys()):
            msg += "Generic - {}mm: {} walls\n".format(t_mm, len(groups[t_mm]))
        if use_up:
            msg += "\nUnpaired: {} walls (Generic - {}mm)\n".format(len(up), self._get_default_thk())
        msg += "\nTotal: {} walls\n".format(total)
        msg += "Level: {}\nHeight: {} mm".format(lv["name"], self.txtHeight.Text.strip())

        if MessageBox.Show(msg, "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes:
            return

        created, failed, skipped, types_created = create_walls_auto(
            cl, up, lv["id"], height, use_up, self._get_default_thk(), self._chk(self.chkStructural))

        result_msg = "Created: {} walls\n".format(created)
        if failed > 0:
            result_msg += "Failed: {}\n".format(failed)
        if skipped > 0:
            result_msg += "Skipped: {}\n".format(skipped)
        if types_created:
            result_msg += "\nWall types created/used:\n"
            for tn in types_created:
                result_msg += "  {}\n".format(tn)

        self.txtSummary.Text = "Done: {} walls, {} types".format(created, len(types_created))
        MessageBox.Show(result_msg, "CAD to Wall", MessageBoxButton.OK, MessageBoxImage.Information)

    def on_close(self, sender, args):
        self.Close()


# ============================================================
# MAIN
# ============================================================
def main():
    try:
        window = CADtoWallWindow()
        window.ShowDialog()
    except Exception as e:
        import traceback
        print("CAD to Wall Error:")
        print(traceback.format_exc())
        forms.alert("Error: {}".format(str(e)), title="CAD to Wall")


if __name__ == "__main__":
    main()
else:
    main()