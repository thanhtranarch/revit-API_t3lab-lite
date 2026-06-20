# -*- coding: utf-8 -*-
"""
Point Cloud to Model
--------------------
Scan-to-BIM wizard: extract point cloud data, detect Walls, Floors,
Ceilings, Doors, Windows, Columns, Stairs, and Roof planes, then
create all elements in a single TransactionGroup.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__title__   = "Point Cloud\nto Model"
__author__  = "Tran Tien Thanh"
__version__  = "1.0.0"

# IMPORT LIBRARIES
# ==============================================================================
import os
import sys
import clr
import math

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState, Visibility
from System.Collections.Generic import List

from rpw import revit, DB
from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    WallType,
    Wall,
    Level,
    Grid,
    FloorType,
    Floor,
    XYZ,
    Line,
    CurveLoop,
    CurveArray,
    BoundingBoxXYZ,
    PointCloudInstance,
    IFailuresPreprocessor,
    FailureProcessingResult,
    ElementId,
    FamilySymbol,
    DetailLine,
    ElementTransformUtils
)
try:
    from Autodesk.Revit.DB import StructuralType
except ImportError:
    from Autodesk.Revit.DB.Structure import StructuralType

from Autodesk.Revit.DB.PointClouds import PointCloudFilterFactory
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, script

# Path setup
# ==============================================================================
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR    = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))))
lib_dir    = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

XAML_FILE  = os.path.join(EXT_DIR, 'lib', 'GUI', 'Tools', 'PointCloudModel.xaml')

# DEFINE VARIABLES
# ==============================================================================
logger        = script.get_logger()
doc           = revit.doc
uidoc         = revit.uidoc
REVIT_VERSION = int(doc.Application.VersionNumber)

# CLASS / FUNCTIONS
# ==============================================================================

# ── Section 1: Unit Helpers ────────────────────────────────────────────────────

MM_PER_FOOT = 304.8

def ft_to_mm(ft):
    return ft * MM_PER_FOOT

def mm_to_ft(mm):
    return mm / MM_PER_FOOT

def internal_to_mm(value_ft):
    """Convert Revit internal (feet) to millimetres, version-safe."""
    try:
        if REVIT_VERSION >= 2022:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId
            return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)
        else:
            from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
            return UnitUtils.ConvertFromInternalUnits(value_ft, DisplayUnitType.DUT_MILLIMETERS)
    except Exception:
        return value_ft * MM_PER_FOOT

def mm_to_internal(value_mm):
    """Convert millimetres to Revit internal (feet), version-safe."""
    try:
        if REVIT_VERSION >= 2022:
            from Autodesk.Revit.DB import UnitUtils, UnitTypeId
            return UnitUtils.ConvertToInternalUnits(value_mm, UnitTypeId.Millimeters)
        else:
            from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
            return UnitUtils.ConvertToInternalUnits(value_mm, DisplayUnitType.DUT_MILLIMETERS)
    except Exception:
        return value_mm / MM_PER_FOOT


# ── Section 2: Point Cloud Extraction ─────────────────────────────────────────

class PointCloudSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, PointCloudInstance)

    def AllowReference(self, reference, position):
        return False


def build_cloud_filter(center_pt, half_x_ft, half_y_ft, half_z_ft):
    """
    Six-plane axis-aligned box filter.
    Each plane is defined by a point on the plane and an inward normal
    pointing into the box interior.
    """
    x, y, z = center_pt.X, center_pt.Y, center_pt.Z
    normals = List[XYZ]()
    points  = List[XYZ]()
    # +X face — inward normal -X
    normals.Add(XYZ(-1.0, 0.0, 0.0)); points.Add(XYZ(x + half_x_ft, y, z))
    # -X face — inward normal +X
    normals.Add(XYZ( 1.0, 0.0, 0.0)); points.Add(XYZ(x - half_x_ft, y, z))
    # +Y face — inward normal -Y
    normals.Add(XYZ(0.0, -1.0, 0.0)); points.Add(XYZ(x, y + half_y_ft, z))
    # -Y face — inward normal +Y
    normals.Add(XYZ(0.0,  1.0, 0.0)); points.Add(XYZ(x, y - half_y_ft, z))
    # +Z face — inward normal -Z
    normals.Add(XYZ(0.0, 0.0, -1.0)); points.Add(XYZ(x, y, z + half_z_ft))
    # -Z face — inward normal +Z
    normals.Add(XYZ(0.0, 0.0,  1.0)); points.Add(XYZ(x, y, z - half_z_ft))
    return PointCloudFilterFactory.CreateMultiPlaneFilter(points, normals)


def extract_full_cloud(pc_instance, density_cap):
    """
    Extract up to density_cap points from the full bounding box of pc_instance.
    Returns list of (x, y, z) tuples in feet. Returns [] on any error.
    """
    try:
        bbox = pc_instance.get_BoundingBox(None)
        if bbox is None:
            return []
        cx = (bbox.Min.X + bbox.Max.X) / 2.0
        cy = (bbox.Min.Y + bbox.Max.Y) / 2.0
        cz = (bbox.Min.Z + bbox.Max.Z) / 2.0
        hx = (bbox.Max.X - bbox.Min.X) / 2.0 + 0.5
        hy = (bbox.Max.Y - bbox.Min.Y) / 2.0 + 0.5
        hz = (bbox.Max.Z - bbox.Min.Z) / 2.0 + 0.5
        vol = (2.0 * hx) * (2.0 * hy) * (2.0 * hz)
        min_dist = max(0.005, (vol / max(density_cap, 1)) ** (1.0 / 3.0))
        pcf = build_cloud_filter(XYZ(cx, cy, cz), hx, hy, hz)
        raw = pc_instance.GetPoints(pcf, min_dist, density_cap)
        return [(p.X, p.Y, p.Z) for p in raw]
    except Exception as ex:
        logger.debug("extract_full_cloud: {}".format(ex))
        return []


def extract_full_cloud_from_region(pc_instance, center, hx, hy, hz, density_cap):
    """
    Extract up to density_cap points from a user-defined axis-aligned region.
    center is an XYZ; hx/hy/hz are half-extents in feet. Returns [] on error.
    """
    try:
        vol = (2.0 * hx) * (2.0 * hy) * (2.0 * hz)
        min_dist = max(0.005, (vol / max(density_cap, 1)) ** (1.0 / 3.0))
        pcf = build_cloud_filter(center, hx, hy, hz)
        raw = pc_instance.GetPoints(pcf, min_dist, density_cap)
        return [(p.X, p.Y, p.Z) for p in raw]
    except Exception as ex:
        logger.debug("extract_full_cloud_from_region: {}".format(ex))
        return []


# ── Section 3: Geometry Math (pure Python) ────────────────────────────────────

def fit_line_2d(pts_xy):
    """
    Principal-axis LSR on 2D points.
    Returns (angle_rad, cx, cy) or None if fewer than 10 points.
    """
    if len(pts_xy) < 10:
        return None
    xs = [p[0] for p in pts_xy]
    ys = [p[1] for p in pts_xy]
    n = len(xs)
    mx = sum(xs) / n
    my = sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    sxy = sum((xs[i] - mx) * (ys[i] - my) for i in range(n))
    syy = sum((y - my) ** 2 for y in ys)
    angle = 0.5 * math.atan2(2.0 * sxy, sxx - syy)
    return angle, mx, my


def line_thickness_2d(pts_xy, angle, cx, cy):
    """97.5th–2.5th percentile span on the wall-normal axis (in feet)."""
    if not pts_xy:
        return 0.0
    nx = -math.sin(angle)
    ny =  math.cos(angle)
    proj = sorted((p[0] - cx) * nx + (p[1] - cy) * ny for p in pts_xy)
    n = len(proj)
    lo = int(round(0.025 * n))
    hi = int(round(0.975 * n)) - 1
    if hi <= lo:
        return abs(proj[-1] - proj[0])
    return abs(proj[hi] - proj[lo])


def line_length_2d(pts_xy, angle, cx, cy, percentile_lo=0.01, percentile_hi=0.99):
    """99th–1st percentile span along the wall direction (in feet)."""
    if not pts_xy:
        return 0.0
    dx = math.cos(angle)
    dy = math.sin(angle)
    proj = sorted((p[0] - cx) * dx + (p[1] - cy) * dy for p in pts_xy)
    n = len(proj)
    lo = int(round(percentile_lo * n))
    hi = int(round(percentile_hi * n)) - 1
    if hi <= lo:
        return abs(proj[-1] - proj[0])
    return abs(proj[hi] - proj[lo])


def cluster_2d(pts_xy, cell_size_ft, min_pts_cell=2, min_cluster_size=20):
    """
    Groups nearby 2D points into clusters using grid cells + BFS connected
    components. Returns list of clusters; each cluster is a list of (x, y).
    """
    grid = {}
    for (x, y) in pts_xy:
        key = (int(x / cell_size_ft), int(y / cell_size_ft))
        grid.setdefault(key, []).append((x, y))

    dense = {k: v for k, v in grid.items() if len(v) >= min_pts_cell}
    visited = set()
    clusters = []

    for cell in list(dense.keys()):
        if cell in visited:
            continue
        cluster_pts = []
        queue = [cell]
        while queue:
            k = queue.pop()
            if k in visited:
                continue
            visited.add(k)
            cluster_pts.extend(dense.get(k, []))
            ck, cl = k
            for dk in [-1, 0, 1]:
                for dl in [-1, 0, 1]:
                    nb = (ck + dk, cl + dl)
                    if nb in dense and nb not in visited:
                        queue.append(nb)
        if len(cluster_pts) >= min_cluster_size:
            clusters.append(cluster_pts)

    return clusters


def z_histogram(all_pts, bin_mm=50.0):
    """
    Builds a Z-value histogram from a list of (x, y, z) tuples.
    Returns (hist dict, z_min_ft, bin_size_ft).
    """
    if not all_pts:
        return {}, 0.0, 0.0
    z_vals = [p[2] for p in all_pts]
    z_min = min(z_vals)
    bin_ft = bin_mm / 304.8
    hist = {}
    for z in z_vals:
        bi = int((z - z_min) / bin_ft)
        hist[bi] = hist.get(bi, 0) + 1
    return hist, z_min, bin_ft


def find_histogram_peaks(hist, z_min, bin_ft, min_ratio=0.05):
    """
    Returns list of (z_ft, count) for local maxima above min_ratio of the
    global max count. Also includes the first bin if it qualifies.
    """
    if not hist:
        return []
    max_count = max(hist.values())
    threshold = max_count * min_ratio
    keys = sorted(hist.keys())
    peaks = []
    for i in range(1, len(keys) - 1):
        k = keys[i]
        c = hist.get(k, 0)
        prev_c = hist.get(keys[i - 1], 0)
        next_c = hist.get(keys[i + 1], 0)
        if c >= threshold and c >= prev_c and c >= next_c:
            z_ft = z_min + (k + 0.5) * bin_ft
            peaks.append((z_ft, c))
    # Also include the first bin if it qualifies
    if keys:
        k0 = keys[0]
        if hist.get(k0, 0) >= threshold:
            peaks.append((z_min + 0.5 * bin_ft, hist[k0]))
    return sorted(peaks, key=lambda p: p[0])


# ── Section 4: DetectedElement Data Classes ────────────────────────────────────

class DetectedElement(object):
    """Base class bound to the DataGrid in the WPF results view."""

    def __init__(self, elem_type, level_name, dimensions_str, confidence, data):
        self.Type           = elem_type
        self.LevelName      = level_name
        self.Dimensions     = dimensions_str
        self.ConfidenceText = u"{}%".format(confidence)
        self.Include        = True
        self._data          = data


class DetectedWall(DetectedElement):
    def __init__(self, angle, cx, cy, length_ft, thickness_ft, level_name, snapped, snap_desc):
        dims = u"L={:.0f} mm  T={:.0f} mm".format(ft_to_mm(length_ft), ft_to_mm(thickness_ft))
        super(DetectedWall, self).__init__(
            'Wall', level_name, dims, 85 if snapped else 70,
            {
                'angle': angle, 'cx': cx, 'cy': cy,
                'length_ft': length_ft, 'thickness_ft': thickness_ft,
            }
        )


class DetectedFloor(DetectedElement):
    def __init__(self, z_ft, corners_xy, surface_type, level_name):
        w = abs(corners_xy[1][0] - corners_xy[0][0])
        h = abs(corners_xy[2][1] - corners_xy[0][1])
        area_m2 = (ft_to_mm(w) / 1000.0) * (ft_to_mm(h) / 1000.0)
        dims = u"~{:.1f} m²  @Z={:.0f} mm".format(area_m2, ft_to_mm(z_ft))
        elem_type = 'Ceiling' if surface_type == 'ceiling' else 'Floor'
        super(DetectedFloor, self).__init__(
            elem_type, level_name, dims, 80,
            {'z_ft': z_ft, 'corners_xy': corners_xy, 'surface_type': surface_type}
        )


class DetectedColumn(DetectedElement):
    def __init__(self, cx, cy, width_ft, depth_ft, z_bot_ft, z_top_ft, level_name):
        dims = u"W={:.0f} D={:.0f} H={:.0f} mm".format(
            ft_to_mm(width_ft), ft_to_mm(depth_ft), ft_to_mm(z_top_ft - z_bot_ft))
        super(DetectedColumn, self).__init__(
            'Column', level_name, dims, 65,
            {
                'cx': cx, 'cy': cy,
                'width_ft': width_ft, 'depth_ft': depth_ft,
                'z_bot_ft': z_bot_ft, 'z_top_ft': z_top_ft,
            }
        )


class DetectedOpening(DetectedElement):
    """Door or Window opening detected in a wall."""

    def __init__(self, elem_type, host_wall_data, u_center, w_bottom, width_ft, height_ft, level_name):
        dims = u"W={:.0f} H={:.0f} mm".format(ft_to_mm(width_ft), ft_to_mm(height_ft))
        super(DetectedOpening, self).__init__(
            elem_type, level_name, dims, 60,
            {
                'host_wall_data': host_wall_data,
                'u_center': u_center, 'w_bottom': w_bottom,
                'width_ft': width_ft, 'height_ft': height_ft,
            }
        )


class DetectedStair(DetectedElement):
    def __init__(self, z_bot_ft, z_top_ft, tread_count, level_name):
        dims = u"{} treads  Rise={:.0f}–{:.0f} mm".format(
            tread_count, ft_to_mm(z_bot_ft), ft_to_mm(z_top_ft))
        super(DetectedStair, self).__init__(
            'Stair', level_name, dims, 55,
            {'z_bot_ft': z_bot_ft, 'z_top_ft': z_top_ft, 'tread_count': tread_count}
        )


class DetectedRoof(DetectedElement):
    def __init__(self, z_ft, corners_xy, slope_deg, level_name):
        dims = u"Slope={:.1f}°  @Z={:.0f} mm".format(slope_deg, ft_to_mm(z_ft))
        super(DetectedRoof, self).__init__(
            'Roof', level_name, dims, 60,
            {'z_ft': z_ft, 'corners_xy': corners_xy, 'slope_deg': slope_deg}
        )


# ── Section 5: PointCloudAnalyzer ─────────────────────────────────────────────

class PointCloudAnalyzer(object):
    """Runs all detection algorithms on extracted point cloud data."""

    RISER_TYPICAL_MM = (140.0, 200.0)

    def __init__(self, all_pts):
        self.pts    = all_pts
        self._z_min = min(p[2] for p in all_pts) if all_pts else 0.0
        self._z_max = max(p[2] for p in all_pts) if all_pts else 0.0
        self._x_min = min(p[0] for p in all_pts) if all_pts else 0.0
        self._x_max = max(p[0] for p in all_pts) if all_pts else 0.0
        self._y_min = min(p[1] for p in all_pts) if all_pts else 0.0
        self._y_max = max(p[1] for p in all_pts) if all_pts else 0.0
        self._grid_angles = self._load_grid_angles()

    def _load_grid_angles(self):
        angles = [0.0, math.pi / 2.0]
        try:
            grids = FilteredElementCollector(doc).OfClass(Grid).ToElements()
            for g in grids:
                curve = g.Curve
                if curve is None:
                    continue
                d = curve.Direction
                a = math.atan2(d.Y, d.X) % math.pi
                angles.append(a)
        except Exception:
            pass
        return angles

    def _snap_angle(self, angle_rad, tol_deg=1.0):
        tol = math.radians(tol_deg)
        a = angle_rad % math.pi
        for ref in self._grid_angles:
            r = ref % math.pi
            for cand in [r, (r + math.pi / 2.0) % math.pi]:
                diff = abs(a - cand)
                if diff > math.pi / 2.0:
                    diff = math.pi - diff
                if diff <= tol:
                    return cand, True, u"Snapped {:.1f}°".format(math.degrees(cand))
        return a, False, u"{:.2f}°".format(math.degrees(a))

    def _best_level(self, z_ft):
        """Find the closest Level element at or below z_ft."""
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        best      = None
        best_diff = float('inf')
        for lv in levels:
            diff = z_ft - lv.Elevation
            if 0.0 <= diff < best_diff:
                best_diff = diff
                best = lv
        if best is None and levels:
            best = sorted(levels, key=lambda l: abs(l.Elevation - z_ft))[0]
        return best

    # ── Wall detection ─────────────────────────────────────────────────────────

    def detect_walls(self, snap_tol_deg=1.0, min_length_mm=500.0, z_bin_mm=50.0):
        """
        For each detected floor level, take a horizontal slice at floor_z + 1.2 m,
        cluster 2D points, fit lines, filter by aspect ratio.
        Returns list of DetectedWall.
        """
        results = []
        min_length_ft = mm_to_ft(min_length_mm)
        cell_size_ft  = mm_to_ft(50.0)

        hist, z_min_h, bin_ft = z_histogram(self.pts, z_bin_mm)
        peaks = find_histogram_peaks(hist, z_min_h, bin_ft)
        floor_zs = [p[0] for p in peaks if p[0] < (self._z_max - mm_to_ft(300.0))]
        if not floor_zs:
            floor_zs = [(self._z_min + self._z_max) / 2.0]

        seen_walls = []  # list of (cx, cy) for deduplication

        for fz in floor_zs:
            scan_z = fz + mm_to_ft(1200.0)
            if scan_z > self._z_max:
                continue
            dz = mm_to_ft(z_bin_mm / 2.0)

            slice_xy = [(p[0], p[1]) for p in self.pts if abs(p[2] - scan_z) <= dz]
            if len(slice_xy) < 30:
                continue

            clusters = cluster_2d(slice_xy, cell_size_ft, min_pts_cell=2, min_cluster_size=20)

            for cluster in clusters:
                fit = fit_line_2d(cluster)
                if fit is None:
                    continue
                angle, cx, cy = fit

                length_ft = line_length_2d(cluster, angle, cx, cy)
                if length_ft < min_length_ft:
                    continue

                thickness_ft  = line_thickness_2d(cluster, angle, cx, cy)
                max_thickness = mm_to_ft(800.0)
                if thickness_ft > max_thickness or thickness_ft <= 0.0:
                    continue

                # Aspect ratio: length must be >= 3x thickness
                if length_ft < thickness_ft * 3.0:
                    continue

                snapped_angle, snapped, snap_desc = self._snap_angle(angle, snap_tol_deg)

                # Dedup: skip if within 200 mm of an already-detected wall centroid
                too_close = False
                for (ex_cx, ex_cy) in seen_walls:
                    if math.sqrt((cx - ex_cx) ** 2 + (cy - ex_cy) ** 2) < mm_to_ft(200.0):
                        too_close = True
                        break
                if too_close:
                    continue
                seen_walls.append((cx, cy))

                lv = self._best_level(fz)
                level_name = lv.Name if lv else u"Level 1"
                results.append(
                    DetectedWall(snapped_angle, cx, cy, length_ft,
                                 thickness_ft, level_name, snapped, snap_desc)
                )

        return results

    # ── Horizontal surface detection ───────────────────────────────────────────

    def detect_horizontal_surfaces(self, z_bin_mm=50.0, min_area_m2=1.0):
        """Detect floors and ceilings via Z-histogram peak analysis."""
        results = []
        hist, z_min_h, bin_ft = z_histogram(self.pts, z_bin_mm)
        peaks = find_histogram_peaks(hist, z_min_h, bin_ft, min_ratio=0.05)
        if not peaks:
            return results

        total = len(peaks)
        for idx, (z_ft, _count) in enumerate(peaks):
            stype = 'ceiling' if idx == total - 1 else 'floor'
            dz = bin_ft / 2.0
            slice_xy = [(p[0], p[1]) for p in self.pts if abs(p[2] - z_ft) <= dz]
            if len(slice_xy) < 20:
                continue

            min_x = min(p[0] for p in slice_xy)
            max_x = max(p[0] for p in slice_xy)
            min_y = min(p[1] for p in slice_xy)
            max_y = max(p[1] for p in slice_xy)

            width_ft = max_x - min_x
            depth_ft = max_y - min_y
            area_m2  = (ft_to_mm(width_ft) / 1000.0) * (ft_to_mm(depth_ft) / 1000.0)
            if area_m2 < min_area_m2:
                continue

            corners = [(min_x, min_y), (max_x, min_y), (max_x, max_y), (min_x, max_y)]
            lv = self._best_level(z_ft)
            level_name = lv.Name if lv else u""
            results.append(DetectedFloor(z_ft, corners, stype, level_name))

        return results

    # ── Column detection ───────────────────────────────────────────────────────

    def detect_columns(self, min_size_mm=100.0, max_size_mm=600.0):
        """
        Find vertical clusters that are small in XY (column cross-section)
        but tall in Z (spanning at least 40% of room height).
        """
        results = []
        min_ft      = mm_to_ft(min_size_mm)
        max_ft      = mm_to_ft(max_size_mm)
        cell_size_ft = mm_to_ft(50.0)
        room_height = self._z_max - self._z_min

        mid_z = (self._z_min + self._z_max) / 2.0
        dz    = mm_to_ft(100.0)
        slice_xy = [(p[0], p[1]) for p in self.pts if abs(p[2] - mid_z) <= dz]
        if len(slice_xy) < 10:
            return results

        clusters = cluster_2d(slice_xy, cell_size_ft, min_pts_cell=2, min_cluster_size=5)

        for cluster in clusters:
            xs = [p[0] for p in cluster]
            ys = [p[1] for p in cluster]
            w = max(xs) - min(xs)
            d = max(ys) - min(ys)

            if w < min_ft or d < min_ft:
                continue
            if w > max_ft or d > max_ft:
                continue
            if max(w, d) < min_ft * 0.5:
                continue

            cx = (max(xs) + min(xs)) / 2.0
            cy = (max(ys) + min(ys)) / 2.0
            r  = max(w, d)
            vert_pts = [p for p in self.pts
                        if abs(p[0] - cx) <= r and abs(p[1] - cy) <= r]
            if not vert_pts:
                continue

            z_bot   = min(p[2] for p in vert_pts)
            z_top   = max(p[2] for p in vert_pts)
            height_ft = z_top - z_bot
            if height_ft < room_height * 0.4:
                continue

            lv = self._best_level(z_bot)
            level_name = lv.Name if lv else u""
            results.append(DetectedColumn(cx, cy, w, d, z_bot, z_top, level_name))

        return results

    # ── Opening detection ──────────────────────────────────────────────────────

    def detect_openings(self, detected_walls,
                        min_door_w_mm=600.0, min_door_h_mm=1800.0,
                        min_win_w_mm=400.0,  min_win_h_mm=400.0):
        """
        For each detected wall, project nearby points onto the wall face plane
        and look for vertical rectangular gaps indicating door/window openings.
        Returns list of DetectedOpening.
        """
        results = []
        if not detected_walls:
            return results
        floor_z = self._z_min
        ceil_z  = self._z_max

        for wall in detected_walls:
            d = wall._data
            angle    = d['angle']
            cx, cy   = d['cx'], d['cy']
            length_ft = d['length_ft']
            thick_ft  = d['thickness_ft']

            nx = -math.sin(angle)
            ny =  math.cos(angle)
            dx_wall = math.cos(angle)
            dy_wall = math.sin(angle)

            face_pts = []
            for p in self.pts:
                rel_x = p[0] - cx
                rel_y = p[1] - cy
                u_proj = rel_x * dx_wall + rel_y * dy_wall
                v_proj = rel_x * nx      + rel_y * ny
                if (abs(u_proj) <= length_ft / 2.0 * 1.1
                        and abs(v_proj) <= thick_ft * 1.5):
                    face_pts.append((u_proj, p[2], v_proj))

            if len(face_pts) < 20:
                continue

            wall_pts_uz = [(p[0], p[1]) for p in face_pts if abs(p[2]) <= thick_ft]
            u_vals = [p[0] for p in wall_pts_uz]
            if not u_vals:
                continue

            u_min, u_max = min(u_vals), max(u_vals)
            u_bin_ft = mm_to_ft(100.0)
            n_ubins  = max(1, int((u_max - u_min) / u_bin_ft) + 1)
            u_hist   = [0] * n_ubins
            for (u, z) in wall_pts_uz:
                bi = min(int((u - u_min) / u_bin_ft), n_ubins - 1)
                u_hist[bi] += 1

            if not u_hist or max(u_hist) == 0:
                continue
            avg_density   = sum(u_hist) / float(len(u_hist))
            gap_threshold = avg_density * 0.2

            in_gap    = False
            gap_start = 0
            for i, cnt in enumerate(u_hist):
                if cnt <= gap_threshold and not in_gap:
                    in_gap    = True
                    gap_start = i
                elif cnt > gap_threshold and in_gap:
                    in_gap   = False
                    gap_end  = i
                    gap_width    = (gap_end - gap_start) * u_bin_ft
                    gap_u_center = u_min + (gap_start + gap_end) / 2.0 * u_bin_ft

                    gap_pts_z = [p[1] for p in face_pts
                                 if u_min + gap_start * u_bin_ft <= p[0] <= u_min + gap_end * u_bin_ft]
                    if gap_pts_z:
                        gap_z_min = min(gap_pts_z)
                        gap_z_max = max(gap_pts_z)
                    else:
                        gap_z_min = floor_z
                        gap_z_max = ceil_z
                    gap_height = gap_z_max - gap_z_min

                    is_door_bottom = (gap_z_min - floor_z) < mm_to_ft(100.0)

                    if (is_door_bottom
                            and gap_width  >= mm_to_ft(min_door_w_mm)
                            and gap_height >= mm_to_ft(min_door_h_mm)):
                        results.append(DetectedOpening(
                            'Door', d, gap_u_center, gap_z_min,
                            gap_width, gap_height, wall.LevelName))
                    elif (not is_door_bottom
                            and gap_width  >= mm_to_ft(min_win_w_mm)
                            and gap_height >= mm_to_ft(min_win_h_mm)):
                        results.append(DetectedOpening(
                            'Window', d, gap_u_center, gap_z_min,
                            gap_width, gap_height, wall.LevelName))

        return results

    # ── Stair detection ────────────────────────────────────────────────────────

    def detect_stairs(self, riser_height_mm=165.0):
        """
        Look for regions where Z increases in discrete steps of ~riser_height.
        Returns list of DetectedStair (confidence ~55%).
        """
        results = []
        riser_ft     = mm_to_ft(riser_height_mm)
        riser_tol_ft = mm_to_ft(30.0)

        hist, z_min_h, bin_ft = z_histogram(self.pts, riser_height_mm * 0.5)
        peaks = find_histogram_peaks(hist, z_min_h, bin_ft, min_ratio=0.02)
        if len(peaks) < 3:
            return results

        stair_sequences = []
        current_seq = [peaks[0]]
        for i in range(1, len(peaks)):
            dz = peaks[i][0] - peaks[i - 1][0]
            if abs(dz - riser_ft) <= riser_tol_ft:
                current_seq.append(peaks[i])
            else:
                if len(current_seq) >= 3:
                    stair_sequences.append(current_seq)
                current_seq = [peaks[i]]
        if len(current_seq) >= 3:
            stair_sequences.append(current_seq)

        for seq in stair_sequences:
            z_bot       = seq[0][0]
            z_top       = seq[-1][0]
            tread_count = len(seq)
            lv          = self._best_level(z_bot)
            level_name  = lv.Name if lv else u""
            results.append(DetectedStair(z_bot, z_top, tread_count, level_name))

        return results

    # ── Roof detection ─────────────────────────────────────────────────────────

    def detect_roof(self, min_slope_deg=5.0):
        """
        Find inclined planes near the top of the cloud using a simplified 3D
        variance check. Returns list of DetectedRoof.
        """
        results = []
        z_range      = self._z_max - self._z_min
        top_z_thresh = self._z_max - z_range * 0.15
        top_pts      = [p for p in self.pts if p[2] >= top_z_thresh]
        if len(top_pts) < 20:
            return results

        n  = len(top_pts)
        mx = sum(p[0] for p in top_pts) / n
        my = sum(p[1] for p in top_pts) / n
        mz = sum(p[2] for p in top_pts) / n

        sxx = sum((p[0] - mx) ** 2 for p in top_pts)
        syy = sum((p[1] - my) ** 2 for p in top_pts)
        szz = sum((p[2] - mz) ** 2 for p in top_pts)

        # If Z variance is negligible relative to XY variance it is a flat ceiling
        if szz < (sxx + syy) * 0.01:
            return results

        sxz = sum((p[0] - mx) * (p[2] - mz) for p in top_pts)
        syz = sum((p[1] - my) * (p[2] - mz) for p in top_pts)

        slope_x     = sxz / max(sxx, 1e-10)
        slope_y     = syz / max(syy, 1e-10)
        slope_total = math.sqrt(slope_x ** 2 + slope_y ** 2)
        slope_deg   = math.degrees(math.atan(slope_total))

        if slope_deg < min_slope_deg:
            return results

        min_x = min(p[0] for p in top_pts)
        max_x = max(p[0] for p in top_pts)
        min_y = min(p[1] for p in top_pts)
        max_y = max(p[1] for p in top_pts)
        corners = [(min_x, min_y), (max_x, min_y), (max_x, max_y), (min_x, max_y)]

        lv = self._best_level(top_z_thresh)
        level_name = lv.Name if lv else u""
        results.append(DetectedRoof(top_z_thresh, corners, slope_deg, level_name))
        return results

    # ── Main run method ────────────────────────────────────────────────────────

    def run(self, settings):
        """
        Run all enabled detectors. settings is a dict populated from the UI.
        Returns a list of DetectedElement sorted by type.
        """
        all_results = []

        if settings.get('detect_wall', True):
            walls = self.detect_walls(
                snap_tol_deg=settings.get('snap_tol', 1.0),
                min_length_mm=settings.get('wall_min_len', 500.0))
            all_results.extend(walls)
        else:
            walls = []

        if settings.get('detect_floor', True) or settings.get('detect_ceiling', True):
            surfaces = self.detect_horizontal_surfaces(
                z_bin_mm=settings.get('floor_zbin', 50.0),
                min_area_m2=settings.get('floor_min_area', 1.0))
            for s in surfaces:
                if s.Type == 'Floor' and settings.get('detect_floor', True):
                    all_results.append(s)
                elif s.Type == 'Ceiling' and settings.get('detect_ceiling', True):
                    all_results.append(s)

        if settings.get('detect_column', False):
            cols = self.detect_columns(
                min_size_mm=settings.get('col_min', 100.0),
                max_size_mm=settings.get('col_max', 600.0))
            all_results.extend(cols)

        if settings.get('detect_door', False) or settings.get('detect_window', False):
            openings = self.detect_openings(
                walls,
                min_door_w_mm=settings.get('door_min_w', 600.0),
                min_door_h_mm=settings.get('door_min_h', 1800.0),
                min_win_w_mm=settings.get('win_min_w', 400.0),
                min_win_h_mm=settings.get('win_min_h', 400.0))
            for o in openings:
                if o.Type == 'Door' and settings.get('detect_door', False):
                    all_results.append(o)
                elif o.Type == 'Window' and settings.get('detect_window', False):
                    all_results.append(o)

        if settings.get('detect_stair', False):
            stairs = self.detect_stairs(
                riser_height_mm=settings.get('stair_riser', 165.0))
            all_results.extend(stairs)

        if settings.get('detect_roof', False):
            roofs = self.detect_roof(
                min_slope_deg=settings.get('roof_slope', 5.0))
            all_results.extend(roofs)

        return all_results


# ── Section 6: ElementBuilder ─────────────────────────────────────────────────

class ElementBuilder(object):
    """Creates Revit elements from DetectedElement data inside named Transactions."""

    def __init__(self, document):
        self.doc         = document
        self._wall_types  = self._load_wall_types()
        self._floor_types = self._load_floor_types()
        self._levels      = self._load_levels()

    def _load_wall_types(self):
        return list(FilteredElementCollector(self.doc).OfClass(WallType).ToElements())

    def _load_floor_types(self):
        return list(FilteredElementCollector(self.doc).OfClass(FloorType).ToElements())

    def _load_levels(self):
        lvs = FilteredElementCollector(self.doc).OfClass(Level).ToElements()
        return {lv.Name: lv for lv in lvs}

    def _best_wall_type(self, thickness_ft):
        """Find the WallType whose compound structure width is closest to thickness_ft."""
        best      = None
        best_diff = float('inf')
        for wt in self._wall_types:
            try:
                cs = wt.GetCompoundStructure()
                w  = cs.GetTotalWidth() if cs else wt.Width
                d  = abs(w - thickness_ft)
                if d < best_diff:
                    best_diff = d
                    best = wt
            except Exception:
                pass
        return best or (self._wall_types[0] if self._wall_types else None)

    def _best_floor_type(self):
        return self._floor_types[0] if self._floor_types else None

    def _get_level(self, level_name):
        return self._levels.get(level_name) or (
            sorted(self._levels.values(), key=lambda l: l.Elevation)[0]
            if self._levels else None)

    def _rect_curve_loop(self, corners_xy, z_ft):
        """Build a closed rectangular CurveLoop from 4 (x, y) corner tuples."""
        cl  = CurveLoop()
        pts = [XYZ(x, y, z_ft) for (x, y) in corners_xy]
        for i in range(len(pts)):
            p1 = pts[i]
            p2 = pts[(i + 1) % len(pts)]
            if p1.DistanceTo(p2) > 0.01:
                cl.Append(Line.CreateBound(p1, p2))
        return cl

    def build_wall(self, elem, height_ft=None):
        d = elem._data
        angle      = d['angle']
        cx, cy     = d['cx'], d['cy']
        length_ft  = d['length_ft']
        thick_ft   = d['thickness_ft']
        if height_ft is None:
            height_ft = mm_to_ft(3000.0)

        half  = length_ft / 2.0
        pt1   = XYZ(cx - math.cos(angle) * half, cy - math.sin(angle) * half, 0.0)
        pt2   = XYZ(cx + math.cos(angle) * half, cy + math.sin(angle) * half, 0.0)
        wline = Line.CreateBound(pt1, pt2)

        lv = self._get_level(elem.LevelName)
        if not lv:
            return None
        wt = self._best_wall_type(thick_ft)
        if not wt:
            return None

        with Transaction(self.doc, "T3Lab: Create Wall") as t:
            t.Start()
            try:
                wall = Wall.Create(self.doc, wline, wt.Id, lv.Id,
                                   height_ft, 0.0, False, False)
                t.Commit()
                return wall
            except Exception as ex:
                logger.error("build_wall: {}".format(ex))
                t.RollBack()
                return None

    def build_floor(self, elem):
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None
        ft = self._best_floor_type()
        if not ft:
            return None
        cl = self._rect_curve_loop(d['corners_xy'], d['z_ft'])

        with Transaction(self.doc, "T3Lab: Create Floor") as t:
            t.Start()
            try:
                if REVIT_VERSION >= 2022:
                    profile = List[CurveLoop]()
                    profile.Add(cl)
                    floor = Floor.Create(self.doc, profile, ft.Id, lv.Id)
                else:
                    ca = CurveArray()
                    for curve in cl:
                        ca.Append(curve)
                    floor = self.doc.Create.NewFloor(ca, ft, lv, False)
                t.Commit()
                return floor
            except Exception as ex:
                logger.error("build_floor: {}".format(ex))
                t.RollBack()
                return None

    def build_ceiling(self, elem):
        """Create a ceiling (Revit 2022+ Ceiling.Create, else floor-based fallback)."""
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None
        cl = self._rect_curve_loop(d['corners_xy'], d['z_ft'])

        with Transaction(self.doc, "T3Lab: Create Ceiling") as t:
            t.Start()
            try:
                if REVIT_VERSION >= 2022:
                    from Autodesk.Revit.DB import Ceiling, CeilingType
                    ceil_types = FilteredElementCollector(self.doc) \
                        .OfClass(CeilingType).ToElements()
                    if not ceil_types:
                        t.RollBack()
                        return None
                    ct = ceil_types[0]
                    profile = List[CurveLoop]()
                    profile.Add(cl)
                    ceiling = Ceiling.Create(self.doc, profile, ct.Id, lv.Id)
                else:
                    ft = self._best_floor_type()
                    if not ft:
                        t.RollBack()
                        return None
                    ca = CurveArray()
                    for curve in cl:
                        ca.Append(curve)
                    ceiling = self.doc.Create.NewFloor(ca, ft, lv, False)
                t.Commit()
                return ceiling
            except Exception as ex:
                logger.error("build_ceiling: {}".format(ex))
                t.RollBack()
                return None

    def build_column(self, elem):
        """Place a structural (or architectural) column family instance."""
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None

        col_symbols = (FilteredElementCollector(self.doc)
                       .OfClass(FamilySymbol)
                       .OfCategory(BuiltInCategory.OST_StructuralColumns)
                       .ToElements())
        if not col_symbols:
            col_symbols = (FilteredElementCollector(self.doc)
                           .OfClass(FamilySymbol)
                           .OfCategory(BuiltInCategory.OST_Columns)
                           .ToElements())
        if not col_symbols:
            return None

        sym = col_symbols[0]
        pt  = XYZ(d['cx'], d['cy'], d['z_bot_ft'])

        with Transaction(self.doc, "T3Lab: Create Column") as t:
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                col = self.doc.Create.NewFamilyInstance(
                    pt, sym, lv, StructuralType.Column)
                t.Commit()
                return col
            except Exception as ex:
                logger.error("build_column: {}".format(ex))
                t.RollBack()
                return None

    def build_opening(self, elem, host_wall_revit_elem):
        """Place a door or window family instance in the given host wall."""
        if host_wall_revit_elem is None:
            return None
        d     = elem._data
        angle = d['host_wall_data']['angle']
        cx    = d['host_wall_data']['cx']
        cy    = d['host_wall_data']['cy']
        u_c   = d['u_center']

        ox = cx + math.cos(angle) * u_c
        oy = cy + math.sin(angle) * u_c
        oz = d['w_bottom'] + d['height_ft'] / 2.0
        pt = XYZ(ox, oy, oz)

        cat = (BuiltInCategory.OST_Doors
               if elem.Type == 'Door' else BuiltInCategory.OST_Windows)
        syms = (FilteredElementCollector(self.doc)
                .OfClass(FamilySymbol)
                .OfCategory(cat)
                .ToElements())
        if not syms:
            return None
        sym = syms[0]
        lv  = self._get_level(elem.LevelName)
        if not lv:
            return None

        with Transaction(self.doc, "T3Lab: Create {}".format(elem.Type)) as t:
            t.Start()
            try:
                if not sym.IsActive:
                    sym.Activate()
                try:
                    from Autodesk.Revit.DB.Structure import StructuralType as ST
                    non_structural = ST.NonStructural
                except ImportError:
                    non_structural = StructuralType.NonStructural
                inst = self.doc.Create.NewFamilyInstance(
                    pt, sym, host_wall_revit_elem, lv, non_structural)
                t.Commit()
                return inst
            except Exception as ex:
                logger.error("build_opening: {}".format(ex))
                t.RollBack()
                return None

    def build_roof(self, elem):
        """Create a basic FootPrintRoof from the detected boundary."""
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None

        roof_types = (FilteredElementCollector(self.doc)
                      .OfCategory(BuiltInCategory.OST_Roofs)
                      .OfClass(DB.RoofType)
                      .ToElements())
        if not roof_types:
            return None
        rt = roof_types[0]

        z_ft    = d['z_ft']
        corners = d['corners_xy']
        pts     = [XYZ(x, y, z_ft) for (x, y) in corners]
        n       = len(pts)

        with Transaction(self.doc, "T3Lab: Create Roof") as t:
            t.Start()
            try:
                ca = CurveArray()
                for i in range(n):
                    p1 = pts[i]
                    p2 = pts[(i + 1) % n]
                    if p1.DistanceTo(p2) > 0.01:
                        ca.Append(Line.CreateBound(p1, p2))
                ma   = DB.ModelCurveArray()
                roof = self.doc.Create.NewFootPrintRoof(ca, lv, rt, ma)
                t.Commit()
                return roof
            except Exception as ex:
                logger.error("build_roof: {}".format(ex))
                t.RollBack()
                return None

    def build_stair_marker(self, elem):
        """Stub — stair creation requires the Architecture API; skipped here."""
        return None

    def build_all(self, elements, default_wall_height_ft=None):
        """
        Create all selected elements in dependency order.
        Returns a summary dict {'counts': {...}, 'errors': int}.
        """
        if default_wall_height_ft is None:
            default_wall_height_ft = mm_to_ft(3000.0)

        counts = {k: 0 for k in ['Wall', 'Floor', 'Ceiling', 'Column',
                                  'Door', 'Window', 'Stair', 'Roof']}
        errors        = 0
        created_walls = {}   # list index → Revit Wall element

        # Pass 1: walls
        for i, elem in enumerate(elements):
            if not elem.Include:
                continue
            if elem.Type == 'Wall':
                w = self.build_wall(elem, default_wall_height_ft)
                if w:
                    counts['Wall'] += 1
                    created_walls[i] = w
                else:
                    errors += 1

        # Pass 2: floors, ceilings, columns, roof
        for elem in elements:
            if not elem.Include:
                continue
            if elem.Type == 'Floor':
                r = self.build_floor(elem)
                if r:   counts['Floor'] += 1
                else:   errors += 1
            elif elem.Type == 'Ceiling':
                r = self.build_ceiling(elem)
                if r:   counts['Ceiling'] += 1
                else:   errors += 1
            elif elem.Type == 'Column':
                r = self.build_column(elem)
                if r:   counts['Column'] += 1
                else:   errors += 1
            elif elem.Type == 'Roof':
                r = self.build_roof(elem)
                if r:   counts['Roof'] += 1
                else:   errors += 1
            elif elem.Type == 'Stair':
                counts['Stair'] += 1  # marker only

        # Pass 3: doors / windows (need a host wall)
        for elem in elements:
            if not elem.Include:
                continue
            if elem.Type not in ('Door', 'Window'):
                continue
            host = None
            for _i, rev_wall in created_walls.items():
                if rev_wall:
                    host = rev_wall
                    break
            r = self.build_opening(elem, host)
            if r:   counts[elem.Type] += 1
            else:   errors += 1

        return {'counts': counts, 'errors': errors}


# ── Section 7: WPF Wizard Window ──────────────────────────────────────────────

class PointCloudModelWindow(forms.WPFWindow):
    """Single-window UI for Point Cloud to Model analysis and element generation."""

    DENSITY_CAPS = [5000, 20000, 50000]

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self._pc_instance       = None
        self._custom_min_pt     = None
        self._custom_max_pt     = None
        self._detected_elements = []
        self._builder           = ElementBuilder(doc)
        self.result             = None

        self.status_count.Text = u"Revit {}".format(REVIT_VERSION)
        self._try_preselect_cloud()

    def _try_preselect_cloud(self):
        try:
            sel_ids = uidoc.Selection.GetElementIds()
            for eid in sel_ids:
                el = doc.GetElement(eid)
                if isinstance(el, PointCloudInstance):
                    self._set_cloud(el)
                    break
        except Exception:
            pass

    def _set_cloud(self, pc_inst):
        self._pc_instance = pc_inst
        try:
            name = pc_inst.Name or u"Point Cloud"
        except Exception:
            name = u"Point Cloud"
        self.lbl_cloud_name.Text       = name
        self.lbl_cloud_name.Foreground = self._brush('#0F172A')
        self.status_text.Text          = u"Cloud selected: {}".format(name)

    def _brush(self, hex_color):
        from System.Windows.Media import BrushConverter
        return BrushConverter().ConvertFromString(hex_color)

    # ── Window chrome ──────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        self.WindowState = (WindowState.Normal
                            if self.WindowState == WindowState.Maximized
                            else WindowState.Maximized)

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── Step 0 handlers ────────────────────────────────────────────────────────

    def btn_pick_cloud_clicked(self, sender, e):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                PointCloudSelectionFilter(),
                "Select a Point Cloud Instance")
            pc = doc.GetElement(ref.ElementId)
            self._set_cloud(pc)
        except Exception:
            pass
        self.Show()

    def rb_custom_region_checked(self, sender, e):
        self.btn_pick_region.IsEnabled = True

    def rb_full_extent_checked(self, sender, e):
        self.btn_pick_region.IsEnabled = False
        self._custom_min_pt = None
        self._custom_max_pt = None

    def btn_pick_region_clicked(self, sender, e):
        self.Hide()
        try:
            pt1 = uidoc.Selection.PickPoint("Pick first corner of scan region")
            pt2 = uidoc.Selection.PickPoint("Pick second corner of scan region")
            self._custom_min_pt = XYZ(
                min(pt1.X, pt2.X), min(pt1.Y, pt2.Y),
                min(pt1.Z, pt2.Z) - mm_to_ft(500.0))
            self._custom_max_pt = XYZ(
                max(pt1.X, pt2.X), max(pt1.Y, pt2.Y),
                max(pt1.Z, pt2.Z) + mm_to_ft(500.0))
            self.status_text.Text = u"Custom region defined."
        except Exception:
            pass
        self.Show()

    def btn_analyze_clicked(self, sender, e):
        if self._pc_instance is None:
            TaskDialog.Show("Point Cloud to Model",
                            "Please select a Point Cloud Instance first.")
            return
        if self.rb_custom_region.IsChecked and self._custom_min_pt is None:
            TaskDialog.Show("Point Cloud to Model",
                            "Custom Region selected but no corners picked.\n"
                            "Click 'Pick Region Corners' first, or use Full Cloud Extent.")
            return
        self.btn_analyze.IsEnabled = False
        self.status_text.Text = u"Extracting points from cloud…"
        try:
            self._run_analysis()
        except Exception as ex:
            logger.error("Analysis error: {}".format(ex))
            TaskDialog.Show("Point Cloud to Model",
                            u"Analysis error: {}".format(ex))
        finally:
            self.btn_analyze.IsEnabled = True

    def _default_settings(self):
        """Auto-detect all element types with standard tolerances."""
        return {
            'detect_wall':    True,
            'snap_tol':       1.0,
            'wall_min_len':   500.0,
            'detect_floor':   True,
            'floor_zbin':     50.0,
            'floor_min_area': 1.0,
            'detect_ceiling': True,
            'ceil_zbin':      50.0,
            'detect_door':    True,
            'door_min_w':     600.0,
            'door_min_h':     1800.0,
            'detect_window':  True,
            'win_min_w':      400.0,
            'win_min_h':      400.0,
            'detect_column':  True,
            'col_min':        100.0,
            'col_max':        600.0,
            'detect_stair':   True,
            'stair_riser':    165.0,
            'detect_roof':    True,
            'roof_slope':     5.0,
        }

    def _run_analysis(self):
        idx = self.cmb_density.SelectedIndex
        if idx < 0:
            idx = 1
        density_cap = self.DENSITY_CAPS[idx]

        self.status_text.Text = u"Extracting {} points…".format(density_cap)

        if self.rb_custom_region.IsChecked and self._custom_min_pt:
            min_pt = self._custom_min_pt
            max_pt = self._custom_max_pt
            half_x = (max_pt.X - min_pt.X) / 2.0
            half_y = (max_pt.Y - min_pt.Y) / 2.0
            half_z = (max_pt.Z - min_pt.Z) / 2.0
            center = XYZ(
                (min_pt.X + max_pt.X) / 2.0,
                (min_pt.Y + max_pt.Y) / 2.0,
                (min_pt.Z + max_pt.Z) / 2.0)
            pts = extract_full_cloud_from_region(
                self._pc_instance, center, half_x, half_y, half_z, density_cap)
        else:
            pts = extract_full_cloud(self._pc_instance, density_cap)

        if len(pts) < 50:
            TaskDialog.Show(
                "Point Cloud to Model",
                "Only {} points extracted. "
                "The cloud may be out of range or empty.".format(len(pts)))
            return

        self.status_text.Text = u"Analyzing {} points…".format(len(pts))
        settings = self._default_settings()
        analyzer = PointCloudAnalyzer(pts)
        self._detected_elements = analyzer.run(settings)

        self._populate_results()
        total = len(self._detected_elements)
        self.status_text.Text = (
            u"Detected {} elements. Review and click Generate.".format(total))

    def _populate_results(self):
        elems = self._detected_elements

        type_counts = {}
        for e in elems:
            type_counts[e.Type] = type_counts.get(e.Type, 0) + 1

        badge_map = {
            'Wall':    self.badge_wall,
            'Floor':   self.badge_floor,
            'Ceiling': self.badge_ceiling,
            'Door':    self.badge_door,
            'Window':  self.badge_window,
            'Column':  self.badge_column,
            'Stair':   self.badge_stair,
            'Roof':    self.badge_roof,
        }
        labels = {
            'Wall': 'Walls', 'Floor': 'Floors', 'Ceiling': 'Ceilings',
            'Door': 'Doors', 'Window': 'Windows', 'Column': 'Columns',
            'Stair': 'Stairs', 'Roof': 'Roofs',
        }
        for t, badge in badge_map.items():
            cnt = type_counts.get(t, 0)
            badge.Text = u"{} {}".format(cnt, labels[t])

        total = len(elems)
        self.pnl_empty_state.Visibility = Visibility.Collapsed
        if total > 0:
            self.results_grid.ItemsSource  = elems
            self.results_grid.Visibility   = Visibility.Visible
            self.pnl_no_results.Visibility = Visibility.Collapsed
            self.btn_generate.IsEnabled    = True
            self.btn_generate.Content      = u"Generate {} Selected".format(total)
        else:
            self.results_grid.Visibility   = Visibility.Collapsed
            self.pnl_no_results.Visibility = Visibility.Visible
            self.btn_generate.IsEnabled    = False
            self.btn_generate.Content      = u"Generate"

    # ── Step 2 handlers ────────────────────────────────────────────────────────

    def btn_generate_clicked(self, sender, e):
        selected = [el for el in self._detected_elements if el.Include]
        if not selected:
            TaskDialog.Show(
                "Point Cloud to Model",
                "No elements selected. Check the Include checkboxes.")
            return

        self.btn_generate.IsEnabled = False
        self.status_text.Text = u"Creating {} elements…".format(len(selected))
        try:
            summary = self._builder.build_all(selected)
            counts  = summary['counts']
            errors  = summary['errors']
            parts   = [u"{} {}".format(v, k)
                       for k, v in counts.items() if v > 0]
            msg = (u"Created: " + u", ".join(parts)) if parts else u"No elements created."
            if errors:
                msg += u"\n{} element(s) failed.".format(errors)
            TaskDialog.Show("Point Cloud to Model", msg)
            self.result = summary
            self.Close()
        except Exception as ex:
            logger.error("Generate error: {}".format(ex))
            TaskDialog.Show("Point Cloud to Model",
                            u"Generation error: {}".format(ex))
        finally:
            self.btn_generate.IsEnabled = True

    def btn_cancel_clicked(self, sender, e):
        self.result = None
        self.Close()


# MAIN
# ==============================================================================

if __name__ == '__main__':
    try:
        window = PointCloudModelWindow()
        window.ShowDialog()
    except SystemExit:
        pass
    except Exception as ex:
        logger.error("Point Cloud to Model error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
        TaskDialog.Show("Point Cloud to Model",
                        u"Unexpected error: {}".format(ex))
