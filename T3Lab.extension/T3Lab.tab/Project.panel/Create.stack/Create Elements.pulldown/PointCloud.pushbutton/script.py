# -*- coding: utf-8 -*-
"""
Point Cloud to Wall
-------------------
Scan-to-BIM: select a point cloud instance, click a sampling point,
and the tool fits a wall centerline via Least-Squares Regression,
auto-snaps to project axes/grids within 1 degree, matches wall types by
compound thickness, then creates the wall after user confirmation.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__title__   = "Point Cloud\nto Wall"
__author__  = "Tran Tien Thanh"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==============================================================================
import os
import sys
import clr
import math
import json

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState
from System.Collections.Generic import List

from rpw import revit, DB
from Autodesk.Revit.DB import (
    Transaction,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    WallType,
    Wall,
    Level,
    Grid,
    XYZ,
    Line,
    BoundingBoxXYZ,
    PointCloudInstance,
    IFailuresPreprocessor,
    FailureProcessingResult,
    ElementId
)
from Autodesk.Revit.DB.PointClouds import PointCloudFilter, PointCloudFilterFactory
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import forms, script

# Path setup
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR    = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))))
lib_dir    = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

XAML_FILE  = os.path.join(EXT_DIR, 'lib', 'GUI', 'Tools', 'PointCloudToWall.xaml')

# DEFINE VARIABLES
# ==============================================================================
logger        = script.get_logger()
doc           = revit.doc
uidoc         = revit.uidoc
REVIT_VERSION = int(doc.Application.VersionNumber)

# CLASS/FUNCTIONS
# ==============================================================================

# ── Unit conversion helpers ────────────────────────────────────────────────────

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


# ── Point cloud sampling constants ────────────────────────────────────────────

SAMPLE_XY_MM = 1000.0   # half-size in XY (total box is 2 x this per side)
SAMPLE_Z_MM  = 3000.0   # full height of sample box
MAX_POINTS   = 30000


# ── ISelectionFilter for PointCloudInstance ───────────────────────────────────

class PointCloudSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, PointCloudInstance)

    def AllowReference(self, reference, position):
        return False


# ── Geometry analysis ─────────────────────────────────────────────────────────

def build_point_cloud_filter(center_pt, half_xy_ft, half_z_ft):
    """
    Returns a PointCloudFilter for an axis-aligned box.
    Uses PointCloudFilterFactory.CreateMultiPlaneFilter with 6 half-space planes.
    Each plane is defined by a point on the plane and an inward normal (pointing
    into the box so the intersection of all half-spaces equals the box interior).
    """
    x = center_pt.X
    y = center_pt.Y
    z = center_pt.Z

    normals = List[XYZ]()
    points  = List[XYZ]()

    # +X face — inward normal points in -X direction
    normals.Add(XYZ(-1.0,  0.0,  0.0))
    points.Add( XYZ(x + half_xy_ft, y, z))

    # -X face — inward normal points in +X direction
    normals.Add(XYZ( 1.0,  0.0,  0.0))
    points.Add( XYZ(x - half_xy_ft, y, z))

    # +Y face — inward normal points in -Y direction
    normals.Add(XYZ( 0.0, -1.0,  0.0))
    points.Add( XYZ(x, y + half_xy_ft, z))

    # -Y face — inward normal points in +Y direction
    normals.Add(XYZ( 0.0,  1.0,  0.0))
    points.Add( XYZ(x, y - half_xy_ft, z))

    # +Z face — inward normal points in -Z direction
    normals.Add(XYZ( 0.0,  0.0, -1.0))
    points.Add( XYZ(x, y, z + half_z_ft))

    # -Z face — inward normal points in +Z direction
    normals.Add(XYZ( 0.0,  0.0,  1.0))
    points.Add( XYZ(x, y, z - half_z_ft))

    return PointCloudFilterFactory.CreateMultiPlaneFilter(points, normals)


def extract_points(pc_instance, center_pt, half_xy_ft, half_z_ft):
    """
    Returns list of (x, y, z) tuples in Revit world coordinates (feet).
    Caps at MAX_POINTS. Returns empty list on any error.
    The minimum-distance argument (1.0/304.8) requests at most 1 mm spacing.
    """
    try:
        pcf = build_point_cloud_filter(center_pt, half_xy_ft, half_z_ft)
        raw = pc_instance.GetPoints(pcf, 1.0 / 304.8, MAX_POINTS)
        pts = []
        for p in raw:
            pts.append((p.X, p.Y, p.Z))
        return pts
    except Exception as ex:
        logger.debug("extract_points error: {}".format(ex))
        return []


def fit_wall_centerline(pts):
    """
    Fits a 2D line to the (x, y) projection of pts using the principal-axis
    method (covariance eigenvector). This avoids the vertical-slope singularity
    of ordinary y-on-x regression.

    Returns (angle_rad, centroid_x, centroid_y) or None if < 10 points.
    angle_rad is measured from the +X axis in the XY plane.
    """
    if len(pts) < 10:
        return None

    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    n = len(xs)
    mean_x = sum(xs) / n
    mean_y = sum(ys) / n

    sxx = sum((x - mean_x) ** 2 for x in xs)
    sxy = sum((xs[i] - mean_x) * (ys[i] - mean_y) for i in range(n))
    syy = sum((y - mean_y) ** 2 for y in ys)

    # Principal axis angle via covariance matrix eigenvector
    diff = sxx - syy
    angle_rad = 0.5 * math.atan2(2.0 * sxy, diff)
    return angle_rad, mean_x, mean_y


def calc_thickness(pts, angle_rad, centroid_x, centroid_y):
    """
    Projects all points onto the wall-normal axis (perpendicular to the wall
    direction) and returns the spread between the 2.5th and 97.5th percentiles
    as the estimated wall thickness (in feet, matching Revit internal units).
    """
    if not pts:
        return 0.0

    nx = -math.sin(angle_rad)
    ny =  math.cos(angle_rad)

    projections = [
        (p[0] - centroid_x) * nx + (p[1] - centroid_y) * ny
        for p in pts
    ]
    projections.sort()
    n = len(projections)
    lo_idx = int(round(0.025 * n))
    hi_idx = int(round(0.975 * n)) - 1
    if hi_idx <= lo_idx:
        return abs(projections[-1] - projections[0])
    return abs(projections[hi_idx] - projections[lo_idx])


# ── Grid / axis snap logic ─────────────────────────────────────────────────────

SNAP_TOLERANCE_DEG = 1.0


def collect_grid_angles(document):
    """
    Returns a list of angles (radians, normalized to [0, pi)) derived from
    all project grids plus the cardinal axes (0 deg and 90 deg).
    """
    angles = [0.0, math.pi / 2.0]
    try:
        grids = FilteredElementCollector(document).OfClass(Grid).ToElements()
        for g in grids:
            curve = g.Curve
            if curve is None:
                continue
            d = curve.Direction
            a = math.atan2(d.Y, d.X)
            # Normalize to [0, pi)
            while a < 0:
                a += math.pi
            while a >= math.pi:
                a -= math.pi
            angles.append(a)
    except Exception:
        pass
    return angles


def snap_angle(angle_rad, grid_angles, tol_deg=SNAP_TOLERANCE_DEG):
    """
    Checks whether angle_rad is within tol_deg of any reference angle or its
    perpendicular. Returns (snapped_angle_rad, snapped_bool, description_str).
    All angles are normalized to [0, pi) before comparison.
    """
    tol_rad = math.radians(tol_deg)
    a = angle_rad % math.pi

    for ref in grid_angles:
        r = ref % math.pi
        for candidate in [r, (r + math.pi / 2.0) % math.pi]:
            diff = abs(a - candidate)
            if diff > math.pi / 2.0:
                diff = math.pi - diff
            if diff <= tol_rad:
                desc = u"Snapped to {:.1f}°".format(math.degrees(candidate))
                return candidate, True, desc

    return a, False, u"No snap ({:.2f}°)".format(math.degrees(a))


# ── Wall type matching ─────────────────────────────────────────────────────────

def get_wall_types_sorted(document):
    """
    Collects all WallType elements and returns a list of
    (display_name_str, wall_type_element, thickness_mm_float) tuples,
    sorted by thickness ascending. Types where thickness cannot be
    determined are skipped.
    """
    wt_elements = FilteredElementCollector(document).OfClass(WallType).ToElements()
    result = []
    for wt in wt_elements:
        try:
            cs = wt.GetCompoundStructure()
            if cs is None:
                width_ft = wt.Width
            else:
                width_ft = cs.GetTotalWidth()
            width_mm = internal_to_mm(width_ft)
            name = wt.Name if wt.Name else u""
            result.append((
                u"{} ({:.0f} mm)".format(name, width_mm),
                wt,
                width_mm
            ))
        except Exception:
            pass
    result.sort(key=lambda x: x[2])
    return result


def find_best_wall_type(sorted_types, target_mm, tolerance_mm=20.0):
    """
    Finds the WallType in sorted_types whose thickness is closest to target_mm.
    Returns (index_in_list, within_tolerance_bool).
    """
    if not sorted_types:
        return 0, False
    best_idx = 0
    best_diff = abs(sorted_types[0][2] - target_mm)
    for i, (_, _, t_mm) in enumerate(sorted_types):
        d = abs(t_mm - target_mm)
        if d < best_diff:
            best_diff = d
            best_idx = i
    return best_idx, best_diff <= tolerance_mm


# ── WPF Window ────────────────────────────────────────────────────────────────

class PointCloudToWallWindow(forms.WPFWindow):
    """
    Verification UI for Scan-to-BIM wall generation.
    Shows detected geometry values, lets the user pick wall type / level /
    height, then either generates the wall or cancels.
    """

    def __init__(self, analysis_result):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.result = None          # populated with a dict on Generate, stays None on Cancel
        self._analysis = analysis_result
        self._wall_types = []
        self._level_map  = {}
        self._populate(analysis_result)

    # ── populate ──────────────────────────────────────────────────────────────

    def _populate(self, ar):
        """Fill all controls from the analysis_result dict."""
        thickness_mm = ar.get('thickness_mm', 0.0)
        angle_deg    = math.degrees(ar.get('angle_rad', 0.0))
        snapped      = ar.get('snapped', False)
        snap_desc    = ar.get('snap_desc', u'')
        pt_count     = ar.get('point_count', 0)
        sample_mm    = ar.get('sample_xy_mm', SAMPLE_XY_MM)

        self.lbl_thickness.Text = u"{:.0f} mm".format(thickness_mm)
        self.lbl_angle.Text     = u"{:.2f}°".format(angle_deg)
        self.lbl_snap.Text      = snap_desc
        self.lbl_snap.Foreground = (
            self._brush('#15803D') if snapped else self._brush('#64748B')
        )
        self.lbl_points.Text = u"{:,} pts".format(pt_count)
        self.lbl_zone.Text   = u"{:.0f} × {:.0f} mm".format(sample_mm, sample_mm)

        # Status bar version info
        self.status_version.Text = u"Revit {}".format(REVIT_VERSION)

        # Wall types
        self._wall_types = get_wall_types_sorted(doc)
        for (display, _, _) in self._wall_types:
            self.cmb_wall_type.Items.Add(display)
        best_idx, within_tol = find_best_wall_type(self._wall_types, thickness_mm)
        if self._wall_types:
            self.cmb_wall_type.SelectedIndex = best_idx
        if not within_tol:
            import System.Windows
            self.lbl_type_warning.Visibility = System.Windows.Visibility.Visible

        # Levels — sorted by elevation
        levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        sorted_levels = sorted(levels, key=lambda lv: lv.Elevation)
        for lv in sorted_levels:
            self._level_map[lv.Name] = lv
            self.cmb_base_level.Items.Add(lv.Name)
        if self.cmb_base_level.Items.Count > 0:
            self.cmb_base_level.SelectedIndex = 0

        self.status_text.Text = u"Ready — review detected values and click Generate Wall"

    def _brush(self, hex_color):
        from System.Windows.Media import BrushConverter
        bc = BrushConverter()
        return bc.ConvertFromString(hex_color)

    # ── Window chrome handlers ────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        self.WindowState = (
            WindowState.Normal
            if self.WindowState == WindowState.Maximized
            else WindowState.Maximized
        )

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── Action handlers ───────────────────────────────────────────────────────

    def cancel_clicked(self, sender, e):
        self.result = None
        self.Close()

    def generate_clicked(self, sender, e):
        # Validate wall type
        idx = self.cmb_wall_type.SelectedIndex
        if idx < 0 or idx >= len(self._wall_types):
            TaskDialog.Show("Point Cloud to Wall", "Please select a Wall Type.")
            return
        _, wall_type_elem, _ = self._wall_types[idx]

        # Validate level
        level_name = self.cmb_base_level.SelectedItem
        if not level_name or level_name not in self._level_map:
            TaskDialog.Show("Point Cloud to Wall", "Please select a Base Level.")
            return
        level_elem = self._level_map[level_name]

        # Validate height
        try:
            height_mm = float(self.txt_height.Text)
            if height_mm <= 0:
                raise ValueError("height must be > 0")
        except (ValueError, TypeError):
            TaskDialog.Show(
                "Point Cloud to Wall",
                "Invalid wall height. Enter a positive number in mm."
            )
            return

        structural = bool(self.chk_structural.IsChecked)

        self.result = {
            'wall_type':  wall_type_elem,
            'level':      level_elem,
            'height_ft':  mm_to_ft(height_mm),
            'structural': structural,
        }
        self.Close()


# ── Wall generator ────────────────────────────────────────────────────────────

class WallGenerator(object):
    """Wraps the Revit Wall.Create call in a named Transaction."""

    def __init__(self, document):
        self.doc = document

    def create_wall(self, center_x, center_y, angle_rad, length_ft,
                    wall_type, level, height_ft, structural=False):
        """
        Builds a wall from center point, direction angle, and length.
        All positional coordinates are in Revit internal units (feet).
        Returns the new Wall element or None on failure.

        Wall.Create signature (Revit 2022+):
            Wall.Create(doc, curve, wallTypeId, levelId, height, offset, flip, structural)
        """
        half = length_ft / 2.0
        dx = math.cos(angle_rad) * half
        dy = math.sin(angle_rad) * half
        pt1 = XYZ(center_x - dx, center_y - dy, 0.0)
        pt2 = XYZ(center_x + dx, center_y + dy, 0.0)
        wall_line = Line.CreateBound(pt1, pt2)

        new_wall = None
        with Transaction(self.doc, "T3Lab: Point Cloud to Wall") as t:
            t.Start()
            try:
                new_wall = Wall.Create(
                    self.doc,
                    wall_line,
                    wall_type.Id,
                    level.Id,
                    height_ft,
                    0.0,        # base offset from level
                    False,      # flip face
                    structural  # structural usage
                )
                t.Commit()
            except Exception as ex:
                logger.error("Wall.Create failed: {}".format(ex))
                t.RollBack()
        return new_wall


# MAIN
# ==============================================================================

if __name__ == '__main__':
    try:

        # ── Step 1: Obtain point cloud instance ───────────────────────────────
        # Try to reuse an element already selected in the canvas.
        sel_ids = uidoc.Selection.GetElementIds()
        pc_instance = None
        for eid in sel_ids:
            el = doc.GetElement(eid)
            if isinstance(el, PointCloudInstance):
                pc_instance = el
                break

        if pc_instance is None:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    PointCloudSelectionFilter(),
                    "Select a Point Cloud Instance"
                )
                pc_instance = doc.GetElement(ref.ElementId)
            except Exception:
                raise SystemExit

        # ── Step 2: Ask the user to click a sampling location ─────────────────
        try:
            sample_pt = uidoc.Selection.PickPoint(
                "Click on the wall location in the point cloud"
            )
        except Exception:
            raise SystemExit

        # ── Step 3: Extract points from the sampling box ──────────────────────
        half_xy = mm_to_ft(SAMPLE_XY_MM / 2.0)
        half_z  = mm_to_ft(SAMPLE_Z_MM  / 2.0)
        pts = extract_points(pc_instance, sample_pt, half_xy, half_z)

        if len(pts) < 10:
            TaskDialog.Show(
                "Point Cloud to Wall",
                "Not enough points in sampling zone ({} found). "
                "Try a different location.".format(len(pts))
            )
            raise SystemExit

        # ── Step 4: Analyse geometry ──────────────────────────────────────────
        fit = fit_wall_centerline(pts)
        if fit is None:
            TaskDialog.Show(
                "Point Cloud to Wall",
                "Could not fit a wall line to the point cloud data."
            )
            raise SystemExit

        raw_angle, cx, cy = fit
        thickness_ft = calc_thickness(pts, raw_angle, cx, cy)
        thickness_mm = internal_to_mm(thickness_ft)

        grid_angles = collect_grid_angles(doc)
        snapped_angle, snapped, snap_desc = snap_angle(raw_angle, grid_angles)

        # ── Step 5: Show confirmation window ──────────────────────────────────
        ar = {
            'thickness_mm': thickness_mm,
            'angle_rad':    snapped_angle,
            'snapped':      snapped,
            'snap_desc':    snap_desc,
            'point_count':  len(pts),
            'sample_xy_mm': SAMPLE_XY_MM,
            'centroid_x':   cx,
            'centroid_y':   cy,
        }

        window = PointCloudToWallWindow(ar)
        window.ShowDialog()

        if window.result is None:
            raise SystemExit

        # ── Step 6: Estimate wall length from point spread ────────────────────
        res = window.result
        angle_rad = ar['angle_rad']
        dx_unit = math.cos(angle_rad)
        dy_unit = math.sin(angle_rad)

        projections_along = [
            (p[0] - cx) * dx_unit + (p[1] - cy) * dy_unit
            for p in pts
        ]
        proj_sorted = sorted(projections_along)
        n = len(proj_sorted)
        lo_i = int(round(0.01 * n))
        hi_i = int(round(0.99 * n)) - 1

        if hi_i > lo_i:
            raw_len = abs(proj_sorted[hi_i] - proj_sorted[lo_i])
        else:
            raw_len = mm_to_ft(1000.0)

        wall_len_ft = max(raw_len, mm_to_ft(1000.0))

        # ── Step 7: Create wall ───────────────────────────────────────────────
        gen = WallGenerator(doc)
        new_wall = gen.create_wall(
            cx, cy,
            angle_rad,
            wall_len_ft,
            res['wall_type'],
            res['level'],
            res['height_ft'],
            res['structural']
        )

        if new_wall:
            try:
                uidoc.Selection.SetElementIds(List[ElementId]([new_wall.Id]))
            except Exception:
                pass
            TaskDialog.Show(
                "Point Cloud to Wall",
                u"Wall created successfully.\n"
                u"Type: {}\n"
                u"Thickness: {:.0f} mm\n"
                u"Height: {:.0f} mm".format(
                    res['wall_type'].Name,
                    thickness_mm,
                    internal_to_mm(res['height_ft'])
                )
            )
        else:
            TaskDialog.Show(
                "Point Cloud to Wall",
                "Wall creation failed. Check the Output Window for details."
            )

    except SystemExit:
        pass
    except Exception as ex:
        logger.error("Point Cloud to Wall error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
        TaskDialog.Show(
            "Point Cloud to Wall",
            u"Unexpected error: {}".format(ex)
        )
