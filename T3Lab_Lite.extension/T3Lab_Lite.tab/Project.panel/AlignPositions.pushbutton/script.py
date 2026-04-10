# -*- coding: utf-8 -*-
"""
Align Positions
---------------
Snap element positions so their distance to a reference element
(Grid, Wall, or Column) becomes a clean multiple of 5 or 10 mm.

Workflow
  1. Pick a reference element (Grid / Wall / Column)
  2. Select surrounding elements
  3. Review the preview table
  4. Apply corrections

Author: T3Lab (Tran Tien Thanh)
"""

__title__  = "Align\nPositions"
__author__ = "T3Lab"

import os
import math
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Data')

from System.Windows import WindowState
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Boolean
from System.Data import DataTable

from Autodesk.Revit.DB import (
    XYZ,
    Transaction,
    ElementTransformUtils,
    Grid,
    Wall,
    FamilyInstance,
    BuiltInCategory,
    LocationPoint,
    LocationCurve,
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, script

doc    = revit.doc
uidoc  = revit.uidoc
logger = script.get_logger()

# ════════════════════════════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════════════════════════════
MM_PER_FOOT = 304.8
TOLERANCE_MM = 0.05        # ignore corrections smaller than this


def feet_to_mm(feet):
    return feet * MM_PER_FOOT


def mm_to_feet(mm):
    return mm / MM_PER_FOOT


def round_to_snap(value_mm, snap_mm):
    """Round *value_mm* to the nearest multiple of *snap_mm*."""
    return round(value_mm / snap_mm) * snap_mm


# ════════════════════════════════════════════════════════════════
# SELECTION FILTER
# ════════════════════════════════════════════════════════════════
class ReferenceFilter(ISelectionFilter):
    """Allow only Grid, Wall, or Column elements."""

    def AllowElement(self, element):
        if isinstance(element, Grid):
            return True
        if isinstance(element, Wall):
            return True
        if isinstance(element, FamilyInstance):
            cat = element.Category
            if cat is None:
                return False
            cat_id = cat.Id.IntegerValue
            if cat_id == int(BuiltInCategory.OST_Columns):
                return True
            if cat_id == int(BuiltInCategory.OST_StructuralColumns):
                return True
        return False

    def AllowReference(self, reference, position):
        return False


# ════════════════════════════════════════════════════════════════
# GEOMETRY HELPERS
# ════════════════════════════════════════════════════════════════
def get_element_location(element):
    """Return an XYZ location point for any element."""
    loc = element.Location
    if loc is not None:
        if isinstance(loc, LocationPoint):
            return loc.Point
        if isinstance(loc, LocationCurve):
            return loc.Curve.Evaluate(0.5, True)

    # Fallback: bounding-box centre
    bb = element.get_BoundingBox(None)
    if bb is not None:
        return XYZ(
            (bb.Min.X + bb.Max.X) / 2.0,
            (bb.Min.Y + bb.Max.Y) / 2.0,
            (bb.Min.Z + bb.Max.Z) / 2.0,
        )
    return None


def project_onto_curve(point, curve):
    """Project *point* onto *curve*; return the projected XYZ or None."""
    result = curve.Project(point)
    if result is not None:
        return result.XYZPoint
    return None


# ════════════════════════════════════════════════════════════════
# CORRECTION CALCULATORS  (one per reference type)
# ════════════════════════════════════════════════════════════════
def _correction_for_grid(elem, elem_loc, curve, direction, snap_mm):
    """Perpendicular distance from *elem_loc* to the Grid line."""
    projected = project_onto_curve(elem_loc, curve)
    if projected is None:
        return None

    # perpendicular unit vector (in XY plane)
    perp = XYZ(-direction.Y, direction.X, 0)

    signed_feet = (elem_loc - projected).DotProduct(perp)
    signed_mm   = feet_to_mm(signed_feet)
    rounded_mm  = round_to_snap(signed_mm, snap_mm)
    corr_mm     = rounded_mm - signed_mm

    if abs(corr_mm) < TOLERANCE_MM:
        return None

    corr_feet = mm_to_feet(corr_mm)
    return dict(
        element        = elem,
        current_dist_mm = signed_mm,
        rounded_dist_mm = rounded_mm,
        correction_mm   = corr_mm,
        move_vector     = XYZ(perp.X * corr_feet, perp.Y * corr_feet, 0),
    )


def _correction_for_wall(elem, elem_loc, curve, direction, half_w, snap_mm):
    """Distance from *elem_loc* to the nearest wall face."""
    projected = project_onto_curve(elem_loc, curve)
    if projected is None:
        return None

    perp = XYZ(-direction.Y, direction.X, 0)
    dist_to_cl = (elem_loc - projected).DotProduct(perp)  # centre-line

    # Signed distance to nearest face
    if dist_to_cl >= 0:
        face_feet = dist_to_cl - half_w
    else:
        face_feet = dist_to_cl + half_w

    face_mm    = feet_to_mm(face_feet)
    rounded_mm = round_to_snap(face_mm, snap_mm)
    corr_mm    = rounded_mm - face_mm

    if abs(corr_mm) < TOLERANCE_MM:
        return None

    corr_feet = mm_to_feet(corr_mm)
    return dict(
        element        = elem,
        current_dist_mm = face_mm,
        rounded_dist_mm = rounded_mm,
        correction_mm   = corr_mm,
        move_vector     = XYZ(perp.X * corr_feet, perp.Y * corr_feet, 0),
    )


def _correction_for_column(elem, elem_loc, bb_min, bb_max, snap_mm):
    """Distance from *elem_loc* to nearest column bounding-box face (X+Y)."""
    # --- X axis: pick nearest face ---
    dx_min = feet_to_mm(elem_loc.X - bb_min.X)
    dx_max = feet_to_mm(elem_loc.X - bb_max.X)
    dx = dx_min if abs(dx_min) <= abs(dx_max) else dx_max

    rx = round_to_snap(dx, snap_mm)
    cx = rx - dx

    # --- Y axis: pick nearest face ---
    dy_min = feet_to_mm(elem_loc.Y - bb_min.Y)
    dy_max = feet_to_mm(elem_loc.Y - bb_max.Y)
    dy = dy_min if abs(dy_min) <= abs(dy_max) else dy_max

    ry = round_to_snap(dy, snap_mm)
    cy = ry - dy

    total = math.sqrt(cx * cx + cy * cy)
    if total < TOLERANCE_MM:
        return None

    return dict(
        element        = elem,
        current_dist_mm = math.sqrt(dx * dx + dy * dy),
        rounded_dist_mm = math.sqrt(rx * rx + ry * ry),
        correction_mm   = total,
        move_vector     = XYZ(mm_to_feet(cx), mm_to_feet(cy), 0),
    )


# ════════════════════════════════════════════════════════════════
# WPF WINDOW
# ════════════════════════════════════════════════════════════════
_GUI_DIR  = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))),
    'lib', 'GUI',
)
_XAML     = os.path.join(_GUI_DIR, 'AlignPositions.xaml')
_LOGO     = os.path.join(_GUI_DIR, 'T3Lab_logo.png')


class AlignPositionsWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)

        # State
        self._ref_element  = None
        self._ref_type     = None   # "Grid" | "Wall" | "Column"
        self._ref_data     = None   # tuple of geometry data
        self._selected_elements = []
        self._results      = []

        # DataTable for the DataGrid
        self._dt = DataTable()
        self._dt.Columns.Add("Apply",     Boolean)
        self._dt.Columns.Add("Category")
        self._dt.Columns.Add("Name")
        self._dt.Columns.Add("ElementId")
        self._dt.Columns.Add("Distance")
        self._dt.Columns.Add("Rounded")
        self._dt.Columns.Add("Correction")
        self.dg_elements.ItemsSource = self._dt.DefaultView

        # Initial button states
        self.btn_select.IsEnabled = False
        self.btn_apply.IsEnabled  = False

        self._load_logo()
        self._update_status("Ready - Pick a reference element to start")

    # ── logo & chrome ────────────────────────────────────────
    def _load_logo(self):
        try:
            if os.path.exists(_LOGO):
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(_LOGO, UriKind.Absolute)
                bmp.EndInit()
                self.logo_image.Source = bmp
                self.Icon = bmp
        except Exception:
            pass

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── helpers ──────────────────────────────────────────────
    def _update_status(self, text):
        self.status_text.Text = text

    def _snap_mm(self):
        return 10.0 if self.rb_snap10.IsChecked else 5.0

    # ── pick reference ───────────────────────────────────────
    def pick_reference_clicked(self, sender, e):
        self.Hide()
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                ReferenceFilter(),
                "Pick a reference element (Grid, Wall, or Column)",
            )
            self._set_reference(doc.GetElement(ref.ElementId))
        except OperationCanceledException:
            pass
        except Exception as ex:
            logger.error(str(ex))
        self.Show()

    def _set_reference(self, element):
        self._ref_element = element

        if isinstance(element, Grid):
            self._ref_type = "Grid"
            curve = element.Curve
            d = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            self._ref_data = (curve, d)
            self.txt_ref_info.Text = "Grid: {}".format(element.Name)
            self.txt_ref_info.Foreground = \
                SolidColorBrush(Color.FromRgb(44, 62, 80))

        elif isinstance(element, Wall):
            self._ref_type = "Wall"
            curve = element.Location.Curve
            d = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            hw = element.Width / 2.0
            self._ref_data = (curve, d, hw)
            w_mm = element.Width * MM_PER_FOOT
            self.txt_ref_info.Text = "Wall: Id {} (w={:.0f} mm)".format(
                element.Id.IntegerValue, w_mm)
            self.txt_ref_info.Foreground = \
                SolidColorBrush(Color.FromRgb(44, 62, 80))

        elif isinstance(element, FamilyInstance):
            self._ref_type = "Column"
            bb = element.get_BoundingBox(None)
            self._ref_data = (bb.Min, bb.Max)
            self.txt_ref_info.Text = "Column: {} (Id:{})".format(
                element.Symbol.Family.Name, element.Id.IntegerValue)
            self.txt_ref_info.Foreground = \
                SolidColorBrush(Color.FromRgb(44, 62, 80))

        # Reset downstream state
        self.btn_select.IsEnabled = True
        self._dt.Clear()
        self._results = []
        self._selected_elements = []
        self.btn_apply.IsEnabled = False
        self.txt_element_count.Text = "No elements selected"
        self._update_status("Reference set - Select elements to analyze")

    # ── select elements ──────────────────────────────────────
    def select_elements_clicked(self, sender, e):
        self.Hide()
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select elements to snap (finish with Enter)",
            )
            elements = [doc.GetElement(r.ElementId) for r in refs]
            self._analyze(elements)
        except OperationCanceledException:
            pass
        except Exception as ex:
            logger.error(str(ex))
        self.Show()

    # ── snap-value changed ───────────────────────────────────
    def snap_changed(self, sender, e):
        if self._selected_elements and self._ref_element:
            self._analyze(self._selected_elements)

    # ── analysis ─────────────────────────────────────────────
    def _analyze(self, elements):
        self._selected_elements = list(elements)
        self._dt.Clear()
        self._results = []
        snap = self._snap_mm()
        count = 0

        for elem in elements:
            if elem.Id == self._ref_element.Id:
                continue
            result = self._compute(elem, snap)
            if result is None:
                continue

            self._results.append(result)
            row = self._dt.NewRow()
            row["Apply"]     = True
            row["Category"]  = elem.Category.Name if elem.Category else "N/A"
            row["Name"]      = getattr(elem, 'Name', '') \
                               or "Id:{}".format(elem.Id.IntegerValue)
            row["ElementId"] = str(elem.Id.IntegerValue)
            row["Distance"]  = "{:.1f}".format(result["current_dist_mm"])
            row["Rounded"]   = "{:.0f}".format(result["rounded_dist_mm"])
            row["Correction"] = "{:+.1f}".format(result["correction_mm"])
            self._dt.Rows.Add(row)
            count += 1

        self.btn_apply.IsEnabled = count > 0
        self.txt_element_count.Text = "{} element(s) need adjustment".format(count)
        self._update_status(
            "{} selected, {} need correction (snap {} mm)".format(
                len(elements), count, int(snap)))

    def _compute(self, elem, snap_mm):
        loc = get_element_location(elem)
        if loc is None:
            return None

        if self._ref_type == "Grid":
            curve, d = self._ref_data
            return _correction_for_grid(elem, loc, curve, d, snap_mm)

        if self._ref_type == "Wall":
            curve, d, hw = self._ref_data
            return _correction_for_wall(elem, loc, curve, d, hw, snap_mm)

        if self._ref_type == "Column":
            bb_min, bb_max = self._ref_data
            return _correction_for_column(elem, loc, bb_min, bb_max, snap_mm)

        return None

    # ── apply ────────────────────────────────────────────────
    def apply_clicked(self, sender, e):
        if not self._results:
            return

        moved  = 0
        failed = 0

        t = Transaction(doc, "Align Positions")
        t.Start()
        try:
            for i, result in enumerate(self._results):
                # Respect the Apply checkbox
                if i < self._dt.Rows.Count:
                    if not self._dt.Rows[i]["Apply"]:
                        continue

                try:
                    ElementTransformUtils.MoveElement(
                        doc,
                        result["element"].Id,
                        result["move_vector"],
                    )
                    moved += 1
                except Exception as ex:
                    logger.warning(
                        "Cannot move Id {}: {}".format(
                            result["element"].Id.IntegerValue, ex))
                    failed += 1

            t.Commit()
        except Exception as ex:
            t.RollBack()
            self._update_status("Error: {}".format(ex))
            return

        msg = "Done! {} element(s) moved".format(moved)
        if failed:
            msg += ", {} failed".format(failed)
        self._update_status(msg)

        # Clear table
        self._dt.Clear()
        self._results = []
        self._selected_elements = []
        self.btn_apply.IsEnabled = False
        self.txt_element_count.Text = "{} element(s) moved".format(moved)


# ════════════════════════════════════════════════════════════════
# ENTRY POINT
# ════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    window = AlignPositionsWindow()
    window.ShowDialog()
