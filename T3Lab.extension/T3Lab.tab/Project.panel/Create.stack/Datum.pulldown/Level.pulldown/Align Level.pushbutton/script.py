# -*- coding: utf-8 -*-
"""
Align Level 2D Extents v1.0
Align level start/end to a reference element (Level, Detail Line,
Model Line, Reference Plane, Grid).

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Align\nLevels"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Align 2D level extents to a reference line."

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
import System
from System.Windows import (
    Window, WindowStartupLocation, Thickness,
    HorizontalAlignment, VerticalAlignment, TextWrapping,
    GridLength, GridUnitType, FontWeights, CornerRadius
)
from System.Windows.Controls import (
    StackPanel, Button, TextBlock, RadioButton, GroupBox,
    Border, CheckBox, Grid as WPFGrid, ColumnDefinition
)
from System.Windows.Media import SolidColorBrush, Color, Colors, FontFamily

# ─── Revit Context ──────────────────────────────────────────
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView

# ─── DQT Color Palette ─────────────────────────────────────
DQT_PRIMARY   = Color.FromRgb(0xF0, 0xCC, 0x88)   # #0F172A Gold
DQT_ACCENT    = Color.FromRgb(0xC8, 0x96, 0x50)   # #C89650 Dark Gold
DQT_BG        = Color.FromRgb(0xFE, 0xF8, 0xE7)   # #F8FAFC Cream
DQT_DARK      = Color.FromRgb(0x3C, 0x3C, 0x3C)   # #3C3C3C Dark
DQT_WHITE     = Colors.White
DQT_BORDER    = Color.FromRgb(0xDD, 0xDD, 0xDD)   # #DDDDDD
DQT_TEXT_DARK = Color.FromRgb(0x33, 0x33, 0x33)   # #0F172A

def B(color):
    return SolidColorBrush(color)


# ─── Selection Filters ──────────────────────────────────────
class LevelSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Level)
    def AllowReference(self, reference, position):
        return False

class RefSelectionFilter(ISelectionFilter):
    """Allow Levels, Grids, CurveElements (Detail/Model Lines), ReferencePlanes."""
    def AllowElement(self, elem):
        if isinstance(elem, Level):
            return True
        if isinstance(elem, Grid):
            return True
        if isinstance(elem, ReferencePlane):
            return True
        if isinstance(elem, CurveElement):
            return True
        return False
    def AllowReference(self, reference, position):
        return False


# ─── Core Logic ─────────────────────────────────────────────
def get_level_2d_extents(level, view):
    """Get the 2D curve of a level in the given view."""
    try:
        curves = level.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        if curves and curves.Count > 0:
            return curves[0]
    except:
        pass
    try:
        curves = level.GetCurvesInView(DatumExtentType.Model, view)
        if curves and curves.Count > 0:
            return curves[0]
    except:
        pass
    return None


def get_reference_line(element):
    """Extract a Line from any supported reference element."""
    if isinstance(element, Level):
        try:
            curves = element.GetCurvesInView(DatumExtentType.Model, active_view)
            if curves and curves.Count > 0:
                return curves[0]
        except:
            pass
        return None
    elif isinstance(element, Grid):
        return element.Curve
    elif isinstance(element, ReferencePlane):
        return Line.CreateBound(element.BubbleEnd, element.FreeEnd)
    elif isinstance(element, CurveElement):
        return element.GeometryCurve
    elif hasattr(element, 'Location') and isinstance(element.Location, LocationCurve):
        return element.Location.Curve
    return None


def intersect_lines_2d(line_a, line_b):
    """
    Find intersection of two infinite lines.
    For levels (horizontal in section/elevation) we work in the
    plane of the view, so we use X-Z for section views and X-Y
    for plan views. We detect which plane the level lives in.
    """
    p1 = line_a.GetEndPoint(0)
    p2 = line_a.GetEndPoint(1)
    p3 = line_b.GetEndPoint(0)
    p4 = line_b.GetEndPoint(1)

    # Determine working plane: levels in section views run along X with constant Z
    # Use the two axes that have the most variation
    da = p2 - p1
    db = p4 - p3

    # Try XZ plane first (section/elevation views — levels are horizontal)
    denom_xz = da.X * db.Z - da.Z * db.X
    if abs(denom_xz) > 1e-10:
        t = ((p3.X - p1.X) * db.Z - (p3.Z - p1.Z) * db.X) / denom_xz
        return XYZ(p1.X + t * da.X, p1.Y + t * da.Y, p1.Z + t * da.Z)

    # Try XY plane (plan views)
    denom_xy = da.X * db.Y - da.Y * db.X
    if abs(denom_xy) > 1e-10:
        t = ((p3.X - p1.X) * db.Y - (p3.Y - p1.Y) * db.X) / denom_xy
        return XYZ(p1.X + t * da.X, p1.Y + t * da.Y, p1.Z + t * da.Z)

    # Try YZ plane
    denom_yz = da.Y * db.Z - da.Z * db.Y
    if abs(denom_yz) > 1e-10:
        t = ((p3.Y - p1.Y) * db.Z - (p3.Z - p1.Z) * db.Y) / denom_yz
        return XYZ(p1.X + t * da.X, p1.Y + t * da.Y, p1.Z + t * da.Z)

    return None  # Parallel


def find_new_endpoint_on_level(level_curve, ref_line):
    """
    Find where level (infinite) intersects reference (infinite).
    For parallel lines: project ref midpoint along level direction.
    """
    intersection = intersect_lines_2d(level_curve, ref_line)
    if intersection is not None:
        return intersection

    # Parallel fallback
    lv_start = level_curve.GetEndPoint(0)
    lv_dir = (level_curve.GetEndPoint(1) - lv_start).Normalize()
    ref_mid = (ref_line.GetEndPoint(0) + ref_line.GetEndPoint(1)) * 0.5
    dist = (ref_mid - lv_start).DotProduct(lv_dir)
    return lv_start + lv_dir.Multiply(dist)


def align_level_to_reference(level, ref_line, align_end_index, view, set_2d):
    """
    Align a level's 2D endpoint to the intersection with reference line.
    New curve stays collinear with the original level axis.
    """
    if set_2d:
        try:
            level.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
            level.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
        except:
            pass

    level_curve = get_level_2d_extents(level, view)
    if level_curve is None:
        return False

    new_point = find_new_endpoint_on_level(level_curve, ref_line)
    if new_point is None:
        return False

    # Force collinearity — project new_point back onto the level's own line
    lv_start = level_curve.GetEndPoint(0)
    lv_end = level_curve.GetEndPoint(1)
    lv_dir = (lv_end - lv_start).Normalize()
    dist = (new_point - lv_start).DotProduct(lv_dir)
    new_pt = lv_start + lv_dir.Multiply(dist)

    fixed_pt = level_curve.GetEndPoint(1 - align_end_index)

    if fixed_pt.DistanceTo(new_pt) < 0.01:
        return False

    if align_end_index == 0:
        new_line = Line.CreateBound(new_pt, fixed_pt)
    else:
        new_line = Line.CreateBound(fixed_pt, new_pt)

    try:
        level.SetCurveInView(DatumExtentType.ViewSpecific, view, new_line)
        return True
    except:
        return False


# ─── WPF UI ─────────────────────────────────────────────────
class AlignLevelsWindow(Window):
    def __init__(self):
        self.Title = "Align Level 2D Extents"
        self.Width = 380
        self.Height = 530
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.Background = B(DQT_BG)
        self.result = None
        self._build_ui()

    def _text(self, text, size=12, bold=False, color=None):
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = size
        tb.FontFamily = FontFamily("Segoe UI")
        tb.Foreground = B(color or DQT_TEXT_DARK)
        if bold:
            tb.FontWeight = FontWeights.Bold
        tb.TextWrapping = TextWrapping.Wrap
        return tb

    def _radio(self, text, group, checked=False):
        rb = RadioButton()
        rb.Content = text
        rb.GroupName = group
        rb.IsChecked = checked
        rb.Foreground = B(DQT_TEXT_DARK)
        rb.FontSize = 11.5
        rb.FontFamily = FontFamily("Segoe UI")
        rb.Margin = Thickness(0, 3, 0, 3)
        return rb

    def _button(self, text, handler, primary=False):
        btn = Button()
        btn.Content = text
        btn.Height = 34
        btn.FontSize = 12
        btn.FontFamily = FontFamily("Segoe UI")
        btn.FontWeight = FontWeights.SemiBold
        btn.Margin = Thickness(0, 3, 0, 3)
        btn.Cursor = System.Windows.Input.Cursors.Hand
        btn.BorderThickness = Thickness(1)
        if primary:
            btn.Background = B(DQT_PRIMARY)
            btn.Foreground = B(DQT_TEXT_DARK)
            btn.BorderBrush = B(DQT_ACCENT)
        else:
            btn.Background = B(DQT_WHITE)
            btn.Foreground = B(DQT_TEXT_DARK)
            btn.BorderBrush = B(DQT_BORDER)
        btn.Click += handler
        return btn

    def _group(self, header, panel):
        gb = GroupBox()
        gb.Margin = Thickness(0, 4, 0, 4)
        gb.Padding = Thickness(8, 6, 8, 6)
        hdr = self._text(header, 11, bold=True, color=DQT_ACCENT)
        gb.Header = hdr
        gb.BorderBrush = B(DQT_BORDER)
        gb.Content = panel
        return gb

    def _build_ui(self):
        root = StackPanel()

        # ── Header ──
        header = Border()
        header.Background = B(DQT_PRIMARY)
        header.CornerRadius = CornerRadius(0, 0, 6, 6)
        header.Padding = Thickness(16, 10, 16, 10)
        header.Margin = Thickness(0, 0, 0, 8)

        hdr_grid = WPFGrid()
        c1 = ColumnDefinition()
        c1.Width = GridLength(1, GridUnitType.Star)
        c2 = ColumnDefinition()
        c2.Width = GridLength(1, GridUnitType.Auto)
        hdr_grid.ColumnDefinitions.Add(c1)
        hdr_grid.ColumnDefinitions.Add(c2)

        left = StackPanel()
        t1 = self._text("ALIGN LEVEL EXTENTS", 17, bold=True, color=DQT_TEXT_DARK)
        left.Children.Add(t1)
        t2 = self._text("Snap 2D level endpoints to reference element", 10, color=DQT_DARK)
        t2.Margin = Thickness(0, 2, 0, 0)
        left.Children.Add(t2)
        WPFGrid.SetColumn(left, 0)
        hdr_grid.Children.Add(left)

        right = StackPanel()
        right.VerticalAlignment = VerticalAlignment.Center
        right.HorizontalAlignment = HorizontalAlignment.Right
        r1 = self._text("v1.0", 11, bold=True, color=DQT_TEXT_DARK)
        r1.HorizontalAlignment = HorizontalAlignment.Right
        right.Children.Add(r1)
        WPFGrid.SetColumn(right, 1)
        hdr_grid.Children.Add(right)

        header.Child = hdr_grid
        root.Children.Add(header)

        # ── Body ──
        body = StackPanel()
        body.Margin = Thickness(14, 0, 14, 0)

        # Which End
        end_panel = StackPanel()
        self.rb_end_0 = self._radio("Align LEFT (Bubble End / End0)", "end", True)
        self.rb_end_1 = self._radio("Align RIGHT (Non-Bubble / End1)", "end")
        self.rb_end_both = self._radio("Align BOTH Ends (pick 2 references)", "end")
        end_panel.Children.Add(self.rb_end_0)
        end_panel.Children.Add(self.rb_end_1)
        end_panel.Children.Add(self.rb_end_both)
        body.Children.Add(self._group("Which End to Align", end_panel))

        # Reference Elements
        ref_panel = StackPanel()
        ref_info = self._text(
            "Supported references:\n"
            "  Level  |  Grid  |  Detail Line  |  Model Line  |  Ref Plane",
            11, color=DQT_DARK)
        ref_info.Margin = Thickness(0, 2, 0, 2)
        ref_panel.Children.Add(ref_info)
        body.Children.Add(self._group("Reference Elements", ref_panel))

        # Options
        opt_panel = StackPanel()
        self.chk_2d = CheckBox()
        self.chk_2d.Content = "Force 2D extent (ViewSpecific) before aligning"
        self.chk_2d.IsChecked = True
        self.chk_2d.Foreground = B(DQT_TEXT_DARK)
        self.chk_2d.FontSize = 11.5
        self.chk_2d.Margin = Thickness(0, 2, 0, 2)
        opt_panel.Children.Add(self.chk_2d)
        body.Children.Add(self._group("Options", opt_panel))

        # Workflow Hint
        hint = Border()
        hint.Background = B(DQT_WHITE)
        hint.BorderBrush = B(DQT_PRIMARY)
        hint.BorderThickness = Thickness(1)
        hint.CornerRadius = CornerRadius(4)
        hint.Padding = Thickness(10, 8, 10, 8)
        hint.Margin = Thickness(0, 6, 0, 6)
        hint_text = self._text(
            "1. Click Run\n"
            "2. Select levels (Enter / Right-click to finish)\n"
            "3. Pick reference element\n"
            "4. Level endpoints snap to intersection",
            11, color=DQT_DARK)
        hint.Child = hint_text
        body.Children.Add(hint)

        # Buttons
        btn_panel = StackPanel()
        btn_panel.Margin = Thickness(0, 4, 0, 0)
        btn_panel.Children.Add(self._button("Run Alignment", self._on_run, primary=True))
        btn_panel.Children.Add(self._button("Cancel", self._on_cancel))
        body.Children.Add(btn_panel)

        root.Children.Add(body)

        # ── Footer ──
        footer = Border()
        footer.Background = B(DQT_PRIMARY)
        footer.CornerRadius = CornerRadius(6, 6, 0, 0)
        footer.Padding = Thickness(14, 6, 14, 6)
        footer.Margin = Thickness(0, 8, 0, 0)

        f_grid = WPFGrid()
        fc1 = ColumnDefinition()
        fc1.Width = GridLength(1, GridUnitType.Star)
        fc2 = ColumnDefinition()
        fc2.Width = GridLength(1, GridUnitType.Auto)
        f_grid.ColumnDefinitions.Add(fc1)
        f_grid.ColumnDefinitions.Add(fc2)

        fl = self._text("View: " + active_view.Name, 9.5, color=DQT_TEXT_DARK)
        fl.VerticalAlignment = VerticalAlignment.Center
        WPFGrid.SetColumn(fl, 0)
        f_grid.Children.Add(fl)

        fr = self._text("Copyright 2025 Dang Quoc Truong (DQT)", 9.5,
                         bold=True, color=DQT_TEXT_DARK)
        fr.HorizontalAlignment = HorizontalAlignment.Right
        WPFGrid.SetColumn(fr, 1)
        f_grid.Children.Add(fr)

        footer.Child = f_grid
        root.Children.Add(footer)

        self.Content = root

    def _on_run(self, sender, args):
        if self.rb_end_0.IsChecked:
            mode = "start"
        elif self.rb_end_1.IsChecked:
            mode = "end"
        else:
            mode = "both"
        self.result = {"align_mode": mode, "set_2d": self.chk_2d.IsChecked}
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# ─── Pick Helpers ───────────────────────────────────────────
def pick_levels():
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            LevelSelectionFilter(),
            "Select levels to align (Enter / Right-click to finish)")
        return [doc.GetElement(r.ElementId) for r in refs]
    except:
        return []


def pick_reference(prompt):
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, RefSelectionFilter(), prompt)
        return doc.GetElement(ref.ElementId)
    except:
        return None


# ─── Main ───────────────────────────────────────────────────
def run():
    # Validate view — levels only visible in section/elevation views
    vtype = active_view.ViewType
    valid_views = [
        ViewType.Section,
        ViewType.Elevation,
        ViewType.Detail,
        ViewType.DraftingView,
    ]
    if vtype not in valid_views and vtype != ViewType.ThreeD:
        TaskDialog.Show(
            "Align Levels",
            "This tool works best in Section or Elevation views\n"
            "where level lines are visible.\n\n"
            "Current view: {} ({})".format(active_view.Name, vtype))

    window = AlignLevelsWindow()
    window.ShowDialog()

    if window.result is None:
        return

    mode = window.result["align_mode"]
    set_2d = window.result["set_2d"]

    # Pick levels
    levels = pick_levels()
    if not levels:
        return

    # Pick reference(s)
    if mode == "both":
        ref_start = pick_reference("Pick reference for LEFT (Bubble End)")
        if not ref_start:
            return
        line_start = get_reference_line(ref_start)
        if not line_start:
            TaskDialog.Show("Align Levels", "Cannot extract line from left reference.")
            return

        ref_end = pick_reference("Pick reference for RIGHT (Non-Bubble End)")
        if not ref_end:
            return
        line_end = get_reference_line(ref_end)
        if not line_end:
            TaskDialog.Show("Align Levels", "Cannot extract line from right reference.")
            return
    else:
        ref_elem = pick_reference("Pick reference (Level / Grid / Line / Ref Plane)")
        if not ref_elem:
            return
        ref_line = get_reference_line(ref_elem)
        if not ref_line:
            TaskDialog.Show("Align Levels", "Cannot extract line from selected element.")
            return

    # Execute
    ok = 0
    fail = 0

    t = Transaction(doc, "Align Level 2D Extents")
    t.Start()

    try:
        for level in levels:
            if mode == "start":
                success = align_level_to_reference(level, ref_line, 0, active_view, set_2d)
            elif mode == "end":
                success = align_level_to_reference(level, ref_line, 1, active_view, set_2d)
            else:
                s1 = align_level_to_reference(level, line_start, 0, active_view, set_2d)
                s2 = align_level_to_reference(level, line_end, 1, active_view, set_2d)
                success = s1 and s2

            if success:
                ok += 1
            else:
                fail += 1

        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Align Levels", "Error: {}".format(str(e)))
        return

    # Result
    msg = "{} level(s) aligned successfully.".format(ok)
    if fail > 0:
        msg += "\n{} level(s) could not be aligned.".format(fail)
    TaskDialog.Show("Align Levels", msg)


run()