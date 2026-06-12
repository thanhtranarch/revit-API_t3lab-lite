# -*- coding: utf-8 -*-
"""
Align Grid 2D Extents v1.0
Align grid start/end to a reference element (Grid, Detail Line,
Model Line, Reference Plane).

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Align\nGrids"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Align 2D grid extents to a reference line."

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
    Border, CheckBox, Grid as WPFGrid, ColumnDefinition, Orientation
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
class GridSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, Grid)
    def AllowReference(self, reference, position):
        return False

class RefSelectionFilter(ISelectionFilter):
    """Allow Grids, CurveElements (Detail/Model Lines), ReferencePlanes."""
    def AllowElement(self, elem):
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
def get_grid_2d_extents(grid, view):
    """Get the 2D curve of a grid in the given view."""
    try:
        curves = grid.GetCurvesInView(DatumExtentType.ViewSpecific, view)
        if curves and curves.Count > 0:
            return curves[0]
    except:
        pass
    try:
        curves = grid.GetCurvesInView(DatumExtentType.Model, view)
        if curves and curves.Count > 0:
            return curves[0]
    except:
        pass
    return None


def get_reference_line(element):
    """Extract a Line from any supported reference element."""
    if isinstance(element, Grid):
        return element.Curve
    elif isinstance(element, ReferencePlane):
        return Line.CreateBound(element.BubbleEnd, element.FreeEnd)
    elif isinstance(element, CurveElement):
        return element.GeometryCurve
    elif hasattr(element, 'Location') and isinstance(element.Location, LocationCurve):
        return element.Location.Curve
    return None


def intersect_lines_2d(line_a, line_b):
    """Find intersection of two infinite lines in XY plane."""
    p1 = line_a.GetEndPoint(0)
    p2 = line_a.GetEndPoint(1)
    p3 = line_b.GetEndPoint(0)
    p4 = line_b.GetEndPoint(1)

    d1x = p2.X - p1.X
    d1y = p2.Y - p1.Y
    d2x = p4.X - p3.X
    d2y = p4.Y - p3.Y

    denom = d1x * d2y - d1y * d2x
    if abs(denom) < 1e-10:
        return None

    t = ((p3.X - p1.X) * d2y - (p3.Y - p1.Y) * d2x) / denom
    return XYZ(p1.X + t * d1x, p1.Y + t * d1y, p1.Z)


def find_new_endpoint_on_grid(grid_curve, ref_line):
    """
    Find where grid (infinite) intersects reference (infinite).
    For parallel lines: project ref midpoint along grid direction.
    """
    intersection = intersect_lines_2d(grid_curve, ref_line)
    if intersection is not None:
        return intersection

    grid_start = grid_curve.GetEndPoint(0)
    grid_dir = (grid_curve.GetEndPoint(1) - grid_start).Normalize()
    ref_mid = (ref_line.GetEndPoint(0) + ref_line.GetEndPoint(1)) * 0.5
    dist = (ref_mid - grid_start).DotProduct(grid_dir)
    return grid_start + grid_dir.Multiply(dist)


def align_grid_to_reference(grid, ref_line, align_end_index, view, set_2d):
    """
    Align a grid's 2D endpoint to the intersection with reference line.
    New curve stays collinear with the original grid axis.
    """
    if set_2d:
        try:
            grid.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
            grid.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
        except:
            pass

    grid_curve = get_grid_2d_extents(grid, view)
    if grid_curve is None:
        return False

    new_point = find_new_endpoint_on_grid(grid_curve, ref_line)
    if new_point is None:
        return False

    # Force collinearity
    grid_start = grid_curve.GetEndPoint(0)
    grid_dir = (grid_curve.GetEndPoint(1) - grid_start).Normalize()
    dist = (new_point - grid_start).DotProduct(grid_dir)
    new_pt = XYZ(
        grid_start.X + grid_dir.X * dist,
        grid_start.Y + grid_dir.Y * dist,
        grid_start.Z
    )

    fixed_pt = grid_curve.GetEndPoint(1 - align_end_index)

    if fixed_pt.DistanceTo(new_pt) < 0.01:
        return False

    if align_end_index == 0:
        new_line = Line.CreateBound(new_pt, fixed_pt)
    else:
        new_line = Line.CreateBound(fixed_pt, new_pt)

    try:
        grid.SetCurveInView(DatumExtentType.ViewSpecific, view, new_line)
        return True
    except:
        return False


# ─── WPF UI ─────────────────────────────────────────────────
class AlignGridsWindow(Window):
    def __init__(self):
        self.Title = "Align Grid 2D Extents"
        self.Width = 380
        self.Height = 520
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.Background = B(DQT_BG)
        self.result = None
        self._build_ui()

    # ── UI Builders ──
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

    # ── Build Layout ──
    def _build_ui(self):
        root = StackPanel()

        # Header
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
        t1 = self._text("ALIGN GRID EXTENTS", 17, bold=True, color=DQT_TEXT_DARK)
        left.Children.Add(t1)
        t2 = self._text("Snap 2D grid endpoints to reference element", 10, color=DQT_DARK)
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

        # Body
        body = StackPanel()
        body.Margin = Thickness(14, 0, 14, 0)

        # ── Which End ──
        end_panel = StackPanel()
        self.rb_end_0 = self._radio("Align START (Bubble End / End0)", "end", True)
        self.rb_end_1 = self._radio("Align END (Non-Bubble / End1)", "end")
        self.rb_end_both = self._radio("Align BOTH Ends (pick 2 references)", "end")
        end_panel.Children.Add(self.rb_end_0)
        end_panel.Children.Add(self.rb_end_1)
        end_panel.Children.Add(self.rb_end_both)
        body.Children.Add(self._group("Which End to Align", end_panel))

        # ── Reference Type Info ──
        ref_panel = StackPanel()
        ref_info = self._text(
            "Supported references:\n"
            "  Grid  |  Detail Line  |  Model Line  |  Reference Plane",
            11, color=DQT_DARK)
        ref_info.Margin = Thickness(0, 2, 0, 2)
        ref_panel.Children.Add(ref_info)
        body.Children.Add(self._group("Reference Elements", ref_panel))

        # ── Options ──
        opt_panel = StackPanel()
        self.chk_2d = CheckBox()
        self.chk_2d.Content = "Force 2D extent (ViewSpecific) before aligning"
        self.chk_2d.IsChecked = True
        self.chk_2d.Foreground = B(DQT_TEXT_DARK)
        self.chk_2d.FontSize = 11.5
        self.chk_2d.Margin = Thickness(0, 2, 0, 2)
        opt_panel.Children.Add(self.chk_2d)
        body.Children.Add(self._group("Options", opt_panel))

        # ── Workflow Hint ──
        hint_border = Border()
        hint_border.Background = B(DQT_WHITE)
        hint_border.BorderBrush = B(DQT_PRIMARY)
        hint_border.BorderThickness = Thickness(1)
        hint_border.CornerRadius = CornerRadius(4)
        hint_border.Padding = Thickness(10, 8, 10, 8)
        hint_border.Margin = Thickness(0, 6, 0, 6)
        hint_text = self._text(
            "1. Click Run\n"
            "2. Select grids (Enter / Right-click to finish)\n"
            "3. Pick reference element\n"
            "4. Grid endpoints snap to intersection",
            11, color=DQT_DARK)
        hint_border.Child = hint_text
        body.Children.Add(hint_border)

        # ── Buttons ──
        btn_panel = StackPanel()
        btn_panel.Margin = Thickness(0, 4, 0, 0)
        btn_panel.Children.Add(self._button("Run Alignment", self._on_run, primary=True))
        btn_panel.Children.Add(self._button("Cancel", self._on_cancel))
        body.Children.Add(btn_panel)

        root.Children.Add(body)

        # Footer
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

    # ── Events ──
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
def pick_grids():
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            GridSelectionFilter(),
            "Select grids to align (Enter / Right-click to finish)")
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


def get_ref_name(elem):
    name = getattr(elem, 'Name', '')
    typ = elem.__class__.__name__
    return "{} '{}'".format(typ, name) if name else typ


# ─── Main ───────────────────────────────────────────────────
def run():
    window = AlignGridsWindow()
    window.ShowDialog()

    if window.result is None:
        return

    mode = window.result["align_mode"]
    set_2d = window.result["set_2d"]

    # Pick grids
    grids = pick_grids()
    if not grids:
        return

    # Pick reference(s)
    if mode == "both":
        ref_start = pick_reference("Pick reference for START (Bubble End)")
        if not ref_start:
            return
        line_start = get_reference_line(ref_start)
        if not line_start:
            TaskDialog.Show("Align Grids", "Cannot extract line from start reference.")
            return

        ref_end = pick_reference("Pick reference for END (Non-Bubble End)")
        if not ref_end:
            return
        line_end = get_reference_line(ref_end)
        if not line_end:
            TaskDialog.Show("Align Grids", "Cannot extract line from end reference.")
            return
    else:
        ref_elem = pick_reference("Pick reference (Grid / Line / Reference Plane)")
        if not ref_elem:
            return
        ref_line = get_reference_line(ref_elem)
        if not ref_line:
            TaskDialog.Show("Align Grids", "Cannot extract line from selected element.")
            return

    # Execute
    ok = 0
    fail = 0

    t = Transaction(doc, "Align Grid 2D Extents")
    t.Start()

    try:
        for grid in grids:
            if mode == "start":
                success = align_grid_to_reference(grid, ref_line, 0, active_view, set_2d)
            elif mode == "end":
                success = align_grid_to_reference(grid, ref_line, 1, active_view, set_2d)
            else:
                s1 = align_grid_to_reference(grid, line_start, 0, active_view, set_2d)
                s2 = align_grid_to_reference(grid, line_end, 1, active_view, set_2d)
                success = s1 and s2

            if success:
                ok += 1
            else:
                fail += 1

        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Align Grids", "Error: {}".format(str(e)))
        return

    # Result
    msg = "{} grid(s) aligned successfully.".format(ok)
    if fail > 0:
        msg += "\n{} grid(s) could not be aligned.".format(fail)
    TaskDialog.Show("Align Grids", msg)


run()