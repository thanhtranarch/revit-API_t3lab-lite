# -*- coding: utf-8 -*-
"""
Model Health Check v1.4 - DQT
Analyzes Revit model health with color-coded metrics dashboard.
Features: Gauge dashboard, Select Elements, Weighted score, Purgeable elements.
Pure code-behind WPF for IronPython stability.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Model\nHealth"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Analyze model health with color-coded metrics and gauge dashboard."

# ============================================================
# IMPORTS
# ============================================================
import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import System
import System.Diagnostics
from System import Environment, Math
from System.Collections.Generic import List
from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, Visibility, TextWrapping, FontWeights,
    GridLength, GridUnitType, MessageBox, MessageBoxButton, MessageBoxImage,
    CornerRadius as WinCornerRadius, Point
)
import System.Windows.Controls as WPFControls
from System.Windows.Controls import (
    StackPanel, Border, TextBlock, Button, ScrollViewer, Canvas,
    ColumnDefinition, RowDefinition, Orientation, ScrollBarVisibility,
    ToolTip, TextBox
)
from System.Windows.Media import (
    SolidColorBrush, Color, BrushConverter, Pen,
    PathGeometry, PathFigure, ArcSegment, SweepDirection,
    PenLineCap
)
from System.Windows.Shapes import Path, Ellipse
from System.Windows.Input import Cursors

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from pyrevit import script

WPFGrid = WPFControls.Grid

import os
import datetime
import codecs
from collections import OrderedDict

# ============================================================
# DQT BRAND COLORS
# ============================================================
DQT_PRIMARY = "#0F172A"
DQT_PRIMARY_DARK = "#CBD5E1"
DQT_BACKGROUND = "#F8FAFC"
DQT_TEXT_DARK = "#5D4E37"
DQT_TEXT = "#0F172A"
DQT_BORDER = "#CBD5E1"

HEALTH_GREEN = "#10B981"
HEALTH_LIGHT_GREEN = "#8BC34A"
HEALTH_YELLOW = "#FFC107"
HEALTH_ORANGE = "#F59E0B"
HEALTH_RED = "#EF4444"
HEALTH_DARK_RED = "#D32F2F"

BC = BrushConverter()
def brush(hex_color):
    return BC.ConvertFromString(hex_color)

# ============================================================
# METRIC THRESHOLDS WITH WEIGHTS
# weight: 5=Critical, 4=High, 3=Medium, 2=Low, 1=Minor
# ============================================================
METRIC_THRESHOLDS = OrderedDict([
    ("file_size_mb", {
        "label": "File Size (MB)",
        "thresholds": [100, 250, 500, 750, 1000],
        "tooltip": "Model file size. Large files slow loading and sync.",
        "unit": "MB",
        "selectable": False,
        "weight": 4
    }),
    ("warnings", {
        "label": "Warnings",
        "thresholds": [100, 500, 1000, 2000, 5000],
        "tooltip": "Total warnings. High count = model instability.",
        "unit": "",
        "selectable": False,
        "weight": 5
    }),
    ("cad_imports", {
        "label": "CAD Imports",
        "thresholds": [0, 2, 5, 7, 10],
        "tooltip": "Imported CAD (not linked). Bloats file size significantly.",
        "unit": "",
        "selectable": True,
        "weight": 5
    }),
    ("in_place_families", {
        "label": "In-Place Families",
        "thresholds": [5, 15, 30, 60, 100],
        "tooltip": "In-Place families can't be reused, increase file size.",
        "unit": "",
        "selectable": True,
        "weight": 4
    }),
    ("rvt_links", {
        "label": "RVT Links",
        "thresholds": [10, 20, 35, 50, 80],
        "tooltip": "Linked Revit files. Too many = slow performance.",
        "unit": "",
        "selectable": True,
        "weight": 2
    }),
    ("worksets", {
        "label": "Worksets",
        "thresholds": [10, 20, 30, 40, 50],
        "tooltip": "User worksets. Excessive worksets complicate management.",
        "unit": "",
        "selectable": False,
        "weight": 1
    }),
    ("cad_links", {
        "label": "CAD Links",
        "thresholds": [10, 25, 50, 80, 120],
        "tooltip": "Linked CAD files. Many links degrade navigation.",
        "unit": "",
        "selectable": True,
        "weight": 3
    }),
    ("views", {
        "label": "Views",
        "thresholds": [200, 500, 1000, 2000, 4000],
        "tooltip": "Total views. Too many slow file open/save.",
        "unit": "",
        "selectable": True,
        "weight": 3
    }),
    ("sheets", {
        "label": "Sheets",
        "thresholds": [100, 200, 400, 600, 1000],
        "tooltip": "Total sheets with placed views increase file size.",
        "unit": "",
        "selectable": True,
        "weight": 2
    }),
    ("groups", {
        "label": "Groups",
        "thresholds": [20, 50, 100, 200, 500],
        "tooltip": "Model and Detail Groups cause performance issues.",
        "unit": "",
        "selectable": True,
        "weight": 3
    }),
    ("design_options", {
        "label": "Design Options",
        "thresholds": [3, 5, 8, 15, 20],
        "tooltip": "Design Options add complexity and memory usage.",
        "unit": "",
        "selectable": True,
        "weight": 1
    }),
    ("reference_planes", {
        "label": "Ref. Planes",
        "thresholds": [100, 200, 500, 800, 1500],
        "tooltip": "Leftover reference planes clutter the model.",
        "unit": "",
        "selectable": True,
        "weight": 1
    }),
    ("detail_lines", {
        "label": "Detail Lines",
        "thresholds": [1000, 5000, 10000, 25000, 50000],
        "tooltip": "Excessive detail lines = drafting overuse.",
        "unit": "",
        "selectable": True,
        "weight": 2
    }),
    ("filled_regions", {
        "label": "Filled Regions",
        "thresholds": [100, 500, 1000, 3000, 5000],
        "tooltip": "Many filled regions slow view rendering.",
        "unit": "",
        "selectable": True,
        "weight": 2
    }),
    ("rooms_unplaced", {
        "label": "Unplaced Rooms",
        "thresholds": [0, 5, 15, 30, 50],
        "tooltip": "Unplaced rooms cause errors in schedules.",
        "unit": "",
        "selectable": True,
        "weight": 2
    }),
    ("linked_dwg_not_pinned", {
        "label": "Unpinned Links",
        "thresholds": [0, 3, 8, 15, 30],
        "tooltip": "Unpinned links can be accidentally moved.",
        "unit": "",
        "selectable": True,
        "weight": 2
    }),
    ("duplicate_elements", {
        "label": "Duplicate Elements",
        "thresholds": [0, 10, 30, 60, 100],
        "tooltip": "Elements of same type overlapping at same location. Cause double counting and visual issues.",
        "unit": "",
        "selectable": True,
        "weight": 4
    }),
])


# ============================================================
# MODEL HEALTH ANALYZER
# ============================================================
class ModelHealthAnalyzer:
    def __init__(self, doc):
        self.doc = doc
        self.metrics = OrderedDict()
        self.element_ids = {}

    def analyze(self):
        self._file_size()
        self._warnings()
        self._cad_imports()
        self._in_place_families()
        self._rvt_links()
        self._worksets()
        self._cad_links()
        self._views()
        self._sheets()
        self._groups()
        self._design_options()
        self._reference_planes()
        self._detail_lines()
        self._filled_regions()
        self._unplaced_rooms()
        self._unpinned_links()
        self._duplicate_elements()
        return self.metrics

    def _store_ids(self, key, elements):
        ids = []
        for elem in elements:
            try:
                ids.append(elem.Id)
            except:
                pass
        self.element_ids[key] = ids

    def _file_size(self):
        try:
            path = self.doc.PathName
            if path and os.path.exists(path):
                self.metrics["file_size_mb"] = round(os.path.getsize(path) / (1024.0 * 1024.0), 1)
            else:
                self.metrics["file_size_mb"] = 0
        except:
            self.metrics["file_size_mb"] = 0

    def _warnings(self):
        try:
            w = self.doc.GetWarnings()
            self.metrics["warnings"] = len(w) if w else 0
        except:
            self.metrics["warnings"] = 0

    def _cad_imports(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(ImportInstance).WhereElementIsNotElementType()
            elems = []
            for inst in col:
                try:
                    if not inst.IsLinked:
                        elems.append(inst)
                except:
                    pass
            self.metrics["cad_imports"] = len(elems)
            self._store_ids("cad_imports", elems)
        except:
            self.metrics["cad_imports"] = 0

    def _in_place_families(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
            elems = []
            for fi in col:
                try:
                    if fi.Symbol and fi.Symbol.Family and fi.Symbol.Family.IsInPlace:
                        elems.append(fi)
                except:
                    pass
            self.metrics["in_place_families"] = len(elems)
            self._store_ids("in_place_families", elems)
        except:
            self.metrics["in_place_families"] = 0

    def _rvt_links(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(RevitLinkInstance).WhereElementIsNotElementType()
            elems = list(col)
            self.metrics["rvt_links"] = len(elems)
            self._store_ids("rvt_links", elems)
        except:
            self.metrics["rvt_links"] = 0

    def _worksets(self):
        try:
            if self.doc.IsWorkshared:
                ws = FilteredWorksetCollector(self.doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
                self.metrics["worksets"] = ws.Count
            else:
                self.metrics["worksets"] = 0
        except:
            self.metrics["worksets"] = 0

    def _cad_links(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(ImportInstance).WhereElementIsNotElementType()
            elems = []
            for inst in col:
                try:
                    if inst.IsLinked:
                        elems.append(inst)
                except:
                    pass
            self.metrics["cad_links"] = len(elems)
            self._store_ids("cad_links", elems)
        except:
            self.metrics["cad_links"] = 0

    def _views(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(View).WhereElementIsNotElementType()
            elems = []
            for v in col:
                try:
                    if not v.IsTemplate and v.ViewType != ViewType.Internal:
                        elems.append(v)
                except:
                    pass
            self.metrics["views"] = len(elems)
            self._store_ids("views", elems)
        except:
            self.metrics["views"] = 0

    def _sheets(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(ViewSheet).WhereElementIsNotElementType()
            elems = list(col)
            self.metrics["sheets"] = len(elems)
            self._store_ids("sheets", elems)
        except:
            self.metrics["sheets"] = 0

    def _groups(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(Group).WhereElementIsNotElementType()
            elems = list(col)
            self.metrics["groups"] = len(elems)
            self._store_ids("groups", elems)
        except:
            self.metrics["groups"] = 0

    def _design_options(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(DesignOption).WhereElementIsNotElementType()
            elems = list(col)
            self.metrics["design_options"] = len(elems)
            self._store_ids("design_options", elems)
        except:
            self.metrics["design_options"] = 0

    def _reference_planes(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(ReferencePlane).WhereElementIsNotElementType()
            elems = list(col)
            self.metrics["reference_planes"] = len(elems)
            self._store_ids("reference_planes", elems)
        except:
            self.metrics["reference_planes"] = 0

    def _detail_lines(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(CurveElement).WhereElementIsNotElementType()
            elems = []
            for ce in col:
                try:
                    cat = ce.Category
                    if cat and "Lines" in cat.Name:
                        elems.append(ce)
                except:
                    pass
            self.metrics["detail_lines"] = len(elems)
            self._store_ids("detail_lines", elems)
        except:
            self.metrics["detail_lines"] = 0

    def _filled_regions(self):
        try:
            col = FilteredElementCollector(self.doc).OfClass(FilledRegion).WhereElementIsNotElementType()
            elems = list(col)
            self.metrics["filled_regions"] = len(elems)
            self._store_ids("filled_regions", elems)
        except:
            self.metrics["filled_regions"] = 0

    def _unplaced_rooms(self):
        try:
            col = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
            elems = []
            for room in col:
                try:
                    if room.Location is None:
                        elems.append(room)
                except:
                    pass
            self.metrics["rooms_unplaced"] = len(elems)
            self._store_ids("rooms_unplaced", elems)
        except:
            self.metrics["rooms_unplaced"] = 0

    def _unpinned_links(self):
        try:
            elems = []
            for link in FilteredElementCollector(self.doc).OfClass(RevitLinkInstance).WhereElementIsNotElementType():
                try:
                    if not link.Pinned:
                        elems.append(link)
                except:
                    pass
            for inst in FilteredElementCollector(self.doc).OfClass(ImportInstance).WhereElementIsNotElementType():
                try:
                    if inst.IsLinked and not inst.Pinned:
                        elems.append(inst)
                except:
                    pass
            self.metrics["linked_dwg_not_pinned"] = len(elems)
            self._store_ids("linked_dwg_not_pinned", elems)
        except:
            self.metrics["linked_dwg_not_pinned"] = 0

    def _duplicate_elements(self):
        """Find duplicate elements using Revit Warnings API.
        Reads warnings with message 'identical instances in the same place'
        which is exactly what Revit uses to detect duplicates.
        This is the most accurate method - matches Revit's own detection.
        """
        try:
            warnings = self.doc.GetWarnings()
            dupe_ids_set = set()
            
            if warnings:
                for warning in warnings:
                    try:
                        desc = warning.GetDescriptionText()
                        # Revit's exact warning for duplicate elements
                        if "identical instances" in desc.lower() and "same place" in desc.lower():
                            # Get all element IDs involved in this warning
                            failing = warning.GetFailingElements()
                            additional = warning.GetAdditionalElements()
                            
                            if failing:
                                for eid in failing:
                                    dupe_ids_set.add(eid)
                            if additional:
                                for eid in additional:
                                    dupe_ids_set.add(eid)
                    except:
                        pass
            
            # Convert to element list for selection
            dupe_ids = list(dupe_ids_set)
            self.metrics["duplicate_elements"] = len(dupe_ids)
            self.element_ids["duplicate_elements"] = dupe_ids
        except:
            self.metrics["duplicate_elements"] = 0


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def get_health_color(key, value):
    if key not in METRIC_THRESHOLDS:
        return "#FFFFFF"
    t = METRIC_THRESHOLDS[key]["thresholds"]
    if value <= t[0]: return HEALTH_GREEN
    elif value <= t[1]: return HEALTH_LIGHT_GREEN
    elif value <= t[2]: return HEALTH_YELLOW
    elif value <= t[3]: return HEALTH_ORANGE
    elif value <= t[4]: return HEALTH_RED
    else: return HEALTH_DARK_RED

def get_health_score(metrics):
    weighted_total = 0
    weight_sum = 0
    for key, value in metrics.items():
        if key in METRIC_THRESHOLDS:
            t = METRIC_THRESHOLDS[key]["thresholds"]
            w = METRIC_THRESHOLDS[key].get("weight", 1)
            if value <= t[0]: s = 100
            elif value <= t[1]: s = 80
            elif value <= t[2]: s = 60
            elif value <= t[3]: s = 40
            elif value <= t[4]: s = 20
            else: s = 0
            weighted_total += s * w
            weight_sum += w
    return round(weighted_total / max(weight_sum, 1), 1)

def get_health_grade(score):
    if score >= 90: return "A", "Excellent", HEALTH_GREEN
    elif score >= 75: return "B", "Good", HEALTH_LIGHT_GREEN
    elif score >= 60: return "C", "Fair", HEALTH_YELLOW
    elif score >= 40: return "D", "Poor", HEALTH_ORANGE
    else: return "F", "Critical", HEALTH_RED

def get_status_text(key, value):
    if key not in METRIC_THRESHOLDS:
        return "N/A"
    t = METRIC_THRESHOLDS[key]["thresholds"]
    if value <= t[0]: return "Good"
    elif value <= t[1]: return "Acceptable"
    elif value <= t[2]: return "Warning"
    elif value <= t[3]: return "Concerning"
    elif value <= t[4]: return "Critical"
    else: return "Severe"

RECOMMENDATIONS = {
    "file_size_mb": "Purge unused families, remove imported CAD files, audit model.",
    "warnings": "Review and resolve warnings. Start with most frequent types.",
    "cad_imports": "Delete imported CAD. Use linked CAD instead.",
    "in_place_families": "Convert In-Place to loadable families.",
    "rvt_links": "Review if all RVT links are necessary. Unload unused.",
    "worksets": "Consolidate worksets if possible.",
    "cad_links": "Minimize CAD links. Convert to native Revit elements.",
    "views": "Delete unused views. Use View Templates.",
    "sheets": "Archive completed sheets. Remove test sheets.",
    "groups": "Ungroup where possible. Use families instead.",
    "design_options": "Finalize and accept primary design options.",
    "reference_planes": "Delete unnamed/unnecessary reference planes.",
    "detail_lines": "Review detail lines. Use line-based detail components.",
    "filled_regions": "Minimize filled regions. Use material hatching.",
    "rooms_unplaced": "Place or delete unplaced rooms.",
    "linked_dwg_not_pinned": "Pin all linked files to prevent accidental movement.",
    "duplicate_elements": "Review and delete overlapping duplicate elements. They cause double counting in schedules and visual artifacts.",
}


# ============================================================
# GAUGE WIDGET - Semi-circle gauge using WPF Path
# ============================================================
def create_gauge(score, size=120):
    """Create a semi-circle gauge widget showing score 0-100"""
    canvas = Canvas()
    canvas.Width = size
    canvas.Height = size * 0.7
    
    cx = size / 2.0
    cy = size * 0.6
    radius = size * 0.42
    stroke_width = size * 0.08
    
    # Background arc (gray)
    bg_arc = _create_arc_path(cx, cy, radius, 180, 360, "#E0E0E0", stroke_width)
    canvas.Children.Add(bg_arc)
    
    # Colored arc based on score
    if score <= 0:
        sweep = 0
    else:
        sweep = min(score / 100.0, 1.0) * 180
    
    if sweep > 0:
        _, _, color = get_health_grade(score)
        fg_arc = _create_arc_path(cx, cy, radius, 180, 180 + sweep, color, stroke_width)
        canvas.Children.Add(fg_arc)
    
    # Score text
    score_tb = TextBlock()
    score_tb.Text = str(score)
    score_tb.FontSize = size * 0.22
    score_tb.FontWeight = FontWeights.Bold
    _, _, sc = get_health_grade(score)
    score_tb.Foreground = brush(sc)
    score_tb.HorizontalAlignment = HorizontalAlignment.Center
    Canvas.SetLeft(score_tb, cx - size * 0.18)
    Canvas.SetTop(score_tb, cy - size * 0.28)
    canvas.Children.Add(score_tb)
    
    return canvas


def _create_arc_path(cx, cy, radius, start_angle, end_angle, color, stroke_width):
    """Create a WPF Path representing an arc"""
    p = Path()
    p.Stroke = brush(color)
    p.StrokeThickness = stroke_width
    p.StrokeStartLineCap = PenLineCap.Round
    p.StrokeEndLineCap = PenLineCap.Round
    p.Fill = None
    
    start_rad = start_angle * Math.PI / 180.0
    end_rad = end_angle * Math.PI / 180.0
    
    start_x = cx + radius * Math.Cos(start_rad)
    start_y = cy + radius * Math.Sin(start_rad)
    end_x = cx + radius * Math.Cos(end_rad)
    end_y = cy + radius * Math.Sin(end_rad)
    
    is_large = (end_angle - start_angle) > 180
    
    fig = PathFigure()
    fig.StartPoint = Point(start_x, start_y)
    fig.IsClosed = False
    
    arc = ArcSegment()
    arc.Point = Point(end_x, end_y)
    arc.Size = System.Windows.Size(radius, radius)
    arc.IsLargeArc = is_large
    arc.SweepDirection = SweepDirection.Clockwise
    
    fig.Segments.Add(arc)
    
    geo = PathGeometry()
    geo.Figures.Add(fig)
    p.Data = geo
    
    return p


# ============================================================
# WPF WINDOW
# ============================================================
class ModelHealthWindow(Window):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.metrics = OrderedDict()
        self.analyzer = None

        self.Title = "Model Health Check v1.4 - DQT"
        self.Height = 900
        self.Width = 1400
        self.MinHeight = 700
        self.MinWidth = 1100
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = brush(DQT_BACKGROUND)

        self._build_ui()
        self._run_analysis()

    def _build_ui(self):
        root = WPFGrid()
        root.Margin = Thickness(14)

        for h in [
            GridLength(1, GridUnitType.Auto),   # header
            GridLength(1, GridUnitType.Auto),   # score + gauge
            GridLength(1, GridUnitType.Auto),   # legend
            GridLength(1, GridUnitType.Star),   # metrics
            GridLength(1, GridUnitType.Auto),   # recommendations
            GridLength(1, GridUnitType.Auto),   # footer
        ]:
            rd = RowDefinition()
            rd.Height = h
            root.RowDefinitions.Add(rd)

        header = self._make_header()
        WPFGrid.SetRow(header, 0)
        root.Children.Add(header)

        score_card = self._make_score_card()
        WPFGrid.SetRow(score_card, 1)
        root.Children.Add(score_card)

        legend = self._make_legend()
        WPFGrid.SetRow(legend, 2)
        root.Children.Add(legend)

        # Metrics
        metrics_border = Border()
        metrics_border.Background = brush("#FFFFFF")
        metrics_border.BorderBrush = brush(DQT_BORDER)
        metrics_border.BorderThickness = Thickness(1)
        metrics_border.CornerRadius = WinCornerRadius(6)
        metrics_border.Margin = Thickness(0, 0, 0, 10)
        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        self.metrics_stack = StackPanel()
        sv.Content = self.metrics_stack
        metrics_border.Child = sv
        WPFGrid.SetRow(metrics_border, 3)
        root.Children.Add(metrics_border)

        # Recommendations
        rec_border = Border()
        rec_border.Background = brush("#FFFFFF")
        rec_border.BorderBrush = brush(DQT_BORDER)
        rec_border.BorderThickness = Thickness(1)
        rec_border.CornerRadius = WinCornerRadius(6)
        rec_border.Padding = Thickness(12, 8, 12, 8)
        rec_border.Margin = Thickness(0, 0, 0, 8)
        rec_border.MaxHeight = 160
        rec_sv = ScrollViewer()
        rec_sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.rec_panel = StackPanel()
        rec_sv.Content = self.rec_panel
        rec_border.Child = rec_sv
        WPFGrid.SetRow(rec_border, 4)
        root.Children.Add(rec_border)

        footer = self._make_footer()
        WPFGrid.SetRow(footer, 5)
        root.Children.Add(footer)

        self.Content = root

    # ---- HEADER ----
    def _make_header(self):
        border = Border()
        border.Background = brush(DQT_PRIMARY)
        border.CornerRadius = WinCornerRadius(6)
        border.Padding = Thickness(16, 12, 16, 12)
        border.Margin = Thickness(0, 0, 0, 10)
        g = WPFGrid()
        c1 = ColumnDefinition()
        c1.Width = GridLength(1, GridUnitType.Star)
        c2 = ColumnDefinition()
        c2.Width = GridLength(1, GridUnitType.Auto)
        g.ColumnDefinitions.Add(c1)
        g.ColumnDefinitions.Add(c2)

        left = StackPanel()
        title = TextBlock()
        title.Text = "MODEL HEALTH CHECK"
        title.FontSize = 22
        title.FontWeight = FontWeights.Bold
        title.Foreground = brush(DQT_TEXT_DARK)
        left.Children.Add(title)
        self.txt_project = TextBlock()
        self.txt_project.Text = "Project: Loading..."
        self.txt_project.FontSize = 12
        self.txt_project.Foreground = brush(DQT_TEXT_DARK)
        self.txt_project.Margin = Thickness(0, 3, 0, 0)
        left.Children.Add(self.txt_project)
        WPFGrid.SetColumn(left, 0)
        g.Children.Add(left)

        right = StackPanel()
        right.HorizontalAlignment = HorizontalAlignment.Right
        right.VerticalAlignment = VerticalAlignment.Center
        t1 = TextBlock()
        t1.Text = "pyDQT Suite"
        t1.FontSize = 14
        t1.FontWeight = FontWeights.SemiBold
        t1.Foreground = brush(DQT_TEXT_DARK)
        t1.HorizontalAlignment = HorizontalAlignment.Right
        right.Children.Add(t1)
        t2 = TextBlock()
        t2.Text = "Copyright by Dang Quoc Truong - DQT"
        t2.FontSize = 9
        t2.Foreground = brush(DQT_TEXT_DARK)
        t2.Opacity = 0.7
        t2.HorizontalAlignment = HorizontalAlignment.Right
        right.Children.Add(t2)
        WPFGrid.SetColumn(right, 1)
        g.Children.Add(right)

        border.Child = g
        return border

    # ---- SCORE CARD WITH GAUGE ----
    def _make_score_card(self):
        border = Border()
        border.Background = brush("#FFFFFF")
        border.BorderBrush = brush(DQT_BORDER)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = WinCornerRadius(6)
        border.Padding = Thickness(14, 10, 14, 10)
        border.Margin = Thickness(0, 0, 0, 10)

        g = WPFGrid()
        # Columns: Gauge | Score Text | Buttons + Mini Gauges
        for w in [GridLength(1, GridUnitType.Auto),
                   GridLength(1, GridUnitType.Star),
                   GridLength(1, GridUnitType.Auto)]:
            cd = ColumnDefinition()
            cd.Width = w
            g.ColumnDefinitions.Add(cd)

        # Col 0: Main gauge
        self.gauge_container = StackPanel()
        self.gauge_container.VerticalAlignment = VerticalAlignment.Center
        self.gauge_container.Margin = Thickness(0, 0, 15, 0)
        
        # Placeholder gauge - will be updated
        self.score_circle = Border()
        self.score_circle.Width = 100
        self.score_circle.Height = 100
        self.score_circle.CornerRadius = WinCornerRadius(50)
        self.score_circle.Background = brush(HEALTH_GREEN)
        cs = StackPanel()
        cs.VerticalAlignment = VerticalAlignment.Center
        cs.HorizontalAlignment = HorizontalAlignment.Center
        self.txt_grade = TextBlock()
        self.txt_grade.Text = "?"
        self.txt_grade.FontSize = 32
        self.txt_grade.FontWeight = FontWeights.Bold
        self.txt_grade.Foreground = brush("#FFFFFF")
        self.txt_grade.HorizontalAlignment = HorizontalAlignment.Center
        cs.Children.Add(self.txt_grade)
        self.txt_score_num = TextBlock()
        self.txt_score_num.Text = "--"
        self.txt_score_num.FontSize = 12
        self.txt_score_num.Foreground = brush("#FFFFFF")
        self.txt_score_num.HorizontalAlignment = HorizontalAlignment.Center
        self.txt_score_num.Opacity = 0.9
        cs.Children.Add(self.txt_score_num)
        self.score_circle.Child = cs
        self.gauge_container.Children.Add(self.score_circle)
        
        WPFGrid.SetColumn(self.gauge_container, 0)
        g.Children.Add(self.gauge_container)

        # Col 1: Score text
        mid = StackPanel()
        mid.VerticalAlignment = VerticalAlignment.Center
        self.txt_score_label = TextBlock()
        self.txt_score_label.Text = "Analyzing..."
        self.txt_score_label.FontSize = 18
        self.txt_score_label.FontWeight = FontWeights.Bold
        self.txt_score_label.Foreground = brush(DQT_TEXT)
        mid.Children.Add(self.txt_score_label)
        self.txt_summary = TextBlock()
        self.txt_summary.FontSize = 12
        self.txt_summary.Foreground = brush("#475569")
        self.txt_summary.Margin = Thickness(0, 4, 0, 0)
        self.txt_summary.TextWrapping = TextWrapping.Wrap
        mid.Children.Add(self.txt_summary)
        self.txt_date = TextBlock()
        self.txt_date.FontSize = 10
        self.txt_date.Foreground = brush("#94A3B8")
        self.txt_date.Margin = Thickness(0, 4, 0, 0)
        mid.Children.Add(self.txt_date)

        # Metric counts
        self.txt_metric_counts = TextBlock()
        self.txt_metric_counts.FontSize = 11
        self.txt_metric_counts.Foreground = brush("#888888")
        self.txt_metric_counts.Margin = Thickness(0, 4, 0, 0)
        mid.Children.Add(self.txt_metric_counts)

        WPFGrid.SetColumn(mid, 1)
        g.Children.Add(mid)

        # Col 2: Buttons
        right_panel = StackPanel()
        right_panel.VerticalAlignment = VerticalAlignment.Center

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.Margin = Thickness(0, 0, 0, 8)

        self.btn_refresh = Button()
        self.btn_refresh.Content = u"\u27F3 Re-Analyze"
        self.btn_refresh.Background = brush(DQT_PRIMARY)
        self.btn_refresh.Foreground = brush(DQT_TEXT_DARK)
        self.btn_refresh.FontWeight = FontWeights.SemiBold
        self.btn_refresh.Padding = Thickness(16, 8, 16, 8)
        self.btn_refresh.BorderBrush = brush(DQT_PRIMARY_DARK)
        self.btn_refresh.BorderThickness = Thickness(1)
        self.btn_refresh.Margin = Thickness(0, 0, 8, 0)
        self.btn_refresh.Cursor = Cursors.Hand
        self.btn_refresh.Click += self._on_refresh
        btn_row.Children.Add(self.btn_refresh)

        self.btn_export = Button()
        self.btn_export.Content = u"Export Report"
        self.btn_export.Background = brush("#FFFFFF")
        self.btn_export.Foreground = brush(DQT_TEXT_DARK)
        self.btn_export.Padding = Thickness(12, 8, 12, 8)
        self.btn_export.BorderBrush = brush(DQT_PRIMARY_DARK)
        self.btn_export.BorderThickness = Thickness(1)
        self.btn_export.Cursor = Cursors.Hand
        self.btn_export.Click += self._on_export
        btn_row.Children.Add(self.btn_export)

        right_panel.Children.Add(btn_row)

        # Score breakdown text
        self.txt_breakdown = TextBlock()
        self.txt_breakdown.FontSize = 10
        self.txt_breakdown.Foreground = brush("#888888")
        self.txt_breakdown.TextWrapping = TextWrapping.Wrap
        self.txt_breakdown.MaxWidth = 220
        right_panel.Children.Add(self.txt_breakdown)

        WPFGrid.SetColumn(right_panel, 2)
        g.Children.Add(right_panel)

        border.Child = g
        return border

    # ---- LEGEND ----
    def _make_legend(self):
        border = Border()
        border.Background = brush("#FFFFFF")
        border.BorderBrush = brush(DQT_BORDER)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = WinCornerRadius(6)
        border.Padding = Thickness(10, 6, 10, 6)
        border.Margin = Thickness(0, 0, 0, 10)
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.HorizontalAlignment = HorizontalAlignment.Center
        lbl = TextBlock()
        lbl.Text = "Health Scale:  "
        lbl.FontSize = 11
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Foreground = brush(DQT_TEXT_DARK)
        lbl.VerticalAlignment = VerticalAlignment.Center
        sp.Children.Add(lbl)
        for color, text in [
            (HEALTH_GREEN, "Good"), (HEALTH_LIGHT_GREEN, "Acceptable"),
            (HEALTH_YELLOW, "Warning"), (HEALTH_ORANGE, "Concerning"),
            (HEALTH_RED, "Critical"), (HEALTH_DARK_RED, "Severe")]:
            b = Border()
            b.Background = brush(color)
            b.CornerRadius = WinCornerRadius(3)
            b.Padding = Thickness(8, 3, 8, 3)
            b.Margin = Thickness(2, 0, 2, 0)
            t = TextBlock()
            t.Text = text
            t.FontSize = 10
            t.Foreground = brush("#FFFFFF")
            t.FontWeight = FontWeights.SemiBold
            b.Child = t
            sp.Children.Add(b)
        border.Child = sp
        return border

    # ---- FOOTER ----
    def _make_footer(self):
        border = Border()
        border.Background = brush(DQT_PRIMARY)
        border.CornerRadius = WinCornerRadius(4)
        border.Padding = Thickness(10, 6, 10, 6)
        t = TextBlock()
        t.Text = u"Copyright by Dang Quoc Truong - DQT \u00A9 2025"
        t.FontSize = 10
        t.Foreground = brush(DQT_TEXT_DARK)
        t.HorizontalAlignment = HorizontalAlignment.Center
        border.Child = t
        return border

    # ----------------------------------------------------------
    # ANALYSIS
    # ----------------------------------------------------------
    def _run_analysis(self):
        try:
            proj = self.doc.ProjectInformation
            proj_name = proj.Name if proj else "Untitled"
            file_name = os.path.basename(self.doc.PathName) if self.doc.PathName else "Unsaved"
            self.txt_project.Text = "Project: {}  |  File: {}".format(proj_name, file_name)

            self.analyzer = ModelHealthAnalyzer(self.doc)
            self.metrics = self.analyzer.analyze()

            self._update_score()
            self._build_heatmap()
            self._build_recommendations()

            self.txt_date.Text = "Last analyzed: {}".format(
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        except Exception as ex:
            self.txt_score_label.Text = "Error: " + str(ex)

    def _update_score(self):
        score = get_health_score(self.metrics)
        grade, label, color = get_health_grade(score)

        self.txt_grade.Text = grade
        self.txt_score_num.Text = str(score)
        self.txt_score_label.Text = "Model Health: {}".format(label)
        self.score_circle.Background = brush(color)

        # Update gauge
        self.gauge_container.Children.Clear()
        try:
            gauge = create_gauge(score, 130)
            self.gauge_container.Children.Add(gauge)
        except:
            # Fallback to circle if gauge fails
            self.gauge_container.Children.Add(self.score_circle)

        # Grade label under gauge
        grade_label = TextBlock()
        grade_label.Text = "{} - {}".format(grade, label)
        grade_label.FontSize = 13
        grade_label.FontWeight = FontWeights.Bold
        grade_label.Foreground = brush(color)
        grade_label.HorizontalAlignment = HorizontalAlignment.Center
        grade_label.Margin = Thickness(0, 0, 0, 0)
        self.gauge_container.Children.Add(grade_label)

        # Count issues
        good = 0
        acceptable = 0
        warn = 0
        concern = 0
        crit = 0
        severe = 0
        for key, value in self.metrics.items():
            if key in METRIC_THRESHOLDS:
                status = get_status_text(key, value)
                if status == "Good": good += 1
                elif status == "Acceptable": acceptable += 1
                elif status == "Warning": warn += 1
                elif status == "Concerning": concern += 1
                elif status == "Critical": crit += 1
                elif status == "Severe": severe += 1

        total = good + acceptable + warn + concern + crit + severe

        if crit + severe > 0:
            self.txt_summary.Text = "{} critical/severe issues, {} warnings. Immediate attention needed.".format(crit + severe, warn + concern)
        elif warn + concern > 0:
            self.txt_summary.Text = "{} warnings found. Review recommended.".format(warn + concern)
        else:
            self.txt_summary.Text = "All metrics within acceptable ranges. Model is healthy."

        self.txt_metric_counts.Text = "Total: {} metrics | Good: {} | Acceptable: {} | Warning: {} | Concerning: {} | Critical: {} | Severe: {}".format(
            total, good, acceptable, warn, concern, crit, severe)

        self.txt_breakdown.Text = "Weighted Score: {}/100\nGrade: {} ({})".format(score, grade, label)

    # ----------------------------------------------------------
    # HEATMAP TABLE
    # ----------------------------------------------------------
    def _build_heatmap(self):
        self.metrics_stack.Children.Clear()

        table = WPFGrid()
        # Wider columns: Metric | Value | Bar | Status | Select
        col_widths = [180, 80, 400, 340, 90]
        for w in col_widths:
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            table.ColumnDefinitions.Add(cd)

        # Sort by severity
        sorted_keys = [k for k in METRIC_THRESHOLDS if k in self.metrics]
        def sev(k):
            v = self.metrics[k]
            t = METRIC_THRESHOLDS[k]["thresholds"]
            if v > t[4]: return 0
            elif v > t[3]: return 1
            elif v > t[2]: return 2
            elif v > t[1]: return 3
            elif v > t[0]: return 4
            return 5
        sorted_keys.sort(key=sev)

        # Header row
        rd = RowDefinition()
        rd.Height = GridLength(36)
        table.RowDefinitions.Add(rd)
        for ci, htxt in enumerate(["METRIC", "VALUE", "HEALTH INDICATOR", "STATUS", "ACTION"]):
            b = Border()
            b.Background = brush(DQT_PRIMARY)
            b.BorderBrush = brush(DQT_PRIMARY_DARK)
            b.BorderThickness = Thickness(0, 0, 1, 2)
            b.Padding = Thickness(10, 6, 10, 6)
            tb = TextBlock()
            tb.Text = htxt
            tb.FontSize = 11
            tb.FontWeight = FontWeights.Bold
            tb.Foreground = brush(DQT_TEXT_DARK)
            tb.VerticalAlignment = VerticalAlignment.Center
            if ci >= 4:
                tb.HorizontalAlignment = HorizontalAlignment.Center
            b.Child = tb
            WPFGrid.SetRow(b, 0)
            WPFGrid.SetColumn(b, ci)
            table.Children.Add(b)

        # Data rows
        row_idx = 0
        for key in sorted_keys:
            row_idx += 1
            rd = RowDefinition()
            rd.Height = GridLength(50)  # Taller rows
            table.RowDefinitions.Add(rd)

            value = self.metrics[key]
            config = METRIC_THRESHOLDS[key]
            h_color = get_health_color(key, value)
            thresholds = config["thresholds"]
            row_bg = "#FFFFFF" if row_idx % 2 == 0 else "#FAF8F0"
            is_selectable = config.get("selectable", False) and value > 0

            # Col 0: Name
            b0 = Border()
            b0.Background = brush(row_bg)
            b0.BorderBrush = brush("#E8E0D0")
            b0.BorderThickness = Thickness(0, 0, 1, 1)
            b0.Padding = Thickness(10, 6, 10, 6)
            b0.ToolTip = config["tooltip"]
            t0 = TextBlock()
            t0.Text = config["label"]
            t0.FontSize = 12
            t0.FontWeight = FontWeights.SemiBold
            t0.Foreground = brush(DQT_TEXT)
            t0.VerticalAlignment = VerticalAlignment.Center
            b0.Child = t0
            WPFGrid.SetRow(b0, row_idx)
            WPFGrid.SetColumn(b0, 0)
            table.Children.Add(b0)

            # Col 1: Value
            b1 = Border()
            b1.Background = brush(h_color)
            b1.BorderBrush = brush("#E8E0D0")
            b1.BorderThickness = Thickness(0, 0, 1, 1)
            b1.Padding = Thickness(8, 6, 8, 6)
            t1 = TextBlock()
            unit = config.get("unit", "")
            t1.Text = "{}{}".format(value, " " + unit if unit else "")
            t1.FontSize = 13
            t1.FontWeight = FontWeights.Bold
            t1.Foreground = brush("#FFFFFF")
            t1.HorizontalAlignment = HorizontalAlignment.Center
            t1.VerticalAlignment = VerticalAlignment.Center
            b1.Child = t1
            WPFGrid.SetRow(b1, row_idx)
            WPFGrid.SetColumn(b1, 1)
            table.Children.Add(b1)

            # Col 2: Bar - using Canvas for pixel-perfect rendering
            b2 = Border()
            b2.Background = brush(row_bg)
            b2.BorderBrush = brush("#E8E0D0")
            b2.BorderThickness = Thickness(0, 0, 1, 1)
            b2.Padding = Thickness(10, 12, 10, 12)
            
            BAR_W = 370
            BAR_H = 16
            bar_canvas = Canvas()
            bar_canvas.Width = BAR_W
            bar_canvas.Height = BAR_H
            
            # Background bar
            bar_bg = Border()
            bar_bg.Width = BAR_W
            bar_bg.Height = BAR_H
            bar_bg.Background = brush("#E0DDD5")
            bar_bg.CornerRadius = WinCornerRadius(3)
            Canvas.SetLeft(bar_bg, 0)
            Canvas.SetTop(bar_bg, 0)
            bar_canvas.Children.Add(bar_bg)
            
            # Fill bar
            max_t = float(thresholds[4])
            if max_t > 0:
                ratio = min(value / max_t, 1.0)
            else:
                ratio = 0
            fill_w = max(int(ratio * BAR_W), 3) if value > 0 else 0
            if fill_w > 0:
                bar_fill = Border()
                bar_fill.Width = fill_w
                bar_fill.Height = BAR_H
                bar_fill.Background = brush(h_color)
                bar_fill.CornerRadius = WinCornerRadius(3)
                Canvas.SetLeft(bar_fill, 0)
                Canvas.SetTop(bar_fill, 0)
                bar_canvas.Children.Add(bar_fill)
            
            # Threshold markers
            for tv in thresholds:
                if tv > 0 and max_t > 0:
                    mr = tv / max_t
                    if mr <= 1.0:
                        mk = Border()
                        mk.Width = 1
                        mk.Height = BAR_H
                        mk.Background = brush("#94A3B8")
                        mk.Opacity = 0.4
                        Canvas.SetLeft(mk, int(mr * BAR_W))
                        Canvas.SetTop(mk, 0)
                        bar_canvas.Children.Add(mk)
            
            b2.Child = bar_canvas
            WPFGrid.SetRow(b2, row_idx)
            WPFGrid.SetColumn(b2, 2)
            table.Children.Add(b2)

            # Col 3: Status (WIDER - full info)
            b3 = Border()
            b3.Background = brush(row_bg)
            b3.BorderBrush = brush("#E8E0D0")
            b3.BorderThickness = Thickness(0, 0, 1, 1)
            b3.Padding = Thickness(10, 5, 10, 5)

            info_sp = StackPanel()
            info_sp.VerticalAlignment = VerticalAlignment.Center

            status = get_status_text(key, value)
            weight = config.get("weight", 1)
            weight_stars = u"\u2605" * weight + u"\u2606" * (5 - weight)

            st = TextBlock()
            st.Text = "Status: {}".format(status)
            st.FontSize = 11
            st.FontWeight = FontWeights.SemiBold
            st.Foreground = brush(h_color)
            info_sp.Children.Add(st)

            wt = TextBlock()
            wt.Text = "Weight: {} ({}/5)".format(weight_stars, weight)
            wt.FontSize = 10
            wt.Foreground = brush("#888888")
            wt.Margin = Thickness(0, 2, 0, 0)
            info_sp.Children.Add(wt)

            tt = TextBlock()
            tt.Text = "Thresholds: {} | {} | {} | {} | {}".format(*thresholds)
            tt.FontSize = 9
            tt.Foreground = brush("#AAAAAA")
            tt.Margin = Thickness(0, 2, 0, 0)
            info_sp.Children.Add(tt)

            b3.Child = info_sp
            WPFGrid.SetRow(b3, row_idx)
            WPFGrid.SetColumn(b3, 3)
            table.Children.Add(b3)

            # Col 4: Select Button
            b4 = Border()
            b4.Background = brush(row_bg)
            b4.BorderBrush = brush("#E8E0D0")
            b4.BorderThickness = Thickness(0, 0, 0, 1)
            b4.Padding = Thickness(4, 4, 4, 4)
            if is_selectable:
                btn = Button()
                btn.Content = u"\u25BA Select"
                btn.FontSize = 10
                btn.Padding = Thickness(8, 4, 8, 4)
                btn.Background = brush("#FFFFFF")
                btn.Foreground = brush(DQT_TEXT_DARK)
                btn.BorderBrush = brush(DQT_PRIMARY_DARK)
                btn.BorderThickness = Thickness(1)
                btn.Cursor = Cursors.Hand
                btn.VerticalAlignment = VerticalAlignment.Center
                btn.HorizontalAlignment = HorizontalAlignment.Center
                btn.Tag = key
                btn.Click += self._on_select_elements
                b4.Child = btn
            else:
                na = TextBlock()
                na.Text = "--"
                na.FontSize = 10
                na.Foreground = brush("#CCCCCC")
                na.HorizontalAlignment = HorizontalAlignment.Center
                na.VerticalAlignment = VerticalAlignment.Center
                b4.Child = na
            WPFGrid.SetRow(b4, row_idx)
            WPFGrid.SetColumn(b4, 4)
            table.Children.Add(b4)

        self.metrics_stack.Children.Add(table)

    # ----------------------------------------------------------
    # RECOMMENDATIONS
    # ----------------------------------------------------------
    def _build_recommendations(self):
        self.rec_panel.Children.Clear()
        title = TextBlock()
        title.Text = "RECOMMENDATIONS"
        title.FontSize = 13
        title.FontWeight = FontWeights.Bold
        title.Foreground = brush(DQT_TEXT_DARK)
        title.Margin = Thickness(0, 0, 0, 8)
        self.rec_panel.Children.Add(title)

        has_rec = False
        for key in METRIC_THRESHOLDS:
            if key in self.metrics:
                value = self.metrics[key]
                t = METRIC_THRESHOLDS[key]["thresholds"]
                if value > t[2]:
                    has_rec = True
                    h_color = get_health_color(key, value)
                    label = METRIC_THRESHOLDS[key]["label"]
                    rec = RECOMMENDATIONS.get(key, "Review and optimize.")
                    row = StackPanel()
                    row.Orientation = Orientation.Horizontal
                    row.Margin = Thickness(0, 2, 0, 2)
                    dot = Border()
                    dot.Width = 8
                    dot.Height = 8
                    dot.CornerRadius = WinCornerRadius(4)
                    dot.Background = brush(h_color)
                    dot.VerticalAlignment = VerticalAlignment.Center
                    dot.Margin = Thickness(0, 0, 8, 0)
                    row.Children.Add(dot)
                    lbl = TextBlock()
                    lbl.Text = "{} ({}): ".format(label, value)
                    lbl.FontSize = 11
                    lbl.FontWeight = FontWeights.SemiBold
                    lbl.Foreground = brush(DQT_TEXT)
                    lbl.VerticalAlignment = VerticalAlignment.Center
                    row.Children.Add(lbl)
                    rtb = TextBlock()
                    rtb.Text = rec
                    rtb.FontSize = 11
                    rtb.Foreground = brush("#475569")
                    rtb.VerticalAlignment = VerticalAlignment.Center
                    rtb.TextWrapping = TextWrapping.Wrap
                    rtb.MaxWidth = 800
                    row.Children.Add(rtb)
                    self.rec_panel.Children.Add(row)

        if not has_rec:
            good = TextBlock()
            good.Text = "All metrics within acceptable ranges. No action required."
            good.FontSize = 12
            good.Foreground = brush(HEALTH_GREEN)
            good.FontWeight = FontWeights.SemiBold
            self.rec_panel.Children.Add(good)

    # ----------------------------------------------------------
    # SELECT ELEMENTS
    # ----------------------------------------------------------
    def _on_select_elements(self, sender, args):
        try:
            metric_key = sender.Tag
            if not self.analyzer or metric_key not in self.analyzer.element_ids:
                MessageBox.Show("No elements available.\nTry Re-Analyze first.",
                    "Select Elements", MessageBoxButton.OK, MessageBoxImage.Information)
                return
            ids = self.analyzer.element_ids[metric_key]
            if not ids:
                MessageBox.Show("No elements found.", "Select Elements",
                    MessageBoxButton.OK, MessageBoxImage.Information)
                return
            MAX_SELECT = 5000
            if len(ids) > MAX_SELECT:
                result = MessageBox.Show(
                    "{} elements found.\nSelect first {}?".format(len(ids), MAX_SELECT),
                    "Large Selection", MessageBoxButton.YesNo, MessageBoxImage.Warning)
                if result != MessageBoxResult.Yes:
                    return
                ids = ids[:MAX_SELECT]
            id_list = List[ElementId]()
            for eid in ids:
                id_list.Add(eid)
            self.Close()
            self.uidoc.Selection.SetElementIds(id_list)
            try:
                if id_list.Count > 0:
                    self.uidoc.ShowElements(id_list)
            except:
                pass
        except Exception as ex:
            MessageBox.Show("Error selecting elements:\n{}".format(str(ex)),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    # ----------------------------------------------------------
    # EVENTS
    # ----------------------------------------------------------
    def _on_refresh(self, sender, args):
        self._run_analysis()

    def _on_export(self, sender, args):
        """Export report as PDF using HTML rendering"""
        try:
            desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            fname = os.path.basename(self.doc.PathName).replace(".rvt", "") if self.doc.PathName else "Model"
            fname = fname.replace(" ", "_")
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            
            proj = self.doc.ProjectInformation
            proj_name = proj.Name if proj else "Untitled"
            file_name = os.path.basename(self.doc.PathName) if self.doc.PathName else "Unsaved"
            
            score = get_health_score(self.metrics)
            grade, label, grade_color = get_health_grade(score)
            
            # Count issues
            good = acceptable = warn = concern = crit = severe = 0
            for key, value in self.metrics.items():
                if key in METRIC_THRESHOLDS:
                    st = get_status_text(key, value)
                    if st == "Good": good += 1
                    elif st == "Acceptable": acceptable += 1
                    elif st == "Warning": warn += 1
                    elif st == "Concerning": concern += 1
                    elif st == "Critical": crit += 1
                    elif st == "Severe": severe += 1
            
            # Sort by severity
            sorted_keys = [k for k in METRIC_THRESHOLDS if k in self.metrics]
            def sev(k):
                v = self.metrics[k]
                t = METRIC_THRESHOLDS[k]["thresholds"]
                if v > t[4]: return 0
                elif v > t[3]: return 1
                elif v > t[2]: return 2
                elif v > t[1]: return 3
                elif v > t[0]: return 4
                return 5
            sorted_keys.sort(key=sev)
            
            # Build metric rows HTML
            rows_html = ""
            for i, key in enumerate(sorted_keys):
                value = self.metrics[key]
                config = METRIC_THRESHOLDS[key]
                h_color = get_health_color(key, value)
                thresholds = config["thresholds"]
                status = get_status_text(key, value)
                weight = config.get("weight", 1)
                weight_stars = "&#9733;" * weight + "&#9734;" * (5 - weight)
                unit = config.get("unit", "")
                val_str = "{}{}".format(value, " " + unit if unit else "")
                row_bg = "#FFFFFF" if i % 2 == 0 else "#FAF8F0"
                
                # Bar width as percentage of threshold[4]
                max_t = float(thresholds[4])
                bar_pct = min(value / max(max_t, 1) * 100, 100) if value > 0 else 0
                empty_pct = 100 - bar_pct
                
                rows_html += """
                <tr style="background:{bg};">
                    <td style="padding:8px 10px;font-weight:600;border-bottom:1px solid #E8E0D0;">{label}</td>
                    <td style="padding:8px;text-align:center;background:{color};color:#FFF;font-weight:700;border-bottom:1px solid #E8E0D0;">{val}</td>
                    <td style="padding:8px 10px;border-bottom:1px solid #E8E0D0;">
                        <table style="width:100%;border-collapse:collapse;height:14px;table-layout:fixed;"><tr>
                            <td style="width:{pct}%;background:{color};border-radius:3px 0 0 3px;height:14px;padding:0;"></td>
                            <td style="width:{epct}%;background:#E0DDD5;border-radius:0 3px 3px 0;height:14px;padding:0;"></td>
                        </tr></table>
                    </td>
                    <td style="padding:6px 10px;border-bottom:1px solid #E8E0D0;">
                        <div style="color:{color};font-weight:600;font-size:11px;">Status: {status}</div>
                        <div style="color:#888;font-size:9px;">Weight: {stars} ({w}/5)</div>
                        <div style="color:#AAA;font-size:8px;">Thresholds: {t0} | {t1} | {t2} | {t3} | {t4}</div>
                    </td>
                </tr>""".format(
                    bg=row_bg, label=config["label"], color=h_color, val=val_str,
                    pct=round(bar_pct, 1), epct=round(empty_pct, 1),
                    status=status, stars=weight_stars, w=weight,
                    t0=thresholds[0], t1=thresholds[1], t2=thresholds[2], t3=thresholds[3], t4=thresholds[4])
            
            # Build recommendations HTML
            rec_html = ""
            has_rec = False
            for key in METRIC_THRESHOLDS:
                if key in self.metrics:
                    value = self.metrics[key]
                    t = METRIC_THRESHOLDS[key]["thresholds"]
                    if value > t[2]:
                        has_rec = True
                        h_color = get_health_color(key, value)
                        label = METRIC_THRESHOLDS[key]["label"]
                        rec = RECOMMENDATIONS.get(key, "Review and optimize.")
                        rec_html += '<div style="margin:4px 0;"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:{};margin-right:8px;vertical-align:middle;"></span><strong>{} ({}):</strong> <span style="color:#666;">{}</span></div>'.format(
                            h_color, label, value, rec)
            if not has_rec:
                rec_html = '<div style="color:#10B981;font-weight:600;">All metrics within acceptable ranges. No action required.</div>'
            
            # Full HTML
            html = u"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Model Health Check Report - DQT</title>
<style>
    @page {{ size: A4 landscape; margin: 12mm; }}
    * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }}
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 15px; background: #F8FAFC; color: #333; font-size: 11px; }}
    .header {{ background: #0F172A; padding: 16px 20px; border-radius: 6px; margin-bottom: 15px; overflow: hidden; }}
    .header h1 {{ margin: 0; font-size: 22px; color: #5D4E37; float: left; }}
    .header .right {{ float: right; text-align: right; color: #5D4E37; }}
    .header .sub {{ font-size: 11px; color: #5D4E37; margin-top: 4px; }}
    .score-card {{ background: #FFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 15px 20px; margin-bottom: 15px; overflow: hidden; }}
    .gauge {{ float: left; width: 90px; height: 90px; border-radius: 50%; background: {grade_color}; text-align: center; margin-right: 20px; }}
    .gauge .grade {{ font-size: 32px; font-weight: bold; color: #FFF; margin-top: 16px; }}
    .gauge .num {{ font-size: 12px; color: #FFF; opacity: 0.9; }}
    .score-text h2 {{ margin: 0 0 5px 0; font-size: 18px; color: #333; }}
    .score-text .summary {{ color: #666; font-size: 12px; }}
    .score-text .counts {{ color: #888; font-size: 10px; margin-top: 4px; }}
    .legend {{ background: #FFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 8px 15px; margin-bottom: 15px; text-align: center; }}
    .legend span {{ display: inline-block; padding: 3px 10px; border-radius: 3px; color: #FFF; font-size: 10px; font-weight: 600; margin: 0 2px; -webkit-print-color-adjust: exact; }}
    table {{ width: 100%; border-collapse: collapse; background: #FFF; border: 1px solid #CBD5E1; }}
    th {{ background: #0F172A !important; color: #5D4E37; padding: 8px 10px; text-align: left; font-size: 11px; border-bottom: 2px solid #CBD5E1; }}
    .rec-box {{ background: #FFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 12px 15px; margin-top: 15px; }}
    .rec-box h3 {{ margin: 0 0 8px 0; color: #5D4E37; font-size: 13px; }}
    .footer {{ background: #0F172A; border-radius: 4px; padding: 8px; text-align: center; margin-top: 15px; font-size: 10px; color: #5D4E37; }}
</style>
</head>
<body>
    <div class="header">
        <h1>MODEL HEALTH CHECK</h1>
        <div class="right">
            <div style="font-size:14px;font-weight:600;">pyDQT Suite</div>
            <div style="font-size:9px;opacity:0.7;">Copyright by Dang Quoc Truong - DQT</div>
        </div>
        <div style="clear:both;"></div>
        <div class="sub">Project: {proj} &nbsp;|&nbsp; File: {file}</div>
    </div>

    <div class="score-card">
        <div class="gauge">
            <div class="grade">{grade}</div>
            <div class="num">{score}</div>
        </div>
        <div class="score-text">
            <h2>Model Health: {label}</h2>
            <div class="summary">{summary}</div>
            <div class="counts">Total: {total} metrics | Good: {good} | Acceptable: {acceptable} | Warning: {warn} | Concerning: {concern} | Critical: {crit} | Severe: {severe}</div>
            <div class="counts">Date: {date}</div>
        </div>
        <div style="clear:both;"></div>
    </div>

    <div class="legend">
        <strong>Health Scale:</strong>
        <span style="background:#10B981;">Good</span>
        <span style="background:#8BC34A;">Acceptable</span>
        <span style="background:#FFC107;">Warning</span>
        <span style="background:#F59E0B;">Concerning</span>
        <span style="background:#EF4444;">Critical</span>
        <span style="background:#D32F2F;">Severe</span>
    </div>

    <table>
        <tr>
            <th style="width:160px;">METRIC</th>
            <th style="width:70px;text-align:center;">VALUE</th>
            <th style="width:250px;">HEALTH INDICATOR</th>
            <th>STATUS</th>
        </tr>
        {rows}
    </table>

    <div class="rec-box">
        <h3>RECOMMENDATIONS</h3>
        {recs}
    </div>

    <div class="footer">Copyright by Dang Quoc Truong - DQT &copy; 2025</div>
</body>
</html>""".format(
                grade_color=grade_color,
                proj=proj_name, file=file_name,
                grade=grade, score=score, label=label,
                summary=self.txt_summary.Text,
                total=good+acceptable+warn+concern+crit+severe,
                good=good, acceptable=acceptable, warn=warn,
                concern=concern, crit=crit, severe=severe,
                date=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                rows=rows_html, recs=rec_html)
            
            # Save HTML with UTF-8 encoding
            html_path = os.path.join(desktop, "ModelHealth_{}_{}.html".format(fname, ts))
            with codecs.open(html_path, 'w', 'utf-8') as f:
                f.write(html)
            
            # Auto-open in default browser
            try:
                System.Diagnostics.Process.Start(html_path)
            except:
                pass
            
            MessageBox.Show(
                "Report exported!\n\n{}\n\nFile opened in browser.\nUse Ctrl+P > Save as PDF to create PDF.".format(html_path),
                "Export Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information)

        except Exception as ex:
            MessageBox.Show("Export error:\n{}".format(str(ex)),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================
# MAIN
# ============================================================
if __name__ == '__main__':
    try:
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        window = ModelHealthWindow(doc, uidoc)
        window.ShowDialog()
    except Exception as e:
        from pyrevit import forms
        forms.alert("Error launching Model Health Check:\n{}".format(str(e)),
                    title="Model Health Check - DQT",
                    exitscript=True)