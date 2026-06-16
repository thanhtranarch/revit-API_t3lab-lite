# -*- coding: utf-8 -*-
"""
Smart Delete Manager v9.1 - DQT
Analyzes element dependencies before deletion.

Safe ListBox approach with column-formatted text.
Compatible with Revit 2024 - 2027 (ElementId.Value / IntegerValue safe).

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Smart\nDelete"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Analyze element dependencies before deletion."

import clr
clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from pyrevit import revit, forms, script

from Autodesk.Revit.DB import (
    Transaction, ElementId, FilteredElementCollector,
    BuiltInCategory, BuiltInParameter,
    ElementClassFilter, BoundingBoxIntersectsFilter, Outline,
    FamilyInstance, Dimension, IndependentTag, XYZ
)
from Autodesk.Revit.UI.Selection import ObjectType

from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, GridLength, GridUnitType,
    FontWeights, TextWrapping, CornerRadius as CR
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, StackPanel, Border,
    TextBlock, TextBox, Button, ListBox, ComboBox, ComboBoxItem,
    Orientation, ScrollViewer, ScrollBarVisibility, SelectionMode
)
from System.Windows.Media import SolidColorBrush, Color, FontFamily
from System.Windows.Input import Cursors

import traceback

output = script.get_output()


# ============================================================================
# REVIT VERSION COMPATIBILITY (2024 - 2027)
# ============================================================================
# Revit 2026+ removed ElementId.IntegerValue in favor of ElementId.Value
# (which returns a 64-bit Int64). These helpers work across 2024-2027.

def eid_int(elem_id):
    """Get integer value of an ElementId, safe for Revit 2024-2027."""
    try:
        return elem_id.Value          # Revit 2024+ (preferred, Int64)
    except AttributeError:
        try:
            return elem_id.IntegerValue   # Revit 2024 and earlier
        except:
            return -1


def make_eid(value):
    """Construct an ElementId from an int, safe for Revit 2024-2027.

    Revit 2024+ added an ElementId(Int64) overload; older versions only
    accept Int32. Try the native value first, fall back to int().
    """
    try:
        return ElementId(value)
    except:
        try:
            from System import Int64
            return ElementId(Int64(value))
        except:
            return ElementId(int(value))


# ============================================================================
# FROZEN BRUSHES
# ============================================================================

def FB(r, g, b):
    brush = SolidColorBrush(Color.FromArgb(255, r, g, b))
    brush.Freeze()
    return brush

B_PRI = FB(240, 204, 136)
B_BG = FB(254, 248, 231)
B_BRD = FB(212, 184, 122)
B_TXT = FB(93, 78, 55)
B_SUB = FB(122, 107, 85)
B_WHT = FB(255, 255, 255)
B_GRY = FB(136, 136, 136)
B_HDR = FB(247, 238, 213)
B_CRIT = FB(211, 47, 47)
B_HIGH = FB(245, 124, 0)
B_SAFE = FB(56, 142, 60)


# ============================================================================
# DEPENDENCY ANALYSIS
# ============================================================================

class DepInfo:
    def __init__(self, eid, name, dep_type, severity, desc="", view_name=""):
        self.eid = eid
        self.name = name
        self.dep_type = dep_type
        self.severity = severity
        self.desc = desc
        self.view_name = view_name


def safe_name(elem):
    try:
        return elem.Name or ""
    except:
        return ""


def safe_cat(elem):
    try:
        return elem.Category.Name if elem.Category else ""
    except:
        return ""


def get_view_name(element, document):
    try:
        ov_id = element.OwnerViewId
        if ov_id and ov_id != ElementId.InvalidElementId:
            v = document.GetElement(ov_id)
            if v:
                return v.Name or ""
    except:
        pass
    return ""


def analyze_element(element, document):
    deps = []
    el_id = element.Id

    # 1. Group
    try:
        gid = element.GroupId
        if gid and gid != ElementId.InvalidElementId:
            g = document.GetElement(gid)
            if g:
                deps.append(DepInfo(eid_int(gid), safe_name(g) or "Group",
                    "Group", "Critical", "In group"))
    except:
        pass

    # 2. Assembly
    try:
        aid = element.AssemblyInstanceId
        if aid and aid != ElementId.InvalidElementId:
            a = document.GetElement(aid)
            if a:
                deps.append(DepInfo(eid_int(aid), safe_name(a) or "Assembly",
                    "Assembly", "High", "In assembly"))
    except:
        pass

    # 3. Hosted elements
    try:
        bb = element.get_BoundingBox(None)
        if bb:
            ol = Outline(bb.Min, bb.Max)
            bbf = BoundingBoxIntersectsFilter(ol)
            nearby = FilteredElementCollector(document).OfClass(FamilyInstance).WherePasses(bbf)
            n = 0
            for fi in nearby:
                if n >= 15 or fi.Id == el_id:
                    continue
                try:
                    h = fi.Host
                    if h and h.Id == el_id:
                        cn = safe_cat(fi)
                        deps.append(DepInfo(eid_int(fi.Id),
                            safe_name(fi) or "Hosted",
                            "Hosted ({})".format(cn) if cn else "Hosted",
                            "Critical", "DELETED with host"))
                        n += 1
                except:
                    continue
    except:
        pass

    # 4. Dimensions
    try:
        dim_filter = ElementClassFilter(Dimension)
        dim_ids = element.GetDependentElements(dim_filter)
        if dim_ids:
            n = 0
            for did in dim_ids:
                if n >= 20 or did == el_id:
                    continue
                try:
                    d = document.GetElement(did)
                    if d:
                        val = ""
                        try:
                            val = d.ValueString or ""
                        except:
                            pass
                        vn = get_view_name(d, document)
                        deps.append(DepInfo(eid_int(did),
                            "Dim: {}".format(val) if val else "Dim",
                            "Dimension", "High", "Will be deleted", vn))
                        n += 1
                except:
                    continue
    except:
        pass

    # 5. Tags
    try:
        tag_filter = ElementClassFilter(IndependentTag)
        tag_ids = element.GetDependentElements(tag_filter)
        if tag_ids:
            n = 0
            for tid in tag_ids:
                if n >= 20 or tid == el_id:
                    continue
                try:
                    t = document.GetElement(tid)
                    if t:
                        tt = ""
                        try:
                            tt = t.TagText or ""
                        except:
                            pass
                        vn = get_view_name(t, document)
                        deps.append(DepInfo(eid_int(tid),
                            tt or "Tag",
                            "Tag", "Medium", "Will be deleted", vn))
                        n += 1
                except:
                    continue
    except:
        pass

    # 6. Room boundaries
    try:
        p = element.get_Parameter(BuiltInParameter.WALL_ATTR_IS_ROOM_BOUNDING)
        if p and p.AsInteger() == 1:
            bb = element.get_BoundingBox(None)
            if bb:
                ol = Outline(bb.Min, bb.Max)
                bbf = BoundingBoxIntersectsFilter(ol)
                rooms = FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Rooms).WherePasses(bbf).WhereElementIsNotElementType()
                n = 0
                for rm in rooms:
                    if n >= 5:
                        break
                    try:
                        rname = ""
                        try:
                            rname = rm.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() or ""
                        except:
                            pass
                        rnum = ""
                        try:
                            rnum = rm.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() or ""
                        except:
                            pass
                        label = "{} {}".format(rnum, rname).strip() or "Room"
                        deps.append(DepInfo(eid_int(rm.Id), label,
                            "Room Boundary", "High", "Boundary may change"))
                        n += 1
                    except:
                        continue
    except:
        pass

    # 7. MEP
    try:
        cm = None
        try:
            mm = getattr(element, 'MEPModel', None)
            if mm:
                cm = mm.ConnectorManager
        except:
            pass
        if not cm:
            try:
                cm = getattr(element, 'ConnectorManager', None)
            except:
                pass
        if cm:
            seen = set()
            for c in cm.Connectors:
                try:
                    for rc in c.AllRefs:
                        try:
                            o = rc.Owner
                            if not o or o.Id == el_id:
                                continue
                            oid = eid_int(o.Id)
                            if oid in seen:
                                continue
                            seen.add(oid)
                            deps.append(DepInfo(oid, safe_name(o) or "MEP",
                                "MEP Connection", "Critical", "System disconnects"))
                        except:
                            continue
                except:
                    continue
    except:
        pass

    return deps


# ============================================================================
# ELEMENT DATA
# ============================================================================

class ElemData:
    SEV = {"Critical": 4, "High": 3, "Medium": 2, "Low": 1, "None": 0}

    def __init__(self, element):
        self.element = element
        self.eid = eid_int(element.Id)
        self.name = safe_name(element) or "Element"
        self.cat = safe_cat(element)
        self.deps = []
        self.dep_count = 0
        self.max_sev = "None"

    def set_deps(self, dep_list):
        self.deps = dep_list
        self.dep_count = len(dep_list)
        ms = "None"
        for d in dep_list:
            if self.SEV.get(d.severity, 0) > self.SEV.get(ms, 0):
                ms = d.severity
        self.max_sev = ms

    @property
    def is_safe(self):
        return self.dep_count == 0

    @property
    def risk_label(self):
        if self.dep_count == 0:
            return "SAFE"
        return "{} ({})".format(self.max_sev.upper(), self.dep_count)

    @property
    def list_text(self):
        """Column-formatted text: Name | Category | ID | Risk"""
        # Fixed widths for pseudo-columns (matching header)
        name = self.name[:28].ljust(28)
        cat = self.cat[:16].ljust(16)
        eid = str(self.eid).ljust(10)
        risk = self.risk_label
        return "{}  {}  {}  {}".format(name, cat, eid, risk)


# ============================================================================
# MAIN WINDOW
# ============================================================================

class SmartDeleteWindow(Window):

    def __init__(self):
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.all_items = []
        self.filtered_items = []
        self._updating = False

        self.Title = "Smart Delete Manager - DQT"
        self.Width = 1000
        self.Height = 680
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = B_BG

        self._build_ui()
        self._load_selection()

    def _build_ui(self):
        root = Grid()
        root.Margin = Thickness(15)

        for h in [GridLength.Auto, GridLength.Auto, GridLength.Auto,
                  GridLength(1, GridUnitType.Star), GridLength.Auto, GridLength.Auto]:
            rd = RowDefinition()
            rd.Height = h
            root.RowDefinitions.Add(rd)

        hdr = self._create_header()
        Grid.SetRow(hdr, 0)
        root.Children.Add(hdr)

        cards = self._create_cards()
        Grid.SetRow(cards, 1)
        root.Children.Add(cards)

        flt = self._create_filters()
        Grid.SetRow(flt, 2)
        root.Children.Add(flt)

        main = self._create_main()
        Grid.SetRow(main, 3)
        root.Children.Add(main)

        act = self._create_actions()
        Grid.SetRow(act, 4)
        root.Children.Add(act)

        ftr = self._create_footer()
        Grid.SetRow(ftr, 5)
        root.Children.Add(ftr)

        self.Content = root

    def _create_header(self):
        b = Border()
        b.Background = B_PRI
        b.CornerRadius = CR(6)
        b.Padding = Thickness(15, 12, 15, 12)
        b.Margin = Thickness(0, 0, 0, 12)

        g = Grid()
        g.ColumnDefinitions.Add(ColumnDefinition())
        cd = ColumnDefinition()
        cd.Width = GridLength.Auto
        g.ColumnDefinitions.Add(cd)

        sp = StackPanel()
        t1 = TextBlock()
        t1.Text = "Smart Delete Manager"
        t1.FontSize = 20
        t1.FontWeight = FontWeights.Bold
        t1.Foreground = B_TXT
        sp.Children.Add(t1)

        t2 = TextBlock()
        t2.Text = "Analyze dependencies before deleting elements"
        t2.FontSize = 11
        t2.Foreground = B_SUB
        t2.Margin = Thickness(0, 2, 0, 0)
        sp.Children.Add(t2)

        Grid.SetColumn(sp, 0)
        g.Children.Add(sp)

        self.btnPick = self._btn("+ Pick Elements", True)
        self.btnPick.Click += self._on_pick
        Grid.SetColumn(self.btnPick, 1)
        g.Children.Add(self.btnPick)

        b.Child = g
        return b

    def _create_cards(self):
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        sp.Margin = Thickness(0, 0, 0, 10)

        self._cards = {}
        for name, color in [("Selected", B_TXT), ("Dependencies", B_HIGH),
                            ("Critical", B_CRIT), ("High Risk", B_HIGH), ("Safe", B_SAFE)]:
            bd = Border()
            bd.Background = B_WHT
            bd.BorderBrush = B_BRD
            bd.BorderThickness = Thickness(1)
            bd.CornerRadius = CR(4)
            bd.Padding = Thickness(15, 8, 15, 8)
            bd.Margin = Thickness(0, 0, 8, 0)
            bd.MinWidth = 100

            inner = StackPanel()
            lbl = TextBlock()
            lbl.Text = name.upper()
            lbl.FontSize = 9
            lbl.Foreground = B_GRY
            inner.Children.Add(lbl)

            vt = TextBlock()
            vt.Text = "0"
            vt.FontSize = 24
            vt.FontWeight = FontWeights.Bold
            vt.Foreground = color
            inner.Children.Add(vt)

            self._cards[name] = vt
            bd.Child = inner
            sp.Children.Add(bd)

        return sp

    def _create_filters(self):
        b = Border()
        b.Background = B_WHT
        b.BorderBrush = B_BRD
        b.BorderThickness = Thickness(1)
        b.CornerRadius = CR(4)
        b.Padding = Thickness(12, 8, 12, 8)
        b.Margin = Thickness(0, 0, 0, 10)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        sp.Children.Add(self._label("Search:"))
        self.txtSearch = TextBox()
        self.txtSearch.Width = 180
        self.txtSearch.Margin = Thickness(5, 0, 15, 0)
        self.txtSearch.Padding = Thickness(5, 3, 5, 3)
        self.txtSearch.TextChanged += self._on_filter
        sp.Children.Add(self.txtSearch)

        sp.Children.Add(self._label("Category:"))
        self.cmbCat = ComboBox()
        self.cmbCat.Width = 150
        self.cmbCat.Margin = Thickness(5, 0, 15, 0)
        self.cmbCat.SelectionChanged += self._on_filter
        sp.Children.Add(self.cmbCat)

        sp.Children.Add(self._label("Risk:"))
        self.cmbRisk = ComboBox()
        self.cmbRisk.Width = 130
        self.cmbRisk.Margin = Thickness(5, 0, 0, 0)
        self.cmbRisk.SelectionChanged += self._on_filter
        sp.Children.Add(self.cmbRisk)

        self._init_filters()

        b.Child = sp
        return b

    def _create_main(self):
        g = Grid()
        g.Margin = Thickness(0, 0, 0, 10)

        c1 = ColumnDefinition()
        c1.Width = GridLength(6, GridUnitType.Star)
        c2 = ColumnDefinition()
        c2.Width = GridLength(4, GridUnitType.Star)
        g.ColumnDefinitions.Add(c1)
        g.ColumnDefinitions.Add(c2)

        # Left: Element list with header
        left_container = StackPanel()
        
        # Header
        hdr = Border()
        hdr.Background = B_HDR
        hdr.Padding = Thickness(10, 8, 10, 8)
        hdr_txt = TextBlock()
        hdr_txt.Text = "Elements to Delete"
        hdr_txt.FontWeight = FontWeights.SemiBold
        hdr_txt.Foreground = B_TXT
        hdr.Child = hdr_txt
        left_container.Children.Add(hdr)

        # Column header row (text-based, same widths as list_text)
        col_hdr = Border()
        col_hdr.Background = B_PRI
        col_hdr.Padding = Thickness(8, 6, 8, 6)
        col_txt = TextBlock()
        col_txt.Text = "{}  {}  {}  {}".format(
            "Name".ljust(28), "Category".ljust(16), "ID".ljust(10), "Risk")
        col_txt.FontWeight = FontWeights.SemiBold
        col_txt.FontSize = 13
        col_txt.Foreground = B_TXT
        col_txt.FontFamily = FontFamily("Segoe UI")
        col_hdr.Child = col_txt
        left_container.Children.Add(col_hdr)

        # ListBox with matching font for alignment
        self.lstElem = ListBox()
        self.lstElem.BorderThickness = Thickness(0)
        self.lstElem.SelectionMode = SelectionMode.Single
        self.lstElem.FontFamily = FontFamily("Segoe UI")
        self.lstElem.FontSize = 13
        left_container.Children.Add(self.lstElem)

        lb = Border()
        lb.Background = B_WHT
        lb.BorderBrush = B_BRD
        lb.BorderThickness = Thickness(1)
        lb.CornerRadius = CR(4)
        lb.Margin = Thickness(0, 0, 5, 0)
        lb.Child = left_container
        Grid.SetColumn(lb, 0)
        g.Children.Add(lb)

        # Right: Details with matching font for column alignment
        right = self._panel("Dependency Details")
        self.txtDeps = TextBlock()
        self.txtDeps.TextWrapping = TextWrapping.NoWrap  # Keep columns aligned
        self.txtDeps.Foreground = B_TXT
        self.txtDeps.FontFamily = FontFamily("Segoe UI")
        self.txtDeps.FontSize = 13
        self.txtDeps.Padding = Thickness(8, 5, 8, 5)
        self.txtDeps.Text = "Select an element and click View Details"

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Content = self.txtDeps
        right.Children.Add(sv)

        rb = Border()
        rb.Background = B_WHT
        rb.BorderBrush = B_BRD
        rb.BorderThickness = Thickness(1)
        rb.CornerRadius = CR(4)
        rb.Margin = Thickness(5, 0, 0, 0)
        rb.Child = right
        Grid.SetColumn(rb, 1)
        g.Children.Add(rb)

        return g

    def _create_actions(self):
        b = Border()
        b.Background = B_WHT
        b.BorderBrush = B_BRD
        b.BorderThickness = Thickness(1)
        b.CornerRadius = CR(4)
        b.Padding = Thickness(12, 10, 12, 10)
        b.Margin = Thickness(0, 0, 0, 10)

        g = Grid()
        g.ColumnDefinitions.Add(ColumnDefinition())
        cd = ColumnDefinition()
        cd.Width = GridLength.Auto
        g.ColumnDefinitions.Add(cd)

        left = StackPanel()
        left.Orientation = Orientation.Horizontal

        btnView = self._btn("View Details", True)
        btnView.Margin = Thickness(0, 0, 8, 0)
        btnView.Click += self._on_view_details
        left.Children.Add(btnView)

        for txt, handler in [("+ Pick More", self._on_pick), ("Remove", self._on_remove),
                             ("Zoom To", self._on_zoom), ("Select in Revit", self._on_select),
                             ("Export Excel", self._on_export)]:
            btn = self._btn(txt, False)
            btn.Margin = Thickness(0, 0, 8, 0)
            btn.Click += handler
            left.Children.Add(btn)

        Grid.SetColumn(left, 0)
        g.Children.Add(left)

        right = StackPanel()
        right.Orientation = Orientation.Horizontal

        self.btnDelSafe = self._btn("Delete Safe Only", True)
        self.btnDelSafe.Margin = Thickness(0, 0, 8, 0)
        self.btnDelSafe.Click += self._on_del_safe
        right.Children.Add(self.btnDelSafe)

        self.btnDelAll = Button()
        self.btnDelAll.Content = "Delete All Selected"
        self.btnDelAll.Background = FB(255, 205, 210)
        self.btnDelAll.Foreground = B_CRIT
        self.btnDelAll.BorderThickness = Thickness(1)
        self.btnDelAll.BorderBrush = B_BRD
        self.btnDelAll.FontWeight = FontWeights.SemiBold
        self.btnDelAll.Padding = Thickness(15, 8, 15, 8)
        self.btnDelAll.Cursor = Cursors.Hand
        self.btnDelAll.Margin = Thickness(0, 0, 8, 0)
        self.btnDelAll.Click += self._on_del_all
        right.Children.Add(self.btnDelAll)

        btnClose = self._btn("Close", False)
        btnClose.Click += lambda s, e: self.Close()
        right.Children.Add(btnClose)

        Grid.SetColumn(right, 1)
        g.Children.Add(right)

        b.Child = g
        return b

    def _create_footer(self):
        b = Border()
        b.Background = B_PRI
        b.CornerRadius = CR(4)
        b.Padding = Thickness(12, 6, 12, 6)

        g = Grid()
        t1 = TextBlock()
        t1.Text = "Select element then click View Details"
        t1.FontSize = 10
        t1.Foreground = B_TXT
        g.Children.Add(t1)

        t2 = TextBlock()
        t2.Text = "Copyright 2026 Dang Quoc Truong (DQT)"
        t2.FontSize = 10
        t2.FontWeight = FontWeights.SemiBold
        t2.HorizontalAlignment = HorizontalAlignment.Right
        t2.Foreground = B_TXT
        g.Children.Add(t2)

        b.Child = g
        return b

    def _btn(self, text, primary):
        bt = Button()
        bt.Content = text
        bt.Background = B_PRI if primary else B_WHT
        bt.Foreground = B_TXT
        bt.BorderThickness = Thickness(1)
        bt.BorderBrush = B_BRD
        bt.Padding = Thickness(15, 8, 15, 8)
        bt.Cursor = Cursors.Hand
        if primary:
            bt.FontWeight = FontWeights.SemiBold
        return bt

    def _label(self, text):
        t = TextBlock()
        t.Text = text
        t.FontWeight = FontWeights.SemiBold
        t.Foreground = B_TXT
        t.VerticalAlignment = VerticalAlignment.Center
        return t

    def _panel(self, title):
        sp = StackPanel()
        hdr = Border()
        hdr.Background = B_HDR
        hdr.Padding = Thickness(10, 8, 10, 8)
        t = TextBlock()
        t.Text = title
        t.FontWeight = FontWeights.SemiBold
        t.Foreground = B_TXT
        hdr.Child = t
        sp.Children.Add(hdr)
        return sp

    def _init_filters(self):
        self._updating = True
        try:
            ci = ComboBoxItem()
            ci.Content = "All Categories"
            self.cmbCat.Items.Add(ci)
            self.cmbCat.SelectedIndex = 0

            for r in ["All Risk Levels", "Critical", "High", "Medium", "Safe"]:
                ri = ComboBoxItem()
                ri.Content = r
                self.cmbRisk.Items.Add(ri)
            self.cmbRisk.SelectedIndex = 0
        finally:
            self._updating = False

    def _load_selection(self):
        sel = self.uidoc.Selection.GetElementIds()
        if sel.Count == 0:
            return
        elems = []
        for eid in sel:
            e = self.doc.GetElement(eid)
            if e:
                elems.append(e)
        if elems:
            self._add_elements(elems)

    def _add_elements(self, elems):
        existing = set(it.eid for it in self.all_items)
        new_items = []

        for e in elems:
            eid = eid_int(e.Id)
            if eid in existing:
                continue
            item = ElemData(e)
            try:
                deps = analyze_element(e, self.doc)
                item.set_deps(deps)
            except Exception as ex:
                print("WARN: {} - {}".format(item.name, str(ex)[:40]))
                item.set_deps([])
            new_items.append(item)
            existing.add(eid)

        self.all_items.extend(new_items)
        self._rebuild_cat_filter()
        self._apply_filters()
        self._update_cards()

    def _rebuild_cat_filter(self):
        self._updating = True
        try:
            cats = sorted(set(it.cat for it in self.all_items if it.cat))
            cur = ""
            if self.cmbCat.SelectedItem:
                cur = self.cmbCat.SelectedItem.Content

            self.cmbCat.Items.Clear()
            ci = ComboBoxItem()
            ci.Content = "All Categories"
            self.cmbCat.Items.Add(ci)

            for cat in cats:
                ci = ComboBoxItem()
                ci.Content = cat
                self.cmbCat.Items.Add(ci)

            self.cmbCat.SelectedIndex = 0
            for i in range(self.cmbCat.Items.Count):
                if self.cmbCat.Items[i].Content == cur:
                    self.cmbCat.SelectedIndex = i
                    break
        finally:
            self._updating = False

    def _apply_filters(self):
        if self._updating:
            return

        search = (self.txtSearch.Text or "").lower().strip()
        cat_sel = "All Categories"
        if self.cmbCat.SelectedItem:
            cat_sel = self.cmbCat.SelectedItem.Content
        risk_sel = "All Risk Levels"
        if self.cmbRisk.SelectedItem:
            risk_sel = self.cmbRisk.SelectedItem.Content

        self.filtered_items = []
        for it in self.all_items:
            if search:
                hay = "{} {} {}".format(it.name, it.cat, it.eid).lower()
                if search not in hay:
                    continue
            if cat_sel != "All Categories" and it.cat != cat_sel:
                continue
            if risk_sel != "All Risk Levels":
                if risk_sel == "Safe" and not it.is_safe:
                    continue
                elif risk_sel == "Critical" and it.max_sev != "Critical":
                    continue
                elif risk_sel == "High" and it.max_sev != "High":
                    continue
                elif risk_sel == "Medium" and it.max_sev != "Medium":
                    continue
            self.filtered_items.append(it)

        self._refresh_list()

    def _refresh_list(self):
        """Simple, safe refresh - just plain strings"""
        self._updating = True
        try:
            self.lstElem.Items.Clear()
            for it in self.filtered_items:
                self.lstElem.Items.Add(it.list_text)
        finally:
            self._updating = False

    def _update_cards(self):
        total = len(self.all_items)
        deps = sum(it.dep_count for it in self.all_items)
        crit = sum(1 for it in self.all_items if it.max_sev == "Critical")
        high = sum(1 for it in self.all_items if it.max_sev in ("Critical", "High"))
        safe = sum(1 for it in self.all_items if it.is_safe)

        self._cards["Selected"].Text = str(total)
        self._cards["Dependencies"].Text = str(deps)
        self._cards["Critical"].Text = str(crit)
        self._cards["High Risk"].Text = str(high)
        self._cards["Safe"].Text = str(safe)

    def _show_deps(self, item):
        try:
            if item.dep_count == 0:
                text = "=" * 60 + "\n"
                text += "ELEMENT: {}\n".format(item.name)
                text += "Category: {}  |  ID: {}\n".format(item.cat, item.eid)
                text += "=" * 60 + "\n\n"
                text += "STATUS: SAFE - No dependencies found\n\n"
                text += "This element can be safely deleted."
                self.txtDeps.Text = text
                return

            # Header section
            text = "=" * 70 + "\n"
            text += "ELEMENT: {}\n".format(item.name)
            text += "Category: {}  |  ID: {}\n".format(item.cat, item.eid)
            text += "=" * 70 + "\n\n"
            
            text += "DEPENDENCIES: {} total  |  Risk: {}\n".format(
                item.dep_count, item.max_sev.upper())
            text += "-" * 70 + "\n\n"

            # Group dependencies by type
            groups = {}
            for d in item.deps:
                if d.dep_type not in groups:
                    groups[d.dep_type] = []
                groups[d.dep_type].append(d)

            # Sort by severity
            sev_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
            sorted_types = sorted(groups.keys(), 
                key=lambda t: sev_order.get(groups[t][0].severity, 99))

            for dep_type in sorted_types:
                dep_list = groups[dep_type]
                if not dep_list:
                    continue
                    
                sev = dep_list[0].severity
                
                # Group header
                text += "[{}] {} ({} items)\n".format(
                    sev.upper(), dep_type, len(dep_list))
                
                # Column header for this group
                text += "  {}  {}  {}\n".format(
                    "Name".ljust(25), "View".ljust(25), "ID".ljust(10))
                text += "  " + "-" * 65 + "\n"
                
                # Dependency items
                for d in dep_list[:15]:
                    d_name = d.name[:25].ljust(25) if d.name else "".ljust(25)
                    d_view = d.view_name[:25].ljust(25) if d.view_name else "-".ljust(25)
                    d_id = str(d.eid).ljust(10)
                    text += "  {}  {}  {}\n".format(d_name, d_view, d_id)
                
                if len(dep_list) > 15:
                    text += "  ... +{} more items\n".format(len(dep_list) - 15)
                
                text += "\n"

            self.txtDeps.Text = text
        except Exception as ex:
            self.txtDeps.Text = "Error: {}".format(str(ex))

    def _on_filter(self, sender, args):
        if self._updating:
            return
        self._apply_filters()

    def _on_view_details(self, sender, args):
        try:
            idx = self.lstElem.SelectedIndex
            if idx < 0 or idx >= len(self.filtered_items):
                self.txtDeps.Text = "Select an element first."
                return
            self._show_deps(self.filtered_items[idx])
        except Exception as ex:
            self.txtDeps.Text = "Error: {}".format(str(ex))

    def _on_pick(self, sender, args):
        try:
            self.Hide()
            try:
                refs = self.uidoc.Selection.PickObjects(ObjectType.Element, "Select elements (ESC to finish)")
                elems = []
                for ref in refs:
                    e = self.doc.GetElement(ref.ElementId)
                    if e:
                        elems.append(e)
                if elems:
                    self._add_elements(elems)
            except:
                pass
            self.Show()
        except:
            self.Show()

    def _on_remove(self, sender, args):
        idx = self.lstElem.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_items):
            self.txtDeps.Text = "Select an element first."
            return
        item = self.filtered_items[idx]
        self.all_items = [i for i in self.all_items if i.eid != item.eid]
        self._rebuild_cat_filter()
        self._apply_filters()
        self._update_cards()
        self.txtDeps.Text = "Element removed from list."

    def _on_zoom(self, sender, args):
        idx = self.lstElem.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_items):
            self.txtDeps.Text = "Select an element first."
            return
        try:
            item = self.filtered_items[idx]
            self.Hide()
            self.uidoc.ShowElements(item.element.Id)
            self.Show()
        except Exception as ex:
            self.Show()
            self.txtDeps.Text = "Zoom error: {}".format(str(ex))

    def _on_select(self, sender, args):
        idx = self.lstElem.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_items):
            self.txtDeps.Text = "Select an element first."
            return
        try:
            item = self.filtered_items[idx]
            from System.Collections.Generic import List
            ids = List[ElementId]()
            ids.Add(make_eid(item.eid))
            self.uidoc.Selection.SetElementIds(ids)
            self.txtDeps.Text = "Selected in Revit: {}".format(item.name)
        except Exception as ex:
            self.txtDeps.Text = "Select error: {}".format(str(ex))

    def _on_export(self, sender, args):
        """Export elements and dependencies to Excel"""
        if not self.all_items:
            self.txtDeps.Text = "No elements to export."
            return
        
        try:
            import os
            import datetime
            
            # Get save path
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = "SmartDelete_Report_{}.xlsx".format(timestamp)
            
            # Use forms to get save path
            save_path = forms.save_file(file_ext='xlsx', default_name=default_name)
            if not save_path:
                return
            
            self.txtDeps.Text = "Exporting to Excel..."
            
            # Create Excel using COM
            import clr
            clr.AddReference("Microsoft.Office.Interop.Excel")
            from Microsoft.Office.Interop import Excel
            
            excel_app = Excel.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            try:
                wb = excel_app.Workbooks.Add()
                
                # Sheet 1: Elements Summary
                ws1 = wb.Worksheets[1]
                ws1.Name = "Elements"
                
                # Header row
                headers = ["Name", "Category", "ID", "Risk Level", "Dep Count", "Max Severity"]
                for col, h in enumerate(headers, 1):
                    ws1.Cells[1, col].Value2 = h
                    ws1.Cells[1, col].Font.Bold = True
                    ws1.Cells[1, col].Interior.Color = 0x0F172A  # DQT Primary color (BGR)
                
                # Data rows
                for row, item in enumerate(self.all_items, 2):
                    ws1.Cells[row, 1].Value2 = item.name
                    ws1.Cells[row, 2].Value2 = item.cat
                    ws1.Cells[row, 3].Value2 = item.eid
                    ws1.Cells[row, 4].Value2 = item.risk_label
                    ws1.Cells[row, 5].Value2 = item.dep_count
                    ws1.Cells[row, 6].Value2 = item.max_sev
                    
                    # Color code risk
                    if item.max_sev == "Critical":
                        ws1.Cells[row, 4].Font.Color = 0x2F2FD3  # Red (BGR)
                    elif item.max_sev == "High":
                        ws1.Cells[row, 4].Font.Color = 0x007CF5  # Orange (BGR)
                    elif item.is_safe:
                        ws1.Cells[row, 4].Font.Color = 0x3C8E38  # Green (BGR)
                
                # Auto-fit columns
                ws1.Columns.AutoFit()
                
                # Sheet 2: All Dependencies
                ws2 = wb.Worksheets.Add()
                ws2.Name = "Dependencies"
                
                # Header
                dep_headers = ["Element Name", "Element ID", "Dep Type", "Dep Name", 
                              "Dep ID", "Severity", "View"]
                for col, h in enumerate(dep_headers, 1):
                    ws2.Cells[1, col].Value2 = h
                    ws2.Cells[1, col].Font.Bold = True
                    ws2.Cells[1, col].Interior.Color = 0x0F172A
                
                # All dependencies
                dep_row = 2
                for item in self.all_items:
                    for dep in item.deps:
                        ws2.Cells[dep_row, 1].Value2 = item.name
                        ws2.Cells[dep_row, 2].Value2 = item.eid
                        ws2.Cells[dep_row, 3].Value2 = dep.dep_type
                        ws2.Cells[dep_row, 4].Value2 = dep.name
                        ws2.Cells[dep_row, 5].Value2 = dep.eid
                        ws2.Cells[dep_row, 6].Value2 = dep.severity
                        ws2.Cells[dep_row, 7].Value2 = dep.view_name or "-"
                        
                        # Color code severity
                        if dep.severity == "Critical":
                            ws2.Cells[dep_row, 6].Font.Color = 0x2F2FD3
                        elif dep.severity == "High":
                            ws2.Cells[dep_row, 6].Font.Color = 0x007CF5
                        
                        dep_row += 1
                
                ws2.Columns.AutoFit()
                
                # Save and close
                wb.SaveAs(save_path)
                wb.Close(False)
                
                self.txtDeps.Text = "Exported successfully!\n\nFile: {}".format(save_path)
                
            finally:
                excel_app.Quit()
                
        except Exception as ex:
            self.txtDeps.Text = "Export error:\n{}".format(str(ex))

    def _on_del_safe(self, sender, args):
        safe = [i for i in self.all_items if i.is_safe]
        if not safe:
            self.txtDeps.Text = "No safe elements to delete."
            return
        self._do_delete(safe, "Safe Delete")

    def _on_del_all(self, sender, args):
        if not self.all_items:
            return
        self._do_delete(self.all_items[:], "Delete All")

    def _do_delete(self, items, label):
        count = len(items)
        crit = sum(1 for i in items if i.max_sev == "Critical")

        msg = "{}: {} element(s)".format(label, count)
        if crit > 0:
            msg += "\n\nWARNING: {} CRITICAL risk!".format(crit)
        msg += "\n\nUndo: Ctrl+Z"

        self.Hide()
        if not forms.alert(msg, title="Confirm - DQT", yes=True, no=True):
            self.Show()
            return

        ok = 0
        t = Transaction(self.doc, "Smart Delete - DQT")
        t.Start()
        try:
            for item in items:
                try:
                    self.doc.Delete(make_eid(item.eid))
                    ok += 1
                except:
                    pass
            if ok > 0:
                t.Commit()
            else:
                t.RollBack()
        except:
            try:
                t.RollBack()
            except:
                pass

        del_ids = set(i.eid for i in items if ok > 0)
        self.all_items = [i for i in self.all_items if i.eid not in del_ids]
        self._rebuild_cat_filter()
        self._apply_filters()
        self._update_cards()

        self.txtDeps.Text = "Deleted {} of {} element(s).\nUndo: Ctrl+Z".format(ok, count)
        self.Show()


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    try:
        win = SmartDeleteWindow()
        win.ShowDialog()
    except Exception as ex:
        forms.alert("Error:\n{}\n\n{}".format(str(ex), traceback.format_exc()),
                    title="Smart Delete - Error")