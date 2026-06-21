# -*- coding: utf-8 -*-
"""
Dim Text

Edit dimension text overrides on selected dimensions,
with optional segment-length filter rules (AND / OR).

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Dim Text"
__version__ = "2.1.0"

# ── IMPORTS ───────────────────────────────────────────────────────────────────
import os
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System')

import System
from System.Windows import WindowState, Thickness, Visibility, VerticalAlignment
from System.Windows.Controls import (
    StackPanel, ComboBox, ComboBoxItem, TextBox, Button, TextBlock
)
from System.Windows.Controls import Orientation as WPFOrientation
from System.Windows.Media import SolidColorBrush, Color, FontFamily as WPFFontFamily
from Autodesk.Revit.DB import Dimension, FilteredElementCollector, Transaction
from pyrevit import revit, forms, script

# ── VARIABLES ─────────────────────────────────────────────────────────────────
uidoc  = revit.uidoc
doc    = revit.doc
logger = script.get_logger()

XAML_PATH = os.path.join(os.path.dirname(__file__), "Tools", "DimText.xaml")

_OPERATORS = [
    "equals",
    "does not equal",
    "is greater than",
    "is greater than or equal to",
    "is less than",
    "is less than or equal to",
    "between",
    "has a value",
    "has no value",
]

# operators that need no value input at all
_NO_VALUE_OPS  = {"has a value", "has no value"}
# operators that need two value inputs
_TWO_VALUE_OPS = {"between"}


# ── HELPERS ───────────────────────────────────────────────────────────────────
def _feet_to_mm(value):
    if value is None:
        return None
    return float(value) * 304.8


def _set_dim_text(dim, prefix, suffix, above, below, override, filter_fn=None):
    """Apply text overrides, optionally filtered by segment length (mm)."""
    if dim.HasOneSegment():
        length_mm = _feet_to_mm(dim.Value)
        if filter_fn is None or (length_mm is not None and filter_fn(length_mm)):
            dim.Prefix        = prefix
            dim.Suffix        = suffix
            dim.Above         = above
            dim.Below         = below
            dim.ValueOverride = override
    else:
        for seg in dim.Segments:
            length_mm = _feet_to_mm(seg.Value)
            if filter_fn is None or (length_mm is not None and filter_fn(length_mm)):
                seg.Prefix        = prefix
                seg.Suffix        = suffix
                seg.Above         = above
                seg.Below         = below
                seg.ValueOverride = override


def _turn_off_leader(dim):
    for para in dim.GetOrderedParameters():
        if para.Definition.Name == "Leader":
            para.Set(0)
            break


def _get_dims_in_view():
    return list(
        FilteredElementCollector(doc, uidoc.ActiveView.Id)
        .OfClass(Dimension)
        .ToElements()
    )


def _get_selected_dims():
    return [
        doc.GetElement(eid)
        for eid in uidoc.Selection.GetElementIds()
        if isinstance(doc.GetElement(eid), Dimension)
    ]


# ── WINDOW ────────────────────────────────────────────────────────────────────
class DimTextWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_PATH)
        self._rules = []  # list of dicts: {panel, combo, txt1, txt2, lbl_and, lbl_mm2}
        # Pre-cache all named controls immediately so they remain accessible after
        # the content grid is detached from this Window and embedded into a parent.
        for _n in ("txt_prefix", "txt_suffix", "txt_above", "txt_below", "txt_override",
                   "wrap_presets", "chk_leader", "rb_selection", "rb_view",
                   "chk_filter_enable", "sp_filter_config", "combo_combine",
                   "sp_rules", "btn_clear_fields", "btn_cancel", "btn_apply", "lbl_status"):
            setattr(self, _n, self.FindName(_n))

    # ── window chrome ──────────────────────────────────────────────────────────
    def minimize_button_clicked(self, sender, args):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def close_button_clicked(self, sender, args):
        self.Close()

    # ── presets ────────────────────────────────────────────────────────────────
    def preset_below_clicked(self, sender, args):
        self.txt_below.Text = sender.Tag

    # ── clear fields ───────────────────────────────────────────────────────────
    def clear_fields_clicked(self, sender, args):
        self.txt_prefix.Text   = ""
        self.txt_suffix.Text   = ""
        self.txt_above.Text    = ""
        self.txt_below.Text    = ""
        self.txt_override.Text = ""
        self.lbl_status.Text   = "Fields cleared."

    # ── filter section ─────────────────────────────────────────────────────────
    def filter_toggle(self, sender, args):
        if self.chk_filter_enable.IsChecked:
            self.sp_filter_config.Visibility = Visibility.Visible
        else:
            self.sp_filter_config.Visibility = Visibility.Collapsed

    def add_rule_clicked(self, sender, args):
        rd = self._create_rule_row()
        self._rules.append(rd)
        self.sp_rules.Children.Add(rd["panel"])

    def _create_rule_row(self):
        """Build one rule row and return a dict of its controls."""
        rd = {}

        row = StackPanel()
        row.Orientation = WPFOrientation.Horizontal
        row.Margin = Thickness(0, 0, 0, 6)
        rd["panel"] = row

        # ── operator combo ──
        combo = ComboBox()
        combo.Width = 185
        combo.Height = 28
        combo.FontFamily = WPFFontFamily("Inter")
        combo.FontSize = 12
        combo.Margin = Thickness(0, 0, 6, 0)
        for op in _OPERATORS:
            item = ComboBoxItem()
            item.Content = op
            combo.Items.Add(item)
        combo.SelectedIndex = 0
        combo.SelectionChanged += self._make_op_handler(rd)
        rd["combo"] = combo
        row.Children.Add(combo)

        # ── value 1 ──
        txt1 = TextBox()
        txt1.Width = 72
        txt1.Height = 28
        txt1.FontSize = 12
        txt1.Padding = Thickness(6, 4, 6, 4)
        txt1.Margin = Thickness(0, 0, 4, 0)
        txt1.BorderBrush = SolidColorBrush(Color.FromRgb(0x54, 0x6E, 0x7A))
        txt1.BorderThickness = Thickness(1)
        rd["txt1"] = txt1
        row.Children.Add(txt1)

        # ── mm label ──
        lbl_mm = TextBlock()
        lbl_mm.Text = "mm"
        lbl_mm.FontSize = 11
        lbl_mm.Foreground = SolidColorBrush(Color.FromRgb(0x7F, 0x8C, 0x8D))
        lbl_mm.Margin = Thickness(0, 0, 8, 0)
        lbl_mm.VerticalAlignment = VerticalAlignment.Center
        rd["lbl_mm"] = lbl_mm
        row.Children.Add(lbl_mm)

        # ── "and" label (between only) ──
        lbl_and = TextBlock()
        lbl_and.Text = "and"
        lbl_and.FontSize = 11
        lbl_and.Foreground = SolidColorBrush(Color.FromRgb(0x2C, 0x3E, 0x50))
        lbl_and.Margin = Thickness(0, 0, 6, 0)
        lbl_and.VerticalAlignment = VerticalAlignment.Center
        lbl_and.Visibility = Visibility.Collapsed
        rd["lbl_and"] = lbl_and
        row.Children.Add(lbl_and)

        # ── value 2 (between only) ──
        txt2 = TextBox()
        txt2.Width = 72
        txt2.Height = 28
        txt2.FontSize = 12
        txt2.Padding = Thickness(6, 4, 6, 4)
        txt2.Margin = Thickness(0, 0, 4, 0)
        txt2.BorderBrush = SolidColorBrush(Color.FromRgb(0x54, 0x6E, 0x7A))
        txt2.BorderThickness = Thickness(1)
        txt2.Visibility = Visibility.Collapsed
        rd["txt2"] = txt2
        row.Children.Add(txt2)

        # ── mm2 label (between only) ──
        lbl_mm2 = TextBlock()
        lbl_mm2.Text = "mm"
        lbl_mm2.FontSize = 11
        lbl_mm2.Foreground = SolidColorBrush(Color.FromRgb(0x7F, 0x8C, 0x8D))
        lbl_mm2.Margin = Thickness(0, 0, 8, 0)
        lbl_mm2.VerticalAlignment = VerticalAlignment.Center
        lbl_mm2.Visibility = Visibility.Collapsed
        rd["lbl_mm2"] = lbl_mm2
        row.Children.Add(lbl_mm2)

        # ── remove button ──
        btn = Button()
        btn.Content = "-"
        btn.Width = 26
        btn.Height = 26
        btn.FontSize = 14
        btn.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
        btn.BorderThickness = Thickness(1)
        btn.BorderBrush = SolidColorBrush(Color.FromRgb(0xD3, 0x2F, 0x2F))
        btn.Foreground = SolidColorBrush(Color.FromRgb(0xD3, 0x2F, 0x2F))
        btn.Click += self._make_remove_handler(rd)
        row.Children.Add(btn)

        return rd

    def _make_op_handler(self, rd):
        def handler(sender, args):
            op = sender.SelectedItem.Content if sender.SelectedItem else ""
            no_val  = op in _NO_VALUE_OPS
            two_val = op in _TWO_VALUE_OPS
            # first value column
            v1_vis = Visibility.Collapsed if no_val else Visibility.Visible
            rd["txt1"].Visibility   = v1_vis
            rd["lbl_mm"].Visibility = v1_vis
            # second value column (between only)
            v2_vis = Visibility.Visible if two_val else Visibility.Collapsed
            rd["lbl_and"].Visibility = v2_vis
            rd["txt2"].Visibility    = v2_vis
            rd["lbl_mm2"].Visibility = v2_vis
        return handler

    def _make_remove_handler(self, rd):
        def handler(sender, args):
            self.sp_rules.Children.Remove(rd["panel"])
            if rd in self._rules:
                self._rules.remove(rd)
        return handler

    # ── build filter function from current rules ────────────────────────────────
    def _build_filter_fn(self):
        if not self.chk_filter_enable.IsChecked or not self._rules:
            return None

        parsed = []
        for rd in self._rules:
            op = rd["combo"].SelectedItem.Content if rd["combo"].SelectedItem else None
            if op is None:
                continue
            v1, v2 = 0.0, 0.0
            if op not in _NO_VALUE_OPS:
                try:
                    v1 = float(rd["txt1"].Text.strip() or "0")
                except ValueError:
                    v1 = 0.0
            if op in _TWO_VALUE_OPS:
                try:
                    v2 = float(rd["txt2"].Text.strip() or "0")
                except ValueError:
                    v2 = 0.0
            parsed.append((op, v1, v2))

        if not parsed:
            return None

        use_and = (self.combo_combine.SelectedIndex == 0)

        def filter_fn(length_mm):
            if length_mm is None:
                return False
            results = []
            for op, v1, v2 in parsed:
                if   op == "equals":                      results.append(abs(length_mm - v1) < 0.5)
                elif op == "does not equal":              results.append(abs(length_mm - v1) >= 0.5)
                elif op == "is greater than":             results.append(length_mm >  v1)
                elif op == "is greater than or equal to": results.append(length_mm >= v1)
                elif op == "is less than":                results.append(length_mm <  v1)
                elif op == "is less than or equal to":    results.append(length_mm <= v1)
                elif op == "between":                     results.append(min(v1, v2) <= length_mm <= max(v1, v2))
                elif op == "has a value":                 results.append(True)
                elif op == "has no value":                results.append(False)
            if not results:
                return True
            return all(results) if use_and else any(results)

        return filter_fn

    # ── apply ──────────────────────────────────────────────────────────────────
    def apply_clicked(self, sender, args):
        prefix   = self.txt_prefix.Text.strip()
        suffix   = self.txt_suffix.Text.strip()
        above    = self.txt_above.Text.strip()
        below    = self.txt_below.Text.strip()
        override = self.txt_override.Text.strip()
        leader_off = self.chk_leader.IsChecked
        filter_fn  = self._build_filter_fn()

        if self.rb_view.IsChecked:
            dims = _get_dims_in_view()
            scope_label = "view"
        else:
            dims = _get_selected_dims()
            scope_label = "selection"

        if not dims:
            self.lbl_status.Text = "No dimensions found in {}.".format(scope_label)
            return

        with Transaction(doc, "Dim Text Override") as t:
            t.Start()
            for dim in dims:
                _set_dim_text(dim, prefix, suffix, above, below, override, filter_fn)
                if leader_off:
                    _turn_off_leader(dim)
            t.Commit()

        filter_note = " (length filter active)" if filter_fn else ""
        self.lbl_status.Text = "Applied to {} dim(s) in {}{}.".format(
            len(dims), scope_label, filter_note
        )
        logger.info("DimText applied: {} dims, scope={}, filter={}".format(
            len(dims), scope_label, filter_fn is not None
        ))


def show_dialog():
    DimTextWindow().ShowDialog()

if __name__ == '__main__':
    show_dialog()
