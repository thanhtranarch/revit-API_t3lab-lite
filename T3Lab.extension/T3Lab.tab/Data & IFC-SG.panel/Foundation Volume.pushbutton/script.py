# -*- coding: utf-8 -*-
"""
Foundation Volume Writer v1.0 - DQT
Writes Revit built-in volume value into a selected instance parameter
on all Structural Foundation elements in the active document.

Workflow:
  1. Tool collects all writable instance parameters from foundations
  2. User searches and selects target parameter
  3. Tool writes HOST_VOLUME_COMPUTED into selected parameter

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Foundation\nVolume"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Write Revit computed volume into a selected shared parameter on Structural Foundation elements."

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

import System
import codecs
import os
from System import Array
from System.IO import MemoryStream
from System.Text import Encoding
from System.Windows.Markup import XamlReader
from System.Windows import Window, Thickness, WindowState
from System.Windows.Controls import ListBoxItem

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, StorageType, UnitUtils
)

try:
    from Autodesk.Revit.DB import UnitTypeId
    HAS_UNIT_TYPE_ID = True
except Exception:
    HAS_UNIT_TYPE_ID = False

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# ── Revit 2025+ compatibility ──────────────────────────────────────────────
def _eid_int(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue

# ── Unit conversion: internal ft³ → m³ ────────────────────────────────────
def ft3_to_m3(value):
    """Convert Revit internal cubic feet to cubic metres."""
    try:
        if HAS_UNIT_TYPE_ID:
            return UnitUtils.ConvertFromInternalUnits(value, DB.UnitTypeId.CubicMeters)
        else:
            from Autodesk.Revit.DB import DisplayUnitType
            return UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_CUBIC_METERS)
    except Exception:
        return value * 0.0283168466  # fallback constant

# ── Collect foundations ────────────────────────────────────────────────────
def get_foundations():
    return list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_StructuralFoundation)
        .WhereElementIsNotElementType()
        .ToElements()
    )

# ── Get volume from element ────────────────────────────────────────────────
def get_volume_m3(element):
    """Try HOST_VOLUME_COMPUTED first, fallback to geometry solid sum."""
    try:
        p = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
        if p and p.HasValue and p.AsDouble() > 0:
            return ft3_to_m3(p.AsDouble())
    except Exception:
        pass
    # Geometry fallback
    try:
        opts = DB.Options()
        opts.ComputeReferences = False
        geom = element.get_Geometry(opts)
        total = 0.0
        for obj in geom:
            if isinstance(obj, DB.Solid) and obj.Volume > 0:
                total += obj.Volume
            elif isinstance(obj, DB.GeometryInstance):
                for sub in obj.GetInstanceGeometry():
                    if isinstance(sub, DB.Solid) and sub.Volume > 0:
                        total += sub.Volume
        if total > 0:
            return ft3_to_m3(total)
    except Exception:
        pass
    return None

# ── Collect writable instance parameters ──────────────────────────────────
def get_writable_params(foundations):
    """Return sorted list of writable instance parameter names from foundations."""
    param_names = set()
    for f in foundations[:20]:  # sample first 20 for speed
        for p in f.Parameters:
            if p.IsReadOnly:
                continue
            if p.StorageType not in (StorageType.Double, StorageType.String):
                continue
            name = p.Definition.Name
            if name:
                param_names.add(name)
    return sorted(param_names)

# ═══════════════════════════════════════════════════════════════════════════
# XAML UI
# ═══════════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════════
# Main dialog class
# ═══════════════════════════════════════════════════════════════════════════
class FoundationVolumeDialog(object):

    def __init__(self):
        self.foundations    = get_foundations()
        self.all_params     = get_writable_params(self.foundations)
        self.selected_param = None
        self._current_tab   = 0

        # Find T3Lab.extension parent folder dynamically
        current_dir = os.path.dirname(__file__)
        while current_dir and not current_dir.endswith('T3Lab.extension'):
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                break
            current_dir = parent
        xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "FoundationVolume.xaml")

        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()
        stream = MemoryStream(Encoding.UTF8.GetBytes(xaml_content))
        self.window = XamlReader.Load(stream)

        # Controls
        self._count_lbl       = self.window.FindName("FoundationCount")
        self._search_box      = self.window.FindName("SearchBox")
        self._param_list      = self.window.FindName("ParamList")
        self._sel_label       = self.window.FindName("SelectedLabel")
        self._status_bdr      = self.window.FindName("StatusBorder")
        self._status_title    = self.window.FindName("StatusTitle")
        self._status_txt      = self.window.FindName("StatusText")
        self._run_btn         = self.window.FindName("RunButton")
        self._close_btn       = self.window.FindName("CloseButton")
        self._btn_minimize    = self.window.FindName("btn_minimize")
        self._btn_maximize    = self.window.FindName("btn_maximize")
        self._btn_close       = self.window.FindName("btn_close")
        self._main_tabs       = self.window.FindName("main_tabs")
        self._back_btn        = self.window.FindName("back_button")
        self._next_txt        = self.window.FindName("next_button_text")
        self._next_icon       = self.window.FindName("next_button_icon")
        self._write_hint      = self.window.FindName("WriteHintBorder")
        self._nav_select      = self.window.FindName("nav_toggle_select")
        self._nav_write       = self.window.FindName("nav_toggle_write")

        # Init
        self._count_lbl.Text = "{0} foundation(s)".format(len(self.foundations))
        self._populate_list(self.all_params)
        self._go_to_tab(0)

        # Events
        self._search_box.TextChanged      += self._on_search
        self._param_list.MouseDoubleClick += self._on_list_double_click
        self._param_list.SelectionChanged += self._on_selection_changed
        self._run_btn.Click               += self._on_next
        self._back_btn.Click              += self._on_back
        if self._nav_select:
            self._nav_select.Click        += self._on_nav_select
        if self._nav_write:
            self._nav_write.Click         += self._on_nav_write
        # Chrome button events (XamlReader doesn't auto-wire Click="...")
        if self._btn_minimize:
            self._btn_minimize.Click += self._on_minimize
        if self._btn_maximize:
            self._btn_maximize.Click += self._on_maximize
        if self._btn_close:
            self._btn_close.Click += self._on_close

    # ── List helpers ──────────────────────────────────────────────────────
    def _populate_list(self, names):
        self._param_list.Items.Clear()
        for n in names:
            item = ListBoxItem()
            item.Content = n
            self._param_list.Items.Add(item)
        if self._param_list.Items.Count > 0:
            self._param_list.SelectedIndex = 0

    def _on_search(self, sender, e):
        query = self._search_box.Text.strip().lower()
        filtered = [n for n in self.all_params if query in n.lower()] if query else self.all_params
        self._populate_list(filtered)

    def _on_selection_changed(self, sender, e):
        item = self._param_list.SelectedItem
        if item:
            self.selected_param = item.Content
            self._sel_label.Text = self.selected_param
        else:
            self.selected_param = None
            self._sel_label.Text = "(none)"

    def _on_list_double_click(self, sender, e):
        if self.selected_param:
            self._go_to_tab(1)

    # ── Status helper ─────────────────────────────────────────────────────
    def _show_status(self, title, detail, success=True):
        conv = System.Windows.Media.BrushConverter()
        if success:
            self._status_bdr.Background  = conv.ConvertFromString("#E8F5E9")
            self._status_bdr.SetValue(
                System.Windows.Controls.Border.BorderBrushProperty,
                conv.ConvertFromString("#A5D6A7"))
            self._status_bdr.SetValue(
                System.Windows.Controls.Border.BorderThicknessProperty,
                System.Windows.Thickness(1))
            self._status_title.Foreground = conv.ConvertFromString("#1B5E20")
        else:
            self._status_bdr.Background  = conv.ConvertFromString("#FFF3E0")
            self._status_bdr.SetValue(
                System.Windows.Controls.Border.BorderBrushProperty,
                conv.ConvertFromString("#FFCC80"))
            self._status_bdr.SetValue(
                System.Windows.Controls.Border.BorderThicknessProperty,
                System.Windows.Thickness(1))
            self._status_title.Foreground = conv.ConvertFromString("#E65100")

        self._status_title.Text = title
        self._status_txt.Text   = detail
        self._status_bdr.Visibility = System.Windows.Visibility.Visible
        if self._write_hint:
            self._write_hint.Visibility = System.Windows.Visibility.Collapsed

    # ── Run logic ─────────────────────────────────────────────────────────
    def _on_run(self, sender, e):
        if not self.selected_param:
            self._show_status(
                "⚠  No parameter selected",
                "Please select a target parameter from the list above.",
                success=False)
            return
        if not self.foundations:
            self._show_status(
                "⚠  No foundations found",
                "No Structural Foundation elements were found in the active document.",
                success=False)
            return

        target_param_name = self.selected_param
        updated = 0
        skipped_no_vol = 0
        skipped_no_param = 0
        skipped_readonly = 0
        errors = 0

        with Transaction(doc, "DQT - Write Foundation Volume") as t:
            t.Start()
            for f in self.foundations:
                try:
                    vol_m3 = get_volume_m3(f)
                    if vol_m3 is None or vol_m3 <= 0:
                        skipped_no_vol += 1
                        continue

                    p = f.LookupParameter(target_param_name)
                    if p is None:
                        skipped_no_param += 1
                        continue
                    if p.IsReadOnly:
                        skipped_readonly += 1
                        continue

                    # Write value according to StorageType
                    if p.StorageType == StorageType.Double:
                        try:
                            is_volume_spec = False
                            try:
                                spec_id = p.Definition.GetSpecTypeId()
                                if HAS_UNIT_TYPE_ID:
                                    is_volume_spec = (spec_id == DB.SpecTypeId.Volume)
                            except Exception:
                                pass

                            if is_volume_spec:
                                try:
                                    if HAS_UNIT_TYPE_ID:
                                        internal_val = UnitUtils.ConvertToInternalUnits(vol_m3, DB.UnitTypeId.CubicMeters)
                                    else:
                                        from Autodesk.Revit.DB import DisplayUnitType
                                        internal_val = UnitUtils.ConvertToInternalUnits(vol_m3, DisplayUnitType.DUT_CUBIC_METERS)
                                    p.Set(internal_val)
                                except Exception:
                                    p.Set(vol_m3)
                            else:
                                p.Set(vol_m3)
                        except Exception:
                            p.Set(vol_m3)

                    elif p.StorageType == StorageType.String:
                        p.Set(str(round(vol_m3, 4)))

                    else:
                        skipped_no_param += 1
                        continue

                    updated += 1

                except Exception as ex:
                    errors += 1

            t.Commit()

        # ── Build result ──────────────────────────────────────────────────
        success = updated > 0
        if success:
            title = "✅  Completed — {0} of {1} foundation(s) updated".format(
                updated, len(self.foundations))
        else:
            title = "⚠  No foundations were updated"

        detail_lines = []
        if skipped_no_vol   > 0: detail_lines.append("• {0} skipped — volume = 0 or unavailable".format(skipped_no_vol))
        if skipped_no_param > 0: detail_lines.append("• {0} skipped — parameter \"{1}\" not found on element".format(skipped_no_param, target_param_name))
        if skipped_readonly > 0: detail_lines.append("• {0} skipped — parameter is read-only".format(skipped_readonly))
        if errors           > 0: detail_lines.append("• {0} error(s) encountered during write".format(errors))
        if not detail_lines:
            detail_lines.append("All foundations processed successfully.")

        self._show_status(title, "\n".join(detail_lines), success=success)

    # ── Tab navigation ────────────────────────────────────────────────────
    def _go_to_tab(self, index):
        self._current_tab = index
        self._main_tabs.SelectedIndex = index
        if self._nav_select:
            self._nav_select.IsChecked = (index == 0)
        if self._nav_write:
            self._nav_write.IsChecked  = (index == 1)
        if index == 0:
            self._back_btn.Visibility = System.Windows.Visibility.Collapsed
            self._next_txt.Text       = "Next"
            self._next_icon.Text      = " →"
        else:
            self._back_btn.Visibility = System.Windows.Visibility.Visible
            self._next_txt.Text       = "Write Volume"
            self._next_icon.Text      = ""

    def _on_next(self, sender, e):
        if self._current_tab == 0:
            if not self.selected_param:
                return
            self._go_to_tab(1)
        else:
            self._on_run(sender, e)

    def _on_back(self, sender, e):
        self._go_to_tab(0)

    def _on_nav_select(self, sender, e):
        self._go_to_tab(0)

    def _on_nav_write(self, sender, e):
        if self.selected_param:
            self._go_to_tab(1)
        else:
            if self._nav_write:
                self._nav_write.IsChecked = False

    def _on_minimize(self, sender, e):
        self.window.WindowState = WindowState.Minimized

    def _on_maximize(self, sender, e):
        if self.window.WindowState == WindowState.Maximized:
            self.window.WindowState = WindowState.Normal
        else:
            self.window.WindowState = WindowState.Maximized

    def _on_close(self, sender, e):
        self.window.Close()

    def show(self):
        self.window.ShowDialog()


# ═══════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    dlg = FoundationVolumeDialog()
    dlg.show()