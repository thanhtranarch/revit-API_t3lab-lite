# -*- coding: utf-8 -*-
"""Ribbon Name Manager Dialog class."""

import os
from pyrevit import forms
import clr
clr.AddReference("System.Data")
from System import String as System_String
from System import DBNull
from System.Data import DataTable
from System.Windows import WindowState

# unicode() shim for IronPython 2 / CPython 3
try:
    _unicode = unicode
except NameError:
    _unicode = str

# Absolute path to XAML
_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'RibbonNames.xaml')

class RibbonNameWindow(forms.WPFWindow):
    def __init__(self, live_tabs, short_map, originals, default_map, on_save_callback, on_state_callback, on_originals_callback):
        # WPFWindow.__init__ loads XAML and registers named controls
        forms.WPFWindow.__init__(self, _XAML)
        
        self.live_tabs = live_tabs
        self.short_map = short_map
        self.originals = originals
        self.default_map = default_map
        
        self.on_save_callback = on_save_callback
        self.on_state_callback = on_state_callback
        self.on_originals_callback = on_originals_callback
        self.message = None

        # Build DataTable and bind to DataGrid (which has x:Name="Grid")
        self.table = DataTable("tabs")
        self.table.Columns.Add("CurrentName", System_String)
        self.table.Columns.Add("ShortName", System_String)
        self._build_rows()
        self.Grid.ItemsSource = self.table.DefaultView

        # Wire buttons
        self.BtnShort.Click += self._on_apply_short
        self.BtnFull.Click += self._on_restore_full
        self.BtnSave.Click += self._on_save
        self.BtnReset.Click += self._on_reset
        self.BtnClose.Click += self._on_close

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

    def _full_name_of(self, tab):
        title = tab.Title
        for full, short in self.short_map.items():
            if short == title:
                return full
        return title

    def _build_rows(self):
        seen = set()
        for tab in self.live_tabs:
            full = self._full_name_of(tab)
            if full in seen:
                continue
            seen.add(full)
            short = self.short_map.get(full, full)
            self.table.Rows.Add(full, short)

    def _collect_map_from_grid(self):
        m = {}
        for row in self.table.Rows:
            full = self._cell(row, "CurrentName")
            short = self._cell(row, "ShortName").strip()
            if not short:
                short = full
            m[full] = short
        return m

    @staticmethod
    def _cell(row, col):
        try:
            val = row[col]
        except Exception:
            return u""
        if val is None or val == DBNull.Value:
            return u""
        try:
            return _unicode(val)
        except Exception:
            return str(val)

    def _commit_grid(self):
        try:
            self.Grid.CommitEdit()
            self.Grid.CommitEdit()
        except Exception:
            pass

    def _on_apply_short(self, sender, args):
        self._commit_grid()
        m = self._collect_map_from_grid()
        self.short_map.update(m)
        applied = 0
        for tab in self.live_tabs:
            full = self._full_name_of(tab)
            short = self.short_map.get(full, full)
            if short and tab.Title != short:
                try:
                    tab.Title = short
                    applied += 1
                except Exception:
                    pass
        if self.on_state_callback:
            self.on_state_callback("short")
        if self.on_save_callback:
            self.on_save_callback(self.short_map)
        self.message = "Applied short names to " + str(applied) + " tab(s)."
        self._update_sub(self.message)

    def _on_restore_full(self, sender, args):
        self._commit_grid()
        restored = 0
        for tab in self.live_tabs:
            full = self._full_name_of(tab)
            if full and tab.Title != full:
                try:
                    tab.Title = full
                    restored += 1
                except Exception:
                    pass
        if self.on_state_callback:
            self.on_state_callback("full")
        self.message = "Restored full names on " + str(restored) + " tab(s)."
        self._update_sub(self.message)

    def _on_save(self, sender, args):
        self._commit_grid()
        m = self._collect_map_from_grid()
        self.short_map.update(m)
        if self.on_save_callback:
            ok = self.on_save_callback(self.short_map)
        else:
            ok = False
        self.message = "Short-name map saved." if ok else "Could not save map."
        self._update_sub(self.message)

    def _on_reset(self, sender, args):
        self._commit_grid()
        for row in self.table.Rows:
            full = self._cell(row, "CurrentName")
            row["ShortName"] = self.default_map.get(full, full)
        self._update_sub("Short names reset to defaults (not yet applied).")

    def _on_close(self, sender, args):
        self.Close()

    def _update_sub(self, text):
        self.HeaderSub.Text = text

def show_ribbon_names_dialog(live_tabs, short_map, originals, default_map, on_save_callback, on_state_callback, on_originals_callback):
    """Factory function to show the Ribbon Name dialog."""
    dlg = RibbonNameWindow(live_tabs, short_map, originals, default_map, on_save_callback, on_state_callback, on_originals_callback)
    dlg.ShowDialog()
    return dlg
