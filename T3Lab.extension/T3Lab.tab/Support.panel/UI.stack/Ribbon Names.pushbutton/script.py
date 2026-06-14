from System.Windows import WindowState
# -*- coding: utf-8 -*-
"""
DQT - Ribbon Name Manager
Shorten / restore Revit ribbon tab names with full control. Unlike the classic
fixed-JSON tool, this reads every live ribbon tab, lets you edit each short
name inline (double-click), toggle Short/Full for all tabs, persists your own
mappings, and needs no external Snippets dependency or per-language files.

Dang Quoc Truong - DQT (c) 2026
"""

__title__     = "Ribbon Name\nManager"
__author__    = "Dang Quoc Truong (DQT)"
__version__   = "1.0.0"
__copyright__ = "Copyright (c) 2026 by Dang Quoc Truong (DQT)"
__doc__       = """DQT - Ribbon Name Manager

Improved ribbon-name tool. Opens a themed window listing every ribbon tab with
its current name and your short name. Double-click a short-name cell to edit it.
Then:
  - Apply Short  -> renames all tabs to their short names
  - Restore Full -> puts the original full names back
  - Save Map     -> remembers your custom short names

No external Snippets._context_manager dependency, no fixed language JSON files.

Works on Revit 2024 / 2025 / 2026 / 2027.
"""

# ------------------------------------------------------------------ IMPORTS
import os
import codecs
import json

from pyrevit import script, forms
from pyrevit.api import AdWindows

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")
clr.AddReference("System.Data")

from System.IO import MemoryStream
from System.Text import Encoding
from System.Windows.Markup import XamlReader
from System.Windows.Media import BrushConverter
from System.Data import DataTable
from System import String as System_String
from System import DBNull

# ------------------------------------------------------------------ GENERAL
app = __revit__.Application

PATH_SCRIPT  = os.path.dirname(__file__)
MAP_PATH     = os.path.join(PATH_SCRIPT, "dqt_ribbon_map.json")
ORIG_PATH    = os.path.join(PATH_SCRIPT, "dqt_ribbon_originals.json")
STATE_PATH   = os.path.join(PATH_SCRIPT, "dqt_ribbon_state.json")

FOOTER_TEXT = "Dang Quoc Truong - DQT (c) 2026"

# unicode() exists in IronPython 2.7 but not CPython3; shim for both engines
try:
    _unicode = unicode  # noqa: F821  (IronPython 2)
except NameError:
    _unicode = str       # CPython 3

# ------------------------------------------------------------------ DQT BRAND
def brush(hex_string):
    return BrushConverter().ConvertFromString(hex_string)


# ------------------------------------------------------------------ DEFAULT MAP
# Sensible defaults for common Revit built-in + popular extension tabs.
# These are only suggestions; the user can edit any of them in the grid and
# save their own version.
DEFAULT_MAP = {
    "Architecture": "Arch",
    "Structure": "Struc",
    "Steel": "Steel",
    "Precast": "Precast",
    "Systems": "MEP",
    "Insert": "Insert",
    "Annotate": "Anno",
    "Analyze": "Analyze",
    "Massing & Site": "Mass&Site",
    "Collaborate": "Collab",
    "View": "View",
    "Manage": "Manage",
    "Add-Ins": "Add-Ins",
    "Modify": "Modify",
    "Create": "Create",
    "Family Editor": "Fam.Editor",
    "BIM Interoperability Tools": "BIM-IOT",
    "Enscape": "Ensc",
    "pyRevit": "pyRevit",
    "Rhino.Inside": "RiR",
}


# ------------------------------------------------------------------ JSON IO
def _read_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def _write_json(path, data):
    try:
        with open(path, "w") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def load_map():
    data = _read_json(MAP_PATH, None)
    if data is None:
        return dict(DEFAULT_MAP)
    merged = dict(DEFAULT_MAP)
    merged.update(data)
    return merged


def save_map(m):
    return _write_json(MAP_PATH, m)


def load_originals():
    return _read_json(ORIG_PATH, {})


def save_originals(m):
    return _write_json(ORIG_PATH, m)


def load_state():
    # state: "full" or "short"
    return _read_json(STATE_PATH, {"mode": "full"}).get("mode", "full")


def save_state(mode):
    _write_json(STATE_PATH, {"mode": mode})


# ------------------------------------------------------------------ RIBBON HELPERS
def get_ribbon_tabs():
    """Return list of live AdWindows ribbon tab objects."""
    tabs = []
    try:
        for tab in AdWindows.ComponentManager.Ribbon.Tabs:
            tabs.append(tab)
    except Exception:
        pass
    return tabs


# ------------------------------------------------------------------ XAML



# ------------------------------------------------------------------ WINDOW
class RibbonNameWindow(object):
    def __init__(self):
        # Find T3Lab.extension parent folder dynamically
        current_dir = os.path.dirname(__file__)
        while current_dir and not current_dir.endswith('T3Lab.extension'):
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                break
            current_dir = parent
        xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "RibbonNames.xaml")

        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()
        stream = MemoryStream(Encoding.UTF8.GetBytes(xaml_content))
        self.win = XamlReader.Load(stream)

        self.short_map = load_map()
        self.originals = load_originals()  # original full names captured once

        # Capture original full names the first time we ever run, keyed by
        # the *short* name too so we can restore regardless of current mode.
        self.live_tabs = get_ribbon_tabs()
        self._capture_originals()

        self.grid = self.win.FindName("Grid")
        self.table = DataTable("tabs")
        self.table.Columns.Add("CurrentName", System_String)
        self.table.Columns.Add("ShortName", System_String)
        self._build_rows()
        self.grid.ItemsSource = self.table.DefaultView

        # Wire buttons (button-based interaction; no SelectionChanged handlers)
        self.win.FindName("BtnShort").Click += self._on_apply_short
        self.win.FindName("BtnFull").Click += self._on_restore_full
        self.win.FindName("BtnSave").Click += self._on_save
        self.win.FindName("BtnReset").Click += self._on_reset
        self.win.FindName("BtnClose").Click += self._on_close

        self.message = None

    def _capture_originals(self):
        """Record the original full name of each tab keyed by a stable id.
        We use the *current* title as the key only if we don't already have
        an original recorded. The map of short->full lets us restore even
        after a previous shorten."""
        # Build reverse lookup short->full from the saved map
        short_to_full = {}
        for full, short in self.short_map.items():
            short_to_full[short] = full

        changed = False
        for tab in self.live_tabs:
            title = tab.Title
            # If the current title is a known short, the original is its full
            if title in short_to_full:
                full = short_to_full[title]
                if full not in self.originals:
                    self.originals[full] = full
                    changed = True
            else:
                # current title is (probably) a full name
                if title not in self.originals:
                    self.originals[title] = title
                    changed = True
        if changed:
            save_originals(self.originals)

    def _full_name_of(self, tab):
        """Resolve the full name for a live tab, accounting for the fact that
        it may currently be displaying a short name."""
        title = tab.Title
        # reverse map: short -> full
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
        """Read edited short names back out of the DataTable."""
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
        """Safe string read from a DataRow cell (handles DBNull/None)."""
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
        """Flush any in-progress edit out of the DataGrid into the DataTable.
        DataGrid often needs two CommitEdit passes (cell then row)."""
        try:
            self.grid.CommitEdit()
            self.grid.CommitEdit()
        except Exception:
            pass

    # --- actions
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
        save_state("short")
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
        save_state("full")
        self.message = "Restored full names on " + str(restored) + " tab(s)."
        self._update_sub(self.message)

    def _on_save(self, sender, args):
        self._commit_grid()
        m = self._collect_map_from_grid()
        self.short_map.update(m)
        ok = save_map(self.short_map)
        self.message = "Short-name map saved." if ok else "Could not save map."
        self._update_sub(self.message)

    def _on_reset(self, sender, args):
        # Reset short names in the grid back to the DEFAULT_MAP / full name
        self._commit_grid()
        for row in self.table.Rows:
            full = self._cell(row, "CurrentName")
            row["ShortName"] = DEFAULT_MAP.get(full, full)
        self._update_sub("Short names reset to defaults (not yet applied).")

    def _on_close(self, sender, args):
        self.win.Close()

    def _update_sub(self, text):
        sub = self.win.FindName("HeaderSub")
        if sub is not None:
            sub.Text = text

    def show(self):
        self.win.ShowDialog()


# ------------------------------------------------------------------ MAIN
def main():
    tabs = get_ribbon_tabs()
    if not tabs:
        forms.alert("No ribbon tabs found.", title="DQT - Ribbon Name Manager")
        return

    dlg = RibbonNameWindow()
    dlg.show()

    if dlg.message:
        script.get_output().print_md(
            "**DQT - Ribbon Name Manager:** " + dlg.message +
            "\n\n*" + FOOTER_TEXT + "*")


if __name__ == "__main__":
    main()
