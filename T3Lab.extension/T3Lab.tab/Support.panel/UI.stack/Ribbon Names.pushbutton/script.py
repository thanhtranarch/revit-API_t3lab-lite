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

import os
import json
from pyrevit import script, forms
from pyrevit.api import AdWindows

# ------------------------------------------------------------------ GENERAL
app = __revit__.Application

PATH_SCRIPT  = os.path.dirname(__file__)
MAP_PATH     = os.path.join(PATH_SCRIPT, "dqt_ribbon_map.json")
ORIG_PATH    = os.path.join(PATH_SCRIPT, "dqt_ribbon_originals.json")
STATE_PATH   = os.path.join(PATH_SCRIPT, "dqt_ribbon_state.json")

FOOTER_TEXT = "Dang Quoc Truong - DQT (c) 2026"

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
    return _read_json(STATE_PATH, {"mode": "full"}).get("mode", "full")

def save_state(mode):
    _write_json(STATE_PATH, {"mode": mode})

def get_ribbon_tabs():
    tabs = []
    try:
        for tab in AdWindows.ComponentManager.Ribbon.Tabs:
            tabs.append(tab)
    except Exception:
        pass
    return tabs

def main():
    live_tabs = get_ribbon_tabs()
    if not live_tabs:
        forms.alert("No ribbon tabs found.", title="DQT - Ribbon Name Manager")
        return

    short_map = load_map()
    originals = load_originals()

    # Capture originals logic
    short_to_full = {short: full for full, short in short_map.items()}
    changed = False
    for tab in live_tabs:
        title = tab.Title
        if title in short_to_full:
            full = short_to_full[title]
            if full not in originals:
                originals[full] = full
                changed = True
        else:
            if title not in originals:
                originals[title] = title
                changed = True
    if changed:
        save_originals(originals)

    from GUI.RibbonNamesDialog import show_ribbon_names_dialog

    dlg = show_ribbon_names_dialog(
        live_tabs=live_tabs,
        short_map=short_map,
        originals=originals,
        default_map=DEFAULT_MAP,
        on_save_callback=save_map,
        on_state_callback=save_state,
        on_originals_callback=save_originals
    )

    if dlg.message:
        script.get_output().print_md(
            "**DQT - Ribbon Name Manager:** " + dlg.message +
            "\n\n*" + FOOTER_TEXT + "*")

if __name__ == "__main__":
    main()
