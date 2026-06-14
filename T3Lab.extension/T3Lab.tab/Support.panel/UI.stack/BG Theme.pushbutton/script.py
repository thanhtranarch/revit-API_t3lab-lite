from System.Windows import WindowState
# -*- coding: utf-8 -*-
"""
DQT - Background Theme
Set the Revit model-view background color from a themed picker with live preview,
ready-made presets (Black / Gray / White / Dark Blue / Studio), RGB sliders,
HEX input and quick-cycle. Replaces the blind 3-color cycle with full control
and a remembered last choice.

Dang Quoc Truong - DQT (c) 2026
"""

__title__     = "Background\nTheme"
__author__    = "Dang Quoc Truong (DQT)"
__version__   = "1.0.0"
__copyright__ = "Copyright (c) 2026 by Dang Quoc Truong (DQT)"
__doc__       = """DQT - Background Theme

Improved background colour tool. Open a themed picker to choose a preset
(Black / Gray / White / Dark Blue / Studio), fine-tune with RGB sliders or a
HEX value, see a live preview, and apply. SHIFT+Click quick-cycles
Black -> Gray -> White -> Black like the classic tool.

Works on Revit 2024 / 2025 / 2026 / 2027.
"""

# ------------------------------------------------------------------ IMPORTS
import os
import codecs
import json

import Autodesk.Revit.DB as DB

from pyrevit import script

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")

from System.IO import MemoryStream
from System.Text import Encoding
from System.Windows.Markup import XamlReader
from System.Windows.Media import BrushConverter

# ------------------------------------------------------------------ GENERAL
app = __revit__.Application

PATH_SCRIPT = os.path.dirname(__file__)
CONFIG_PATH = os.path.join(PATH_SCRIPT, "dqt_bg_config.json")

# Detect SHIFT+Click for quick-cycle (classic B/W/G behaviour)
try:
    SHIFT_CLICK = __shiftclick__  # provided by pyRevit when SHIFT is held
except Exception:
    SHIFT_CLICK = False

# ------------------------------------------------------------------ DQT BRAND
HEADER_BG   = "#0F172A"
MAIN_BG     = "#FFFFFF"
ACCENT      = "#CBD5E1"
DARK_ACCENT = "#5D4E37"
TEXT        = "#0F172A"
BORDER      = "#E0E0E0"
FOOTER_TEXT = "Dang Quoc Truong - DQT (c) 2026"


def brush(hex_string):
    """IronPython-safe brush from hex (never Color.FromRgb in IronPython)."""
    return BrushConverter().ConvertFromString(hex_string)


# ------------------------------------------------------------------ PRESETS
# name -> (R, G, B)
PRESETS = [
    ("Black",      (0,   0,   0)),
    ("Charcoal",   (45,  45,  45)),
    ("Gray",       (190, 190, 190)),
    ("White",      (255, 255, 255)),
    ("Dark Blue",  (28,  40,  64)),
    ("Studio",     (54,  61,  74)),
]
PRESET_MAP = dict(PRESETS)


def clamp(v):
    v = int(round(v))
    if v < 0:
        return 0
    if v > 255:
        return 255
    return v


def _hex2(v):
    """Two-digit uppercase hex for a 0-255 int (IronPython-safe, no format())."""
    v = clamp(v)
    digits = "0123456789ABCDEF"
    return digits[v // 16] + digits[v % 16]


def to_hex(r, g, b):
    return "#" + _hex2(r) + _hex2(g) + _hex2(b)


def from_hex(hex_string):
    """Parse #RRGGBB -> (r, g, b). Returns None if invalid."""
    if not hex_string:
        return None
    s = hex_string.strip().lstrip("#")
    if len(s) != 6:
        return None
    try:
        return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
    except Exception:
        return None


# ------------------------------------------------------------------ CONFIG IO
def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                data = json.load(f)
                return (
                    clamp(data.get("r", 0)),
                    clamp(data.get("g", 0)),
                    clamp(data.get("b", 0)),
                )
        except Exception:
            pass
    # default: current Revit background, fallback black
    try:
        c = app.BackgroundColor
        return (c.Red, c.Green, c.Blue)
    except Exception:
        return (0, 0, 0)


def save_config(r, g, b):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump({"r": clamp(r), "g": clamp(g), "b": clamp(b)}, f)
    except Exception:
        pass


# ------------------------------------------------------------------ APPLY
def apply_background(r, g, b):
    """Set the Revit model view background. No transaction needed -
    BackgroundColor is an application-session property."""
    app.BackgroundColor = DB.Color(clamp(r), clamp(g), clamp(b))


# ------------------------------------------------------------------ QUICK CYCLE
def quick_cycle():
    """Classic Black -> Gray -> White -> Black cycle (SHIFT+Click)."""
    try:
        c = app.BackgroundColor
        cr, cg, cb = c.Red, c.Green, c.Blue
    except Exception:
        cr = cg = cb = 0

    if cr >= 250 and cg >= 250 and cb >= 250:        # white -> black
        nr, ng, nb = 0, 0, 0
        name = "Black"
    elif cr <= 5 and cg <= 5 and cb <= 5:            # black -> gray
        nr, ng, nb = 190, 190, 190
        name = "Gray"
    else:                                            # anything else -> white
        nr, ng, nb = 255, 255, 255
        name = "White"

    apply_background(nr, ng, nb)
    save_config(nr, ng, nb)
    script.get_output().print_md(
        "**DQT - Background Theme:** switched to **" + name + "** (" +
        to_hex(nr, ng, nb) + ").\n\n*" + FOOTER_TEXT + "*")


# ------------------------------------------------------------------ XAML



# ------------------------------------------------------------------ WINDOW
class BackgroundThemeWindow(object):
    def __init__(self, r, g, b):
        # Find T3Lab.extension parent folder dynamically
        current_dir = os.path.dirname(__file__)
        while current_dir and not current_dir.endswith('T3Lab.extension'):
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                break
            current_dir = parent
        xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "BGTheme.xaml")

        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()
        stream = MemoryStream(Encoding.UTF8.GetBytes(xaml_content))
        self.win = XamlReader.Load(stream)

        self.applied = False
        self._suspend = False  # guard against slider/hex feedback loops

        # Named controls
        self.preset_grid = self.win.FindName("PresetGrid")
        self.slider_r = self.win.FindName("SliderR")
        self.slider_g = self.win.FindName("SliderG")
        self.slider_b = self.win.FindName("SliderB")
        self.lbl_r = self.win.FindName("LblR")
        self.lbl_g = self.win.FindName("LblG")
        self.lbl_b = self.win.FindName("LblB")
        self.hex_box = self.win.FindName("HexBox")
        self.hex_apply = self.win.FindName("HexApply")
        self.preview_box = self.win.FindName("PreviewBox")
        self.preview_lbl = self.win.FindName("PreviewLbl")
        self.btn_apply = self.win.FindName("BtnApply")
        self.btn_close = self.win.FindName("BtnClose")

        # Build preset buttons
        self._build_presets()

        # Wire events
        self.slider_r.ValueChanged += self._on_slider
        self.slider_g.ValueChanged += self._on_slider
        self.slider_b.ValueChanged += self._on_slider
        self.hex_apply.Click += self._on_hex_apply
        self.btn_apply.Click += self._on_apply
        self.btn_close.Click += self._on_close

        # Initial state
        self.set_rgb(r, g, b)

    # --- preset buttons (button-based, not list selection -> IronPython-safe)
    def _build_presets(self):
        import System.Windows.Controls as WC
        preset_style = self.win.Resources["PresetButton"]
        for name, rgb in PRESETS:
            btn = WC.Button()
            btn.Content = name
            btn.Style = preset_style
            r, g, b = rgb
            # store the rgb on the button so the click handler can read it back
            btn.Tag = str(r) + "," + str(g) + "," + str(b)
            btn.Click += self._on_preset_click
            self.preset_grid.Children.Add(btn)

    def _on_preset_click(self, sender, args):
        try:
            parts = sender.Tag.split(",")
            r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
            self.set_rgb(r, g, b)
        except Exception:
            pass

    # --- central state setter (keeps sliders / hex / preview in sync)
    def set_rgb(self, r, g, b):
        r, g, b = clamp(r), clamp(g), clamp(b)
        self._suspend = True
        try:
            self.slider_r.Value = r
            self.slider_g.Value = g
            self.slider_b.Value = b
            self.lbl_r.Text = str(r)
            self.lbl_g.Text = str(g)
            self.lbl_b.Text = str(b)
            self.hex_box.Text = to_hex(r, g, b)
            self._update_preview(r, g, b)
        finally:
            self._suspend = False

    def _update_preview(self, r, g, b):
        hex_str = to_hex(r, g, b)
        self.preview_box.Background = brush(hex_str)
        # readable label colour: light text on dark bg, dark text on light bg
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        self.preview_lbl.Foreground = brush("#FFFFFF") if luminance < 140 else brush("#0F172A")
        self.preview_lbl.Text = hex_str

    # --- events
    def _on_slider(self, sender, args):
        if self._suspend:
            return
        r = int(self.slider_r.Value)
        g = int(self.slider_g.Value)
        b = int(self.slider_b.Value)
        self.lbl_r.Text = str(r)
        self.lbl_g.Text = str(g)
        self.lbl_b.Text = str(b)
        self._suspend = True
        try:
            self.hex_box.Text = to_hex(r, g, b)
        finally:
            self._suspend = False
        self._update_preview(r, g, b)

    def _on_hex_apply(self, sender, args):
        rgb = from_hex(self.hex_box.Text)
        if rgb is None:
            self.preview_lbl.Text = "Invalid HEX"
            return
        self.set_rgb(*rgb)

    def current_rgb(self):
        return (int(self.slider_r.Value),
                int(self.slider_g.Value),
                int(self.slider_b.Value))

    def _on_apply(self, sender, args):
        r, g, b = self.current_rgb()
        apply_background(r, g, b)
        save_config(r, g, b)
        self.applied = True
        # leave window open so the user can keep trying colours

    def _on_close(self, sender, args):
        self.win.Close()

    def show(self):
        self.win.ShowDialog()


# ------------------------------------------------------------------ MAIN
def main():
    if SHIFT_CLICK:
        quick_cycle()
        return

    r, g, b = load_config()
    dlg = BackgroundThemeWindow(r, g, b)
    dlg.show()

    if dlg.applied:
        cr, cg, cb = dlg.current_rgb()
        script.get_output().print_md(
            "**DQT - Background Theme:** background set to **" +
            to_hex(cr, cg, cb) + "**.\n\n*" + FOOTER_TEXT + "*")


if __name__ == "__main__":
    main()
