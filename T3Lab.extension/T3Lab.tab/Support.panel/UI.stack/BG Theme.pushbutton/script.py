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

import os
import json
import Autodesk.Revit.DB as DB

# Detect SHIFT+Click for quick-cycle (classic B/W/G behaviour)
try:
    SHIFT_CLICK = __shiftclick__
except Exception:
    SHIFT_CLICK = False

PATH_SCRIPT = os.path.dirname(__file__)
CONFIG_PATH = os.path.join(PATH_SCRIPT, "dqt_bg_config.json")

PRESETS = [
    ("Black",      (0,   0,   0)),
    ("Charcoal",   (45,  45,  45)),
    ("Gray",       (190, 190, 190)),
    ("White",      (255, 255, 255)),
    ("Dark Blue",  (28,  40,  64)),
    ("Studio",     (54,  61,  74)),
]

def clamp(v):
    v = int(round(v))
    return max(0, min(v, 255))

def _hex2(v):
    v = clamp(v)
    digits = "0123456789ABCDEF"
    return digits[v // 16] + digits[v % 16]

def to_hex(r, g, b):
    return "#" + _hex2(r) + _hex2(g) + _hex2(b)

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
    try:
        c = __revit__.Application.BackgroundColor
        return (c.Red, c.Green, c.Blue)
    except Exception:
        return (0, 0, 0)

def save_config(r, g, b):
    try:
        with open(CONFIG_PATH, "w") as f:
            json.dump({"r": clamp(r), "g": clamp(g), "b": clamp(b)}, f)
    except Exception:
        pass

def apply_background(r, g, b):
    __revit__.Application.BackgroundColor = DB.Color(clamp(r), clamp(g), clamp(b))

def quick_cycle():
    try:
        c = __revit__.Application.BackgroundColor
        cr, cg, cb = c.Red, c.Green, c.Blue
    except Exception:
        cr = cg = cb = 0

    if cr >= 250 and cg >= 250 and cb >= 250:
        nr, ng, nb = 0, 0, 0
    elif cr <= 5 and cg <= 5 and cb <= 5:
        nr, ng, nb = 190, 190, 190
    else:
        nr, ng, nb = 255, 255, 255

    apply_background(nr, ng, nb)
    save_config(nr, ng, nb)

def main():
    if SHIFT_CLICK:
        quick_cycle()
        return

    r, g, b = load_config()

    # Import the custom WPF Dialog from lib/GUI
    from GUI.BGThemeDialog import show_bg_theme_dialog

    def on_apply(new_r, new_g, new_b):
        apply_background(new_r, new_g, new_b)
        save_config(new_r, new_g, new_b)

    show_bg_theme_dialog(r, g, b, PRESETS, on_apply)

if __name__ == "__main__":
    main()
