# -*- coding: utf-8 -*-
"""
Minimal-line icon redesign for the Standards & Settings panel, finished with
the same soft rainbow aura glow used on the T3Lab Assistant icon
(see draw_assistant.py): a blurred 4-colour halo sits behind a crisp dark
(or light, for the dark-theme variant) line glyph.
Renders at 8x supersampling then downsamples to 64x64 for crisp anti-aliased strokes.
"""

import os
from PIL import Image, ImageDraw, ImageFilter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PANEL_DIR = os.path.join(BASE, "T3Lab.extension", "T3Lab.tab", "Standards & Settings.panel")

INK = (15, 23, 42, 255)        # #0F172A - Lumina ink (light theme icon)
PALE = (248, 250, 252, 255)    # #F8FAFC - Lumina surface (dark theme icon)
ACCENT = (59, 130, 246, 255)   # #3B82F6 - Lumina accent blue

# Same 4-hue aura palette as T3LabAssistant icon (draw_assistant.py), swept
# around the icon like a colour wheel so the halo reads as a smooth rainbow.
# Gold/orange/pink are all warm hues clustered together in hue-space, so a
# plain linear sweep over the full 90deg between stops let them bleed into
# one muddy orange-black smear (esp. near the bottom of the icon, where the
# dark glyph stroke sits right on top of the gold->orange band). Fix: keep
# each stop's pure colour across most of its arc and only blend within a
# narrow band at the boundary, and push orange/pink apart in hue/brightness
# so they read as distinct colours instead of one warm blob.
GLOW_STOPS = [
    (0, (0, 170, 255)),     # blue
    (90, (255, 200, 0)),    # gold
    (180, (255, 60, 0)),    # orange (kept apart from gold by the blue/pink neighbours)
    (270, (255, 0, 140)),   # magenta/pink - pushed away from orange's hue
]
BLEND_BAND = 18  # degrees of cross-fade at each boundary; rest stays pure

SCALE = 8
SIZE = 64 * SCALE
SW = 5 * SCALE        # crisp glyph stroke width
GLOW_SW = 13 * SCALE  # thicker stroke for the glow silhouette mask - more
                       # visible colour ring beyond the dark glyph edge
GLOW_BLUR = 8 * SCALE / 8.0


def _lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def make_angular_gradient(size):
    """Build a 64x64 conic colour-wheel sweep, then let callers upscale it."""
    import math
    grad = Image.new("RGB", (size, size))
    cx = cy = size / 2.0
    stops = GLOW_STOPS + [(360, GLOW_STOPS[0][1])]
    for y in range(size):
        for x in range(size):
            angle = (math.degrees(math.atan2(y - cy, x - cx)) + 360) % 360
            for (a0, c0), (a1, c1) in zip(stops, stops[1:]):
                if a0 <= angle <= a1:
                    seg_len = a1 - a0
                    band = min(BLEND_BAND, seg_len)
                    if angle < a1 - band:
                        color = c0
                    else:
                        t = (angle - (a1 - band)) / band if band else 0
                        color = _lerp(c0, c1, t)
                    grad.putpixel((x, y), color)
                    break
    return grad


def new_canvas():
    return Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))


# ---------------------------------------------------------------------------
# Glyph shapes (each takes draw, color, accent, width)
# ---------------------------------------------------------------------------

def draw_model_auditor(draw, color, accent, w):
    cx, cy, r = 28 * SCALE, 28 * SCALE, 16 * SCALE
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], outline=color, width=w)
    draw.line([cx + r * 0.7, cy + r * 0.7, 56 * SCALE, 56 * SCALE], fill=color, width=w, joint="curve")
    aw = max(1, int(w * 0.85))
    draw.line([cx - 7 * SCALE, cy, cx - 1 * SCALE, cy + 6 * SCALE], fill=accent, width=aw, joint="curve")
    draw.line([cx - 1 * SCALE, cy + 6 * SCALE, cx + 8 * SCALE, cy - 6 * SCALE], fill=accent, width=aw, joint="curve")


def draw_mana_loca(draw, color, accent, w):
    cx, top, bottom = 32 * SCALE, 8 * SCALE, 50 * SCALE
    r = 18 * SCALE
    draw.pieslice([cx - r, top, cx + r, top + 2 * r], 180, 360, outline=color, width=w)
    draw.line([cx - r, top + r, cx, bottom], fill=color, width=w, joint="curve")
    draw.line([cx, bottom, cx + r, top + r], fill=color, width=w, joint="curve")
    draw.ellipse([cx - 6 * SCALE, top + r - 6 * SCALE, cx + 6 * SCALE, top + r + 6 * SCALE], fill=accent)


def draw_mana_styles(draw, color, accent, w):
    rows = [18, 32, 46]
    dashes = [None, (10, 6), (3, 5)]
    lw = max(1, int(w * 0.7))
    for y, dash in zip(rows, dashes):
        yy = y * SCALE
        if dash is None:
            draw.line([10 * SCALE, yy, 54 * SCALE, yy], fill=color, width=lw)
        else:
            on, off = dash
            x = 10 * SCALE
            while x < 54 * SCALE:
                x2 = min(x + on * SCALE, 54 * SCALE)
                draw.line([x, yy, x2, yy], fill=color, width=lw)
                x = x2 + off * SCALE
    rr = 6 * SCALE
    draw.ellipse([34 * SCALE - rr, 18 * SCALE - rr, 34 * SCALE + rr, 18 * SCALE + rr], fill=accent)
    draw.ellipse([24 * SCALE - rr, 32 * SCALE - rr, 24 * SCALE + rr, 32 * SCALE + rr], fill=color)
    draw.ellipse([42 * SCALE - rr, 46 * SCALE - rr, 42 * SCALE + rr, 46 * SCALE + rr], fill=color)


def draw_mana_workset(draw, color, accent, w):
    cx = 32 * SCALE
    rx, ry = 20 * SCALE, 7 * SCALE
    levels = [16, 30, 44]
    for i, y in enumerate(levels):
        yy = y * SCALE
        draw.ellipse([cx - rx, yy - ry, cx + rx, yy + ry], outline=color, width=w)
        if i < len(levels) - 1:
            y_next = levels[i + 1] * SCALE
            draw.line([cx - rx, yy, cx - rx, y_next], fill=color, width=w)
            draw.line([cx + rx, yy, cx + rx, y_next], fill=color, width=w)
    draw.pieslice([cx - rx, levels[0] * SCALE - ry, cx + rx, levels[0] * SCALE + ry], 0, 180, outline=accent, width=w)


ICONS = {
    "ModelAuditor.pushbutton": draw_model_auditor,
    "ManaLoca.pushbutton": draw_mana_loca,
    "ManaStyles.pushbutton": draw_mana_styles,
    "ManaWorkset.pushbutton": draw_mana_workset,
}

_GRADIENT_64 = make_angular_gradient(64)
_GRADIENT_FULL = _GRADIENT_64.resize((SIZE, SIZE), Image.Resampling.BILINEAR).convert("RGBA")


def make_glow(draw_fn):
    mask_layer = new_canvas()
    d = ImageDraw.Draw(mask_layer)
    white = (255, 255, 255, 255)
    draw_fn(d, white, white, GLOW_SW)
    mask = mask_layer.split()[3]
    mask = mask.filter(ImageFilter.GaussianBlur(radius=GLOW_BLUR))

    glow = _GRADIENT_FULL.copy()
    glow.putalpha(mask)
    return glow


def render(draw_fn, glyph_color, accent):
    canvas = make_glow(draw_fn)
    crisp = new_canvas()
    d = ImageDraw.Draw(crisp)
    draw_fn(d, glyph_color, accent, SW)
    canvas = Image.alpha_composite(canvas, crisp)
    return canvas.resize((64, 64), Image.Resampling.LANCZOS)


def main():
    for folder, draw_fn in ICONS.items():
        dest = os.path.join(PANEL_DIR, folder)
        if not os.path.isdir(dest):
            print(f"SKIP (not found): {folder}")
            continue

        light_img = render(draw_fn, INK, ACCENT)
        light_img.save(os.path.join(dest, "icon.png"), "PNG")

        dark_img = render(draw_fn, PALE, ACCENT)
        dark_img.save(os.path.join(dest, "icon.dark.png"), "PNG")

        for svg_name in ("icon.svg", "icon.dark.svg"):
            svg_path = os.path.join(dest, svg_name)
            if os.path.exists(svg_path):
                os.remove(svg_path)

        print(f"OK: {folder}")

    print("Done.")


if __name__ == "__main__":
    main()
