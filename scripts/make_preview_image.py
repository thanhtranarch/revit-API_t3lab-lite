#!/usr/bin/env python3
"""
Render all Swiss SVG icons to a combined PNG grid using cairosvg (if available)
or PIL XML parsing fallback. Outputs one PNG per panel.
"""
import os, sys, base64, io

# Try importing cairosvg for proper SVG rendering
try:
    import cairosvg
    HAS_CAIRO = True
except ImportError:
    HAS_CAIRO = False

try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

if not HAS_PIL:
    print("PIL not available — install Pillow first.")
    sys.exit(1)

TAB_DIR = os.path.join("T3Lab.extension", "T3Lab.tab")
OUT_DIR = "scripts"

# Collect all icons grouped by panel
panels = {}
for root, dirs, files in os.walk(TAB_DIR):
    dn = os.path.basename(root)
    if dn.endswith((".pushbutton", ".urlbutton", ".pulldown")):
        rel = os.path.relpath(root, TAB_DIR).split(os.sep)
        panel = rel[0]
        name  = dn.rsplit(".", 1)[0]
        light = os.path.join(root, "icon.png")
        dark  = os.path.join(root, "icon.dark.png")
        if os.path.exists(light) or os.path.exists(dark):
            panels.setdefault(panel, []).append({
                "name": name,
                "light": light if os.path.exists(light) else None,
                "dark":  dark  if os.path.exists(dark)  else None,
            })

# Build one big composite
CELL      = 64   # icon size
PADDING   = 14
CARD_W    = 200
CARD_H    = 140
COLS      = 6
BG        = (240, 240, 242)
CARD_BG   = (255, 255, 255)
DARK_BG   = (15,  23,  42)
NAVY      = (24,  42,  62)
AMBER     = (208, 120, 24)
TEXT_COL  = (15,  23,  42)
GREY      = (100, 116, 139)
HEADER_H  = 80

all_panels = sorted(panels.keys())

# Total height calculation
def panel_height(icons):
    rows = (len(icons) + COLS - 1) // COLS
    return HEADER_H + rows * (CARD_H + PADDING) + PADDING

total_h = sum(panel_height(panels[p]) for p in all_panels) + 40
total_w = COLS * (CARD_W + PADDING) + PADDING

img  = Image.new("RGB", (total_w, total_h), BG)
draw = ImageDraw.Draw(img)

# Try loading a font
try:
    font_bold  = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf",  11)
    font_small = ImageFont.truetype("C:/Windows/Fonts/arial.ttf",     9)
    font_panel = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf",  13)
except:
    font_bold = font_small = font_panel = ImageFont.load_default()

y_cursor = 20

for panel_name in all_panels:
    icons = sorted(panels[panel_name], key=lambda x: x["name"])

    # Panel label
    label = panel_name.replace(".panel", "").upper()
    draw.text((PADDING, y_cursor + 4), label, fill=NAVY, font=font_panel)
    draw.line([(PADDING + len(label) * 8 + 10, y_cursor + 10),
               (total_w - PADDING, y_cursor + 10)], fill=(203, 213, 225), width=1)
    y_cursor += HEADER_H

    for idx, ic in enumerate(icons):
        col = idx % COLS
        row = idx // COLS
        cx  = PADDING + col * (CARD_W + PADDING)
        cy  = y_cursor + row * (CARD_H + PADDING)

        # Card background
        draw.rounded_rectangle([cx, cy, cx + CARD_W, cy + CARD_H],
                                radius=10, fill=CARD_BG,
                                outline=(226, 232, 240), width=1)

        # Name text
        name = ic["name"]
        if len(name) > 18:
            name = name[:17] + "…"
        draw.text((cx + 8, cy + 8), name, fill=TEXT_COL, font=font_bold)

        # Light preview box
        lx, ly = cx + 8, cy + 26
        draw.rounded_rectangle([lx, ly, lx + 84, ly + 82],
                                radius=6, fill=(248, 250, 252),
                                outline=(226, 232, 240), width=1)
        if ic["light"]:
            try:
                icon_img = Image.open(ic["light"]).convert("RGBA")
                icon_img = icon_img.resize((64, 64), Image.LANCZOS)
                img.paste(icon_img, (lx + 10, ly + 10), icon_img)
            except:
                pass
        draw.text((lx + 4, ly + 68), "LIGHT", fill=GREY, font=font_small)

        # Dark preview box
        dx, dy = cx + 104, cy + 26
        draw.rounded_rectangle([dx, dy, dx + 84, dy + 82],
                                radius=6, fill=DARK_BG,
                                outline=(30, 41, 59), width=1)
        if ic["dark"]:
            try:
                icon_img = Image.open(ic["dark"]).convert("RGBA")
                icon_img = icon_img.resize((64, 64), Image.LANCZOS)
                img.paste(icon_img, (dx + 10, dy + 10), icon_img)
            except:
                pass
        draw.text((dx + 4, dy + 68), "DARK", fill=(71, 85, 105), font=font_small)

    rows_used = (len(icons) + COLS - 1) // COLS
    y_cursor += rows_used * (CARD_H + PADDING) + PADDING

out_path = os.path.join(OUT_DIR, "swiss_icons_preview.png")
img.save(out_path, "PNG")
print(f"Preview image saved: {out_path}  ({total_w}x{total_h}px, {len(all_panels)} panels, {sum(len(v) for v in panels.values())} icons)")
