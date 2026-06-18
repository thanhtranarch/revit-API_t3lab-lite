#!/usr/bin/env python3
import os

SVG_DIR = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\scripts\svg_icons"
OUT = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\scripts\icon_preview.html"

panels = {}
for panel in sorted(os.listdir(SVG_DIR)):
    panel_dir = os.path.join(SVG_DIR, panel)
    if os.path.isdir(panel_dir):
        icons = sorted(f for f in os.listdir(panel_dir) if f.endswith(".svg"))
        panels[panel] = icons

total = sum(len(v) for v in panels.values())

parts = []
parts.append("""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>T3Lab SVG Icon Preview</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #ECECEF; font-family: "Hanken Grotesk", sans-serif; color: #27272A; }
.hero { background: #161618; color: #fff; padding: 40px 56px 36px; }
.hero h1 { margin: 0 0 8px; font-size: 40px; font-weight: 800; letter-spacing: -0.03em; }
.hero p { margin: 0; font-size: 15px; color: rgba(255,255,255,0.55); }
.badge { font-family: "JetBrains Mono"; font-size: 12px; color: rgba(255,255,255,0.5); padding: 4px 12px; border: 1px solid rgba(255,255,255,0.18); border-radius: 999px; display: inline-block; margin-bottom: 18px; }
.content { max-width: 1180px; margin: 0 auto; padding: 40px 56px 80px; }
.section-label { display: flex; align-items: center; gap: 14px; margin: 40px 0 18px; }
.section-label span.num { font-family: "JetBrains Mono"; font-size: 12px; font-weight: 600; color: #9A9AA2; }
.section-label h2 { margin: 0; font-size: 12px; font-weight: 700; letter-spacing: 0.14em; text-transform: uppercase; color: #71717A; }
.section-label .line { flex: 1; height: 1px; background: #DCDCE0; }
.card { background: #fff; border: 1px solid #E4E4E8; border-radius: 20px; padding: 22px 24px; margin-bottom: 16px; box-shadow: 0 1px 2px rgba(24,24,27,0.03); }
.grid { display: flex; flex-wrap: wrap; gap: 10px; }
.icon-cell { display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 14px 12px; background: #F7F7F9; border-radius: 14px; width: 92px; cursor: pointer; transition: background 0.15s; }
.icon-cell:hover { background: #EDEDEF; }
.icon-cell svg { width: 32px; height: 32px; }
.icon-cell span { font-size: 10px; font-weight: 600; color: #52525B; text-align: center; line-height: 1.3; word-break: break-word; }
</style>
</head>
<body>
<div class="hero">
  <div class="badge">T3Lab UI Standard · v2.4 · 2026</div>
  <h1>SVG Icon Set</h1>
  <p>""")
parts.append(str(total))
parts.append(""" icons &middot; Monochrome stroke style &middot; 24x24 viewBox &middot; stroke-width 1.7 &middot; #18181B</p>
</div>
<div class="content">
""")

for i, (panel, icons) in enumerate(panels.items()):
    num = str(i + 1).zfill(2)
    parts.append(f'<div class="section-label"><span class="num">{num}</span><h2>{panel}</h2><span class="line"></span></div>\n')
    parts.append('<div class="card"><div class="grid">\n')
    for icon_file in icons:
        name = icon_file.replace(".svg", "")
        svg_path = os.path.join(SVG_DIR, panel, icon_file)
        with open(svg_path, "r", encoding="utf-8") as f:
            svg_content = f.read()
        parts.append(f'  <div class="icon-cell" title="{name}">{svg_content}<span>{name}</span></div>\n')
    parts.append("</div></div>\n")

parts.append("</div>\n</body>\n</html>")

with open(OUT, "w", encoding="utf-8") as f:
    f.write("".join(parts))

print(f"Preview created: {OUT}")
print(f"Total icons: {total}")
