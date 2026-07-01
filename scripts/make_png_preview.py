#!/usr/bin/env python3
import os

TAB_DIR = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\T3Lab.extension\T3Lab.tab"
OUT = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\scripts\png_preview.html"

panels = {}

# Scan all buttons
for root, dirs, files in os.walk(TAB_DIR):
    dir_name = os.path.basename(root)
    if dir_name.endswith(".pushbutton") or dir_name.endswith(".urlbutton") or dir_name.endswith(".pulldown"):
        # Get panel name
        rel_parts = os.path.relpath(root, TAB_DIR).split(os.sep)
        panel = rel_parts[0]
        
        # Check files
        has_light = "icon.png" in files
        has_dark = "icon.dark.png" in files
        
        if has_light or has_dark:
            if panel not in panels:
                panels[panel] = []
            panels[panel].append({
                "name": dir_name.split(".")[0],
                "folder": os.path.relpath(root, os.path.dirname(OUT)).replace(os.sep, "/"),
                "has_light": has_light,
                "has_dark": has_dark
            })

parts = []
parts.append("""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>T3Lab Icon Preview - Aura Glow & Flat Styles</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: #F1F1F4; font-family: "Hanken Grotesk", sans-serif; color: #1E293B; }
.hero { background: #0F172A; color: #fff; padding: 48px 56px 40px; border-bottom: 4px solid #F58220; }
.hero h1 { margin: 0 0 8px; font-size: 38px; font-weight: 800; letter-spacing: -0.03em; }
.hero p { margin: 0; font-size: 15px; color: #94A3B8; }
.badge { font-family: "JetBrains Mono"; font-size: 11px; color: #F58220; padding: 4px 14px; border: 1px solid rgba(245, 130, 32, 0.4); border-radius: 999px; display: inline-block; margin-bottom: 18px; font-weight: 600; text-transform: uppercase; }
.content { max-width: 1280px; margin: 0 auto; padding: 40px 56px 80px; }
.section-label { display: flex; align-items: center; gap: 14px; margin: 40px 0 18px; }
.section-label span.num { font-family: "JetBrains Mono"; font-size: 13px; font-weight: 700; color: #94A3B8; }
.section-label h2 { margin: 0; font-size: 13px; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; color: #64748B; }
.section-label .line { flex: 1; height: 1px; background: #CBD5E1; }
.grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; }
.card { background: #fff; border: 1px solid #E2E8F0; border-radius: 20px; padding: 20px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); display: flex; flex-direction: column; gap: 16px; }
.card-header { display: flex; flex-direction: column; gap: 4px; }
.card-title { font-size: 16px; font-weight: 700; color: #0F172A; }
.card-path { font-family: "JetBrains Mono"; font-size: 10px; color: #64748B; word-break: break-all; }
.previews { display: flex; gap: 16px; width: 100%; }
.preview-box { flex: 1; display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 16px 8px; border-radius: 12px; border: 1px solid #F1F5F9; }
.preview-box.light { background: #F8FAFC; }
.preview-box.dark { background: #1E293B; color: #F1F5F9; border-color: #334155; }
.preview-box img { width: 64px; height: 64px; object-fit: contain; }
.preview-label { font-size: 10px; font-weight: 700; letter-spacing: 0.05em; text-transform: uppercase; opacity: 0.6; }
</style>
</head>
<body>
<div class="hero">
  <div class="badge">T3Lab Revit Extension · Icon Dashboard</div>
  <h1>Active Ribbon Icons</h1>
  <p>Overview of updated PNG icons (Aura Glow and Standard Flat Styles) rendered inside the active button bundles.</p>
</div>
<div class="content">
""")

for i, panel in enumerate(sorted(panels.keys())):
    num = str(i + 1).zfill(2)
    parts.append(f'<div class="section-label"><span class="num">{num}</span><h2>{panel}</h2><span class="line"></span></div>\n')
    parts.append('<div class="grid">\n')
    for icon in sorted(panels[panel], key=lambda x: x["name"]):
        parts.append(f'  <div class="card">\n')
        parts.append(f'    <div class="card-header">\n')
        parts.append(f'      <div class="card-title">{icon["name"]}</div>\n')
        parts.append(f'      <div class="card-path">{icon["folder"].replace("../T3Lab.extension/T3Lab.tab/", "")}</div>\n')
        parts.append(f'    </div>\n')
        parts.append(f'    <div class="previews">\n')
        
        # Light Preview
        if icon["has_light"]:
            img_src = f'{icon["folder"]}/icon.png'
            parts.append(f'      <div class="preview-box light">\n')
            parts.append(f'        <img src="{img_src}">\n')
            parts.append(f'        <div class="preview-label">Light</div>\n')
            parts.append(f'      </div>\n')
        else:
            parts.append(f'      <div class="preview-box light" style="opacity: 0.3;">\n')
            parts.append(f'        <div style="width:64px;height:64px;display:flex;align-items:center;justify-content:center;font-size:10px;">N/A</div>\n')
            parts.append(f'        <div class="preview-label">Light</div>\n')
            parts.append(f'      </div>\n')
            
        # Dark Preview
        if icon["has_dark"]:
            img_src = f'{icon["folder"]}/icon.dark.png'
            parts.append(f'      <div class="preview-box dark">\n')
            parts.append(f'        <img src="{img_src}">\n')
            parts.append(f'        <div class="preview-label">Dark</div>\n')
            parts.append(f'      </div>\n')
        else:
            parts.append(f'      <div class="preview-box dark" style="opacity: 0.3;">\n')
            parts.append(f'        <div style="width:64px;height:64px;display:flex;align-items:center;justify-content:center;font-size:10px;">N/A</div>\n')
            parts.append(f'        <div class="preview-label">Dark</div>\n')
            parts.append(f'      </div>\n')
            
        parts.append(f'    </div>\n')
        parts.append(f'  </div>\n')
    parts.append("</div>\n")

parts.append("""
</div>
</body>
</html>
""")

with open(OUT, "w", encoding="utf-8") as f:
    f.write("".join(parts))

print(f"Preview created: {OUT}")
print(f"Total panels: {len(panels)}")
