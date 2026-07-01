#!/usr/bin/env python3
"""
Generate a self-contained HTML preview of all PNG icons, 
embedding them as base64 data URIs so it works without a server.
"""
import os
import base64

TAB_DIR = os.path.join("T3Lab.extension", "T3Lab.tab")
OUT = os.path.join("scripts", "png_preview_standalone.html")

panels = {}

for root, dirs, files in os.walk(TAB_DIR):
    dir_name = os.path.basename(root)
    if dir_name.endswith((".pushbutton", ".urlbutton", ".pulldown")):
        rel_parts = os.path.relpath(root, TAB_DIR).split(os.sep)
        panel = rel_parts[0]

        has_light = "icon.png" in files
        has_dark  = "icon.dark.png" in files

        if has_light or has_dark:
            if panel not in panels:
                panels[panel] = []

            def b64(path):
                try:
                    with open(path, "rb") as f:
                        return "data:image/png;base64," + base64.b64encode(f.read()).decode()
                except:
                    return ""

            light_b64 = b64(os.path.join(root, "icon.png"))      if has_light else ""
            dark_b64  = b64(os.path.join(root, "icon.dark.png")) if has_dark  else ""

            panels[panel].append({
                "name":  dir_name.rsplit(".", 1)[0],
                "light": light_b64,
                "dark":  dark_b64,
            })

# ── HTML ──────────────────────────────────────────────────────────────────────
html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>T3Lab Icon Preview — Swiss Style</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;600;700;800&family=JetBrains+Mono:wght@500&display=swap" rel="stylesheet">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { background: #F0F0F2; font-family: "Hanken Grotesk", system-ui, sans-serif; color: #1A1A2E; }

/* ── Hero ── */
.hero {
  background: #0D1B2A;
  padding: 52px 64px 44px;
  border-bottom: 4px solid #D07818;
}
.badge {
  display: inline-block;
  font-family: "JetBrains Mono", monospace;
  font-size: 11px; font-weight: 500;
  color: #D07818;
  border: 1.5px solid rgba(208,120,24,.45);
  border-radius: 999px;
  padding: 4px 14px;
  margin-bottom: 20px;
  letter-spacing: .08em;
  text-transform: uppercase;
}
.hero h1 { font-size: 40px; font-weight: 800; color: #F8FAFC; letter-spacing: -.03em; margin-bottom: 8px; }
.hero p  { font-size: 14px; color: #64748B; }

/* ── Content ── */
.content { max-width: 1300px; margin: 0 auto; padding: 48px 64px 96px; }

/* ── Section label ── */
.section-label {
  display: flex; align-items: center; gap: 14px;
  margin: 48px 0 20px;
}
.section-label .num {
  font-family: "JetBrains Mono", monospace;
  font-size: 12px; font-weight: 500;
  color: #94A3B8;
}
.section-label h2 {
  font-size: 11px; font-weight: 800;
  letter-spacing: .14em;
  text-transform: uppercase;
  color: #64748B;
}
.section-label .rule { flex: 1; height: 1px; background: #CBD5E1; }

/* ── Grid ── */
.grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 16px;
}

/* ── Card ── */
.card {
  background: #fff;
  border: 1px solid #E2E8F0;
  border-radius: 18px;
  padding: 20px;
  display: flex; flex-direction: column; gap: 14px;
  box-shadow: 0 2px 8px rgba(0,0,0,.04);
  transition: box-shadow .2s;
}
.card:hover { box-shadow: 0 6px 20px rgba(0,0,0,.08); }

.card-name {
  font-size: 15px; font-weight: 700;
  color: #0F172A;
}

/* ── Previews ── */
.previews { display: flex; gap: 12px; }
.preview {
  flex: 1; display: flex; flex-direction: column;
  align-items: center; gap: 8px;
  padding: 16px 8px 12px;
  border-radius: 12px;
  border: 1px solid transparent;
}
.preview.light { background: #F8FAFC; border-color: #E2E8F0; }
.preview.dark  { background: #0F172A; border-color: #1E293B; }

.preview img { width: 64px; height: 64px; object-fit: contain; }
.preview .lbl {
  font-family: "JetBrains Mono", monospace;
  font-size: 9px; font-weight: 500;
  letter-spacing: .08em; text-transform: uppercase;
}
.preview.light .lbl { color: #94A3B8; }
.preview.dark  .lbl { color: #475569; }

.na {
  width: 64px; height: 64px;
  display: flex; align-items: center; justify-content: center;
  font-size: 10px; color: #94A3B8;
  border: 1px dashed currentColor; border-radius: 8px;
}
</style>
</head>
<body>

<div class="hero">
  <div class="badge">T3Lab Revit Extension · Swiss / Minimal Style</div>
  <h1>Ribbon Icon Preview</h1>
  <p>All active ribbon icons — flat geometry, strict 2-colour palette (Navy #182A3E + Amber #D07818).</p>
</div>

<div class="content">
"""

for i, panel_name in enumerate(sorted(panels)):
    icons = sorted(panels[panel_name], key=lambda x: x["name"])
    num = str(i + 1).zfill(2)
    html += f"""
  <div class="section-label">
    <span class="num">{num}</span>
    <h2>{panel_name.replace('.panel','')}</h2>
    <span class="rule"></span>
  </div>
  <div class="grid">
"""
    for ic in icons:
        light_img = f'<img src="{ic["light"]}" alt="{ic["name"]} light">' if ic["light"] else '<div class="na">N/A</div>'
        dark_img  = f'<img src="{ic["dark"]}"  alt="{ic["name"]} dark">'  if ic["dark"]  else '<div class="na">N/A</div>'
        html += f"""    <div class="card">
      <div class="card-name">{ic["name"]}</div>
      <div class="previews">
        <div class="preview light">{light_img}<span class="lbl">Light</span></div>
        <div class="preview dark">{dark_img}<span class="lbl">Dark</span></div>
      </div>
    </div>
"""
    html += "  </div>\n"

html += "\n</div>\n</body>\n</html>\n"

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)

total = sum(len(v) for v in panels.values())
print(f"Standalone preview created: {OUT}")
print(f"Total icons: {total}")
