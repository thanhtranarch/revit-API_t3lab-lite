#!/usr/bin/env python3
"""
Convert all icon.svg files in the T3Lab extension to icon.png (96x96 px).
Uses svglib + reportlab for rasterisation.
After conversion, removes the icon.svg files so pyRevit loads cleanly.
"""

import os
import sys

# Needed for svglib
try:
    from svglib.svglib import svg2rlg
    from reportlab.graphics import renderPM
    from PIL import Image
    import io
    HAS_SVGLIB = True
except ImportError:
    HAS_SVGLIB = False

BASE = r"c:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\T3Lab.extension\T3Lab.tab"
SIZE = 96  # Revit large icon size in px

def convert_svg_to_png(svg_path, png_path, size=96):
    """Convert a single SVG file to PNG at the given pixel size."""
    drawing = svg2rlg(svg_path)
    if drawing is None:
        print(f"  WARNING: Could not parse {svg_path}")
        return False
    
    # Render to PNG bytes via reportlab
    buf = io.BytesIO()
    renderPM.drawToFile(drawing, buf, fmt="PNG", dpi=96)
    buf.seek(0)
    
    # Use PIL to resize to exact target size with high quality
    img = Image.open(buf).convert("RGBA")
    img = img.resize((size, size), Image.LANCZOS)
    img.save(png_path, "PNG", optimize=True)
    return True

def main():
    if not HAS_SVGLIB:
        print("ERROR: svglib/reportlab/Pillow not installed.")
        print("Run: python -m pip install svglib reportlab Pillow")
        sys.exit(1)
    
    svg_files = []
    for root, dirs, files in os.walk(BASE):
        for f in files:
            if f == "icon.svg":
                svg_files.append(os.path.join(root, f))
    
    print(f"\nFound {len(svg_files)} icon.svg files to convert\n")
    ok = 0
    fail = 0
    
    for svg_path in svg_files:
        png_path = os.path.join(os.path.dirname(svg_path), "icon.png")
        rel = os.path.relpath(svg_path, BASE)
        try:
            if convert_svg_to_png(svg_path, png_path, SIZE):
                print(f"  OK: {rel}")
                ok += 1
                # Remove the SVG so pyRevit doesn't get confused
                os.remove(svg_path)
            else:
                fail += 1
        except Exception as e:
            print(f"  FAIL: {rel} -> {e}")
            fail += 1
    
    print(f"\nDone: {ok} converted, {fail} failed.")
    print(f"PNG icons saved at {SIZE}x{SIZE}px.")

if __name__ == "__main__":
    main()
