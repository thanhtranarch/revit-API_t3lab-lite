import os
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM

panel = os.path.join(os.path.dirname(__file__), "..", "T3Lab.extension", "T3Lab.tab", "Views & Sheets.panel")
tools = ["ManaViews.pushbutton", "ManaSheets.pushbutton", "BatchOut.pushbutton", "SheetGen.pushbutton"]

for tool in tools:
    svg = os.path.join(panel, tool, "icon.svg")
    png = os.path.join(panel, tool, "icon.png")
    drawing = svg2rlg(svg)
    scale = 64.0 / drawing.width
    drawing.width = 64
    drawing.height = 64
    drawing.transform = (scale, 0, 0, scale, 0, 0)
    renderPM.drawToFile(drawing, png, fmt="PNG", bg=0xFFFFFF00)
    print("OK  " + tool)

print("Done.")
