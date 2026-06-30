import os
from PIL import Image, ImageDraw
import math

NAVY = (24, 48, 76, 255) # Dark Navy Blue
ORANGE = (243, 140, 18, 255) # Vibrant Orange
TRANS = (0, 0, 0, 0)
SIZE = 96

out_dir = "draft_icons"

def new_img():
    return Image.new('RGBA', (SIZE, SIZE), TRANS)

# 1. T3Lab Assistant
img = new_img()
draw = ImageDraw.Draw(img)
draw.ellipse([16, 16, 80, 80], outline=NAVY, width=6)
draw.ellipse([38, 38, 58, 58], fill=NAVY) # eye
draw.polygon([(70, 10), (75, 20), (85, 25), (75, 30), (70, 40), (65, 30), (55, 25), (65, 20)], fill=ORANGE)
img.save(os.path.join(out_dir, "T3LabAssistant.png"))

# 2. MCP Control
img = new_img()
draw = ImageDraw.Draw(img)
for y in [16, 40, 64]:
    draw.rounded_rectangle([16, y, 80, y+16], radius=4, outline=NAVY, width=4)
draw.ellipse([64, 68, 72, 76], fill=ORANGE)
img.save(os.path.join(out_dir, "MCPControl.png"))

# 3. Feedback
img = new_img()
draw = ImageDraw.Draw(img)
draw.rounded_rectangle([12, 20, 84, 68], radius=8, outline=NAVY, width=6)
draw.polygon([(40, 68), (56, 68), (40, 84)], fill=NAVY) # Tail
draw.polygon([(48, 30), (52, 42), (64, 42), (54, 50), (58, 62), (48, 54), (38, 62), (42, 50), (32, 42), (44, 42)], fill=ORANGE)
img.save(os.path.join(out_dir, "Feedback.png"))

# 4. ManaTabs
img = new_img()
draw = ImageDraw.Draw(img)
draw.rounded_rectangle([12, 32, 84, 80], radius=4, outline=NAVY, width=4)
draw.rounded_rectangle([16, 16, 44, 32], radius=4, fill=ORANGE)
draw.rounded_rectangle([48, 16, 76, 32], radius=4, outline=NAVY, width=4)
img.save(os.path.join(out_dir, "ManaTabs.png"))

# 5. Ribbon Names
img = new_img()
draw = ImageDraw.Draw(img)
draw.rounded_rectangle([8, 24, 88, 72], radius=6, outline=NAVY, width=4)
draw.line([24, 48, 72, 48], fill=NAVY, width=4) # text line
draw.rounded_rectangle([16, 40, 24, 56], radius=2, fill=ORANGE)
img.save(os.path.join(out_dir, "RibbonNames.png"))

# 6. BG Theme
img = new_img()
draw = ImageDraw.Draw(img)
draw.chord([16, 16, 80, 80], 90, 270, fill=NAVY) # Left half moon
draw.chord([16, 16, 80, 80], 270, 90, fill=ORANGE) # Right half sun
img.save(os.path.join(out_dir, "BGTheme.png"))

# 7. IFC-SG Suite
img = new_img()
draw = ImageDraw.Draw(img)
cx, cy, r = 48, 48, 40
hex_pts = [(cx + r * math.sin(i * math.pi/3), cy - r * math.cos(i * math.pi/3)) for i in range(6)]
draw.polygon(hex_pts, outline=NAVY, width=4)
draw.polygon([(48, 24), (32, 40), (32, 64), (64, 64), (64, 40)], outline=NAVY, width=4)
draw.ellipse([64, 64, 80, 80], fill=ORANGE)
img.save(os.path.join(out_dir, "IFCSGSuite.png"))

# 8. BCF Reader
img = new_img()
draw = ImageDraw.Draw(img)
draw.polygon([(24, 16), (56, 16), (72, 32), (72, 80), (24, 80)], outline=NAVY, width=4)
draw.line([(56, 16), (56, 32), (72, 32)], fill=NAVY, width=4) # fold
draw.rounded_rectangle([52, 56, 84, 76], radius=4, fill=ORANGE)
draw.polygon([(60, 76), (68, 76), (60, 84)], fill=ORANGE)
img.save(os.path.join(out_dir, "BCFReader.png"))

# 9. Foundation Volume
img = new_img()
draw = ImageDraw.Draw(img)
draw.polygon([(40, 24), (56, 24), (56, 64), (80, 64), (80, 80), (16, 80), (16, 64), (40, 64)], fill=NAVY)
draw.line([(16, 88), (80, 88)], fill=ORANGE, width=4)
draw.line([(16, 84), (16, 92)], fill=ORANGE, width=4)
draw.line([(80, 84), (80, 92)], fill=ORANGE, width=4)
img.save(os.path.join(out_dir, "FoundationVolume.png"))

print("Icons successfully generated to draft_icons folder.")
