import os
from PIL import Image, ImageDraw, ImageFilter, ImageChops

source_path = r'T3Lab.extension\T3Lab.tab\Standards & Settings.panel\ModelAuditor.pushbutton\icon.png'
dest_path = r'C:\Users\tran_tienthanh\.gemini\antigravity\brain\d88865cb-91e9-4c4b-a1d2-a7e48cf693c9\demo_model_auditor_blend.png'

if not os.path.exists(source_path):
    print("Source not found")
    exit()

img = Image.open(source_path).convert("RGBA")
img = img.resize((64, 64), Image.Resampling.LANCZOS)
alpha = img.split()[-1]

size = 64
bg = Image.new('RGBA', (size, size), (0,0,0,0))
draw = ImageDraw.Draw(bg)

draw.ellipse([-10, -10, 40, 40], fill=(0, 180, 255, 255))
draw.ellipse([30, -10, 74, 40], fill=(255, 220, 0, 255))
draw.ellipse([-10, 20, 40, 74], fill=(255, 50, 100, 255))
draw.ellipse([20, 20, 74, 74], fill=(255, 80, 0, 255))

bg = bg.filter(ImageFilter.GaussianBlur(radius=15))

# Convert original to grayscale then multiply
gray = img.convert('L').convert('RGB')
bg_rgb = bg.convert('RGB')

# Wait, the icons have white backgrounds? No, they are transparent. But their colors are dark.
# If they are very dark (e.g., black), multiply will just keep them black!
# Let's see if we should invert them or use Screen blending, or just use the gradient as a color overlay.
# To colorize a dark icon, using Screen or Additive is better, or we can use the grayscale as a brightness map.
# If we do "blend color on icon", maybe a pure gradient overlay keeping the original alpha mask is the cleanest.
# The user wants "blend color on icon", which means replacing the color of the icon with the gradient, 
# while keeping the transparency (alpha).
# What about the internal transparent details? They are encoded in the alpha mask.
# Let's try to just use the gradient masked by the alpha channel! This makes a "flat" icon with a gradient fill.
# Wait, some icons have multiple opacities or internal lines drawn in white. 
# If we want to keep internal lines (which are white or light grey), we can multiply the inverted grayscale?

# Let's just create a version that maps the icon's lightness to the gradient.
# Actually, mapping lightness to gradient:
# pixel = gradient_pixel * (grayscale_pixel / 255)
# Wait, if the icon is dark, grayscale is 0, so pixel becomes black. 
# To make a dark icon colorful, we should use the inverse:
# If the icon is mostly dark blue/black, maybe the user just wants a solid gradient silhouette!
# Let's do a pure gradient silhouette masked by alpha first.

result = Image.new('RGBA', (size, size))
result.paste(bg_rgb, (0,0))
result.putalpha(alpha)

# If the icon had inner lines that were transparent, they are preserved by the alpha!
# Let's save this as the primary demo.
result.save(dest_path)
print("Demo icon saved at", dest_path)
