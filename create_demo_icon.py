import os
from PIL import Image, ImageDraw, ImageFilter

source_path = r'T3Lab.extension\T3Lab.tab\Standards & Settings.panel\ModelAuditor.pushbutton\icon.png'
dest_path = r'C:\Users\tran_tienthanh\.gemini\antigravity\brain\d88865cb-91e9-4c4b-a1d2-a7e48cf693c9\demo_model_auditor.png'

if not os.path.exists(source_path):
    print("Source not found")
    exit()

img = Image.open(source_path).convert("RGBA")
# Resize to 64x64 if not already
img = img.resize((64, 64), Image.Resampling.LANCZOS)

# Create an alpha mask from the original image
alpha = img.split()[-1]

# We will create an aura based on the shape of the alpha mask
size = 64
base = Image.new('RGBA', (size, size), (0,0,0,0))

# To make a colorful glowing shape, we can shift the alpha mask and colorize it
def colorize_alpha(alpha_img, color):
    color_layer = Image.new('RGBA', alpha_img.size, color)
    color_layer.putalpha(alpha_img)
    return color_layer

# Shift offsets
dx_list = [0, 4, -4, 0]
dy_list = [-4, 0, 0, 4]
colors = [
    (0, 180, 255, 200),  # Cyan top
    (255, 220, 0, 200),  # Yellow right
    (255, 50, 100, 200), # Red left
    (255, 80, 0, 200)    # Orange bottom
]

aura = Image.new('RGBA', (size, size), (0,0,0,0))

for dx, dy, col in zip(dx_list, dy_list, colors):
    shifted_alpha = Image.new('L', (size, size), 0)
    shifted_alpha.paste(alpha, (dx, dy))
    col_img = colorize_alpha(shifted_alpha, col)
    aura = Image.alpha_composite(aura, col_img)

# Blur the aura
aura = aura.filter(ImageFilter.GaussianBlur(radius=3))

# For the core, use the original alpha but filled with a very dark color
dark_core = colorize_alpha(alpha, (20, 20, 20, 240))

# Make the center of the dark core slightly eroded or hollow if we want, 
# but the tools have specific shapes (like a magnifying glass), so a solid dark silhouette is better to retain meaning.
# Let's composite dark core over aura
result = Image.alpha_composite(base, aura)
result = Image.alpha_composite(result, dark_core)

result.save(dest_path)
print("Demo icon saved at", dest_path)
