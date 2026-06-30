import os
from PIL import Image, ImageDraw, ImageFilter

path = r'T3Lab.extension\T3Lab.tab\Annotation & Select.panel\ManaAnno.pushbutton\icon.png'

def colorize_alpha(alpha_img, color):
    color_layer = Image.new('RGBA', alpha_img.size, color)
    color_layer.putalpha(alpha_img)
    return color_layer

dx_list = [0, 4, -4, 0]
dy_list = [-4, 0, 0, 4]
colors = [
    (0, 180, 255, 200),
    (255, 220, 0, 200),
    (255, 50, 100, 200),
    (255, 80, 0, 200)
]

size = 64

if os.path.exists(path):
    img = Image.open(path).convert("RGBA")
    img = img.resize((64, 64), Image.Resampling.LANCZOS)
    alpha = img.split()[-1]
    
    base = Image.new('RGBA', (size, size), (0,0,0,0))
    aura = Image.new('RGBA', (size, size), (0,0,0,0))

    for dx, dy, col in zip(dx_list, dy_list, colors):
        shifted_alpha = Image.new('L', (size, size), 0)
        shifted_alpha.paste(alpha, (dx, dy))
        col_img = colorize_alpha(shifted_alpha, col)
        aura = Image.alpha_composite(aura, col_img)

    aura = aura.filter(ImageFilter.GaussianBlur(radius=3))
    dark_core = colorize_alpha(alpha, (20, 20, 20, 240))

    result = Image.alpha_composite(base, aura)
    result = Image.alpha_composite(result, dark_core)
    
    result.save(path)
    print("Updated:", path)
