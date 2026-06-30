import os
from PIL import Image, ImageDraw, ImageFilter

size = 64
img = Image.new('RGBA', (size, size), (0, 0, 0, 0))

aura = Image.new('RGBA', (size, size), (0, 0, 0, 0))
draw = ImageDraw.Draw(aura)

r = 18
w = 8
c = 32

draw.arc([c-r, c-r-3, c+r, c+r-3], start=180, end=360, fill=(0, 180, 255, 200), width=w)
draw.arc([c-r+3, c-r, c+r+3, c+r], start=270, end=450, fill=(255, 220, 0, 200), width=w)
draw.arc([c-r, c-r+3, c+r, c+r+3], start=0, end=180, fill=(255, 80, 0, 200), width=w)
draw.arc([c-r-3, c-r, c+r-3, c+r], start=90, end=270, fill=(255, 50, 100, 200), width=w)

aura = aura.filter(ImageFilter.GaussianBlur(radius=3))

dark = Image.new('RGBA', (size, size), (0, 0, 0, 0))
draw_dark = ImageDraw.Draw(dark)
draw_dark.ellipse([c-r+1, c-r+1, c+r-1, c+r-1], outline=(10, 10, 15, 230), width=4)
dark = dark.filter(ImageFilter.GaussianBlur(radius=1.5))

img = Image.alpha_composite(img, aura)
img = Image.alpha_composite(img, dark)

mask = Image.new('L', (size, size), 255)
draw_mask = ImageDraw.Draw(mask)
hole_r = r - 5
draw_mask.ellipse([c-hole_r, c-hole_r, c+hole_r, c+hole_r], fill=0)
mask = mask.filter(ImageFilter.GaussianBlur(radius=1.5))

result = Image.new('RGBA', (size, size))
for y in range(size):
    for x in range(size):
        r_p, g_p, b_p, a_p = img.getpixel((x, y))
        m = mask.getpixel((x, y))
        new_a = int(a_p * (m / 255.0))
        result.putpixel((x, y), (r_p, g_p, b_p, new_a))

output_path = r"T3Lab.extension\T3Lab.tab\Support.panel\T3LabAssistant.pushbutton\icon.png"
result.save(output_path)
print("Icon saved!")
