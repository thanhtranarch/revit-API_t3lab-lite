import os
from PIL import Image, ImageDraw, ImageFilter

tools = [
    r'T3Lab.extension\T3Lab.tab\Standards & Settings.panel\ModelAuditor.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Standards & Settings.panel\ManaLoca.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Standards & Settings.panel\ManaStyles.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Standards & Settings.panel\ManaWorkset.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Annotation & Select.panel\ManaAnno.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Annotation & Select.panel\Mana.stack\ManaDWG.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Annotation & Select.panel\Mana.stack\ManaAlign.pushbutton\icon.png',
    r'T3Lab.extension\T3Lab.tab\Annotation & Select.panel\Mana.stack\ManaSelect.pushbutton\icon.png'
]

size = 64
bg = Image.new('RGBA', (size, size), (0,0,0,0))
draw = ImageDraw.Draw(bg)
draw.ellipse([-10, -10, 40, 40], fill=(0, 180, 255, 255))
draw.ellipse([30, -10, 74, 40], fill=(255, 220, 0, 255))
draw.ellipse([-10, 20, 40, 74], fill=(255, 50, 100, 255))
draw.ellipse([20, 20, 74, 74], fill=(255, 80, 0, 255))
bg = bg.filter(ImageFilter.GaussianBlur(radius=15))
bg_rgb = bg.convert('RGB')

for path in tools:
    if not os.path.exists(path):
        print("Not found:", path)
        continue
        
    img = Image.open(path).convert("RGBA")
    img = img.resize((64, 64), Image.Resampling.LANCZOS)
    alpha = img.split()[-1]
    
    result = Image.new('RGBA', (size, size))
    result.paste(bg_rgb, (0,0))
    result.putalpha(alpha)
    
    result.save(path)
    print("Updated:", path)
