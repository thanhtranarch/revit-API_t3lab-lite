import os
import sys
import subprocess
import urllib.request

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_and_import('PIL')
from PIL import Image, ImageDraw

base_url = "https://img.icons8.com/fluency-systems-regular/64/0F172A/"

icons = {
    "Feedback.pushbutton": {"icon": "chat-message", "accent": "notification_dot", "color": "#F59E0B"},
    "MCPControl.pushbutton": {"icon": "console", "accent": "terminal_prompt", "color": "#3B82F6"},
    "PDF import.pushbutton": {"icon": "pdf", "accent": "pdf_badge", "color": "#EF4444"},
    "T3LabAssistant.pushbutton": {"icon": "bot", "accent": "ai_glow", "color": "#3B82F6"},
    "CloudLinks.stack/Autodesk Forma.urlbutton": {"icon": "cloud", "accent": "acc_logo", "color": "#3B82F6"},
    "CloudLinks.stack/Autodesk Health.urlbutton": {"icon": "cloud", "accent": "health_tick", "color": "#10B981"},
    "CloudLinks.stack/Bluebeam Status.urlbutton": {"icon": "cloud", "accent": "health_tick", "color": "#10B981"}
}

output_base = r"C:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\T3Lab.extension\T3Lab.tab\Support.panel"

for rel_path, config in icons.items():
    icon_name = config["icon"]
    accent = config["accent"]
    color = config["color"]
    
    # Download
    url = f"{base_url}{icon_name}.png"
    tmp_path = f"tmp_{icon_name}.png"
    print(f"Downloading {icon_name}...")
    
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        with urllib.request.urlopen(req) as response, open(tmp_path, 'wb') as out_file:
            out_file.write(response.read())
    except Exception as e:
        print(f"Failed to download {icon_name} from Icons8: {e}")
        continue
        
    # Process image with Pillow
    img = Image.open(tmp_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    
    if accent == "notification_dot":
        # Amber dot top right
        draw.ellipse([46, 12, 54, 20], fill=color)
    elif accent == "terminal_prompt":
        # Draw a tiny blue prompt in the bottom
        draw.line([16, 42, 26, 42], fill=color, width=4)
    elif accent == "pdf_badge":
        # Red rectangle
        draw.rounded_rectangle([20, 44, 46, 56], radius=3, fill=color)
        draw.line([25, 50, 41, 50], fill="white", width=3)
    elif accent == "ai_glow":
        # Blue circle glow
        draw.ellipse([44, 10, 52, 18], fill=color)
    elif accent == "acc_logo":
        # Blue 'A' abstraction inside cloud
        draw.line([28, 42, 32, 32, 36, 42], fill=color, width=3)
        draw.line([30, 38, 34, 38], fill=color, width=3)
    elif accent == "health_tick":
        # Green tick badge
        draw.ellipse([36, 36, 52, 52], fill="#F8FAFC")
        draw.ellipse([38, 38, 50, 50], fill=color)
        draw.line([42, 44, 44, 46, 47, 41], fill="white", width=2)
        
    out_dir = os.path.join(output_base, rel_path)
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "icon.png")
    
    img.save(out_file)
    print(f"Saved customized icon to {out_file}")
    
    os.remove(tmp_path)

print("Icon processing completed!")
