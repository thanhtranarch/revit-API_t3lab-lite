import os
import sys
import subprocess
import urllib.request
from pathlib import Path

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_and_import('PIL')
from PIL import Image, ImageDraw

def get_icon_mapping(folder_name):
    name = folder_name.lower()
    
    # Default fallback
    icon = "plugin"
    accent = "accent_blue"
    color = "#3B82F6"
    
    # Annotation
    if any(k in name for k in ["align", "distribute", "horizontal", "vertical", "snap"]):
        return "align-left", "accent_blue", "#3B82F6"
    if any(k in name for k in ["text", "tag", "dimtext", "annotation", "upperall"]):
        return "text", "accent_blue", "#3B82F6"
    if any(k in name for k in ["select", "cursor"]):
        return "cursor", "accent_blue", "#3B82F6"
    if any(k in name for k in ["dimension", "measure"]):
        return "ruler", "accent_blue", "#3B82F6"
        
    # Data
    if any(k in name for k in ["excel", "schedule"]):
        return "microsoft-excel", "accent_green", "#10B981"
    if any(k in name for k in ["parameter", "data", "value", "list"]):
        return "database", "accent_blue", "#3B82F6"
    if any(k in name for k in ["ifc", "export", "transfer", "batchout", "cadto"]):
        return "export", "accent_blue", "#3B82F6"
    if any(k in name for k in ["bcf", "feedback", "comment"]):
        return "chat-message", "accent_amber", "#F59E0B"
        
    # Project
    if any(k in name for k in ["room", "area", "tile"]):
        return "floor-plan", "accent_blue", "#3B82F6"
    if any(k in name for k in ["level", "grid", "datum"]):
        return "grid", "accent_blue", "#3B82F6"
    if any(k in name for k in ["family", "element", "category", "component"]):
        return "categorize", "accent_blue", "#3B82F6"
    if any(k in name for k in ["split", "cut", "profile"]):
        return "cut", "accent_red", "#EF4444"
    if any(k in name for k in ["join", "link", "workset", "central"]):
        return "link", "accent_blue", "#3B82F6"
        
    # Settings & Cleanup
    if any(k in name for k in ["clean", "purge", "delete", "remove"]):
        return "broom", "accent_red", "#EF4444"
    if any(k in name for k in ["warning", "error"]):
        return "warning", "accent_amber", "#F59E0B"
    if any(k in name for k in ["health", "check", "auditor"]):
        return "health-check", "accent_green", "#10B981"
    if any(k in name for k in ["style", "color", "graphic", "hatch", "theme", "pattern"]):
        return "paint-palette", "accent_blue", "#3B82F6"
    if any(k in name for k in ["location", "point", "foundation"]):
        return "map-marker", "accent_blue", "#3B82F6"
    
    # Views & Sheets
    if any(k in name for k in ["sheet", "view", "pdf", "page", "drafting", "hub"]):
        return "document", "accent_blue", "#3B82F6"
        
    # App-level
    if any(k in name for k in ["assistant", "smart", "auto", "ui", "t3lab"]):
        return "bot", "accent_blue", "#3B82F6"
    if any(k in name for k in ["manager", "control", "setting"]):
        return "settings", "accent_blue", "#3B82F6"
    if any(k in name for k in ["cloud", "b360", "acc ", "platform"]):
        return "cloud", "accent_blue", "#3B82F6"
        
    return icon, accent, color

base_url = "https://img.icons8.com/fluency-systems-regular/64/0F172A/"
target_dir = r"C:\Users\tran_tienthanh\OneDrive - Woh Hup (Private) Limited\Documents\T3Lab Architect\t3lab-revit-api\T3Lab.extension\T3Lab.tab"

count = 0
for root, dirs, files in os.walk(target_dir):
    for d in dirs:
        if d.endswith(('.pushbutton', '.urlbutton', '.pulldown')):
            tool_path = os.path.join(root, d)
            tool_name = d.split('.')[0]
            
            icon_name, accent, color = get_icon_mapping(tool_name)
            url = f"{base_url}{icon_name}.png"
            tmp_path = os.path.join(tool_path, f"tmp_{icon_name}.png")
            out_file = os.path.join(tool_path, "icon.png")
            
            print(f"Processing {tool_name} -> {icon_name}.png")
            
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            try:
                with urllib.request.urlopen(req) as response, open(tmp_path, 'wb') as f:
                    f.write(response.read())
            except Exception as e:
                print(f"Failed {icon_name}, falling back to plugin icon. ({e})")
                icon_name = "plugin"
                url = f"{base_url}{icon_name}.png"
                try:
                    with urllib.request.urlopen(req) as response, open(tmp_path, 'wb') as f:
                        f.write(response.read())
                except:
                    continue
            
            try:
                img = Image.open(tmp_path).convert("RGBA")
                draw = ImageDraw.Draw(img)
                # Universal badge signature for all icons in Lumina style
                draw.ellipse([46, 10, 54, 18], fill=color)
                
                img.save(out_file)
                os.remove(tmp_path)
                count += 1
            except Exception as e:
                print(f"Failed to process image for {tool_name}: {e}")
                
print(f"Successfully batch generated {count} icons!")
