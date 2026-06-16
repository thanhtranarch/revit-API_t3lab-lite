import json
import os

registry_path = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\lib\config\tool_registry.json"
tab_dir = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\T3Lab.tab"

with open(registry_path, "r") as f:
    data = json.load(f)

tools = data.get("tools", {})

# Scan tab_dir for all script.py files
script_map = {}
for root, dirs, files in os.walk(tab_dir):
    if "script.py" in files:
        script_path = os.path.join(root, "script.py")
        parent_folder = os.path.basename(root)
        script_map[parent_folder] = script_path

def slugify(s):
    # Strip spaces, dots, dashes, pushbuttons, pulldowns, etc., and lowercase
    s = s.lower()
    for clean in [".pushbutton", ".pulldown", ".stack", " "]:
        s = s.replace(clean, "")
    return s

# Create a slugified script map
slug_script_map = {slugify(k): (k, v) for k, v in script_map.items()}

# Update paths
updated_count = 0
for key, val in list(tools.items()):
    button_name = val.get("button", key)
    
    # Clean key name
    slug_button = slugify(button_name)
    
    if slug_button in slug_script_map:
        matched_folder, matched_path = slug_script_map[slug_button]
        old_path = val.get("script_path")
        old_button = val.get("button")
        
        # Update if changed
        if old_path != matched_path or old_button != matched_folder:
            val["script_path"] = matched_path
            val["button"] = matched_folder
            print("Updated '{}':".format(key))
            print("  Button: {} -> {}".format(old_button, matched_folder))
            print("  Path: {} -> {}".format(old_path, matched_path))
            updated_count += 1
    else:
        # If the file really doesn't exist anymore (like consolidated smart align tools), 
        # let's remove it from registry or keep it as-is. We should keep it for safety but warn.
        pass

data["tools"] = tools

with open(registry_path, "w") as f:
    json.dump(data, f, indent=2)

print("Total tools updated: {}".format(updated_count))
