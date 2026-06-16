import json
import os

registry_path = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\lib\config\tool_registry.json"

with open(registry_path, "r") as f:
    data = json.load(f)

tools = data.get("tools", {})
cleaned_count = 0
keys_to_delete = []

for key, val in tools.items():
    if "script_path" in val:
        path = os.path.normpath(val["script_path"])
        if not os.path.exists(path):
            keys_to_delete.append(key)

for key in keys_to_delete:
    del tools[key]
    print("Removed stale registry key:", key)
    cleaned_count += 1

data["tools"] = tools

with open(registry_path, "w") as f:
    json.dump(data, f, indent=2)

print("Total stale keys removed: {}".format(cleaned_count))
