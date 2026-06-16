import json
import os

registry_path = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\lib\config\tool_registry.json"

with open(registry_path, "r") as f:
    data = json.load(f)

tools = data.get("tools", {})

# Define the new tools to add
new_tools = {
    "ModelAuditor.pushbutton": {
        "button": "ModelAuditor.pushbutton",
        "panel": "Settings.panel",
        "title": "Model Auditor",
        "keywords": [
            "model auditor",
            "health check",
            "model checker",
            "warnings",
            "inplace model",
            "material list",
            "auditor"
        ],
        "intent": "open_model_auditor",
        "script_path": r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\T3Lab.tab\Settings.panel\ModelAuditor.pushbutton\script.py"
    },
    "SmartAlign.pushbutton": {
        "button": "SmartAlign.pushbutton",
        "panel": "Annotation.panel",
        "title": "Smart Align",
        "keywords": [
            "smart align",
            "align",
            "distribute",
            "horizontal",
            "vertical"
        ],
        "intent": "open_smart_align",
        "script_path": r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\T3Lab.tab\Annotation.panel\SmartAlign.pushbutton\script.py"
    },
    "FamilyBatchCreator.pushbutton": {
        "button": "FamilyBatchCreator.pushbutton",
        "panel": "Project.panel",
        "title": "Family Batch Creator",
        "keywords": [
            "family batch creator",
            "cad to family",
            "json to family",
            "batch",
            "family"
        ],
        "intent": "open_family_batch_creator",
        "script_path": r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\T3Lab.tab\Project.panel\FamilyBatchCreator.pushbutton\script.py"
    }
}

# Add or update the tools in the registry
added_count = 0
for key, val in new_tools.items():
    if key not in tools:
        tools[key] = val
        print("Added new tool:", key)
        added_count += 1
    else:
        # Update existing
        tools[key].update(val)
        print("Updated existing tool:", key)

data["tools"] = tools

with open(registry_path, "w") as f:
    json.dump(data, f, indent=2)

print("Total new tools added: {}".format(added_count))
