import os
import shutil

src_root = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\T3Lab.tab\Settings.panel"

# 1. Ensure target stack folder exists
stack_dir = os.path.join(src_root, "Settings.stack")
if not os.path.exists(stack_dir):
    os.makedirs(stack_dir)
    print("Created stack directory:", stack_dir)

# 2. Move files for Settings.stack
stack_moves = [
    ("ParaManager.pushbutton", "ParaManager.pushbutton"),
    ("ColorSplasher.pushbutton", "ColorSplasher.pushbutton"),
    ("StyleManager.pushbutton", "StyleManager.pushbutton"),
]

for src_name, dst_name in stack_moves:
    src_path = os.path.join(src_root, src_name)
    dst_path = os.path.join(stack_dir, dst_name)
    if os.path.exists(src_path):
        shutil.move(src_path, dst_path)
        print("Moved stack item {} to {}".format(src_name, dst_path))
    else:
        print("Source path not found for stack item:", src_path)

# 3. Move/Rename cleanup tools
cleanup_pulldown = os.path.join(src_root, "ModelCleanup.pulldown")
cleanup_moves = [
    ("CleanupManager.pushbutton", os.path.join(src_root, "CleanupManager.pushbutton")),
    ("SmartPurge.pushbutton", os.path.join(src_root, "SmartPurge")),
    ("AdvancedPurge.pushbutton", os.path.join(src_root, "AdvancedPurge")),
    ("SmartDelete.pushbutton", os.path.join(src_root, "SmartDelete")),
]

for src_name, dst_path in cleanup_moves:
    src_path = os.path.join(cleanup_pulldown, src_name)
    if os.path.exists(src_path):
        if os.path.exists(dst_path):
            shutil.rmtree(dst_path)
        shutil.move(src_path, dst_path)
        print("Moved cleanup item {} to {}".format(src_name, dst_path))
    else:
        print("Source path not found for cleanup item:", src_path)

# 4. Move/Rename health tools
health_pulldown = os.path.join(src_root, "ModelHealth.pulldown")
health_moves = [
    ("LocationManager.pushbutton", os.path.join(src_root, "LocationManager.pushbutton")),
    ("HealthCheck.pushbutton", os.path.join(src_root, "HealthCheck")),
    ("ModelChecker.pushbutton", os.path.join(src_root, "ModelChecker")),
    ("Warning.pushbutton", os.path.join(src_root, "Warning")),
    ("InPlaceModel.pushbutton", os.path.join(src_root, "InPlaceModel")),
    ("Material List.pushbutton", os.path.join(src_root, "MaterialList")),
]

for src_name, dst_path in health_moves:
    src_path = os.path.join(health_pulldown, src_name)
    if os.path.exists(src_path):
        if os.path.exists(dst_path):
            shutil.rmtree(dst_path)
        shutil.move(src_path, dst_path)
        print("Moved health item {} to {}".format(src_name, dst_path))
    else:
        print("Source path not found for health item:", src_path)

# 5. Remove pulldown folders if empty/unused
for pulldown in [cleanup_pulldown, health_pulldown]:
    if os.path.exists(pulldown):
        shutil.rmtree(pulldown)
        print("Removed pulldown directory:", pulldown)
