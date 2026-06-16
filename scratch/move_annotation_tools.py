import os
import shutil

panel_dir = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\T3Lab.tab\Annotation.panel"

# 1. Ensure target Graphics.stack folder exists
graphics_stack = os.path.join(panel_dir, "Graphics.stack")
if not os.path.exists(graphics_stack):
    os.makedirs(graphics_stack)
    print("Created Graphics.stack directory:", graphics_stack)

# 2. Move items to Graphics.stack
moves = [
    (os.path.join(panel_dir, "Graphic.stack", "AutoDimension.pushbutton"), 
     os.path.join(graphics_stack, "AutoDimension.pushbutton")),
    (os.path.join(panel_dir, "Graphic.stack", "Snap Dimension.pushbutton"), 
     os.path.join(graphics_stack, "SnapDimension.pushbutton")),
    (os.path.join(panel_dir, "Graphic 2.stack", "Reset Overrides.pushbutton"), 
     os.path.join(graphics_stack, "ResetOverrides.pushbutton")),
]

for src, dst in moves:
    if os.path.exists(src):
        if os.path.exists(dst):
            shutil.rmtree(dst)
        shutil.move(src, dst)
        print("Moved {} to {}".format(src, dst))
    else:
        print("Source not found:", src)

# 3. Move Grids.pulldown to Annotation.panel root
src_grids = os.path.join(panel_dir, "Graphic 2.stack", "Grids.pulldown")
dst_grids = os.path.join(panel_dir, "Grids.pulldown")
if os.path.exists(src_grids):
    if os.path.exists(dst_grids):
        shutil.rmtree(dst_grids)
    shutil.move(src_grids, dst_grids)
    print("Moved Grids.pulldown to Annotation.panel root")
else:
    print("Grids.pulldown not found:", src_grids)

# 4. Clean up old stacks
folders_to_delete = [
    os.path.join(panel_dir, "Graphic.stack"),
    os.path.join(panel_dir, "Graphic 2.stack"),
    os.path.join(panel_dir, "SmartAlign.stack"),
]

for folder in folders_to_delete:
    if os.path.exists(folder):
        shutil.rmtree(folder)
        print("Removed old directory:", folder)
