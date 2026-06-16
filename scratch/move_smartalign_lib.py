import os
import shutil

ext_dir = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension"
src_smartalign = os.path.join(ext_dir, "T3Lab.tab", "Annotation.panel", "SmartAlign.stack", "Lib", "smartalign")
dst_smartalign = os.path.join(ext_dir, "lib", "smartalign")

if os.path.exists(src_smartalign):
    if os.path.exists(dst_smartalign):
        shutil.rmtree(dst_smartalign)
    shutil.move(src_smartalign, dst_smartalign)
    print("Successfully moved smartalign package to lib/smartalign")
else:
    print("Source smartalign folder not found at:", src_smartalign)
