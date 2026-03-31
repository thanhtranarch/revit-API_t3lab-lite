# -*- coding: utf-8 -*-
"""
Rename Dimmension
Auto rename Dimension Types theo chuẩn LB_STR, LB_ARC...

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "ReName Dimmension"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.0'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Snippets._context_manager import ef_Transaction, try_except

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

# CLASS/FUNCTIONS
# ==================================================
def get_new_name(old_name):
    # Skip if already valid format
    parts = old_name.split("_")
    if len(parts) >= 5:
        if parts[0] in ["LB_ARC", "LB_STR"] and parts[2] == "Arial" and parts[3] == "Transparent":
            return None

    # Set base
    base = "LB_ARC"
    if "STR" in old_name.upper():
        base = "LB_STR"

    # Font size
    if "1.8mm" in old_name:
        size = "1.8mm"
    elif "2.0mm" in old_name:
        size = "2.0mm"
    elif "2.5mm" in old_name:
        size = "2.5mm"
    else:
        return None

    # Initial tag (color)
    tag = ""
    if "FURNITURE" in old_name.upper():
        tag = "Magenta"
    elif "BLUE" in old_name.upper():
        tag = "Blue"
    elif "GREEN" in old_name.upper():
        tag = "Green"
    elif "CENTER" in old_name.upper():
        tag = "Center"
    elif "TILING" in old_name.upper():
        tag = "Tiling"

    # Extra suffix (Checking, Red if not already handled)
    suffix = []
    if "CHECK" in old_name.upper():
        suffix.append("Checking")

    # If '*' in name, inject Red_* pattern if not already present
    has_star = "*" in old_name
    if has_star:
        # Ensure Red_* is explicitly included if not already
        if "Red_*" not in old_name:
            star_part = "Red_*"
        else:
            star_part = "Red_*"

        # Combine everything: Transparent + Red_* + [tag] + [suffix]
        tag_parts = [star_part]
        if tag:
            tag_parts.append(tag)
        if suffix:
            tag_parts.extend(suffix)
        full_tag = "_".join(tag_parts)
        return "{}_{}_{}_{}_{}".format(base, size, "Arial", "Transparent", full_tag)
    
    # Nếu không có *, xử lý như cũ
    if suffix:
        if tag:
            tag = "{}_{}".format(tag, "_".join(suffix))
        else:
            tag = "_".join(suffix)

    if tag:
        return "{}_{}_{}_{}_{}".format(base, size, "Arial", "Transparent", tag)
    else:
        return "{}_{}_{}_{}".format(base, size, "Arial", "Transparent")


# MAIN SCRIPT
# ==================================================
renamed_count = 0

with ef_Transaction(doc, "Auto Rename Dimension Types"):
    dim_types = FilteredElementCollector(doc).OfClass(DimensionType).WhereElementIsElementType().ToElements()

    for dim_type in dim_types:
        with try_except():
            current_name = dim_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            new_name = get_new_name(current_name)

            if new_name and new_name != current_name:
                dim_type.Name = new_name
                print("Renamed: {} → {}".format(current_name, new_name))
                renamed_count += 1

print("✅ Done. Total dimension types renamed: {}".format(renamed_count))
