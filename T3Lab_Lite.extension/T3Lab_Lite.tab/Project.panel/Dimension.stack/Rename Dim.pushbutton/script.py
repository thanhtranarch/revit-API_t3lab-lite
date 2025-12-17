# -*- coding: utf-8 -*-

"""
Rename Dimmension
Auto rename Dimension Types
LB_<discipline>_<textsize>_<textfont>_<textbackground>_<color>_<dimsuffix>

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Ren Dim"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 3.0'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
import re

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# CLASS/FUNCTIONS
# ==================================================
def sanitize_string(value):
    """
    Remove illegal characters from a string (e.g. for naming types).
    """
    if not value:
        return "N/A"
    # Remove illegal characters: \ / : * ? " < > | and leading/trailing spaces
    return re.sub(r'[\\/:?"<>|=]', '', value).strip()



def get_textsize_only(dim_type):
    try:
        param = dim_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if param:
            raw = param.AsDouble()
            mm_val = round(raw * 304.8, 2)
            return "{:.2f}mm".format(mm_val)
    except Exception as e:
        print("    [ERROR] TEXT_SIZE: {}".format(e))
    return "N/A"

def get_textfont_only(dim_type):
    try:
        param = dim_type.get_Parameter(BuiltInParameter.TEXT_FONT)
        if param:
            return param.AsString()
    except Exception as e:
        print("    [ERROR] TEXT_FONT: {}".format(e))
    return "N/A"

def get_textbackground_only(dim_type):
    try:
        param = dim_type.get_Parameter(BuiltInParameter.DIM_TEXT_BACKGROUND)
        if param:
            return param.AsValueString()
    except Exception as e:
        print("    [ERROR] TEXT_BACKGROUND: {}".format(e))
    return "N/A"

def get_textcolor_only(dim_type):
    def extract_rgb_from_int(color_int):
        r = color_int & 255
        g = (color_int >> 8) & 255
        b = (color_int >> 16) & 255
        return (r, g, b)

    standard_colors = {
        (255, 128, 128): "Light Coral",
        (255, 255, 128): "Light Yellow",
        (128, 255, 128): "Pale Green",
        (128, 255, 255): "Pale Cyan",
        (128, 128, 255): "Light Slate Blue",
        (255, 128, 255): "Orchid",
        (255, 0, 0): "Red",
        (255, 255, 0): "Yellow",
        (0, 255, 0): "Lime",
        (0, 255, 255): "Cyan",
        (0, 0, 255): "Blue",
        (255, 0, 255): "Magenta",
        (128, 64, 64): "Brown",
        (255, 192, 128): "Light Salmon",
        (128, 255, 192): "Aquamarine",
        (192, 192, 255): "Lavender",
        (192, 128, 255): "Medium Orchid",
        (128, 0, 0): "Maroon",
        (255, 128, 0): "Orange",
        (0, 128, 0): "Green",
        (0, 128, 128): "Teal",
        (0, 0, 128): "Navy",
        (128, 0, 128): "Purple",
        (128, 64, 0): "Saddle Brown",
        (192, 128, 64): "Peru",
        (0, 128, 64): "Dark Sea Green",
        (0, 128, 192): "Steel Blue",
        (64, 128, 255): "Dodger Blue",
        (128, 0, 192): "Dark Orchid",
        (0, 0, 0): "Black",
        (128, 128, 0): "Olive",
        (128, 128, 128): "Gray128",
        (0, 192, 192): "Medium Turquoise",
        (192, 192, 192): "Silver",
        (255, 255, 255): "White",
        (255, 128, 64): "Orange",
        (0, 128, 255): "Blue",
        (70, 70, 70): "Gray70",
        (128, 0, 64): "Dark Raspberry",
        (77, 77, 77): "Gray77"
    }
    try:
        param = dim_type.get_Parameter(BuiltInParameter.LINE_COLOR)
        if param:
            rgb = extract_rgb_from_int(param.AsInteger())
            return standard_colors.get(rgb, "RGB")
    except Exception as e:
        print("    [ERROR] LINE_COLOR: {}".format(e))

    return "N/A"

def get_dim_prefix_only(dim_type):
    try:
        param = dim_type.get_Parameter(BuiltInParameter.DIM_PREFIX)
        if param:
            value = param.AsString()
            clearvalue = sanitize_string(value)
            return clearvalue if clearvalue else "N/A"
    except Exception as e:
        print("    [ERROR] DIM_PREFIX: {}".format(e))
    return "N/A"

def get_centerline_symbol_status(dim_type):
    """
    Check if the DimensionType has a Centerline Symbol.
    Returns 'Center' if assigned, otherwise 'N/A'.
    """
    try:
        param = dim_type.get_Parameter(BuiltInParameter.DIM_STYLE_CENTERLINE_SYMBOL)
        if param:
            element_id = param.AsElementId()
            if element_id != ElementId.InvalidElementId:
                return "Center"
    except Exception as e:
        print("    [ERROR] DIM_STYLE_CENTERLINE_SYMBOL: {}".format(e))
    return "N/A"
# def get_bold_italic_status(dim_type):
#     try:
#         is_bold = dim_type.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD)
#         is_italic = dim_type.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC)

#         bold = is_bold.AsInteger() == 1 if is_bold else False
#         italic = is_italic.AsInteger() == 1 if is_italic else False

#         if bold and italic:
#             return "BoldItalic"
#         elif bold:
#             return "Bold"
#         elif italic:
#             return "Italic"
#         else:
#             return "None"
#     except Exception as e:
#         print("    [ERROR] TEXT_STYLE_BOLD/ITALIC: {}".format(e))
#         return "N/A"

def get_top_indicator_only(spot_type):
    try:
        param = spot_type.get_Parameter(BuiltInParameter.SPOT_ELEV_IND_TOP)
        if param:
            value = param.AsString()
            clearvalue = sanitize_string(value)
            return clearvalue if clearvalue else "N/A"
    except Exception as e:
        print("    [ERROR] SPOT_ELEV_IND_TOP: {}".format(e))
    return "N/A"

def get_bottom_indicator_only(spot_type):
    try:
        param = spot_type.get_Parameter(BuiltInParameter.SPOT_ELEV_IND_BOTTOM)
        if param:
            value = param.AsString()
            clearvalue = sanitize_string(value)
            return clearvalue if clearvalue else "N/A"
    except Exception as e:
        print("    [ERROR] SPOT_ELEV_IND_BOTTOM: {}".format(e))
    return "N/A"

def get_elevation_indicator_only(spot_type):
    try:
        param = spot_type.get_Parameter(BuiltInParameter.SPOT_ELEV_IND_ELEVATION)
        
        if param:
            value = param.AsString()
            clearvalue = sanitize_string(value)
            return clearvalue if clearvalue else "N/A"
    except Exception as e:
        print("    [ERROR] SPOT_ELEV_IND_ELEVATION: {}".format(e))
    return "N/A"



def build_dimtype_name(dim_type,origin_name):
    discipline = "STR" if "STR" in origin_name.upper() else "ARC"
    textsize = get_textsize_only(dim_type)
    textfont = get_textfont_only(dim_type)
    textbg = get_textbackground_only(dim_type)
    textcolor = get_textcolor_only(dim_type)
    dimprefix = get_dim_prefix_only(dim_type)
    dimcenter = get_centerline_symbol_status(dim_type)
    topindicator = get_top_indicator_only(dim_type)
    bottomindicator = get_bottom_indicator_only(dim_type)
    elevationindicator = get_elevation_indicator_only(dim_type)
    # dimstyle = get_bold_italic_status(dim_type)

    name_parts = ["LB",discipline, textsize, textfont, textbg]

    # if dimstyle != "N/A":
    #     name_parts.append(dimstyle)  

    if textcolor != "Black":
        name_parts.append(textcolor)
    
    if dimcenter != "N/A":
        name_parts.append(dimcenter)

    if dimprefix != "N/A":
        name_parts.append(dimprefix)
        
    if elevationindicator != "N/A":
        name_parts.append(elevationindicator) 
    else:   
        if topindicator != "N/A":
            name_parts.append(topindicator)
        
        if bottomindicator != "N/A":
            name_parts.append(bottomindicator)

    return "_".join(name_parts)



# MAIN SCRIPT
# ==================================================
renamed_count = 0
transaction = Transaction(doc, "Rename Dimension Types")
transaction.Start()

try:
    dim_types = FilteredElementCollector(doc).OfClass(DimensionType).WhereElementIsElementType().ToElements()
    for dim_type in dim_types:
        try:
            origin_name = dim_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            result_name = build_dimtype_name(dim_type,origin_name)
            dim_type.Name = result_name
            print("{} >> {}".format(origin_name, result_name))
        except Exception as e:
            print("{} >> erorrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr".format(e))
except Exception as outer_err:
    print("Transaction failed: {}".format(outer_err))
finally:
    transaction.Commit()

print("Done.")