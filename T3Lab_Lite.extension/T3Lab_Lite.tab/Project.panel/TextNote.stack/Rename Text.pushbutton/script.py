# -*- coding: utf-8 -*-
"""
Rename TextNoteType
Auto rename TextNote Types:
LB_<discipline>_<textsize>_<textfont>_<textbackground>_<textfactor>_<textcolor>[_Border][_B][_U][_I]

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Ren TextType"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.3'

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, script

# DEFINE VARIABLES
# ==================================================
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# CLASS/FUNCTIONS
# ==================================================
standard_colors = {
    (255, 0, 0): "Red",
    (0, 255, 0): "Lime",
    (0, 0, 255): "Blue",
    (255, 255, 0): "Yellow",
    (0, 255, 255): "Cyan",
    (255, 0, 255): "Magenta",
    (0, 0, 0): "Black",
    (255, 255, 255): "White",
    (128, 128, 128): "Gray",
    (128, 0, 0): "Maroon",
    (0, 128, 0): "Green",
    (0, 0, 128): "Navy",
    (128, 128, 0): "Olive",
    (0, 128, 128): "Teal",
    (128, 0, 128): "Purple",
    (255, 128, 0): "Orange",
    (128, 128, 255): "LightBlue",
    (192, 192, 192): "Silver",
}

def get_textsize_only(txt_type):
    try:
        param = txt_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if param:
            raw = param.AsDouble()
            mm_val = round(raw * 304.8, 2)
            return "{:.2f}mm".format(mm_val)
    except Exception as e:
        print("    [ERROR] TEXT_SIZE: {}".format(e))
    return "N/A"

def get_textfont_only(txt_type):
    try:
        param = txt_type.get_Parameter(BuiltInParameter.TEXT_FONT)
        if param:
            return param.AsString().replace(" ", "")
    except Exception as e:
        print("    [ERROR] TEXT_FONT: {}".format(e))
    return "N/A"

def get_textbackground_only(txt_type):
    try:
        param = txt_type.get_Parameter(BuiltInParameter.TEXT_BACKGROUND)
        if param:
            val = param.AsInteger()
            return "Opaque" if val == 0 else "Transparent"
    except Exception as e:
        print("    [ERROR] TEXT_BACKGROUND: {}".format(e))
    return "N/A"

def get_textfactor_only(txt_type):
    try:
        param = txt_type.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE)
        if param:
            return str(round(param.AsDouble(), 2))
    except Exception as e:
        print("    [ERROR] TEXT_WIDTH_SCALE: {}".format(e))
    return "N/A"

def get_textcolor_only(txt_type):
    try:
        param = txt_type.get_Parameter(BuiltInParameter.LINE_COLOR)
        if param:
            color_int = param.AsInteger()
            r = color_int & 255
            g = (color_int >> 8) & 255
            b = (color_int >> 16) & 255
            return standard_colors.get((r, g, b), "RGB")
    except Exception as e:
        print("    [ERROR] TEXT_COLOR: {}".format(e))
    return "N/A"

def check_show_border(txt_type):
    try:
        param = txt_type.get_Parameter(BuiltInParameter.TEXT_BOX_VISIBILITY)
        if param and param.AsInteger() == 1:
            return True
    except Exception as e:
        print("    [ERROR] TEXT_BOX_VISIBILITY: {}".format(e))
    return False

def check_text_formatting(txt_type):
    """
    Check for Bold, Underline, and Italic formatting
    Returns tuple: (is_bold, is_underline, is_italic)
    """
    is_bold = False
    is_underline = False
    is_italic = False
    
    try:
        # Check for Bold
        bold_param = txt_type.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD)
        if bold_param and bold_param.AsInteger() == 1:
            is_bold = True
    except Exception as e:
        print("    [ERROR] TEXT_STYLE_BOLD: {}".format(e))
    
    try:
        # Check for Underline
        underline_param = txt_type.get_Parameter(BuiltInParameter.TEXT_STYLE_UNDERLINE)
        if underline_param and underline_param.AsInteger() == 1:
            is_underline = True
    except Exception as e:
        print("    [ERROR] TEXT_STYLE_UNDERLINE: {}".format(e))
    
    try:
        # Check for Italic
        italic_param = txt_type.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC)
        if italic_param and italic_param.AsInteger() == 1:
            is_italic = True
    except Exception as e:
        print("    [ERROR] TEXT_STYLE_ITALIC: {}".format(e))
    
    return (is_bold, is_underline, is_italic)

def build_texttype_name(txt_type, origin_name):
    discipline = "STR" if "STR" in origin_name.upper() else "ARC"
    textsize = get_textsize_only(txt_type)
    textfont = get_textfont_only(txt_type)
    textbg = get_textbackground_only(txt_type)
    textfactor = get_textfactor_only(txt_type)
    textcolor = get_textcolor_only(txt_type)

    name_parts = ["LB", discipline, textsize, textfont, textbg, textfactor]

    # Add color if not black
    if textcolor != "Black":
        name_parts.append(textcolor)

    # Add border if visible
    if check_show_border(txt_type):
        name_parts.append("Border")

    # Check for text formatting and add to the end
    is_bold, is_underline, is_italic = check_text_formatting(txt_type)
    
    if is_bold:
        name_parts.append("B")
    if is_underline:
        name_parts.append("U")
    if is_italic:
        name_parts.append("I")

    return "_".join(name_parts)

# MAIN SCRIPT
# ==================================================
transaction = Transaction(doc, "Rename TextNoteTypes")
transaction.Start()

try:
    textnote_types = FilteredElementCollector(doc).OfClass(TextNoteType).WhereElementIsElementType().ToElements()
    for txt_type in textnote_types:
        origin_name = txt_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        new_name = build_texttype_name(txt_type, origin_name)
        txt_type.Name = new_name
        print("{} >> {}".format(origin_name, new_name))
except Exception as e:
    print("Transaction failed: {}".format(e))
finally:
    transaction.Commit()

print("Done.")