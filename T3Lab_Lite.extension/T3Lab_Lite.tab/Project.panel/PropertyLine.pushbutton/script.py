# -*- coding: utf-8 -*-
"""
Property Line

Create and manage property lines from survey data.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Property Line"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import os
import sys

# Add the lib directory to sys.path so GUI module is importable
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

from GUI.PropertyLineDialog import show_property_line_dialog
from pyrevit import script

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
# ==================================================

if __name__ == '__main__':
    logger = script.get_logger()
    try:
        show_property_line_dialog()
    except Exception as ex:
        logger.error("Property Line Tool error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
