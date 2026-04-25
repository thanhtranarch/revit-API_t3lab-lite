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
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==================================================
import os
import sys
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')

from pyrevit import revit, script

# Path setup — 4 levels up: script.py → pushbutton → stack → panel → tab → extension
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

from GUI.PropertyLineDialog import show_property_line_dialog

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()
output = script.get_output()
REVIT_VERSION = int(revit.doc.Application.VersionNumber)

# CLASS/FUNCTIONS
# ==================================================

# MAIN SCRIPT
# ==================================================

if __name__ == '__main__':
    if not revit.doc:
        from pyrevit import forms
        forms.alert("Please open a Revit document first.", exitscript=True)
    try:
        show_property_line_dialog()
    except Exception as ex:
        logger.error("Property Line Tool error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
