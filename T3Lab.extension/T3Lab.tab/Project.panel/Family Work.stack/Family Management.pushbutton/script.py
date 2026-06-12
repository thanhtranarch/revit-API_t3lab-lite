# -*- coding: utf-8 -*-
"""
Family Management

Batch rename families/types, modify case, customize prefix/suffix, and assign worksets.

Author: Tran Tien Thanh & Dang Quoc Truong
"""

__author__  = "Tran Tien Thanh & Dang Quoc Truong"
__title__   = "Family\nManagement"
__version__ = "1.1.0"

# IMPORT LIBRARIES
# ==============================================================================
import os
import sys
import traceback

from pyrevit import revit, script

# Setup environment paths
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
lib_dir = os.path.join(EXT_DIR, 'lib')

if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from GUI.FamilyManagementDialog import show_family_management

# DEFINE VARIABLES
# ==============================================================================
logger        = script.get_logger()
REVIT_VERSION = int(revit.doc.Application.VersionNumber)

# MAIN SCRIPT
# ==============================================================================
if __name__ == '__main__':
    try:
        show_family_management()
    except Exception as ex:
        logger.error("Error running Family Management: {}".format(ex))
        logger.error(traceback.format_exc())
