# -*- coding: utf-8 -*-
"""
Load Family Tool
Load Revit families from folders with category organization
"""
__title__ = "Load Family"
__author__ = "T3Lab"
__doc__ = """Load Revit families from folders with category organization.

Features:
- Select folder to browse families
- Categories organized by folder structure
- Search functionality for quick filtering
- Preview family thumbnails
- Batch loading of multiple families
"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
import os
import sys

# Add lib directory to path
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Import the dialog
from GUI.FamilyLoaderDialog import show_family_loader

# pyRevit Imports
from pyrevit import script

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#====================================================================================================

if __name__ == '__main__':
    logger = script.get_logger()

    try:
        # Show the family loader dialog
        loaded_families = show_family_loader()

        if loaded_families:
            logger.info("Successfully loaded {} families".format(len(loaded_families)))
        else:
            logger.info("No families were loaded")

    except Exception as ex:
        logger.error("Error in Load Family tool: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
