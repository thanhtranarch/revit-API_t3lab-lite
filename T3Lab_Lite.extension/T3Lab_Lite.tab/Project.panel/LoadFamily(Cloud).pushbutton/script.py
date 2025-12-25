# -*- coding: utf-8 -*-
"""
Load Family Cloud Tool
Load Revit families from Vercel cloud API
"""
__title__ = "Load Family (Cloud)"
__author__ = "T3Lab"
__doc__ = """Load Revit families from Vercel cloud API.

Features:
- Cloud-based family library hosted on Vercel
- Categories organized by API metadata
- Search functionality for quick filtering
- Preview family thumbnails
- Batch loading with automatic download
- Vercel deployment protection bypass support
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

# Import the cloud dialog
from GUI.FamilyLoaderCloudDialog import show_family_loader_cloud

# pyRevit Imports
from pyrevit import script

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
#====================================================================================================

if __name__ == '__main__':
    logger = script.get_logger()

    try:
        # Show the family loader cloud dialog
        loaded_families = show_family_loader_cloud()

        if loaded_families:
            logger.info("Successfully loaded {} families from cloud".format(len(loaded_families)))
        else:
            logger.info("No families were loaded from cloud")

    except Exception as ex:
        logger.error("Error in Load Family Cloud tool: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
