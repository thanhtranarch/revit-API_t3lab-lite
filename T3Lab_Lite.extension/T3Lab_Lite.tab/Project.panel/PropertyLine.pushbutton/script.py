# -*- coding: utf-8 -*-
"""
Property Line Tool
Create US property boundary lines in Revit from Lightbox RE parcel API data.

Usage:
1. Enter your Lightbox API key and save it
2. Type a US property address and click Search
3. Select a parcel from the results
4. Configure elevation and line type
5. Click "Create Property Lines in Revit"
"""
__title__ = "Property\nLine"
__author__ = "T3Lab"
__doc__ = """Property Line Tool - Powered by Lightbox RE API

Create accurate US property boundary lines directly in Revit by
searching for any US address. The tool retrieves official parcel
data (GeoJSON boundaries) from the Lightbox RE API and converts
geographic coordinates (WGS84) into Revit internal feet.

Features:
  - Address-based parcel search (Lightbox RE API)
  - Automatic WGS84 -> Revit feet coordinate conversion
  - Native PropertyLine, ModelLine, or DetailLine output
  - Placement relative to Project Base Point, Survey Point, or World Origin
  - API key stored locally in ~/.t3lab/property_line_config.json

Requirements:
  - Lightbox RE API key (https://lightboxre.com)
  - Active Revit document
"""

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
