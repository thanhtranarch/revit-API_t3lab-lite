# -*- coding: utf-8 -*-
"""
PDF Import

Import PDF pages into Revit views sequentially.
Opens a dialog to pick a PDF and map each page to a target view.
Page 1 → View 1, Page 2 → View 2, etc.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__title__   = "PDF\nImport"
__author__  = "Tran Tien Thanh"
__version__ = "2.0.0"

# IMPORTS
# ==============================================================================
import os
import sys

extension_dir = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
)
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# MAIN
# ==============================================================================
if __name__ == '__main__':
    from GUI.PDFImportDialog import show_pdf_import
    show_pdf_import()
