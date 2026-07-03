# -*- coding: utf-8 -*-
"""
Sheet Manager
Unified tool to manage sheets, sets, views on sheets, parameters, and re-number sheets.

Copyright (c) 2026 T3Lab
"""
__title__ = "Sheet\nManager"
__author__ = "Dang Quoc Truong & Antigravity"

import os
import sys

# Ensure lib directory is in sys.path
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from GUI.ManaSheetsDialog import show_sheet_manager

if __name__ == '__main__':
    show_sheet_manager()
