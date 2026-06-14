# -*- coding: utf-8 -*-
"""Advanced View Manager with Sheet Manager Style UI
Enhanced view management with summary cards and modern layout
Copyright: Dang Quoc Truong (DQT)
"""

__title__ = "Advanced\nView Manager"
__author__ = "Dang Quoc Truong (DQT)"

import os
import sys

# Ensure lib directory is in sys.path
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from GUI.AdvancedViewManagerDialog import show_advanced_view_manager

if __name__ == "__main__":
    show_advanced_view_manager()