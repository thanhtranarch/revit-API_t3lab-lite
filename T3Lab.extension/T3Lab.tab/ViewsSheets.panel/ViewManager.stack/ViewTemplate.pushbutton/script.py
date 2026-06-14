# -*- coding: utf-8 -*-
"""
View Template Manager (Launcher)
Manage View Templates with beautiful UI - Sheet Manager Style

Compatible with Revit 2024 / 2025 / 2026
Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""
__title__ = "View Template\nManage"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Manage View Templates - By DQT"

import os
import sys

# Ensure lib directory is in sys.path
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from GUI.ViewTemplateDialog import show_view_template_manager

if __name__ == '__main__':
    show_view_template_manager()