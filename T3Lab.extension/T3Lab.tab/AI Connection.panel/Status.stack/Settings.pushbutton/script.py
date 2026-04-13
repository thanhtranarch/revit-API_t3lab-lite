# -*- coding: utf-8 -*-
"""
AI Settings

Configure T3Lab AI connection settings and API keys.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "AI Settings"

from pyrevit import script
import sys
import os

# Add lib folder to path - find extension lib folder
script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from ui.settings_dialog import show_settings_dialog

if __name__ == '__main__':
    show_settings_dialog()
