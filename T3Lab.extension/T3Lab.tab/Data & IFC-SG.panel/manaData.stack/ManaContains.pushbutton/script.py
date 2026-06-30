# -*- coding: utf-8 -*-
"""
Contains Manager - Find elements in spatial containers or collect element data into spatial elements.
Copyright (c) 2026 Dang Quoc Truong (DQT)
"""

__title__ = "Contains\nManager"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Find elements in spatial containers or collect element data into spatial elements."

import os
import sys

# Define extension directory and add lib to sys.path
SCRIPT_DIR = os.path.dirname(__file__)
# Stacked pushbutton (tab/panel/stack/pushbutton) -> 4 levels up to extension root
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

from GUI.ManaContainsDialog import main

if __name__ == "__main__":
    main()
