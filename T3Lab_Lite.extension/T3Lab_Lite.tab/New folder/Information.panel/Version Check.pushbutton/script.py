# -*- coding: utf-8 -*-
"""
Check Python Engine in pyRevit
Detect whether pyRevit is using IronPython or CPython

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
--------------------------------------------------------
"""
__title__ = "Check Python Engine"
__author__ = "Tran Tien Thanh"
__version__ = 'Version: 1.0'

# IMPORT LIBRARIES
# ==================================================
import sys
import platform

# DEFINE VARIABLES
# ==================================================
engine_name = platform.python_implementation()
engine_version = sys.version

# CLASS/FUNCTIONS
# ==================================================
def check_engine():
    print("Python Engine:", engine_name)
    print("Python Version:", engine_version)
    if "IronPython" in engine_name:
        print("Running on IronPython (Default for pyRevit <= 4.x)")
    elif "CPython" in engine_name:
        print("Running on CPython (pyRevitLabs with .NET 7/8 support)")
    else:
        print("Unknown Python Engine")

# MAIN SCRIPT
# ==================================================
check_engine()
