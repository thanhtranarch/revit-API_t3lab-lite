# -*- coding: utf-8 -*-
"""
Supabase Configuration
Configure connection to Supabase for family sync
"""
__title__ = "Supabase\nConfig"
__author__ = "T3Lab"
__doc__ = "Configure Supabase connection for syncing Revit families"

import sys
import os

# Add lib path
lib_path = os.path.join(
    os.path.dirname(__file__),
    "..", "..", "..",
    "lib"
)
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Import and show config dialog
try:
    from GUI.SupabaseConfigDialog import show_supabase_config
    show_supabase_config()
except Exception as ex:
    from pyrevit import forms
    forms.alert("Error: {}".format(ex))
