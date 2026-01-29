# -*- coding: utf-8 -*-
"""
Advanced Purge - Power User Operations
Main entry point

⚠️ WARNING: This tool contains dangerous operations!
Always preview before executing.

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__title__ = "Advanced\nPurge"
__author__ = "Dang Quoc Truong (DQT)"

from pyrevit import forms, script

# Get current document
doc = __revit__.ActiveUIDocument.Document

# Check if document is valid
if not doc:
    forms.alert("No active document!", title="Advanced Purge Error", exitscript=True)

# Show warning dialog first
warning_msg = (
    "⚠️ ADVANCED PURGE - POWER USER TOOL ⚠️\n\n"
    "This tool contains DANGEROUS operations that can:\n"
    "- Delete large amounts of data\n"
    "- Modify worksets\n"
    "- Remove critical model elements\n\n"
    "ALWAYS:\n"
    "✓ Create a backup first\n"
    "✓ Use Dry Run mode\n"
    "✓ Preview before executing\n"
    "✓ Understand what you're deleting\n\n"
    "Do you want to continue?"
)

result = forms.alert(
    warning_msg,
    title="⚠️ Advanced Purge Warning",
    warn_icon=True,
    yes=True,
    no=True
)

if not result:
    script.exit()

# Import main window with error handling
try:
    import sys
    import os
    
    # Add lib to path
    lib_path = os.path.join(os.path.dirname(__file__), 'lib')
    if lib_path not in sys.path:
        sys.path.insert(0, lib_path)
    
    # Import window
    from advanced_purge_window import AdvancedPurgeWindow
    
    # Show window
    window = AdvancedPurgeWindow(doc)
    window.ShowDialog()
    
except Exception as ex:
    import traceback
    error_msg = "Error loading Advanced Purge:\n\n"
    error_msg += str(ex) + "\n\n"
    error_msg += "Traceback:\n"
    error_msg += traceback.format_exc()
    
    forms.alert(error_msg, title="Advanced Purge Error")
    print("="*80)
    print("ADVANCED PURGE ERROR:")
    print("="*80)
    traceback.print_exc()
    print("="*80)
