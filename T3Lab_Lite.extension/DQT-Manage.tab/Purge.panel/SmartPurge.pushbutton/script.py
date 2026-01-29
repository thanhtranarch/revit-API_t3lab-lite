# -*- coding: utf-8 -*-
"""
Smart Purge v2.0
Main entry point

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"
__title__ = "Smart Purge"

# Add lib to path
import sys
import os

script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Import and run
from smart_purge_window_v2 import SmartPurgeWindowV2

# Get current document
doc = __revit__.ActiveUIDocument.Document

# Create and show window
try:
    window = SmartPurgeWindowV2(doc)
    window.ShowDialog()
except Exception as e:
    from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
    MessageBox.Show(
        "Error opening Smart Purge:\n\n{}".format(str(e)),
        "Smart Purge Error",
        MessageBoxButton.OK,
        MessageBoxImage.Error
    )
    # Print full traceback for debugging
    import traceback
    print("=" * 70)
    print("SMART PURGE ERROR:")
    print("=" * 70)
    traceback.print_exc()
    print("=" * 70)