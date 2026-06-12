# -*- coding: utf-8 -*-
"""
Sheet Manage - Main Script
CLEANED - Direct to Sheet List

Copyright (c) Dang Quoc Truong (DQT)
"""

import sys
import os

# Add lib path
script_dir = os.path.dirname(__file__)
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Revit imports
from Autodesk.Revit.DB import Transaction
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')  # For file dialogs

# Get document
doc = __revit__.ActiveUIDocument.Document

try:
    # Import core modules
    from core.revit_service import RevitService
    from core.data_models import ChangeTracker
    from core.main_window import MainWindow
    
    # Initialize services
    revit_service = RevitService(doc)
    change_tracker = ChangeTracker()
    
    # Create and show main window
    window = MainWindow(doc, revit_service, change_tracker)
    window.ShowDialog()
    
except Exception as e:
    print("Error: {}".format(str(e)))
    import traceback
    traceback.print_exc()
