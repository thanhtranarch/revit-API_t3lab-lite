"""
Settings Button for T3LabAI MCP
Opens the configuration dialog for API keys and modules
"""

from pyrevit import script
import sys
import os

# Add lib folder to path - find extension lib folder
script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from ui.settings_dialog import show_settings_dialog

__title__ = 'Settings'
__author__ = 'T3LabAI Team'
__doc__ = 'Configure MCP API keys and modules'

if __name__ == '__main__':
    show_settings_dialog()
