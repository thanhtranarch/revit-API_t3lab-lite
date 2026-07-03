# -*- coding: utf-8 -*-
"""MCP Control

Unified control panel for the T3Lab MCP server.
Start / Stop the server and manage connection settings in one dialog.
"""
__title__ = "MCP\nControl"
__author__ = "T3Lab & Dang Quoc Truong"

import os
import sys

# Path setup — script.py lives 3 levels below T3Lab.extension/
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from GUI.MCPControlDialog import show_mcp_control_dialog

def main():
    show_mcp_control_dialog()

if __name__ == '__main__':
    main()
