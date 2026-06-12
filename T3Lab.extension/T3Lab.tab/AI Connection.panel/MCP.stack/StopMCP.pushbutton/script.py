# -*- coding: utf-8 -*-
"""Stop MCP Server

Stop the running MCP server inside Revit and free the socket port.
"""
__title__ = "Stop MCP"
__author__ = "T3Lab & Dang Quoc Truong"

import os
import sys

# Setup paths
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import script, forms
try:
    from core.server import get_t3labai_server
    HAS_SERVER = True
except Exception as e:
    forms.alert("Error loading MCP server: {}".format(e))
    HAS_SERVER = False

logger = script.get_logger()

def main():
    if not HAS_SERVER:
        return
    
    server = get_t3labai_server()
    if not server.is_running:
        forms.alert("MCP Server is not currently running.", title="MCP Server Info", warn=False)
        return

    if server.stop_server():
        logger.info("MCP Server stopped successfully.")
        forms.alert("MCP Server stopped successfully and port is freed.", title="MCP Server Stopped", warn=False)
    else:
        logger.error("Failed to stop MCP Server.")
        forms.alert("Failed to stop MCP Server cleanly.", title="MCP Server Error")

if __name__ == '__main__':
    main()
