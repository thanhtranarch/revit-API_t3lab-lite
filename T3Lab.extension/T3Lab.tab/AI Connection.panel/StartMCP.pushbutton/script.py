# -*- coding: utf-8 -*-
"""Start MCP Server

Start the local MCP server inside Revit for AI-to-Revit communication.
"""
__title__ = "Start MCP"
__author__ = "T3Lab & Dang Quoc Truong"

import os
import sys

# Setup paths
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
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
    if server.is_running:
        forms.alert("MCP Server is already running on port {}".format(server.port), title="MCP Server Info", warn=False)
        return

    # Start the server
    if server.start_server():
        logger.info("MCP Server started successfully on port {}".format(server.port))
        
        # Display instructions and config path
        bridge_path = os.path.join(LIB_DIR, "core", "bridge.py").replace("\\", "/")
        config_snippet = {
            "t3lab-revit": {
                "command": "python",
                "args": [bridge_path, str(server.port)]
            }
        }
        
        msg = "MCP Server started successfully on port {}.\n\n".format(server.port)
        msg += "Add this to your Claude Desktop config:\n\n"
        msg += '"t3lab-revit": {\n'
        msg += '  "command": "python",\n'
        msg += '  "args": ["{}", "{}"]\n'.format(bridge_path, server.port)
        msg += '}'
        
        forms.alert(msg, title="MCP Server Started", warn=False)
    else:
        logger.error("Failed to start MCP Server.")
        forms.alert("Failed to start MCP Server. Check port availability.", title="MCP Server Error")

if __name__ == '__main__':
    main()
