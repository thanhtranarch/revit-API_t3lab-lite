# -*- coding: utf-8 -*-
"""
Start MCP

Start the MCP server for AI-assisted Revit automation.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Start MCP"

from pyrevit import script, forms
import sys
import os

# Add lib folder to path - find extension lib folder
script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from core.server import get_t3labai_server

if __name__ == '__main__':
    # This is a status indicator, just show current status when clicked
    try:
        server = get_t3labai_server()
        if server.is_running:
            forms.alert(
                "Server Status: CONNECTED\n\n"
                "Port: {}\n"
                "Clients: {}".format(server.port, server.client_count),
                title="Connection Status"
            )
        else:
            forms.alert(
                "Server Status: NOT CONNECTED\n\n"
                "Click 'Connect' to start the server.",
                title="Connection Status"
            )
    except Exception as e:
        forms.alert(
            "Unable to get server status\n\n"
            "Error: {}".format(str(e)),
            title="Status Error"
        )
