"""
Status Connected Indicator
Shows that the server is connected (non-clickable status indicator)
"""

from pyrevit import script, forms
import sys
import os

# Add lib folder to path - find extension lib folder
script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from core.server import get_t3labai_server

__title__ = 'Connected'
__author__ = 'T3LabAI Team'
__doc__ = 'Status indicator - Server is connected'

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
