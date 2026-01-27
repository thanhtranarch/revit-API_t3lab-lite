"""
Status Not Starting Indicator
Shows that the server has not started (non-clickable status indicator)
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

__title__ = 'Not Starting'
__author__ = 'T3LabAI Team'
__doc__ = 'Status indicator - Server not started'

if __name__ == '__main__':
    # This is a status indicator, show startup status when clicked
    try:
        server = get_t3labai_server()
        if server.is_running:
            forms.alert(
                "Server has started successfully.\n\n"
                "Port: {}\n"
                "Status: Running".format(server.port),
                title="Startup Status"
            )
        else:
            forms.alert(
                "Server has NOT started.\n\n"
                "Click 'Connect' to start the server.",
                title="Startup Status"
            )
    except Exception as e:
        forms.alert(
            "Unable to get startup status\n\n"
            "Error: {}".format(str(e)),
            title="Status Error"
        )
