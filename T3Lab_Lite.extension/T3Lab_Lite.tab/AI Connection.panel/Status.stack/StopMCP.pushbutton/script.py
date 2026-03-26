"""
Status Error Indicator
Shows that there is a connection error (non-clickable status indicator)
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
from ui.button_state import get_last_error

__title__ = 'Error'
__author__ = 'T3LabAI Team'
__doc__ = 'Status indicator - Connection error'

if __name__ == '__main__':
    # This is a status indicator, show error info when clicked
    try:
        last_error = get_last_error()
        if last_error:
            forms.alert(
                "Connection Error:\n\n"
                "{}".format(last_error),
                title="Error Status"
            )
        else:
            forms.alert(
                "No errors detected.\n\n"
                "Connection is working normally.",
                title="Error Status"
            )
    except Exception as e:
        forms.alert(
            "Unable to get error status\n\n"
            "Error: {}".format(str(e)),
            title="Status Error"
        )
