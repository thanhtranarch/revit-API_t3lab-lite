"""
Disconnect Button
Disconnects from the T3LabAI MCP server
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
from ui.button_state import update_connection_buttons

__title__ = 'Disconnect'
__author__ = 'T3LabAI Team'
__doc__ = 'Disconnect from the T3LabAI MCP server'

if __name__ == '__main__':
    output = script.get_output()

    try:
        server = get_t3labai_server()

        if not server.is_running:
            forms.alert(
                "T3LabAI MCP Server is not connected.",
                title="Not Connected"
            )
        else:
            stats = server.get_server_stats()
            server.stop_server()

            output.print_md("## T3LabAI MCP Server Disconnected")
            output.print_md("**Total clients served:** {}".format(stats['total_clients']))
            output.print_md("**Commands processed:** {}".format(stats['commands_processed']))
            output.print_md("")
            output.print_md("Server stopped successfully")

            # Update button states
            update_connection_buttons(connected=False)

            forms.alert(
                "T3LabAI MCP Server Disconnected\n\n"
                "Total clients: {}\n"
                "Commands processed: {}\n\n"
                "Server shut down successfully.".format(
                    stats['total_clients'],
                    stats['commands_processed']
                ),
                title="Disconnected"
            )

    except Exception as e:
        output.print_md("## Error Disconnecting")
        output.print_md("```\n{}\n```".format(str(e)))
        import traceback
        output.print_md("```\n{}\n```".format(traceback.format_exc()))

        forms.alert(
            "Error disconnecting from server\n\n"
            "Error: {}".format(str(e)),
            title="Disconnect Error"
        )
