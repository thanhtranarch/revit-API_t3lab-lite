"""
Stop MCP Server Button
Stops the T3LabAI MCP server
"""

from pyrevit import script, forms
import sys

# Add lib folder to path
lib_path = script.get_bundle_file('lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from core.server import get_t3labai_server

__title__ = 'Stop\nMCP'
__author__ = 'T3LabAI Team'
__doc__ = 'Stop the T3LabAI MCP server'

if __name__ == '__main__':
    output = script.get_output()

    try:
        server = get_t3labai_server()

        if not server.is_running:
            forms.alert(
                "< T3LabAI MCP Server is not running.",
                title="Server Not Running"
            )
        else:
            stats = server.get_server_stats()
            server.stop_server()

            output.print_md("## =� T3LabAI MCP Server Stopped")
            output.print_md("**Total clients served:** {}".format(stats['total_clients']))
            output.print_md("**Commands processed:** {}".format(stats['commands_processed']))
            output.print_md("")
            output.print_md(" Server stopped successfully")

            forms.alert(
                "=� T3LabAI MCP Server Stopped\n\n"
                "Total clients: {}\n"
                "Commands processed: {}\n\n"
                "Server shut down successfully.".format(
                    stats['total_clients'],
                    stats['commands_processed']
                ),
                title="Server Stopped"
            )

    except Exception as e:
        output.print_md("## L Error Stopping Server")
        output.print_md("```\n{}\n```".format(str(e)))
        import traceback
        output.print_md("```\n{}\n```".format(traceback.format_exc()))

        forms.alert(
            "L Error stopping server\n\n"
            "Error: {}".format(str(e)),
            title="Server Error"
        )
