"""
Start MCP Server Button
Starts the T3LabAI MCP server for AI communication
"""

from pyrevit import script, forms
from System.Threading import Thread, ThreadStart
import sys

# Add lib folder to path
lib_path = script.get_bundle_file('lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from core.server import get_t3labai_server
from config.settings import get_settings

__title__ = 'Start\nMCP'
__author__ = 'T3LabAI Team'
__doc__ = 'Start the T3LabAI MCP server for AI communication'

def start_server_thread():
    """Start server in background thread"""
    settings = get_settings()
    server_config = settings.get_server_config()
    port = server_config.get("port", 8080)

    server = get_t3labai_server()
    server.port = port
    server.start_server()


if __name__ == '__main__':
    output = script.get_output()

    try:
        server = get_t3labai_server()

        if server.is_running:
            forms.alert(
                "< T3LabAI MCP Server is already running!\n\n"
                "Port: {}\n"
                "Clients connected: {}".format(server.port, server.client_count),
                title="Server Already Running"
            )
        else:
            # Start server in background thread
            server_thread = Thread(ThreadStart(start_server_thread))
            server_thread.IsBackground = True
            server_thread.Start()

            output.print_md("## < T3LabAI MCP Server Started")
            output.print_md("**Port:** {}".format(server.port))
            output.print_md("**Status:** Waiting for AI connections...")
            output.print_md("")
            output.print_md(" Server is running in the background")
            output.print_md("")
            output.print_md("### Next Steps:")
            output.print_md("1. Open Claude Desktop or your MCP client")
            output.print_md("2. Verify the hammer icon appears")
            output.print_md("3. Start using AI commands with Revit!")

            forms.alert(
                "< T3LabAI MCP Server Started!\n\n"
                "Port: {}\n"
                "Ready for AI connections\n\n"
                "Check the output window for details.".format(server.port),
                title="Server Started"
            )

    except Exception as e:
        output.print_md("## L Error Starting Server")
        output.print_md("```\n{}\n```".format(str(e)))
        import traceback
        output.print_md("```\n{}\n```".format(traceback.format_exc()))

        forms.alert(
            "L Failed to start T3LabAI MCP Server\n\n"
            "Error: {}\n\n"
            "Check output window for details.".format(str(e)),
            title="Server Error"
        )
