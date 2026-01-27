"""
Connect Button
Connects to the T3LabAI MCP server for AI communication with Claude
"""

from pyrevit import script, forms
from System.Threading import Thread, ThreadStart
import sys
import os

# Add lib folder to path - find extension lib folder
script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_path = os.path.join(script_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from core.server import get_t3labai_server
from config.settings import get_settings
from ui.button_state import update_connection_buttons

__title__ = 'Connect'
__author__ = 'T3LabAI Team'
__doc__ = 'Connect to the T3LabAI MCP server for AI communication with Claude'


def start_server_thread():
    """Start server in background thread"""
    settings = get_settings()
    server_config = settings.get_server_config()
    port = server_config.get("port", 8080)

    server = get_t3labai_server()
    server.port = port
    server.start_server()


def get_claude_config_path():
    """Get Claude Desktop config path"""
    if sys.platform == 'win32':
        appdata = os.environ.get('APPDATA', '')
        return os.path.join(appdata, 'Claude', 'claude_desktop_config.json')
    elif sys.platform == 'darwin':
        return os.path.expanduser('~/Library/Application Support/Claude/claude_desktop_config.json')
    else:
        return os.path.expanduser('~/.config/claude/claude_desktop_config.json')


def get_mcp_config_snippet(port):
    """Generate MCP config snippet for Claude Desktop"""
    return '''{
  "mcpServers": {
    "t3labai-revit": {
      "url": "http://localhost:%d/sse"
    }
  }
}''' % port


if __name__ == '__main__':
    output = script.get_output()

    try:
        server = get_t3labai_server()
        settings = get_settings()
        server_config = settings.get_server_config()
        port = server_config.get("port", 8080)

        if server.is_running:
            forms.alert(
                "T3LabAI MCP Server is already connected!\n\n"
                "Port: {}\n"
                "Clients connected: {}\n\n"
                "SSE Endpoint: http://localhost:{}/sse".format(
                    server.port, server.client_count, server.port
                ),
                title="Already Connected"
            )
        else:
            # Set port and start server
            server.port = port

            # Start server in background thread
            server_thread = Thread(ThreadStart(start_server_thread))
            server_thread.IsBackground = True
            server_thread.Start()

            # Wait for server to start
            import time
            time.sleep(1)

            output.print_md("# T3LabAI MCP Server Connected")
            output.print_md("")
            output.print_md("## Server Information")
            output.print_md("- **Port:** {}".format(port))
            output.print_md("- **SSE Endpoint:** `http://localhost:{}/sse`".format(port))
            output.print_md("- **Message Endpoint:** `http://localhost:{}/message`".format(port))
            output.print_md("- **Status:** Running and waiting for Claude connection...")
            output.print_md("")
            output.print_md("---")
            output.print_md("")
            output.print_md("## Claude Desktop Configuration")
            output.print_md("")
            output.print_md("Add this to your `claude_desktop_config.json`:")
            output.print_md("")
            output.print_md("**Config file location:** `{}`".format(get_claude_config_path()))
            output.print_md("")
            output.print_md("```json")
            output.print_md(get_mcp_config_snippet(port))
            output.print_md("```")
            output.print_md("")
            output.print_md("---")
            output.print_md("")
            output.print_md("## Available Revit Tools")
            output.print_md("")
            output.print_md("| Tool | Description |")
            output.print_md("|------|-------------|")
            output.print_md("| `revit_get_active_view` | Get current active view info |")
            output.print_md("| `revit_get_selected_elements` | Get selected elements |")
            output.print_md("| `revit_get_project_info` | Get project information |")
            output.print_md("| `revit_list_views` | List all views in document |")
            output.print_md("| `revit_list_sheets` | List all sheets in document |")
            output.print_md("| `revit_get_element_info` | Get element details by ID |")
            output.print_md("")
            output.print_md("---")
            output.print_md("")
            output.print_md("## Next Steps")
            output.print_md("1. Copy the configuration above to Claude Desktop config")
            output.print_md("2. Restart Claude Desktop")
            output.print_md("3. Look for the hammer icon in Claude")
            output.print_md("4. Start chatting with Claude about your Revit project!")

            # Update button states
            update_connection_buttons(connected=True)

            forms.alert(
                "T3LabAI MCP Server Connected!\n\n"
                "Port: {}\n"
                "SSE Endpoint: http://localhost:{}/sse\n\n"
                "Add this to Claude Desktop config:\n"
                "claude_desktop_config.json\n\n"
                "Check the output window for full details.".format(port, port),
                title="Connected - Ready for Claude"
            )

    except Exception as e:
        output.print_md("## Error Connecting")
        output.print_md("```\n{}\n```".format(str(e)))
        import traceback
        output.print_md("```\n{}\n```".format(traceback.format_exc()))

        # Update status to error
        update_connection_buttons(connected=False, error=True, error_message=str(e))

        forms.alert(
            "Failed to connect to T3LabAI MCP Server\n\n"
            "Error: {}\n\n"
            "Check output window for details.".format(str(e)),
            title="Connection Error"
        )
