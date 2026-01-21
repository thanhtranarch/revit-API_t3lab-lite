"""
MCP Server Status Button
Shows the current status of the T3LabAI MCP server
"""

from pyrevit import script, forms
import sys

# Add lib folder to path
lib_path = script.get_bundle_file('lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from core.server import get_t3labai_server
from core.registry import t3labai_registry
from config.settings import get_settings

__title__ = 'Status\nMCP'
__author__ = 'T3LabAI Team'
__doc__ = 'Show T3LabAI MCP server status and statistics'

if __name__ == '__main__':
    output = script.get_output()

    try:
        server = get_t3labai_server()
        settings = get_settings()
        stats = server.get_server_stats()
        cmd_stats = t3labai_registry.get_command_statistics()

        output.print_md("# < T3LabAI MCP Server Status")
        output.print_md("")

        # Server Status
        output.print_md("## Server Status")
        status_icon = "=� Running" if stats['running'] else "=4 Stopped"
        output.print_md("**Status:** {}".format(status_icon))
        output.print_md("**Port:** {}".format(stats['port']))
        output.print_md("**Total Clients:** {}".format(stats['total_clients']))
        output.print_md("**Commands Processed:** {}".format(stats['commands_processed']))
        output.print_md("")

        # MCP Providers
        output.print_md("## MCP Providers")
        providers = settings.get_enabled_providers()
        if providers:
            output.print_md("**Configured Providers:** {}".format(len(providers)))
            for provider in providers:
                output.print_md("- {} (Enabled: {})".format(
                    provider['name'],
                    "" if provider.get('enabled', True) else "L"
                ))
        else:
            output.print_md("*No providers configured. Use Settings to add.*")
        output.print_md("")

        # Command Statistics
        output.print_md("## Command Statistics")
        output.print_md("**Total Commands:** {}".format(cmd_stats['total_commands']))
        output.print_md("**AI-Enhanced Commands:** {}".format(cmd_stats['ai_enhanced_count']))
        output.print_md("")

        if cmd_stats['most_used']:
            output.print_md("**Most Used Command:** {}".format(cmd_stats['most_used']))
            output.print_md("")

        # Available Commands
        output.print_md("## Available Commands")
        available_commands = t3labai_registry.get_available_commands()
        ai_enhanced = t3labai_registry.get_ai_enhanced_commands()

        for cmd in available_commands:
            is_enhanced = "> AI" if cmd in ai_enhanced else "=� Standard"
            output.print_md("- `{}` - {}".format(cmd, is_enhanced))

        output.print_md("")
        output.print_md("---")
        output.print_md("*Use Settings button to configure API keys and modules*")

        # Show summary alert
        forms.alert(
            "< T3LabAI MCP Server\n\n"
            "Status: {}\n"
            "Port: {}\n"
            "Clients: {}\n"
            "Commands: {}\n\n"
            "Check output window for full details.".format(
                "Running" if stats['running'] else "Stopped",
                stats['port'],
                stats['total_clients'],
                stats['commands_processed']
            ),
            title="Server Status"
        )

    except Exception as e:
        output.print_md("## L Error Getting Status")
        output.print_md("```\n{}\n```".format(str(e)))
        import traceback
        output.print_md("```\n{}\n```".format(traceback.format_exc()))

        forms.alert(
            "L Error getting server status\n\n"
            "Error: {}".format(str(e)),
            title="Status Error"
        )
