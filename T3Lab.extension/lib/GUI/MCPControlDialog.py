# -*- coding: utf-8 -*-
"""MCP Control Dialog class."""

import os
import sys
from pyrevit import forms, script
from System.Windows import WindowState, Clipboard
from System.Windows.Media import BrushConverter

# Absolute path to XAML
_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'MCPControl.xaml')

try:
    from core.server import get_t3labai_server
    HAS_SERVER = True
except Exception as _server_err:
    HAS_SERVER = False
    _SERVER_ERR_MSG = str(_server_err)

logger = script.get_logger()

def brush(hex_string):
    return BrushConverter().ConvertFromString(hex_string)

class MCPControlWindow(forms.WPFWindow):
    def __init__(self):
        # WPFWindow.__init__ loads XAML
        forms.WPFWindow.__init__(self, _XAML)

        # Wire events
        self.toggle_btn.Click += self._on_toggle
        self.copy_btn.Click += self._on_copy
        self.port_tb.TextChanged += self._port_changed

        # Populate initial values and refresh status
        self._init_port()
        self._refresh_status()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def close_button_clicked(self, sender, e):
        self.Close()

    def _init_port(self):
        """Set the port TextBox to the server's current port (or default)."""
        if HAS_SERVER:
            try:
                server = get_t3labai_server()
                self.port_tb.Text = str(server.port)
                return
            except Exception:
                pass
        self.port_tb.Text = "48884"

    def _make_config_snippet(self, port_str):
        gui_dir = os.path.dirname(__file__)  # lib/GUI
        lib_dir = os.path.dirname(gui_dir)  # lib
        bridge_path = os.path.join(lib_dir, "core", "bridge.py").replace("\\", "/")
        return (
            '{\n'
            '  "mcpServers": {\n'
            '    "t3lab-revit": {\n'
            '      "command": "python",\n'
            '      "args": [\n'
            '        "' + bridge_path + '",\n'
            '        "' + port_str + '"\n'
            '      ]\n'
            '    }\n'
            '  }\n'
            '}'
        )

    def _refresh_status(self):
        """Update status indicators, labels, and toggles."""
        if not HAS_SERVER:
            self.status_indicator.Background = brush("#94A3B8")
            self.status_label.Text = "Server module unavailable"
            self.toggle_btn.Content = "Start Server"
            self.toggle_btn.Style = self.Resources["SecondaryButton"]
            self.toggle_btn.IsEnabled = False
            return

        try:
            server = get_t3labai_server()
            running = server.is_running
        except Exception as e:
            self.status_indicator.Background = brush("#EF4444")
            self.status_label.Text = "Error: " + str(e)
            self.toggle_btn.Content = "Start Server"
            self.toggle_btn.Style = self.Resources["PrimaryButton"]
            self.toggle_btn.IsEnabled = True
            return

        if running:
            self.status_indicator.Background = brush("#10B981")
            self.status_label.Text = "Connected — port " + str(server.port)
            self.toggle_btn.Content = "Stop Server"
            self.toggle_btn.Style = self.Resources["DangerButton"]
        else:
            self.status_indicator.Background = brush("#EF4444")
            self.status_label.Text = "Disconnected"
            self.toggle_btn.Content = "Start Server"
            self.toggle_btn.Style = self.Resources["SuccessButton"]

        self.toggle_btn.IsEnabled = True

        # Keep config snippet in sync with current port
        self.config_box.Text = self._make_config_snippet(self.port_tb.Text)

    def _on_toggle(self, sender, e):
        if not HAS_SERVER:
            return
        try:
            server = get_t3labai_server()
            if server.is_running:
                if server.stop_server():
                    logger.info("MCP Server stopped.")
                else:
                    logger.error("Failed to stop MCP Server.")
            else:
                # Update port if modified in text box before starting
                try:
                    port_val = int(self.port_tb.Text.strip())
                    server.port = port_val
                except Exception:
                    pass
                if server.start_server():
                    logger.info("MCP Server started on port {}.".format(server.port))
                else:
                    logger.error("Failed to start MCP Server.")
        except Exception as ex:
            logger.error("Toggle error: {}".format(ex))
        self._refresh_status()

    def _on_copy(self, sender, e):
        try:
            Clipboard.SetText(self.config_box.Text)
            logger.info("Configuration copied to clipboard.")
        except Exception as ex:
            logger.error("Clipboard error: {}".format(ex))

    def _port_changed(self, sender, e):
        self.config_box.Text = self._make_config_snippet(self.port_tb.Text)

def show_mcp_control_dialog():
    """Show the MCP Control dialog."""
    dlg = MCPControlWindow()
    dlg.ShowDialog()
