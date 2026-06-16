# -*- coding: utf-8 -*-
"""MCP Control

Unified control panel for the T3Lab MCP server.
Start / Stop the server and manage connection settings in one dialog.
"""
__title__ = "MCP\nControl"
__author__ = "T3Lab & Dang Quoc Truong"

import os
import sys
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

import System
import System.Windows.Input

from System.Windows import (
    Window, WindowStartupLocation, Thickness, CornerRadius,
    HorizontalAlignment, VerticalAlignment,
    FontWeights, TextWrapping, ResizeMode, Clipboard
)
from System.Windows.Controls import (
    Grid, RowDefinition, StackPanel, Border, TextBlock,
    Button, TextBox, ScrollBarVisibility, Orientation
)
from System.Windows.Media import BrushConverter, FontFamily

# ---------------------------------------------------------------------------
# Path setup — script.py lives 4 levels below T3Lab.extension/
#   MCPControl.pushbutton/script.py
#   AI Connection.panel/
#   T3Lab.tab/
#   T3Lab.extension/   ← EXT_DIR
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import script

try:
    from core.server import get_t3labai_server
    HAS_SERVER = True
except Exception as _server_err:
    HAS_SERVER = False
    _SERVER_ERR_MSG = str(_server_err)

logger = script.get_logger()

# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------
def _brush(hex_color):
    return BrushConverter().ConvertFromString(hex_color)


# ---------------------------------------------------------------------------
# MCPControlWindow
# ---------------------------------------------------------------------------
class MCPControlWindow(Window):

    def __init__(self):
        self.Title = "T3Lab AI — MCP Control"
        self.Width = 480
        self.Height = 460
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = _brush("#F8FAFC")
        self.FontFamily = FontFamily("Hanken Grotesk")

        # Root StackPanel
        root = StackPanel()
        root.Margin = Thickness(20)
        root.Orientation = Orientation.Vertical

        # ── Header ──────────────────────────────────────────────────────────
        header = TextBlock()
        header.Text = "MCP Control"
        header.FontSize = 18
        header.FontWeight = FontWeights.Bold
        header.Foreground = _brush("#0F172A")
        header.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(header)

        subtitle = TextBlock()
        subtitle.Text = "AI Connection & Server Settings"
        subtitle.FontSize = 11
        subtitle.Foreground = _brush("#64748B")
        subtitle.Margin = Thickness(0, 0, 0, 20)
        root.Children.Add(subtitle)

        # ── Status Row ──────────────────────────────────────────────────────
        status_row = StackPanel()
        status_row.Orientation = Orientation.Horizontal
        status_row.Margin = Thickness(0, 0, 0, 16)

        self.status_indicator = Border()
        self.status_indicator.Width = 12
        self.status_indicator.Height = 12
        self.status_indicator.CornerRadius = CornerRadius(6)
        self.status_indicator.VerticalAlignment = VerticalAlignment.Center
        self.status_indicator.Margin = Thickness(0, 0, 8, 0)

        self.status_label = TextBlock()
        self.status_label.VerticalAlignment = VerticalAlignment.Center
        self.status_label.FontSize = 12
        self.status_label.Foreground = _brush("#0F172A")

        status_row.Children.Add(self.status_indicator)
        status_row.Children.Add(self.status_label)
        root.Children.Add(status_row)

        # ── Toggle Button ───────────────────────────────────────────────────
        self.toggle_btn = Button()
        self.toggle_btn.Height = 36
        self.toggle_btn.HorizontalAlignment = HorizontalAlignment.Stretch
        self.toggle_btn.BorderThickness = Thickness(0)
        self.toggle_btn.FontSize = 13
        self.toggle_btn.FontWeight = FontWeights.SemiBold
        self.toggle_btn.Foreground = _brush("#FFFFFF")
        self.toggle_btn.Margin = Thickness(0, 0, 0, 20)
        self.toggle_btn.Cursor = System.Windows.Input.Cursors.Hand
        self.toggle_btn.Click += self._on_toggle
        root.Children.Add(self.toggle_btn)

        # ── Separator ───────────────────────────────────────────────────────
        sep = Border()
        sep.Height = 1
        sep.Background = _brush("#E2E8F0")
        sep.Margin = Thickness(0, 0, 0, 16)
        root.Children.Add(sep)

        # ── Port ────────────────────────────────────────────────────────────
        port_label = TextBlock()
        port_label.Text = "MCP Server Port:"
        port_label.FontWeight = FontWeights.SemiBold
        port_label.Foreground = _brush("#0F172A")
        port_label.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(port_label)

        self.port_tb = TextBox()
        self.port_tb.Width = 120
        self.port_tb.HorizontalAlignment = HorizontalAlignment.Left
        self.port_tb.Padding = Thickness(6, 4, 6, 4)
        self.port_tb.BorderBrush = _brush("#CBD5E1")
        self.port_tb.FontFamily = FontFamily("Inter")
        self.port_tb.FontSize = 12
        self.port_tb.Margin = Thickness(0, 0, 0, 16)
        self.port_tb.TextChanged += self._port_changed
        root.Children.Add(self.port_tb)

        # ── Config Snippet ──────────────────────────────────────────────────
        config_label = TextBlock()
        config_label.Text = "Claude Desktop Configuration:"
        config_label.FontWeight = FontWeights.SemiBold
        config_label.Foreground = _brush("#0F172A")
        config_label.Margin = Thickness(0, 0, 0, 4)
        root.Children.Add(config_label)

        self.config_box = TextBox()
        self.config_box.IsReadOnly = True
        self.config_box.Height = 130
        self.config_box.FontFamily = FontFamily("Consolas")
        self.config_box.FontSize = 11
        self.config_box.Background = _brush("#F8FAFC")
        self.config_box.BorderBrush = _brush("#CBD5E1")
        self.config_box.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.config_box.Padding = Thickness(8, 6, 8, 6)
        self.config_box.Margin = Thickness(0, 0, 0, 8)
        root.Children.Add(self.config_box)

        # ── Copy Button ─────────────────────────────────────────────────────
        copy_btn = Button()
        copy_btn.Content = "Copy to Clipboard"
        copy_btn.Height = 30
        copy_btn.Width = 160
        copy_btn.HorizontalAlignment = HorizontalAlignment.Left
        copy_btn.Background = _brush("#3B82F6")
        copy_btn.Foreground = _brush("#FFFFFF")
        copy_btn.BorderThickness = Thickness(0)
        copy_btn.FontSize = 12
        copy_btn.FontWeight = FontWeights.SemiBold
        copy_btn.Cursor = System.Windows.Input.Cursors.Hand
        copy_btn.Click += self._on_copy
        root.Children.Add(copy_btn)

        self.Content = root

        # Populate initial values then refresh status
        self._init_port()
        self._refresh_status()

    # -----------------------------------------------------------------------
    # Internal helpers
    # -----------------------------------------------------------------------
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
        bridge_path = os.path.join(LIB_DIR, "core", "bridge.py").replace("\\", "/")
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
        """Update the status indicator, label, and toggle button to match server state."""
        if not HAS_SERVER:
            self.status_indicator.Background = _brush("#94A3B8")
            self.status_label.Text = "Server module unavailable"
            self.toggle_btn.Content = "Start Server"
            self.toggle_btn.Background = _brush("#94A3B8")
            self.toggle_btn.IsEnabled = False
            return

        try:
            server = get_t3labai_server()
            running = server.is_running
        except Exception as e:
            self.status_indicator.Background = _brush("#EF4444")
            self.status_label.Text = "Error: " + str(e)
            self.toggle_btn.Content = "Start Server"
            self.toggle_btn.Background = _brush("#10B981")
            self.toggle_btn.IsEnabled = True
            return

        if running:
            self.status_indicator.Background = _brush("#10B981")
            self.status_label.Text = "Connected — port " + str(server.port)
            self.toggle_btn.Content = "Stop Server"
            self.toggle_btn.Background = _brush("#EF4444")
        else:
            self.status_indicator.Background = _brush("#EF4444")
            self.status_label.Text = "Disconnected"
            self.toggle_btn.Content = "Start Server"
            self.toggle_btn.Background = _brush("#10B981")

        self.toggle_btn.IsEnabled = True

        # Keep config snippet in sync with current port
        self.config_box.Text = self._make_config_snippet(self.port_tb.Text)

    # -----------------------------------------------------------------------
    # Event handlers
    # -----------------------------------------------------------------------
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
        except Exception as ex:
            logger.error("Clipboard error: {}".format(ex))

    def _port_changed(self, sender, e):
        self.config_box.Text = self._make_config_snippet(self.port_tb.Text)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    window = MCPControlWindow()
    window.ShowDialog()


if __name__ == '__main__':
    main()
