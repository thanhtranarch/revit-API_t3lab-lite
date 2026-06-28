# -*- coding: utf-8 -*-
"""MCP Control Dialog class."""

from __future__ import unicode_literals

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

try:
    from core.file_watcher import get_task_watcher, T3LAB_DATA_DIR
    HAS_WATCHER = True
except Exception:
    HAS_WATCHER = False
    T3LAB_DATA_DIR = os.path.join(os.path.expanduser('~'), 'T3Lab_AI_Data')

logger = script.get_logger()


def brush(hex_string):
    return BrushConverter().ConvertFromString(hex_string)


class MCPControlWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)

        # Wire MCP server events
        self.toggle_btn.Click   += self._on_toggle
        self.copy_btn.Click     += self._on_copy
        self.port_tb.TextChanged += self._port_changed

        # Wire file watcher events
        try:
            self.watcher_toggle_btn.Click += self._on_watcher_toggle
        except Exception:
            pass
        try:
            self.open_dir_btn.Click += self._on_open_dir
        except Exception:
            pass

        self._init_port()
        self._refresh_status()

    # ── Window chrome ──────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── MCP server ─────────────────────────────────────────────────────────────

    def _init_port(self):
        if HAS_SERVER:
            try:
                server = get_t3labai_server()
                self.port_tb.Text = str(server.port)
                return
            except Exception:
                pass
        self.port_tb.Text = '48884'

    def _make_config_snippet(self, port_str):
        gui_dir     = os.path.dirname(__file__)
        lib_dir     = os.path.dirname(gui_dir)
        bridge_path = os.path.join(lib_dir, 'core', 'bridge.py').replace('\\', '/')
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

    def _refresh_mcp_status(self):
        if not HAS_SERVER:
            self.status_indicator.Background = brush('#94A3B8')
            self.status_label.Text           = 'Server module unavailable'
            self.toggle_btn.Content          = 'Start Server'
            self.toggle_btn.Style            = self.Resources['SecondaryButton']
            self.toggle_btn.IsEnabled        = False
            return

        try:
            server  = get_t3labai_server()
            running = server.is_running
        except Exception as ex:
            self.status_indicator.Background = brush('#EF4444')
            self.status_label.Text           = 'Error: ' + str(ex)
            self.toggle_btn.Content          = 'Start Server'
            self.toggle_btn.Style            = self.Resources['PrimaryButton']
            self.toggle_btn.IsEnabled        = True
            return

        if running:
            self.status_indicator.Background = brush('#10B981')
            self.status_label.Text           = 'Connected — port ' + str(server.port)
            self.toggle_btn.Content          = 'Stop Server'
            self.toggle_btn.Style            = self.Resources['DangerButton']
        else:
            self.status_indicator.Background = brush('#EF4444')
            self.status_label.Text           = 'Disconnected'
            self.toggle_btn.Content          = 'Start Server'
            self.toggle_btn.Style            = self.Resources['SuccessButton']

        self.toggle_btn.IsEnabled = True
        self.config_box.Text      = self._make_config_snippet(self.port_tb.Text)

    def _on_toggle(self, sender, e):
        if not HAS_SERVER:
            return
        try:
            server = get_t3labai_server()
            if server.is_running:
                server.stop_server()
                logger.info('MCP Server stopped.')
            else:
                try:
                    port_val   = int(self.port_tb.Text.strip())
                    server.port = port_val
                except Exception:
                    pass
                if server.start_server():
                    logger.info('MCP Server started on port {}.'.format(server.port))
                else:
                    logger.error('Failed to start MCP Server.')
        except Exception as ex:
            logger.error('Toggle error: {}'.format(ex))
        self._refresh_status()

    def _on_copy(self, sender, e):
        try:
            Clipboard.SetText(self.config_box.Text)
            logger.info('Configuration copied to clipboard.')
        except Exception as ex:
            logger.error('Clipboard error: {}'.format(ex))

    def _port_changed(self, sender, e):
        self.config_box.Text = self._make_config_snippet(self.port_tb.Text)

    # ── File watcher ───────────────────────────────────────────────────────────

    def _refresh_watcher_status(self):
        try:
            watcher_ind = self.FindName('watcher_indicator')
            watcher_lbl = self.FindName('watcher_label')
            watcher_btn = self.FindName('watcher_toggle_btn')
            dir_lbl     = self.FindName('data_dir_label')
        except Exception:
            return

        if dir_lbl:
            try:
                dir_lbl.Text = T3LAB_DATA_DIR
            except Exception:
                pass

        if not HAS_WATCHER:
            if watcher_ind: watcher_ind.Background = brush('#94A3B8')
            if watcher_lbl: watcher_lbl.Text       = 'File watcher module unavailable'
            if watcher_btn: watcher_btn.IsEnabled   = False
            return

        try:
            watcher = get_task_watcher()
            running = watcher.is_running
        except Exception as ex:
            if watcher_ind: watcher_ind.Background = brush('#EF4444')
            if watcher_lbl: watcher_lbl.Text       = 'Error: ' + str(ex)
            if watcher_btn: watcher_btn.IsEnabled   = True
            return

        if running:
            if watcher_ind: watcher_ind.Background = brush('#10B981')
            if watcher_lbl: watcher_lbl.Text       = 'File watcher active — monitoring task.json'
            if watcher_btn:
                watcher_btn.Content = 'Stop Watcher'
                watcher_btn.Style   = self.Resources['DangerButton']
        else:
            if watcher_ind: watcher_ind.Background = brush('#EF4444')
            if watcher_lbl: watcher_lbl.Text       = 'File watcher stopped'
            if watcher_btn:
                watcher_btn.Content = 'Start Watcher'
                watcher_btn.Style   = self.Resources['SuccessButton']

        if watcher_btn: watcher_btn.IsEnabled = True

    def _on_watcher_toggle(self, sender, e):
        if not HAS_WATCHER:
            return
        try:
            watcher = get_task_watcher()
            if watcher.is_running:
                watcher.stop()
                logger.info('File watcher stopped.')
            else:
                watcher.start()
                logger.info('File watcher started.')
        except Exception as ex:
            logger.error('Watcher toggle error: {}'.format(ex))
        self._refresh_status()

    def _on_open_dir(self, sender, e):
        """Open the T3Lab_AI_Data directory in Explorer."""
        try:
            import subprocess
            if not os.path.isdir(T3LAB_DATA_DIR):
                os.makedirs(T3LAB_DATA_DIR)
            subprocess.Popen(['explorer', T3LAB_DATA_DIR])
        except Exception as ex:
            logger.error('Could not open directory: {}'.format(ex))

    # ── Combined refresh ───────────────────────────────────────────────────────

    def _refresh_status(self):
        self._refresh_mcp_status()
        self._refresh_watcher_status()


def show_mcp_control_dialog():
    """Show the MCP Control dialog."""
    dlg = MCPControlWindow()
    dlg.ShowDialog()
