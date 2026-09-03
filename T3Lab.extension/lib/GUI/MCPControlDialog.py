# -*- coding: utf-8 -*-
"""
MCP Control Dialog

Thin WPF wrapper around MCPService. All backend logic lives in
Services/mcp_service.py and can be reused by any other tool.
"""

from __future__ import unicode_literals

import os
import sys
from pyrevit import forms, script
from GUI.WPF_Base import T3WPFWindow
from System.Windows import WindowState
from System.Windows.Media import BrushConverter

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'MCPControl.xaml')

# ─── Backend service ───────────────────────────────────────────────────────────
try:
    _LIB_DIR = os.path.dirname(os.path.dirname(__file__))
    if _LIB_DIR not in sys.path:
        sys.path.insert(0, _LIB_DIR)
    from Services.mcp_service import MCPService
    HAS_SERVICE = True
except Exception as _svc_err:
    HAS_SERVICE  = False
    _SVC_ERR_MSG = str(_svc_err)

logger = script.get_logger()


def _brush(hex_color):
    return BrushConverter().ConvertFromString(hex_color)


# ─── Status helpers shared with embedded widgets ───────────────────────────────

def apply_server_status(status, indicator, label, btn, resources):
    """
    Update server status widgets from an MCPService.server_status() dict.
    All widget args may be None (skipped gracefully).
    """
    if status.get('error'):
        color, text = '#D23B3B', 'Error: {}'.format(status['error'])
        btn_content, btn_style_key = 'Start Server', 'T3.Button.Primary'
        enabled = True
    elif status['running']:
        color       = '#157038'
        text        = 'Connected — port {}'.format(status['port'])
        btn_content = 'Stop Server'
        btn_style_key = 'T3.Button.Danger'
        enabled     = True
    else:
        color, text = '#71717A', 'Disconnected'
        btn_content, btn_style_key = 'Start Server', 'T3.Button.Primary'
        enabled = True

    if indicator: indicator.Background = _brush(color)
    if label:     label.Text           = text
    if btn:
        btn.Content   = btn_content
        btn.IsEnabled = enabled
        if resources and btn_style_key in resources:
            btn.Style = resources[btn_style_key]


def apply_watcher_status(status, indicator, label, btn, resources):
    """
    Update watcher status widgets from an MCPService.watcher_status() dict.
    """
    if not HAS_SERVICE or status.get('error'):
        err = status.get('error', 'Service unavailable') if status else 'Service unavailable'
        if indicator: indicator.Background = _brush('#9A9AA2')
        if label:     label.Text           = err
        if btn:       btn.IsEnabled        = False
        return

    if status['running']:
        color = '#157038'
        text  = 'File watcher active — monitoring task.json'
        btn_content, btn_style_key = 'Stop Watcher', 'T3.Button.Danger'
    else:
        color = '#71717A'
        text  = 'File watcher stopped'
        btn_content, btn_style_key = 'Start Watcher', 'T3.Button.Secondary'

    if indicator: indicator.Background = _brush(color)
    if label:     label.Text           = text
    if btn:
        btn.Content   = btn_content
        btn.IsEnabled = True
        if resources and btn_style_key in resources:
            btn.Style = resources[btn_style_key]


# ─── Dialog ────────────────────────────────────────────────────────────────────

class MCPControlWindow(T3WPFWindow):
    """
    MCP Control dialog — thin UI layer over MCPService.
    """

    def __init__(self):
        T3WPFWindow.__init__(self, _XAML)

        # MCP server events
        self.toggle_btn.Click    += self._on_toggle
        self.copy_btn.Click      += self._on_copy
        self.port_tb.TextChanged += self._on_port_changed

        # Active document widgets (read-only status — switching is done from
        # the AI client via the switch_active_document / open_document tools)
        self._doc_indicator = self.FindName('active_doc_indicator')
        self._doc_label     = self.FindName('active_doc_label')

        # File watcher events (FindName so missing elements don't crash)
        self._watcher_indicator = self.FindName('watcher_indicator')
        self._watcher_label     = self.FindName('watcher_label')
        self._watcher_btn       = self.FindName('watcher_toggle_btn')
        self._dir_label         = self.FindName('data_dir_label')
        open_dir_btn            = self.FindName('open_dir_btn')

        if self._watcher_btn:
            self._watcher_btn.Click += self._on_watcher_toggle
        if open_dir_btn:
            open_dir_btn.Click += self._on_open_dir

        # Claude Desktop auto-configure widgets
        self._claude_cfg_indicator = self.FindName('claude_cfg_indicator')
        self._claude_cfg_label     = self.FindName('claude_cfg_label')
        self._claude_cfg_path      = self.FindName('claude_cfg_path')
        configure_claude_btn       = self.FindName('configure_claude_btn')
        if configure_claude_btn:
            configure_claude_btn.Click += self._on_configure_claude

        # Teaching capture widgets (Opus distils to local Qwen via MCP)
        self._teaching_toggle    = self.FindName('teaching_toggle')
        self._teaching_indicator = self.FindName('teaching_indicator')
        self._teaching_label     = self.FindName('teaching_status_label')
        self._sandbox_label      = self.FindName('sandbox_label')
        mark_sandbox_btn         = self.FindName('mark_sandbox_btn')
        # Bind Click (not Checked/Unchecked): _refresh_teaching sets IsChecked
        # from server state, and Checked would re-fire the handler and fight it.
        if self._teaching_toggle:
            self._teaching_toggle.Click += self._on_teaching_toggle
        if mark_sandbox_btn:
            mark_sandbox_btn.Click += self._on_mark_sandbox

        self._init_port()
        self._refresh_all()

    # ── Window chrome ──────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── Init helpers ───────────────────────────────────────────────────────────

    def _init_port(self):
        port = 48884
        if HAS_SERVICE:
            try:
                port = MCPService.server_status().get('port', 48884)
            except Exception:
                pass
        self.port_tb.Text = str(port)

    # ── Refresh ────────────────────────────────────────────────────────────────

    def _refresh_all(self):
        self._refresh_server()
        self._refresh_documents()
        self._refresh_watcher()
        self._refresh_claude_config()
        self._refresh_teaching()

    def _refresh_server(self):
        if not HAS_SERVICE:
            self.status_indicator.Background = _brush('#94A3B8')
            self.status_label.Text           = 'Service unavailable: ' + _SVC_ERR_MSG
            self.toggle_btn.IsEnabled        = False
            return
        status = MCPService.server_status()
        apply_server_status(
            status,
            self.status_indicator,
            self.status_label,
            self.toggle_btn,
            self.Resources,
        )
        self.config_box.Text = MCPService.config_snippet(
            port=self.port_tb.Text or status.get('port')
        )

    def _refresh_documents(self):
        if not self._doc_label:
            return
        if not HAS_SERVICE:
            if self._doc_indicator: self._doc_indicator.Background = _brush('#94A3B8')
            self._doc_label.Text = 'Service unavailable'
            return

        docs, err = MCPService.list_open_documents()
        if err:
            if self._doc_indicator: self._doc_indicator.Background = _brush('#EF4444')
            self._doc_label.Text = 'Error: {}'.format(err)
            return

        active_title = next((d['title'] for d in docs if d['is_active']), None)
        if active_title:
            if self._doc_indicator: self._doc_indicator.Background = _brush('#10B981')
            self._doc_label.Text = 'Active: {}'.format(active_title)
        elif docs:
            if self._doc_indicator: self._doc_indicator.Background = _brush('#F59E0B')
            self._doc_label.Text = '{} open — none active (click a tab in Revit)'.format(len(docs))
        else:
            if self._doc_indicator: self._doc_indicator.Background = _brush('#94A3B8')
            self._doc_label.Text = 'No document open'

    def _refresh_watcher(self):
        if not HAS_SERVICE:
            apply_watcher_status(None, self._watcher_indicator,
                                 self._watcher_label, self._watcher_btn, self.Resources)
            return
        status = MCPService.watcher_status()
        apply_watcher_status(
            status,
            self._watcher_indicator,
            self._watcher_label,
            self._watcher_btn,
            self.Resources,
        )
        if self._dir_label:
            self._dir_label.Text = status.get('data_dir', MCPService.data_dir())

    # ── Event handlers ─────────────────────────────────────────────────────────

    def _on_toggle(self, sender, e):
        if not HAS_SERVICE:
            return
        try:
            port = int(self.port_tb.Text.strip())
        except Exception:
            port = None
        new_state, err = MCPService.toggle_server(current_port=port)
        if err:
            logger.error('MCP server toggle error: {}'.format(err))
        else:
            logger.info('MCP server: {}'.format(new_state))
            # We're on the UI thread here (button click) — a valid Revit API
            # context — so guarantee the ExternalEvent exists for model-editing
            # tools even if the server was auto-started from a background probe.
            if new_state == 'running':
                ok, ee_err = MCPService.ensure_external_event()
                if not ok:
                    logger.error('MCP ExternalEvent init error: {}'.format(ee_err))
        self._refresh_server()

    def _on_watcher_toggle(self, sender, e):
        if not HAS_SERVICE:
            return
        new_state, err = MCPService.toggle_watcher()
        if err:
            logger.error('File watcher toggle error: {}'.format(err))
        else:
            logger.info('File watcher: {}'.format(new_state))
        self._refresh_watcher()

    def _on_copy(self, sender, e):
        try:
            from System.Windows import Clipboard
            Clipboard.SetText(self.config_box.Text)
            logger.info('Configuration copied to clipboard.')
        except Exception as ex:
            logger.error('Clipboard error: {}'.format(ex))

    def _on_open_dir(self, sender, e):
        ok, err = MCPService.open_data_dir()
        if not ok:
            logger.error('Could not open data dir: {}'.format(err))

    def _refresh_claude_config(self):
        if not HAS_SERVICE:
            if self._claude_cfg_indicator:
                self._claude_cfg_indicator.Background = _brush('#94A3B8')
            if self._claude_cfg_label:
                self._claude_cfg_label.Text = 'Service unavailable'
            return
        status = MCPService.claude_desktop_status()
        if status.get('error'):
            color = '#EF4444'
            text  = 'Error: {}'.format(status['error'])
        elif not status['file_exists']:
            color = '#F59E0B'
            text  = 'Config not found — will be created on Configure'
        elif status['configured']:
            color = '#10B981'
            text  = 'Configured — t3lab-revit entry present'
        else:
            color = '#EF4444'
            text  = 'Not configured — click Configure to add entry'
        if self._claude_cfg_indicator:
            self._claude_cfg_indicator.Background = _brush(color)
        if self._claude_cfg_label:
            self._claude_cfg_label.Text = text
        if self._claude_cfg_path:
            self._claude_cfg_path.Text = status.get('path', '')

    def _on_configure_claude(self, sender, e):
        if not HAS_SERVICE:
            return
        try:
            port = int(self.port_tb.Text.strip())
        except Exception:
            port = None
        ok, msg = MCPService.configure_claude_desktop(port=port)
        if ok:
            logger.info('Claude Desktop configured: {}'.format(msg))
        else:
            logger.error('Claude Desktop configure error: {}'.format(msg))
        self._refresh_claude_config()

    def _on_port_changed(self, sender, e):
        if HAS_SERVICE:
            self.config_box.Text = MCPService.config_snippet(
                port=self.port_tb.Text or None
            )

    # ── Teaching capture ────────────────────────────────────────────────────────

    def _refresh_teaching(self):
        if not HAS_SERVICE or self._teaching_label is None:
            return
        try:
            status = MCPService.teaching_status()
        except Exception as ex:
            status = {'enabled': False, 'error': str(ex)}
        enabled = bool(status.get('enabled'))
        if self._teaching_toggle is not None:
            self._teaching_toggle.IsChecked = enabled
        if self._teaching_indicator is not None:
            self._teaching_indicator.Background = _brush(
                '#10B981' if enabled else '#94A3B8')
        recorded = status.get('sessions_recorded', 0)
        if status.get('error'):
            self._teaching_label.Text = 'Teaching unavailable'
        elif enabled:
            self._teaching_label.Text = (
                'Recording MCP sessions - {} captured'.format(recorded))
        else:
            self._teaching_label.Text = 'Teaching capture off'
        if self._sandbox_label is not None:
            sb = status.get('sandbox')
            self._sandbox_label.Text = (
                'Sandbox: {}'.format(sb) if sb
                else 'Sandbox: none - mark a scratch model')

    def _on_teaching_toggle(self, sender, e):
        if not HAS_SERVICE:
            return
        # The toggle already flipped IsChecked on click; push that to the server.
        want = bool(self._teaching_toggle.IsChecked) \
            if self._teaching_toggle is not None else False
        _new, err = MCPService.set_teaching_mode(want)
        if err:
            logger.error('Teaching mode error: {}'.format(err))
        self._refresh_teaching()

    def _on_mark_sandbox(self, sender, e):
        if not HAS_SERVICE:
            return
        info, err = MCPService.mark_active_document_as_sandbox()
        if err:
            logger.error('Mark sandbox error: {}'.format(err))
        else:
            logger.info('Sandbox document set: {}'.format(
                info.get('title') if info else ''))
        self._refresh_teaching()


def show_mcp_control_dialog():
    """Show the MCP Control dialog."""
    MCPControlWindow().ShowDialog()
