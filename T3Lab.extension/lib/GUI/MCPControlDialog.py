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
        color, text = '#EF4444', 'Error: {}'.format(status['error'])
        btn_content, btn_style_key = 'Start Server', 'PrimaryButton'
        enabled = True
    elif status['running']:
        color       = '#10B981'
        text        = 'Connected — port {}'.format(status['port'])
        btn_content = 'Stop Server'
        btn_style_key = 'DangerButton'
        enabled     = True
    else:
        color, text = '#EF4444', 'Disconnected'
        btn_content, btn_style_key = 'Start Server', 'SuccessButton'
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
        if indicator: indicator.Background = _brush('#94A3B8')
        if label:     label.Text           = err
        if btn:       btn.IsEnabled        = False
        return

    if status['running']:
        color = '#10B981'
        text  = 'File watcher active — monitoring task.json'
        btn_content, btn_style_key = 'Stop Watcher', 'DangerButton'
    else:
        color = '#EF4444'
        text  = 'File watcher stopped'
        btn_content, btn_style_key = 'Start Watcher', 'SuccessButton'

    if indicator: indicator.Background = _brush(color)
    if label:     label.Text           = text
    if btn:
        btn.Content   = btn_content
        btn.IsEnabled = True
        if resources and btn_style_key in resources:
            btn.Style = resources[btn_style_key]


# ─── Dialog ────────────────────────────────────────────────────────────────────

class MCPControlWindow(forms.WPFWindow):
    """
    MCP Control dialog — thin UI layer over MCPService.
    """

    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)

        # MCP server events
        self.toggle_btn.Click    += self._on_toggle
        self.copy_btn.Click      += self._on_copy
        self.port_tb.TextChanged += self._on_port_changed

        # Active document pinning widgets
        self._pin_indicator = self.FindName('pin_indicator')
        self._pin_label     = self.FindName('pin_label')
        self._doc_combo     = self.FindName('doc_combo')
        self._pin_btn       = self.FindName('pin_btn')
        if self._pin_btn:
            self._pin_btn.Click += self._on_pin_toggle

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
        if not self._doc_combo:
            return
        if not HAS_SERVICE:
            if self._pin_indicator: self._pin_indicator.Background = _brush('#94A3B8')
            if self._pin_label:     self._pin_label.Text = 'Service unavailable'
            return

        docs, err = MCPService.list_open_documents()
        if err:
            if self._pin_indicator: self._pin_indicator.Background = _brush('#EF4444')
            if self._pin_label:     self._pin_label.Text = 'Error: {}'.format(err)
            return

        titles = [d['title'] for d in docs]
        pinned_title = next((d['title'] for d in docs if d['is_pinned']), None)
        active_title = next((d['title'] for d in docs if d['is_active']), None)

        selected = self._doc_combo.SelectedItem
        keep = selected if selected in titles else None

        self._doc_combo.Items.Clear()
        for title in titles:
            self._doc_combo.Items.Add(title)

        target = keep or pinned_title or active_title
        if target is not None:
            self._doc_combo.SelectedItem = target
        elif titles:
            self._doc_combo.SelectedIndex = 0

        if pinned_title:
            if self._pin_indicator: self._pin_indicator.Background = _brush('#3B82F6')
            if self._pin_label:     self._pin_label.Text = 'Pinned: {}'.format(pinned_title)
            if self._pin_btn:       self._pin_btn.Content = 'Unpin'
        else:
            if self._pin_indicator: self._pin_indicator.Background = _brush('#94A3B8')
            if self._pin_label:     self._pin_label.Text = 'Following active window'
            if self._pin_btn:       self._pin_btn.Content = 'Pin'

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
        self._refresh_server()

    def _on_pin_toggle(self, sender, e):
        if not HAS_SERVICE or not self._doc_combo:
            return
        pinned_title, _ = MCPService.pinned_document()
        if pinned_title:
            ok, err = MCPService.unpin_document()
        else:
            selected = self._doc_combo.SelectedItem
            if selected is None:
                return
            ok, err = MCPService.pin_document(selected)
        if err:
            logger.error('Document pin error: {}'.format(err))
        self._refresh_documents()

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


def show_mcp_control_dialog():
    """Show the MCP Control dialog."""
    MCPControlWindow().ShowDialog()
