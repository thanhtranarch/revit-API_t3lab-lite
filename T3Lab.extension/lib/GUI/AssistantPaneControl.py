# -*- coding: utf-8 -*-
"""
T3Lab Assistant — Dockable Pane Provider and Controller

Registers and manages the T3Lab AI Assistant as a native Revit DockablePane,
allowing it to dock alongside the Properties panel and Project Browser.
"""

from __future__ import unicode_literals

import os
import sys
import json
import re

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
clr.AddReference('RevitAPIUI')

from System import Guid
from System.Windows import Visibility
from System.Windows.Markup import XamlReader
from System.IO import FileStream, FileMode
from Autodesk.Revit.UI import IDockablePaneProvider, DockablePaneProviderData, DockablePaneState

# ─── Path bootstrap ────────────────────────────────────────────────────────────
_GUI_DIR  = os.path.dirname(__file__)                         # lib/GUI
_LIB_DIR  = os.path.dirname(_GUI_DIR)                        # lib
_EXT_DIR  = os.path.dirname(_LIB_DIR)                        # T3Lab.extension
for _p in (_LIB_DIR, _EXT_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

_XAML_PATH = os.path.join(_GUI_DIR, 'Tools', 'AssistantPane.xaml')

# ─── Shared pane GUID (must match startup.py) ──────────────────────────────────
ASSISTANT_PANE_GUID = Guid('7F3A9B2E-C4D1-4E8F-A6B5-1234567890AB')

# ─── Singleton controller reference (set when pane is first created) ───────────
_pane_controller = None


def get_pane_controller():
    """Return the singleton AssistantPaneController, if it has been created."""
    return _pane_controller


# ─── AI helpers (lazy-imported so pane can load even without full AI stack) ────

def _try_import_ai():
    try:
        from Intelligence.t3lab_assistant import (
            parse_command, has_api_key,
            get_active_provider_name, get_provider_display_label
        )
        from Intelligence.llm_router import route_message
        return True, parse_command, has_api_key, route_message, get_active_provider_name, get_provider_display_label
    except Exception:
        return False, None, None, None, None, None


def _try_import_history():
    try:
        # Reuse the history functions from the main script module if available
        from config.settings import get_setting
        return True, get_setting
    except Exception:
        return False, None


# ─── Pane Controller ───────────────────────────────────────────────────────────

class AssistantPaneController(object):
    """
    Manages the T3Lab Assistant UserControl hosted inside Revit's DockablePane.
    Handles chat messages, AI responses, and MCP command dispatch.
    """

    PROVIDER_COLORS = {
        'claude':    '#D97706',   # amber
        'openai':    '#10B981',   # green
        'deepseek':  '#3B82F6',   # blue
        'ollama':    '#8B5CF6',   # purple
        'lmstudio':  '#EC4899',   # pink
    }

    def __init__(self, control):
        """
        Args:
            control: WPF UserControl loaded from AssistantPane.xaml
        """
        global _pane_controller
        self._control = control
        self._messages = []        # list of {'role': 'user'|'assistant', 'content': str}
        self._thinking = False

        # Wire named controls
        self._chat_panel   = control.FindName('pane_chat_panel')
        self._chat_scroll  = control.FindName('pane_chat_scroll')
        self._chat_input   = control.FindName('pane_chat_input')
        self._send_btn     = control.FindName('pane_send_btn')
        self._clear_btn    = control.FindName('pane_clear_btn')
        self._open_full    = control.FindName('pane_open_full_btn')
        self._provider_dot = control.FindName('pane_provider_dot')
        self._provider_lbl = control.FindName('pane_provider_label')
        self._loading_bar  = control.FindName('pane_loading_bar')

        # Wire events
        if self._send_btn:
            self._send_btn.Click += self._on_send
        if self._clear_btn:
            self._clear_btn.Click += self._on_clear
        if self._open_full:
            self._open_full.Click += self._on_open_full
        if self._chat_input:
            self._chat_input.KeyDown += self._on_key_down

        # Refresh provider badge
        self._refresh_provider()

        _pane_controller = self

    # ── Provider badge ─────────────────────────────────────────────────────────

    def _refresh_provider(self):
        try:
            ok, _, _, _, get_name, get_label = _try_import_ai()
            if ok:
                name  = get_name() or 'claude'
                label = get_label() or 'Claude'
                color = self.PROVIDER_COLORS.get(name, '#64748B')
            else:
                name, label, color = 'ai', 'AI', '#64748B'

            if self._provider_dot:
                from System.Windows.Media import BrushConverter
                self._provider_dot.Fill = BrushConverter().ConvertFromString(color)
            if self._provider_lbl:
                self._provider_lbl.Text = label
        except Exception:
            pass

    # ── Loading bar ────────────────────────────────────────────────────────────

    def _set_loading(self, visible):
        if self._loading_bar:
            self._loading_bar.Visibility = Visibility.Visible if visible else Visibility.Collapsed
        if self._send_btn:
            self._send_btn.IsEnabled = not visible
        if self._chat_input:
            self._chat_input.IsEnabled = not visible
        self._thinking = visible

    # ── Message rendering ──────────────────────────────────────────────────────

    def _add_message_bubble(self, text, is_user):
        """Append a styled chat bubble to the chat panel."""
        from System.Windows.Controls import Border, TextBlock, StackPanel
        from System.Windows import HorizontalAlignment, TextWrapping, Thickness
        from System.Windows.Media import BrushConverter

        bc = BrushConverter()
        bubble = Border()
        bubble.CornerRadius = _corner_radius(10)
        bubble.Margin = Thickness(4, 3, 4, 3)
        bubble.Padding = Thickness(10, 7, 10, 7)
        bubble.MaxWidth = 320

        txt = TextBlock()
        txt.Text = text
        txt.TextWrapping = TextWrapping.Wrap
        txt.FontSize = 12
        txt.LineHeight = 18

        if is_user:
            bubble.Background   = bc.ConvertFromString('#0F172A')
            txt.Foreground      = bc.ConvertFromString('#FFFFFF')
            bubble.HorizontalAlignment = HorizontalAlignment.Right
        else:
            bubble.Background   = bc.ConvertFromString('#F1F5F9')
            txt.Foreground      = bc.ConvertFromString('#0F172A')
            bubble.HorizontalAlignment = HorizontalAlignment.Left

        bubble.Child = txt
        if self._chat_panel:
            self._chat_panel.Children.Add(bubble)
            self._scroll_to_bottom()

    def _scroll_to_bottom(self):
        if self._chat_scroll:
            self._chat_scroll.ScrollToBottom()

    # ── Event handlers ─────────────────────────────────────────────────────────

    def _on_key_down(self, sender, e):
        from System.Windows.Input import Key
        if e.Key == Key.Return and not e.KeyboardDevice.Modifiers:
            e.Handled = True
            self._send_message()

    def _on_send(self, sender, e):
        self._send_message()

    def _on_clear(self, sender, e):
        self._messages = []
        if self._chat_panel:
            self._chat_panel.Children.Clear()

    def _on_open_full(self, sender, e):
        """Open the full floating T3Lab Assistant window."""
        try:
            import imp
            tab_dir = os.path.join(_EXT_DIR, 'T3Lab.tab')
            script_path = os.path.join(
                tab_dir, 'Support.panel', 'T3LabAssistant.pushbutton', 'script.py'
            )
            if os.path.isfile(script_path):
                mod = imp.load_source('t3lab_assistant_script', script_path)
                if hasattr(mod, 'T3LabAssistantWindow'):
                    win = mod.T3LabAssistantWindow()
                    win.ShowDialog()
        except Exception as ex:
            self._add_message_bubble(
                u'Could not open full window: {}'.format(ex), is_user=False
            )

    # ── Core chat logic ────────────────────────────────────────────────────────

    def _send_message(self):
        if self._thinking or not self._chat_input:
            return
        text = (self._chat_input.Text or '').strip()
        if not text:
            return

        self._chat_input.Text = ''
        self._add_message_bubble(text, is_user=True)
        self._messages.append({'role': 'user', 'content': text})
        self._set_loading(True)

        # Run AI call on a background thread to keep UI responsive
        from System.Threading import Thread, ThreadStart
        t = Thread(ThreadStart(lambda: self._ai_call(text)))
        t.IsBackground = True
        t.Start()

    def _ai_call(self, user_text):
        """Called on a background thread. Posts result back to UI thread."""
        reply = self._get_ai_reply(user_text)

        # Marshal back to UI thread
        if self._control and self._control.Dispatcher:
            self._control.Dispatcher.BeginInvoke(
                System_Action(lambda: self._on_reply_received(reply))
            )

    def _on_reply_received(self, reply):
        self._set_loading(False)
        self._add_message_bubble(reply, is_user=False)
        self._messages.append({'role': 'assistant', 'content': reply})

    def _get_ai_reply(self, user_text):
        """Call the AI provider and return a response string."""
        try:
            ok, parse_cmd, has_key, route_msg, _, _ = _try_import_ai()

            if not ok:
                return u'AI modules not available. Check your installation.'

            if not has_key():
                return (u'No API key configured. Click "Open full assistant" '
                        u'and set your API key in Settings.')

            # Build history for context
            history = list(self._messages[:-1])  # exclude the just-added user message

            result = route_msg(user_text, history=history)

            if result and isinstance(result, dict):
                return result.get('message') or result.get('answer') or str(result)
            elif isinstance(result, str):
                return result
            else:
                return u'(no response)'

        except Exception as ex:
            return u'Error: {}'.format(ex)

    # ── Public API ─────────────────────────────────────────────────────────────

    def add_message(self, text, is_user=True):
        """External API to inject a message (e.g. from MCP command results)."""
        self._add_message_bubble(text, is_user=is_user)
        self._messages.append({
            'role': 'user' if is_user else 'assistant',
            'content': text
        })


# ─── Corner radius helper (IronPython 2.7 compat) ──────────────────────────────

def _corner_radius(r):
    from System.Windows import CornerRadius
    return CornerRadius(r)


# ─── System.Action wrapper ──────────────────────────────────────────────────────

try:
    from System import Action as System_Action
except ImportError:
    System_Action = None


# ─── IDockablePaneProvider ─────────────────────────────────────────────────────

class AssistantPaneProvider(IDockablePaneProvider):
    """
    Revit calls SetupDockablePane() the first time the pane is shown.
    We load the UserControl XAML here and attach the controller.
    """

    def SetupDockablePane(self, data):
        try:
            # Load the UserControl XAML
            stream = FileStream(_XAML_PATH, FileMode.Open)
            try:
                control = XamlReader.Load(stream)
            finally:
                stream.Close()

            # Attach the Python controller (wires all events)
            AssistantPaneController(control)

            data.FrameworkElement = control

            # Keep the pane alive between document switches
            from Autodesk.Revit.UI import EditorInteraction, EditorInteractionType
            data.EditorInteraction = EditorInteraction(EditorInteractionType.KeepAlive)

        except Exception as ex:
            # Fallback: show an error label so the pane isn't empty
            from System.Windows.Controls import TextBlock
            from System.Windows import HorizontalAlignment, VerticalAlignment
            lbl = TextBlock()
            lbl.Text = u'T3Lab Assistant pane could not load:\n{}'.format(ex)
            lbl.HorizontalAlignment = HorizontalAlignment.Center
            lbl.VerticalAlignment   = VerticalAlignment.Center
            data.FrameworkElement = lbl
