# -*- coding: utf-8 -*-
"""
T3Lab Assistant — Dockable Pane Provider and Controller

Registers and manages the T3Lab AI Assistant as a native Revit DockablePane,
allowing it to dock alongside the Properties panel and Project Browser.
"""

from __future__ import unicode_literals

import os
import sys

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


# ─── System.Action wrapper (needed for Dispatcher.BeginInvoke) ─────────────────
try:
    from System import Action as System_Action
except ImportError:
    System_Action = None


# ─── AI helpers (lazy-imported so pane can load even without full AI stack) ────

def _try_import_ai():
    """
    Return (ok, has_api_key_fn, route_message_fn, route_stream_fn,
            get_provider_name_fn, get_provider_label_fn).
    All callables are None when ok=False.
    """
    try:
        from Intelligence.t3lab_assistant import (
            has_api_key, get_active_provider_name, get_provider_display_label
        )
        from Intelligence.llm_router import route_message, route_message_stream
        return True, has_api_key, route_message, route_message_stream, \
               get_active_provider_name, get_provider_display_label
    except Exception:
        return False, None, None, None, None, None


# ─── Pane Controller ───────────────────────────────────────────────────────────

class AssistantPaneController(object):
    """
    Manages the T3Lab Assistant UserControl hosted inside Revit's DockablePane.
    Handles chat messages, AI responses (with streaming), and BIM context injection.
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
        self._control  = control
        self._messages = []        # [{role: user|assistant, content: str}]
        self._thinking = False
        self._stream_tb = None     # live TextBlock being updated during streaming

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
        self._ctx_lbl      = control.FindName('pane_context_label')   # optional

        # Wire events
        if self._send_btn:
            self._send_btn.Click  += self._on_send
        if self._clear_btn:
            self._clear_btn.Click += self._on_clear
        if self._open_full:
            self._open_full.Click += self._on_open_full
        if self._chat_input:
            self._chat_input.KeyDown += self._on_key_down

        # Initial provider badge
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
            self._loading_bar.Visibility = (
                Visibility.Visible if visible else Visibility.Collapsed
            )
        if self._send_btn:
            self._send_btn.IsEnabled = not visible
        if self._chat_input:
            self._chat_input.IsEnabled = not visible
        self._thinking = visible

    # ── Context label ──────────────────────────────────────────────────────────

    def _update_context_label(self, text):
        """Update the optional context label (e.g. 'View: Level 1 • 3 selected')."""
        if self._ctx_lbl:
            try:
                self._ctx_lbl.Text = text
            except Exception:
                pass

    # ── BIM context injection ──────────────────────────────────────────────────

    def _get_bim_context(self):
        """
        Return a brief string summarising the current Revit state.
        Injected as a prefix so the AI knows what view/selection is active.
        Returns '' on failure (graceful degradation).
        """
        try:
            from pyrevit import HOST_APP
            uidoc = HOST_APP.uiapp.ActiveUIDocument
            if not uidoc:
                return u''
            doc  = uidoc.Document
            view = doc.ActiveView if doc else None
            parts = []
            if view:
                parts.append(u'Active view: {}'.format(view.Name))
            sel = uidoc.Selection.GetElementIds()
            if sel and sel.Count > 0:
                parts.append(u'{} element(s) selected'.format(sel.Count))
            ctx = u', '.join(parts)
            # Also update the visible label if present
            if ctx:
                if self._control and self._control.Dispatcher:
                    disp_ctx = ctx
                    self._control.Dispatcher.BeginInvoke(
                        System_Action(lambda: self._update_context_label(disp_ctx))
                    )
            return ctx
        except Exception:
            return u''

    # ── Message rendering ──────────────────────────────────────────────────────

    def _add_message_bubble(self, text, is_user, live=False):
        """
        Append a styled chat bubble to the panel.

        Args:
            text (str): Message content.
            is_user (bool): True for user bubbles (right-aligned, dark bg).
            live (bool): If True, keep a reference in self._stream_tb for
                         subsequent live updates (streaming mode).
        Returns:
            The TextBlock so the caller can update it during streaming.
        """
        from System.Windows.Controls import Border, TextBlock
        from System.Windows import HorizontalAlignment, TextWrapping, Thickness
        from System.Windows.Media import BrushConverter

        bc     = BrushConverter()
        bubble = Border()
        bubble.CornerRadius = _corner_radius(10)
        bubble.Margin       = Thickness(4, 3, 4, 3)
        bubble.Padding      = Thickness(10, 7, 10, 7)
        bubble.MaxWidth     = 320

        txt = TextBlock()
        txt.Text        = text
        txt.TextWrapping = TextWrapping.Wrap
        txt.FontSize    = 12
        txt.LineHeight  = 18

        if is_user:
            bubble.Background          = bc.ConvertFromString('#0F172A')
            txt.Foreground             = bc.ConvertFromString('#FFFFFF')
            bubble.HorizontalAlignment = HorizontalAlignment.Right
        else:
            bubble.Background          = bc.ConvertFromString('#F1F5F9')
            txt.Foreground             = bc.ConvertFromString('#0F172A')
            bubble.HorizontalAlignment = HorizontalAlignment.Left

        bubble.Child = txt
        if self._chat_panel:
            self._chat_panel.Children.Add(bubble)
            self._scroll_to_bottom()

        if live:
            self._stream_tb = txt
        return txt

    def _scroll_to_bottom(self):
        if self._chat_scroll:
            self._chat_scroll.ScrollToBottom()

    # ── Event handlers ─────────────────────────────────────────────────────────

    def _on_key_down(self, sender, e):
        from System.Windows.Input import Key, ModifierKeys
        if e.Key == Key.Return and not (e.KeyboardDevice.Modifiers & ModifierKeys.Shift):
            e.Handled = True
            self._send_message()

    def _on_send(self, sender, e):
        self._send_message()

    def _on_clear(self, sender, e):
        self._messages = []
        if self._chat_panel:
            self._chat_panel.Children.Clear()
        self._stream_tb = None

    def _on_open_full(self, sender, e):
        """Open the full floating T3Lab Assistant window."""
        try:
            import imp
            tab_dir     = os.path.join(_EXT_DIR, 'T3Lab.tab')
            script_path = os.path.join(
                tab_dir, 'Support.panel', 'T3LabAssistant.pushbutton', 'script.py'
            )
            if os.path.isfile(script_path):
                mod = imp.load_source('t3lab_assistant_full', script_path)
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
        text = (self._chat_input.Text or u'').strip()
        if not text:
            return

        self._chat_input.Text = u''
        self._add_message_bubble(text, is_user=True)
        self._messages.append({'role': 'user', 'content': text})
        self._set_loading(True)

        # Gather BIM context on current thread before handing off
        bim_ctx = self._get_bim_context()

        from System.Threading import Thread, ThreadStart
        t = Thread(ThreadStart(lambda: self._ai_call(text, bim_ctx)))
        t.IsBackground = True
        t.Start()

    def _ai_call(self, user_text, bim_ctx):
        """Background thread: call AI (with streaming if supported), post result to UI."""
        ok, has_key, route_msg, route_stream, _, _ = _try_import_ai()

        if not ok:
            self._post_to_ui(u'AI modules not available. Check your installation.')
            return

        if not has_key():
            self._post_to_ui(
                u'No API key configured. Click "Open full assistant" and set your key in Settings.'
            )
            return

        # Prefix BIM context so the AI knows what's happening in Revit
        if bim_ctx:
            augmented = u'[Revit context: {}]\n\n{}'.format(bim_ctx, user_text)
        else:
            augmented = user_text

        history = list(self._messages[:-1])   # exclude the just-appended user turn

        # ── Try streaming first ────────────────────────────────────────────────
        if route_stream and System_Action:
            full_text = [u'']
            bubble_created = [False]

            def _on_delta(chunk):
                if not chunk:
                    return
                full_text[0] += chunk

                def _update_ui():
                    if not bubble_created[0]:
                        bubble_created[0] = True
                        self._set_loading(False)
                        self._add_message_bubble(chunk, is_user=False, live=True)
                    else:
                        if self._stream_tb is not None:
                            self._stream_tb.Text = full_text[0]
                            self._scroll_to_bottom()

                try:
                    if self._control and self._control.Dispatcher:
                        self._control.Dispatcher.BeginInvoke(System_Action(_update_ui))
                except Exception:
                    pass

            reply = route_stream(augmented, _on_delta, history=history, max_tokens=600)

            if reply:
                # Finalise: ensure bubble shows the complete text
                final = reply

                def _finalise():
                    if not bubble_created[0]:
                        self._set_loading(False)
                        self._add_message_bubble(final, is_user=False)
                    elif self._stream_tb is not None:
                        self._stream_tb.Text = final
                        self._scroll_to_bottom()
                    self._stream_tb = None
                    self._messages.append({'role': 'assistant', 'content': final})

                if self._control and self._control.Dispatcher:
                    self._control.Dispatcher.BeginInvoke(System_Action(_finalise))
                return

        # ── Fallback: non-streaming ────────────────────────────────────────────
        try:
            result = route_msg(augmented, history=history, max_tokens=600)
            if isinstance(result, dict):
                reply = result.get('message') or result.get('answer') or str(result)
            elif isinstance(result, str):
                reply = result
            else:
                reply = u'(no response)'
        except Exception as ex:
            reply = u'Error: {}'.format(ex)

        self._post_to_ui(reply)

    def _post_to_ui(self, reply):
        """Marshal reply back to UI thread and display it."""
        def _show():
            self._set_loading(False)
            self._add_message_bubble(reply, is_user=False)
            self._messages.append({'role': 'assistant', 'content': reply})

        if self._control and self._control.Dispatcher and System_Action:
            self._control.Dispatcher.BeginInvoke(System_Action(_show))
        else:
            _show()

    # ── Public API ─────────────────────────────────────────────────────────────

    def add_message(self, text, is_user=True):
        """External API to inject a message (e.g. from MCP command results)."""
        self._add_message_bubble(text, is_user=is_user)
        self._messages.append({
            'role': 'user' if is_user else 'assistant',
            'content': text
        })

    def clear(self):
        """Clear the chat history and panel."""
        self._on_clear(None, None)


# ─── Corner radius helper (IronPython 2.7 compat) ──────────────────────────────

def _corner_radius(r):
    from System.Windows import CornerRadius
    return CornerRadius(r)


# ─── IDockablePaneProvider ─────────────────────────────────────────────────────

class AssistantPaneProvider(IDockablePaneProvider):
    """
    Revit calls SetupDockablePane() the first time the pane is shown.
    We load the UserControl XAML here and attach the controller.
    """

    def SetupDockablePane(self, data):
        try:
            import imp
            
            # Load the pushbutton script.py as a module to get T3LabAssistantWindow
            tab_dir = os.path.join(_EXT_DIR, 'T3Lab.tab')
            script_path = os.path.join(
                tab_dir, 'Support.panel', 'T3LabAssistant.pushbutton', 'script.py'
            )
            if os.path.isfile(script_path):
                if _LIB_DIR not in sys.path:
                    sys.path.insert(0, _LIB_DIR)
                if _EXT_DIR not in sys.path:
                    sys.path.insert(0, _EXT_DIR)
                    
                mod = imp.load_source('t3lab_assistant_full', script_path)
                if hasattr(mod, 'T3LabAssistantWindow'):
                    # Instantiate on UI thread as docked
                    win = mod.T3LabAssistantWindow(is_docked=True)
                    
                    # Detach visual content
                    content = win.Content
                    win.Content = None
                    
                    # Keep the window class instance alive
                    self._win_ref = win
                    
                    data.FrameworkElement = content
                    
                    from Autodesk.Revit.UI import EditorInteraction, EditorInteractionType
                    data.EditorInteraction = EditorInteraction(EditorInteractionType.KeepAlive)
                    return
            
            raise Exception("pushbutton script.py not found")

        except Exception as ex:
            import logging
            logging.basicConfig()
            logger = logging.getLogger("T3LabAssistant")
            logger.error("Error setting up DockablePane: %s", ex, exc_info=True)

            from System.Windows.Controls import Border, TextBlock
            from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness, TextWrapping
            from System.Windows.Media import Brushes

            border = Border()
            border.Background = Brushes.Crimson
            border.Padding = Thickness(20)

            lbl = TextBlock()
            lbl.Text = u'T3Lab Assistant pane could not load:\n{}'.format(ex)
            lbl.Foreground = Brushes.White
            lbl.TextWrapping = TextWrapping.Wrap
            lbl.HorizontalAlignment = HorizontalAlignment.Center
            lbl.VerticalAlignment   = VerticalAlignment.Center

            border.Child = lbl
            data.FrameworkElement   = border
