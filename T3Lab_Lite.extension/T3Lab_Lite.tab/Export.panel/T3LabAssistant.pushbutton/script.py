# -*- coding: utf-8 -*-
"""T3Lab Assistant - Standalone AI chatbox for T3Lab tools.

Provides a floating chatbox window that understands natural-language commands
(Tiếng Việt or English) and quickly opens any T3Lab tool.
"""

__title__ = "T3Lab\nAssistant"
__author__ = "T3Lab"
__version__ = "1.0.0"

import json
import os
import sys
import re
import datetime

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility, WindowState, GridLength
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Action
from System.Threading import Thread, ThreadStart, ApartmentState

from pyrevit import revit, forms, script

logger = script.get_logger()

# ─── Lib path setup ───────────────────────────────────────────────────────────
# __file__ → .../T3LabAssistant.pushbutton/script.py
# extension_dir → .../T3Lab_Lite.extension
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# ─── NLP module ───────────────────────────────────────────────────────────────
try:
    from t3lab_assistant import (parse_command, has_api_key, keyword_parse,
                                  learn_pattern, find_learned_match,
                                  has_local_llm, parse_command_local,
                                  get_local_model_name, parse_command_nlu,
                                  inject_discovered_tools)
    HAS_NLP = True
except Exception as e:
    logger.warning("Could not import t3lab_assistant: {}".format(e))
    HAS_NLP = False
    def learn_pattern(*a, **kw): pass
    def find_learned_match(*a, **kw): return None
    def has_local_llm(*a, **kw): return False
    def parse_command_local(*a, **kw): return None
    def get_local_model_name(*a, **kw): return None
    def parse_command_nlu(*a, **kw): return None
    def inject_discovered_tools(*a, **kw): pass

# ─── Tool discovery module ────────────────────────────────────────────────────
try:
    from tool_discovery import (discover_new_tools, get_registered_tools,
                                 make_generic_launcher)
    HAS_DISCOVERY = True
except Exception as e:
    logger.warning("Could not import tool_discovery: {}".format(e))
    HAS_DISCOVERY = False
    def discover_new_tools(): return []
    def get_registered_tools(): return []
    def make_generic_launcher(script_path, title): return lambda: False

# ─── BatchOut executor (configure + direct export) ────────────────────────────
try:
    from batchout_executor import configure_batchout_window, direct_export
    HAS_EXECUTOR = True
except Exception as e:
    logger.warning("Could not import batchout_executor: {}".format(e))
    HAS_EXECUTOR = False

# ─── Tool launchers ───────────────────────────────────────────────────────────
# Each function opens the corresponding T3Lab tool.

def _get_tool_script_dir(panel, pushbutton):
    """Return the path to a pushbutton script.py given panel and pushbutton names."""
    # __file__ = .../T3Lab_Lite.tab/Export.panel/T3LabAssistant.pushbutton/script.py
    # dirname x1 = T3LabAssistant.pushbutton/
    # dirname x2 = Export.panel/
    # dirname x3 = T3Lab_Lite.tab/
    tab_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    return os.path.join(tab_dir, panel, pushbutton, 'script.py')


def _load_script(name, script_path):
    """Load a tool script as a module. Works in both CPython and IronPython."""
    try:
        import imp
        return imp.load_source(name, script_path)
    except ImportError:
        pass
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(name, script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    except Exception:
        pass
    return None


def _load_batchout_mod():
    """Load the BatchOut script module, raising RuntimeError on failure."""
    script_path = _get_tool_script_dir('Export.panel', 'BatchOut.pushbutton')
    mod = _load_script('batchout_script', script_path)
    if mod is None:
        raise RuntimeError("Could not load BatchOut module from: {}".format(script_path))
    return mod


def launch_batchout():
    """Open the BatchOut export dialog (no pre-configuration)."""
    try:
        mod = _load_batchout_mod()
        window = mod.ExportManagerWindow()
        window.ShowDialog()
        return True
    except Exception as ex:
        logger.error("Error launching BatchOut: {}".format(ex))
        return False


def launch_batchout_configured(config, progress_cb=None):
    """Open BatchOut pre-configured: sheets selected, format set, tab = Create.

    Args:
        config: dict with keys format, filter (from batchout_executor / NLP).
        progress_cb: optional callable(str) for status messages.
    Returns:
        bool success
    """
    try:
        mod = _load_batchout_mod()
        window = mod.ExportManagerWindow()

        if HAS_EXECUTOR:
            configure_batchout_window(window, config)
            fmt    = (config.get('format') or 'pdf').upper()
            filt   = config.get('filter') or ''
            filt_s = u" {} sheet".format(filt) if filt else u" tất cả sheet"
            if progress_cb:
                progress_cb(u"BatchOut đã chọn{}, format {} — nhấn Export để xuất.".format(
                    filt_s, fmt))

        window.ShowDialog()
        return True
    except Exception as ex:
        logger.error("Error launching configured BatchOut: {}".format(ex))
        if progress_cb:
            progress_cb(u"Lỗi: {}".format(ex))
        return False


def launch_export_direct(config, progress_cb=None):
    """Export sheets directly without showing BatchOut UI.

    Args:
        config: dict with format, filter, folder (optional).
        progress_cb: optional callable(str) for chat progress updates.
    Returns:
        bool success
    """
    try:
        if not HAS_EXECUTOR:
            raise RuntimeError("batchout_executor not available")
        mod = _load_batchout_mod()
        ok, count, msg = direct_export(mod, config, progress_cb)
        return ok
    except Exception as ex:
        logger.error("Error in direct export: {}".format(ex))
        if progress_cb:
            progress_cb(u"Lỗi xuất file: {}".format(ex))
        return False


def launch_parasync():
    """Open the ParaSync parameter sync tool."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'ParaSync.pushbutton')
        mod = _load_script('parasync_script', script_path)
        if mod is None:
            raise RuntimeError("Could not load ParaSync module from: {}".format(script_path))
        window = mod.ParaSyncWindow()
        window.ShowDialog()
        return True
    except Exception as ex:
        logger.error("Error launching ParaSync: {}".format(ex))
        return False


def launch_loadfamily():
    """Open the Load Family dialog."""
    try:
        from GUI.FamilyLoaderDialog import show_family_loader
        show_family_loader()
        return True
    except Exception as ex:
        logger.error("Error launching LoadFamily: {}".format(ex))
        return False


def launch_loadfamily_cloud():
    """Open the Load Family (Cloud) dialog."""
    try:
        from GUI.FamilyLoaderCloudDialog import show_family_loader_cloud
        show_family_loader_cloud()
        return True
    except Exception as ex:
        logger.error("Error launching LoadFamily Cloud: {}".format(ex))
        return False


def launch_projectname():
    """Open the Project Name tool."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'ProjectName.pushbutton')
        mod = _load_script('projectname_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching ProjectName: {}".format(ex))
        return False


def launch_workset():
    """Open the Workset manager."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'Workset.pushbutton')
        mod = _load_script('workset_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching Workset: {}".format(ex))
        return False


def launch_dimtext():
    """Run the Dim Text tool on current selection."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'DimText.pushbutton')
        mod = _load_script('dimtext_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching DimText: {}".format(ex))
        return False


def launch_upperdimtext():
    """Run the Upper Dim Text tool on current selection."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'UpperDimText.pushbutton')
        mod = _load_script('upperdimtext_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching UpperDimText: {}".format(ex))
        return False


def launch_resetoverrides():
    """Run the Reset Overrides tool on the active view."""
    try:
        script_path = _get_tool_script_dir('Graphic.panel', 'Reset Overrides.pushbutton')
        mod = _load_script('resetoverrides_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching Reset Overrides: {}".format(ex))
        return False


# Map intent → launcher function
def _is_viet_text(text):
    viet_chars = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
                  u"ụủứừửữựỳỵỷỹ")
    return any(c in viet_chars for c in text.lower())


TOOL_LAUNCHERS = {
    "open_batchout":         launch_batchout,
    "open_parasync":         launch_parasync,
    "open_loadfamily":       launch_loadfamily,
    "open_loadfamily_cloud": launch_loadfamily_cloud,
    "open_projectname":      launch_projectname,
    "open_workset":          launch_workset,
    "open_dimtext":          launch_dimtext,
    "open_upperdimtext":     launch_upperdimtext,
    "open_resetoverrides":   launch_resetoverrides,
}


def _register_discovered_launchers(tools):
    """
    For each auto-discovered tool, add a generic launcher to TOOL_LAUNCHERS
    and update the NLP module's system prompt.

    Args:
        tools: list of tool dicts from discover_new_tools() / get_registered_tools()
    """
    for tool in tools:
        intent = tool.get('intent')
        if not intent or intent in TOOL_LAUNCHERS:
            continue
        launcher = make_generic_launcher(tool['script_path'], tool['title'])
        TOOL_LAUNCHERS[intent] = launcher

    # Inject all registered tools (new + old) into the NLP system prompt
    if HAS_NLP:
        try:
            inject_discovered_tools(get_registered_tools())
        except Exception:
            pass


# ─── Chat history persistence ─────────────────────────────────────────────────

def _get_doc_key():
    """Return a filesystem-safe key for the current Revit document."""
    try:
        title = revit.doc.Title or "untitled"
        # Strip chars that are invalid in filenames
        safe = re.sub(r'[\\/:*?"<>|]', '_', title)
        return safe[:80]   # cap at 80 chars
    except Exception:
        return "default"


def _history_file(doc_key):
    """Return path to the JSON history file for doc_key."""
    config_dir = os.path.join(lib_dir, 'config', 'chat_history')
    if not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir)
        except Exception:
            pass
    return os.path.join(config_dir, '{}.json'.format(doc_key))


def save_chat_history(doc_key, messages):
    """Persist the last N messages to disk for this document.

    Args:
        doc_key  : identifier returned by _get_doc_key()
        messages : list of {role, content, ts} dicts
    """
    try:
        path = _history_file(doc_key)
        # Keep only the last 60 messages
        to_save = messages[-60:]
        with open(path, 'w') as f:
            json.dump({"doc_key": doc_key, "messages": to_save}, f,
                      ensure_ascii=False, indent=2)
    except Exception as ex:
        logger.debug("Could not save chat history: {}".format(ex))


def load_chat_history(doc_key):
    """Load saved messages for doc_key.  Returns [] if none / error."""
    try:
        path = _history_file(doc_key)
        if not os.path.exists(path):
            return []
        with open(path, 'r') as f:
            data = json.load(f)
        return data.get("messages", [])
    except Exception as ex:
        logger.debug("Could not load chat history: {}".format(ex))
        return []


def clear_chat_history(doc_key):
    """Delete the saved history file for doc_key."""
    try:
        path = _history_file(doc_key)
        if os.path.exists(path):
            os.remove(path)
    except Exception as ex:
        logger.debug("Could not clear chat history: {}".format(ex))


# ─── WPF Window ───────────────────────────────────────────────────────────────

class T3LabAssistantWindow(forms.WPFWindow):
    """Standalone T3Lab Assistant chatbox window."""

    # ── Quick-button names (for bulk enable/disable) ──────────────────────────
    _QUICK_BTNS = ['btn_batchout', 'btn_parasync', 'btn_loadfamily',
                   'btn_projectname', 'btn_workset', 'btn_dimtext']
    # Dynamic buttons added by _bootstrap_discovered_tools
    _DYNAMIC_BTNS = []   # list of Button WPF objects (not names)

    def __init__(self):
        xaml_path = os.path.join(lib_dir, 'GUI', 'T3LabAssistant.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        # ── Session state ─────────────────────────────────────────────────────
        self._busy             = False          # concurrency guard
        self._typing_row       = None           # reference to typing indicator element
        self._conversation_history = []         # [{role, content}, ...] multi-turn context
        self._last_raw         = ''             # last user input (for learning)
        self._doc_key          = _get_doc_key() # document identifier for history
        self._persisted_msgs   = []             # flat list with timestamps, for save/load

        # ── Logo ──────────────────────────────────────────────────────────────
        try:
            logo_path = os.path.join(lib_dir, 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(logo_path, UriKind.Absolute)
                bmp.EndInit()
                self.logo_image.Source = bmp
                self.Icon = bmp
        except Exception:
            pass

        # ── Restore conversation from previous session ─────────────────────────
        self._restore_history()

        # ── Tool discovery (background, then inject chips into UI) ─────────────
        def _discover_and_update():
            import time; time.sleep(0.3)
            self.Dispatcher.Invoke(Action(self._bootstrap_discovered_tools))

        _dt = Thread(ThreadStart(_discover_and_update))
        _dt.IsBackground = True
        _dt.SetApartmentState(ApartmentState.STA)
        _dt.Start()

        # Update AI badge (run async so Ollama probe doesn't delay window open)
        def _probe_and_update():
            import time; time.sleep(0.5)   # let window render first
            self.Dispatcher.Invoke(Action(self._update_ai_badge))

        _t = Thread(ThreadStart(_probe_and_update))
        _t.IsBackground = True
        _t.SetApartmentState(ApartmentState.STA)
        _t.Start()

    # ─── Window controls ──────────────────────────────────────────────────────

    def minimize_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def close_clicked(self, sender, e):
        self.Close()

    # ─── Tool discovery bootstrap ──────────────────────────────────────────────

    def _bootstrap_discovered_tools(self):
        """
        Run on startup (UI thread):
          1. Discover new tools → register launchers → update NLP prompt.
          2. Render chip buttons for ALL registered tools (new + existing).
          3. Show the "CÔNG CỤ MỚI" section if anything is found.
          4. Post a chat notification for truly NEW tools.
        Must be called from the UI thread (via Dispatcher.Invoke).
        """
        try:
            if not HAS_DISCOVERY:
                return

            # ── Discover (writes registry) ────────────────────────────────────
            new_tools = discover_new_tools()

            # ── Register launchers + inject into NLP ─────────────────────────
            _register_discovered_launchers(new_tools)

            # ── Render all registered tools in discovered_tools_wrap ─────────
            all_tools = get_registered_tools()
            if not all_tools:
                return

            # Also register launchers for tools already in registry from previous runs
            _register_discovered_launchers(all_tools)

            try:
                self.discovered_tools_wrap.Children.Clear()
                for tool in all_tools:
                    btn = self._create_tool_chip(tool)
                    self.discovered_tools_wrap.Children.Add(btn)
                    self._DYNAMIC_BTNS.append(btn)

                # Show the section
                self.discovered_section.Visibility = Visibility.Visible

                # Update badge count
                self.new_tools_count_text.Text = str(len(all_tools))
            except Exception as ui_ex:
                logger.debug("UI chip error: {}".format(ui_ex))

            # ── Chat notification for NEW tools only ──────────────────────────
            if new_tools:
                names = u', '.join(t['title'] for t in new_tools[:5])
                if len(new_tools) > 5:
                    names += u'...'
                self._append_bot_message(
                    u"🔍 Phát hiện {} công cụ mới: {}.\n"
                    u"Tôi đã tự học và có thể mở chúng bằng lệnh tự nhiên.".format(
                        len(new_tools), names)
                )
        except Exception as ex:
            logger.debug("_bootstrap_discovered_tools error: {}".format(ex))

    def _create_tool_chip(self, tool):
        """Build a WPF Button chip for an auto-discovered tool."""
        from System.Windows.Controls import Button, StackPanel, TextBlock
        from System.Windows import Thickness
        from System.Windows.Media import SolidColorBrush, Color

        intent = tool['intent']
        title  = tool['title']

        btn = Button()
        btn.Cursor   = System.Windows.Input.Cursors.Hand
        btn.Margin   = Thickness(0, 0, 6, 6)
        btn.Padding  = Thickness(10, 5, 10, 5)
        btn.FontSize = 12

        # Style: green border (new tool)
        from System.Windows import CornerRadius
        from System.Windows.Controls import Border, ContentPresenter
        # Use the NewChipBtn style via FindResource
        try:
            btn.Style = self.FindResource('NewChipBtn')
        except Exception:
            btn.Background = SolidColorBrush(Color.FromRgb(240, 251, 244))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(39, 174, 96))
            btn.Foreground  = SolidColorBrush(Color.FromRgb(39, 174, 96))

        sp = StackPanel()
        sp.Orientation = System.Windows.Controls.Orientation.Horizontal

        icon = TextBlock()
        icon.Text    = u"✦"
        icon.FontSize = 10
        icon.Margin   = Thickness(0, 0, 5, 0)
        icon.VerticalAlignment = System.Windows.VerticalAlignment.Center
        sp.Children.Add(icon)

        lbl = TextBlock()
        lbl.Text = title
        lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
        sp.Children.Add(lbl)

        btn.Content  = sp
        btn.ToolTip  = u"Mở {} (auto-discovered)".format(title)

        # Capture intent in closure
        _intent = intent

        def _on_click(s, e, _i=_intent, _t=title):
            self._run_tool(_i, u"Đang mở {}...".format(_t))

        btn.Click += _on_click
        return btn

    # ─── History persistence ───────────────────────────────────────────────────

    def _restore_history(self):
        """Load saved conversation from disk and replay bubbles + context."""
        try:
            saved = load_chat_history(self._doc_key)
            if not saved:
                return

            # Replay the last 30 messages (15 exchanges) as bubbles
            for msg in saved[-30:]:
                role    = msg.get("role", "")
                content = msg.get("content", "")
                if not content:
                    continue
                if role == "user":
                    self._append_user_message(content)
                elif role == "assistant":
                    self._append_bot_message(content)
                # Re-populate NLP context (last 16 messages = 8 exchanges)
                self._conversation_history.append(
                    {"role": role, "content": content}
                )

            self._persisted_msgs = list(saved)

            # Show a separator so user knows this is a restored session
            self._append_bot_message(
                u"── Đã khôi phục cuộc trò chuyện trước ──\n"
                u"Nhấn ↺ để bắt đầu cuộc hội thoại mới."
            )
        except Exception as ex:
            logger.debug("Could not restore history: {}".format(ex))

    def _persist_message(self, role, content):
        """Append one message to the in-memory list and save to disk."""
        try:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._persisted_msgs.append(
                {"role": role, "content": content, "ts": ts}
            )
            save_chat_history(self._doc_key, self._persisted_msgs)
        except Exception as ex:
            logger.debug("Could not persist message: {}".format(ex))

    def reset_chat_clicked(self, sender, e):
        """Clear the chat history for this document and reset the UI."""
        try:
            # Remove all chat bubbles
            self.chat_history_panel.Children.Clear()
            # Clear in-memory state
            self._conversation_history = []
            self._persisted_msgs = []
            # Delete saved file
            clear_chat_history(self._doc_key)
            # Show fresh welcome message
            self._append_bot_message(
                u"Cuộc trò chuyện đã được làm mới! 👋\n"
                u"Tôi có thể giúp gì cho bạn?"
            )
        except Exception as ex:
            logger.debug("reset_chat error: {}".format(ex))

    # ─── AI badge ─────────────────────────────────────────────────────────────

    def _update_ai_badge(self):
        """Show badge: LOCAL (Ollama) > AI (Claude) > hidden."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            if HAS_NLP and has_local_llm():
                model = get_local_model_name() or "LOCAL"
                # Show model's short name: "qwen2.5:0.5b" → "qwen0.5b"
                short = model.replace("qwen2.5", "qwen").replace("llama3.2", "llama").replace("phi3", "phi")
                self.ai_status_badge.Visibility = Visibility.Visible
                self.ai_status_badge.Background = SolidColorBrush(
                    Color.FromRgb(39, 174, 96))   # #27AE60 green
                self.ai_status_text.Text = "LOCAL"
                self.ai_status_badge.ToolTip = u"Local LLM: {}".format(model)
            elif HAS_NLP and has_api_key():
                self.ai_status_badge.Visibility = Visibility.Visible
                self.ai_status_badge.Background = SolidColorBrush(
                    Color.FromRgb(52, 152, 219))  # #3498DB blue
                self.ai_status_text.Text = "AI"
                self.ai_status_badge.ToolTip = "Claude AI"
            else:
                self.ai_status_badge.Visibility = Visibility.Collapsed
        except Exception:
            pass

    # ─── Session guard & UI state ─────────────────────────────────────────────

    def _set_busy(self, busy):
        """Lock/unlock the whole input area. Call from UI thread only."""
        self._busy = busy
        try:
            self.send_button.IsEnabled = not busy
            self.chat_input.IsEnabled  = not busy
        except Exception:
            pass
        for btn_name in self._QUICK_BTNS:
            try:
                getattr(self, btn_name).IsEnabled = not busy
            except Exception:
                pass
        for btn in self._DYNAMIC_BTNS:
            try:
                btn.IsEnabled = not busy
            except Exception:
                pass
        if busy:
            self._show_typing_indicator()
        else:
            self._hide_typing_indicator()

    def _show_typing_indicator(self):
        """Add an animated '● ● ●' bubble to the chat."""
        try:
            if self._typing_row is not None:
                return  # already shown
            self._typing_row = self._make_typing_row()
            self.chat_history_panel.Children.Add(self._typing_row)
            self._scroll_to_bottom()
        except Exception:
            pass

    def _hide_typing_indicator(self):
        """Remove the typing indicator bubble."""
        try:
            if self._typing_row is not None:
                self.chat_history_panel.Children.Remove(self._typing_row)
                self._typing_row = None
        except Exception:
            pass

    def _make_typing_row(self):
        """Build the '● ● ●' typing indicator WPF element."""
        from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
        from System.Windows import Thickness, CornerRadius, GridLength, HorizontalAlignment
        from System.Windows.Media import SolidColorBrush, Color

        row = Grid()
        row.Margin = Thickness(0, 0, 60, 6)

        col_av = ColumnDefinition()
        col_av.Width = GridLength.Auto
        col_msg = ColumnDefinition()
        col_msg.Width = GridLength(1, System.Windows.GridUnitType.Star)
        row.ColumnDefinitions.Add(col_av)
        row.ColumnDefinitions.Add(col_msg)

        av = self._make_avatar("T3")
        Grid.SetColumn(av, 0)
        row.Children.Add(av)

        bubble = Border()
        bubble.Background      = SolidColorBrush(Color.FromRgb(255, 255, 255))
        bubble.CornerRadius    = CornerRadius(3, 8, 8, 8)
        bubble.Padding         = Thickness(14, 10, 14, 10)
        bubble.BorderBrush     = SolidColorBrush(Color.FromRgb(189, 195, 199))  # #BDC3C7
        bubble.BorderThickness = Thickness(1)
        bubble.HorizontalAlignment = HorizontalAlignment.Left

        dots = TextBlock()
        dots.Text      = u"● ● ●"
        dots.FontSize  = 10
        dots.Foreground = SolidColorBrush(Color.FromRgb(127, 140, 141))  # #7F8C8D

        bubble.Child = dots
        Grid.SetColumn(bubble, 1)
        row.Children.Add(bubble)
        return row

    def _safe_append_bot(self, msg):
        """Thread-safe bot message append (can be called from background threads)."""
        try:
            self.Dispatcher.Invoke(Action(lambda: self._append_bot_message(msg)))
        except Exception:
            pass

    def _add_to_history(self, role, content):
        """Add a message to conversation history and persist to disk."""
        self._conversation_history.append({"role": role, "content": content})
        if len(self._conversation_history) > 16:
            self._conversation_history = self._conversation_history[-16:]
        # Persist to disk so it survives window close/reopen
        self._persist_message(role, content)

    # ─── Quick tool buttons ───────────────────────────────────────────────────

    def tool_batchout_clicked(self, sender, e):
        self._run_tool("open_batchout", u"Đang mở BatchOut...")

    def tool_parasync_clicked(self, sender, e):
        self._run_tool("open_parasync", "Đang mở ParaSync...")

    def tool_loadfamily_clicked(self, sender, e):
        self._run_tool("open_loadfamily", "Đang mở Load Family...")

    def tool_projectname_clicked(self, sender, e):
        self._run_tool("open_projectname", "Đang mở Project Name...")

    def tool_workset_clicked(self, sender, e):
        self._run_tool("open_workset", "Đang mở Workset...")

    def tool_dimtext_clicked(self, sender, e):
        self._run_tool("open_dimtext", "Đang mở Dim Text...")

    # ─── Chat input ───────────────────────────────────────────────────────────

    def send_clicked(self, sender, e):
        self._process_input()

    def input_keydown(self, sender, e):
        from System.Windows.Input import Key
        if e.Key == Key.Return or e.Key == Key.Enter:
            self._process_input()

    def _process_input(self):
        """Read input, dispatch to NLP or keyword fallback, execute result."""
        try:
            raw = self.chat_input.Text.strip()
            if not raw:
                return

            # ── Concurrency guard ─────────────────────────────────────────────
            if self._busy:
                self._append_bot_message(
                    u"⏳ Đang xử lý lệnh trước, vui lòng chờ một chút..."
                )
                return

            self.chat_input.Text = ""
            self._last_raw = raw

            # Show user message in chat + add to history
            self._append_user_message(raw)
            self._add_to_history("user", raw)

            # Lock UI
            self._set_busy(True)

            # ── 1. Learned patterns (instant, no service needed) ──────────────
            if HAS_NLP:
                learned = find_learned_match(raw)
                if learned:
                    self._execute_result(learned)
                    return

            captured = raw
            history  = list(self._conversation_history[:-1])

            use_local  = HAS_NLP and has_local_llm()
            use_claude = HAS_NLP and has_api_key()

            # ── 2. Built-in NLU (always available, no service needed) ─────────
            # Try synchronously first — it's instant (pure Python, no I/O).
            if HAS_NLP:
                nlu_result = parse_command_nlu(captured, history)
                if nlu_result and nlu_result.get("intent") not in (None, "unknown"):
                    # NLU succeeded → execute immediately
                    # But if Ollama or Claude is also available, let them handle
                    # open-ended "chat" / "help" intents for richer responses.
                    if nlu_result["intent"] not in ("chat", "help") \
                            or not (use_local or use_claude):
                        self._execute_result(nlu_result)
                        return

            if use_local or use_claude:
                # ── 3/4. Async LLM path for chat/help or NLU miss ─────────────
                self._show_typing_indicator()
                nlu_hint = nlu_result if HAS_NLP else None   # pass NLU result as hint

                def do_nlp():
                    result = None

                    # ── Priority A: Local Ollama LLM ──────────────────────────
                    if use_local:
                        try:
                            result = parse_command_local(captured, history)
                        except Exception as le:
                            logger.debug("local_llm error: {}".format(le))

                    # ── Priority B: Claude API ─────────────────────────────────
                    if (result is None or result.get("intent") in (None, "unknown")) \
                            and use_claude:
                        try:
                            result = parse_command(captured, history)
                        except Exception as ce:
                            logger.debug("claude api error: {}".format(ce))

                    def finish():
                        try:
                            self._hide_typing_indicator()
                            if result and result.get("intent") not in (None, "unknown"):
                                self._execute_result(result)
                            elif nlu_hint and nlu_hint.get("intent") not in (None, "unknown"):
                                # Fall back to the NLU result we computed earlier
                                self._execute_result(nlu_hint)
                            else:
                                fb = keyword_parse(captured)
                                if fb:
                                    self._execute_result(fb)
                                else:
                                    msg = (u"Không hiểu lệnh. Thử: 'mở batchout', 'xuất pdf G sheet'..."
                                           if _is_viet_text(captured) else
                                           "I didn't understand. Try: 'open batchout', 'export pdf G sheets'...")
                                    self._append_bot_message(msg)
                                    self._set_busy(False)
                        except Exception as finish_ex:
                            logger.error("finish error: {}".format(finish_ex))
                            self._hide_typing_indicator()
                            self._set_busy(False)

                    self.Dispatcher.Invoke(Action(finish))

                t = Thread(ThreadStart(do_nlp))
                t.IsBackground = True
                t.SetApartmentState(ApartmentState.STA)
                t.Start()
            else:
                # ── 5. Keyword fallback if NLU also missed ────────────────────
                fb = keyword_parse(raw)
                if fb:
                    self._execute_result(fb)
                else:
                    self._append_bot_message(
                        u"Không hiểu lệnh.\n"
                        u"Ví dụ: 'mở batchout', 'xuất pdf G sheet', 'parasync'"
                    )
                    self._set_busy(False)

        except Exception as ex:
            logger.error("Error in _process_input: {}".format(ex))
            self._set_busy(False)

    # ─── Execute intent ────────────────────────────────────────────────────────

    def _execute_result(self, result):
        """Execute the action described by a parsed result dict.

        Responsibilities:
        - Display bot message
        - Add bot reply to conversation history
        - Learn successful patterns
        - Release busy state when done (including after background exports)
        """
        intent  = result.get("intent", "unknown")
        message = result.get("message", "")
        params  = result.get("params", {})
        raw     = self._last_raw

        def _bot(msg):
            """Show message and record in conversation history."""
            self._append_bot_message(msg)
            self._add_to_history("assistant", msg)

        def _learn(msg=''):
            """Record successful command→intent mapping."""
            learn_pattern(raw, intent, params, msg)

        # ── Conversation (no action needed) ──────────────────────────────────
        if intent in ("help", "chat", "greet"):
            reply = params.get("answer", message) if intent == "help" else message
            _bot(reply or u"Có thể giúp gì thêm không?")
            self._set_busy(False)
            return

        # ── Export directly — runs on background thread ───────────────────────
        if intent == "export_direct":
            confirm = message or u"Đang xuất file, vui lòng chờ..."
            _bot(confirm)
            _learn(confirm)

            def do_export():
                ok = launch_export_direct(params, self._safe_append_bot)
                if not ok:
                    self._safe_append_bot(u"Xuất thất bại. Xem console để biết lỗi.")
                self.Dispatcher.Invoke(Action(lambda: self._set_busy(False)))

            t = Thread(ThreadStart(do_export))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
            return

        # ── Open BatchOut pre-configured ──────────────────────────────────────
        if intent == "open_batchout_configured":
            confirm = message or u"Đang mở BatchOut đã cấu hình..."
            _bot(confirm)
            _learn(confirm)
            ok = launch_batchout_configured(params, self._safe_append_bot)
            if not ok:
                self._append_bot_message(u"Không thể mở BatchOut. Xem console.")
            self._set_busy(False)
            return

        # ── Simple tool launchers ─────────────────────────────────────────────
        if intent in TOOL_LAUNCHERS:
            confirm = message or u"Đang mở công cụ..."
            _bot(confirm)
            _learn(confirm)
            ok = TOOL_LAUNCHERS[intent]()
            if not ok:
                self._append_bot_message(u"Không thể mở công cụ. Xem console.")
            self._set_busy(False)
            return

        # ── Unknown / fallthrough ─────────────────────────────────────────────
        if intent == "unknown":
            _bot(params.get("message", u"Lệnh không rõ. Thử: 'mở batchout', 'xuất pdf G sheet'..."))
        else:
            _bot(message or u"Đã thực hiện.")
        self._set_busy(False)

    def _run_tool(self, intent, default_msg):
        """Helper for quick-button clicks: guard, show message, run launcher."""
        if self._busy:
            self._append_bot_message(u"⏳ Đang xử lý lệnh trước, vui lòng chờ...")
            return
        self._set_busy(True)
        self._last_raw = default_msg
        self._append_bot_message(default_msg)
        self._add_to_history("assistant", default_msg)
        launcher = TOOL_LAUNCHERS.get(intent)
        if launcher:
            ok = launcher()
            if not ok:
                self._append_bot_message(u"Không thể mở công cụ. Xem console.")
        self._set_busy(False)

    # ─── Chat UI helpers ──────────────────────────────────────────────────────

    def _make_avatar(self, letter, _unused_start=None, _unused_end=None):
        """Create a circular avatar Border with initials (BatchOut blue #3498DB)."""
        from System.Windows.Controls import Border, TextBlock
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows import HorizontalAlignment, VerticalAlignment

        av = Border()
        av.Width = 36
        av.Height = 36
        av.CornerRadius = CornerRadius(18)
        av.Background = SolidColorBrush(Color.FromRgb(52, 152, 219))   # #3498DB
        av.Margin = Thickness(0, 2, 10, 0)
        av.VerticalAlignment = VerticalAlignment.Top

        lbl = TextBlock()
        lbl.Text = letter
        lbl.FontSize = 12
        lbl.FontWeight = System.Windows.FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment = VerticalAlignment.Center
        av.Child = lbl
        return av

    def _append_user_message(self, text):
        """Add a right-aligned user bubble (BatchOut #3498DB)."""
        try:
            from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength, HorizontalAlignment
            from System.Windows.Media import SolidColorBrush, Color

            row = Grid()
            row.Margin = Thickness(60, 0, 0, 10)
            col0 = ColumnDefinition()
            col0.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col0)

            bubble = Border()
            bubble.Background   = SolidColorBrush(Color.FromRgb(52, 152, 219))   # #3498DB
            bubble.CornerRadius = CornerRadius(8, 3, 8, 8)
            bubble.Padding      = Thickness(12, 8, 12, 8)
            bubble.HorizontalAlignment = HorizontalAlignment.Right

            msg_text = TextBlock()
            msg_text.Text        = text
            msg_text.FontSize    = 13
            msg_text.Foreground  = SolidColorBrush(Color.FromRgb(255, 255, 255))
            msg_text.TextWrapping = TextWrapping.Wrap
            bubble.Child = msg_text

            Grid.SetColumn(bubble, 0)
            row.Children.Add(bubble)
            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("Error adding user message: {}".format(ex))

    def _append_bot_message(self, text):
        """Add a left-aligned bot bubble with avatar (BatchOut color scheme)."""
        try:
            from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength
            from System.Windows.Media import SolidColorBrush, Color

            row = Grid()
            row.Margin = Thickness(0, 0, 60, 10)
            col_av = ColumnDefinition()
            col_av.Width = GridLength.Auto
            col_msg = ColumnDefinition()
            col_msg.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col_av)
            row.ColumnDefinitions.Add(col_msg)

            # Avatar
            av = self._make_avatar("T3")
            Grid.SetColumn(av, 0)
            row.Children.Add(av)

            # Bubble — white with BatchOut border
            bubble = Border()
            bubble.Background     = SolidColorBrush(Color.FromRgb(255, 255, 255))
            bubble.CornerRadius   = CornerRadius(3, 8, 8, 8)
            bubble.Padding        = Thickness(14, 10, 14, 10)
            bubble.BorderBrush    = SolidColorBrush(Color.FromRgb(189, 195, 199))  # #BDC3C7
            bubble.BorderThickness = Thickness(1)

            msg_text = TextBlock()
            msg_text.Text        = text
            msg_text.FontSize    = 13
            msg_text.Foreground  = SolidColorBrush(Color.FromRgb(44, 62, 80))     # #2C3E50
            msg_text.TextWrapping = TextWrapping.Wrap
            bubble.Child = msg_text

            Grid.SetColumn(bubble, 1)
            row.Children.Add(bubble)
            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("Error adding bot message: {}".format(ex))

    def _scroll_to_bottom(self):
        try:
            self.chat_scroll.ScrollToBottom()
        except Exception:
            pass


# ─── MAIN ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)

    window = T3LabAssistantWindow()
    window.ShowDialog()
