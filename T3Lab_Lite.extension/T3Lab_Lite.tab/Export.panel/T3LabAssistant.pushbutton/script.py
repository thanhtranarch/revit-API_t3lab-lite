# -*- coding: utf-8 -*-
"""T3Lab Assistant - Standalone AI chatbox for T3Lab tools.

Provides a floating chatbox window that understands natural-language commands
(Tiếng Việt or English) and quickly opens any T3Lab tool.
"""

__title__ = "T3Lab\nAssistant"
__author__ = "T3Lab"
__version__ = "1.0.0"

import os
import sys

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility, WindowState, GridLength
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Action
from System.Threading import Thread, ThreadStart

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
    from t3lab_assistant import parse_command, has_api_key, keyword_parse
    HAS_NLP = True
except Exception as e:
    logger.warning("Could not import t3lab_assistant: {}".format(e))
    HAS_NLP = False

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


# ─── WPF Window ───────────────────────────────────────────────────────────────

class T3LabAssistantWindow(forms.WPFWindow):
    """Standalone T3Lab Assistant chatbox window."""

    def __init__(self):
        xaml_path = os.path.join(lib_dir, 'GUI', 'T3LabAssistant.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        # Set logo
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

        # Update AI badge
        self._update_ai_badge()

    # ─── Window controls ──────────────────────────────────────────────────────

    def minimize_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def close_clicked(self, sender, e):
        self.Close()

    # ─── AI badge ─────────────────────────────────────────────────────────────

    def _update_ai_badge(self):
        try:
            if HAS_NLP and has_api_key():
                self.ai_status_badge.Visibility = Visibility.Visible
                self.ai_status_text.Text = "AI"
            else:
                self.ai_status_badge.Visibility = Visibility.Collapsed
        except Exception:
            pass

    # ─── Quick tool buttons ───────────────────────────────────────────────────

    def tool_batchout_clicked(self, sender, e):
        self._run_tool("open_batchout", "Đang mở BatchOut...")

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
            self.chat_input.Text = ""

            # Show user message in chat
            self._append_user_message(raw)

            if HAS_NLP and has_api_key():
                # Async NLP path
                self.send_button.IsEnabled = False
                captured = raw

                def do_nlp():
                    result = parse_command(captured)

                    def finish():
                        try:
                            self.send_button.IsEnabled = True
                            if result and result.get("intent") not in (None, "unknown"):
                                self._execute_result(result)
                            else:
                                # Fallback to keywords
                                fb = keyword_parse(captured)
                                if fb:
                                    self._execute_result(fb)
                                else:
                                    msg = (result or {}).get("params", {}).get("message", "")
                                    if not msg:
                                        msg = u"Không hiểu lệnh. Thử: 'mở batchout', 'parasync', 'load family'..."
                                    self._append_bot_message(msg)
                        except Exception as finish_ex:
                            logger.error("finish error: {}".format(finish_ex))

                    self.Dispatcher.Invoke(Action(finish))

                t = Thread(ThreadStart(do_nlp))
                t.IsBackground = True
                t.Start()
            else:
                # Synchronous keyword fallback
                fb = keyword_parse(raw)
                if fb:
                    self._execute_result(fb)
                else:
                    self._append_bot_message(
                        u"Không hiểu lệnh.\n"
                        u"Ví dụ:\n"
                        u"• 'xuất pdf toàn bộ G sheet'\n"
                        u"• 'export all A sheets DWG'\n"
                        u"• 'mở batchout G sheet pdf'\n"
                        u"• 'parasync'  •  'load family'"
                    )

        except Exception as ex:
            logger.error("Error in _process_input: {}".format(ex))

    # ─── Execute intent ────────────────────────────────────────────────────────

    def _execute_result(self, result):
        """Execute the action described by a parsed result dict."""
        intent  = result.get("intent", "unknown")
        message = result.get("message", "")
        params  = result.get("params", {})

        # ── Help ──────────────────────────────────────────────────────────────
        if intent == "help":
            answer = params.get("answer", message)
            self._append_bot_message(answer or message)
            return

        # ── Export directly (no UI) ───────────────────────────────────────────
        if intent == "export_direct":
            self._append_bot_message(message or u"Đang xuất file...")
            ok = launch_export_direct(params, self._append_bot_message)
            if not ok and not params:
                self._append_bot_message(u"Xuất thất bại. Xem console để biết lỗi.")
            return

        # ── Open BatchOut pre-configured ─────────────────────────────────────
        if intent == "open_batchout_configured":
            self._append_bot_message(message or u"Đang mở BatchOut đã cấu hình...")
            ok = launch_batchout_configured(params, self._append_bot_message)
            if not ok:
                self._append_bot_message(u"Không thể mở BatchOut. Xem console.")
            return

        # ── Simple tool launchers ─────────────────────────────────────────────
        if intent in TOOL_LAUNCHERS:
            self._append_bot_message(message or u"Đang mở công cụ...")
            ok = TOOL_LAUNCHERS[intent]()
            if not ok:
                self._append_bot_message(u"Không thể mở công cụ. Xem console để biết lỗi.")
            return

        # ── Unknown ───────────────────────────────────────────────────────────
        if intent == "unknown":
            self._append_bot_message(params.get("message", u"Lệnh không rõ."))
        else:
            self._append_bot_message(message or u"Đã xử lý lệnh.")

    def _run_tool(self, intent, default_msg):
        """Helper: show confirmation message and run a tool launcher on the UI thread."""
        self._append_bot_message(default_msg)
        launcher = TOOL_LAUNCHERS.get(intent)
        if not launcher:
            return
        # Run synchronously on UI thread — WPF dialogs require the UI thread
        ok = launcher()
        if not ok:
            self._append_bot_message(u"Không thể mở công cụ. Xem console để biết lỗi.")

    # ─── Chat UI helpers ──────────────────────────────────────────────────────

    def _make_avatar(self, letter, from_rgb_start, from_rgb_end):
        """Create a circular avatar Border with initials."""
        from System.Windows.Controls import Border, TextBlock
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Media import SolidColorBrush, Color, LinearGradientBrush, GradientStop
        from System.Windows import HorizontalAlignment, VerticalAlignment

        grad = LinearGradientBrush()
        grad.StartPoint = System.Windows.Point(0, 0)
        grad.EndPoint = System.Windows.Point(1, 1)
        gs1 = GradientStop()
        gs1.Color = Color.FromRgb(*from_rgb_start)
        gs1.Offset = 0.0
        gs2 = GradientStop()
        gs2.Color = Color.FromRgb(*from_rgb_end)
        gs2.Offset = 1.0
        grad.GradientStops.Add(gs1)
        grad.GradientStops.Add(gs2)

        av = Border()
        av.Width = 32
        av.Height = 32
        av.CornerRadius = CornerRadius(16)
        av.Background = grad
        av.Margin = Thickness(0, 2, 8, 0)
        av.VerticalAlignment = VerticalAlignment.Top

        lbl = TextBlock()
        lbl.Text = letter
        lbl.FontSize = 11
        lbl.FontWeight = System.Windows.FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment = VerticalAlignment.Center
        av.Child = lbl
        return av

    def _append_user_message(self, text):
        """Add a right-aligned user bubble to the chat history."""
        try:
            from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength, HorizontalAlignment
            from System.Windows.Media import SolidColorBrush, Color

            row = Grid()
            row.Margin = Thickness(48, 0, 0, 10)
            col0 = ColumnDefinition()
            col0.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col0)

            bubble = Border()
            bubble.Background = SolidColorBrush(Color.FromRgb(37, 99, 235))  # #2563EB
            bubble.CornerRadius = CornerRadius(12, 4, 12, 12)
            bubble.Padding = Thickness(12, 8, 12, 8)
            bubble.HorizontalAlignment = HorizontalAlignment.Right

            msg_text = TextBlock()
            msg_text.Text = text
            msg_text.FontSize = 12
            msg_text.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
            msg_text.TextWrapping = TextWrapping.Wrap
            bubble.Child = msg_text

            Grid.SetColumn(bubble, 0)
            row.Children.Add(bubble)
            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("Error adding user message: {}".format(ex))

    def _append_bot_message(self, text):
        """Add a left-aligned bot bubble with avatar to the chat history."""
        try:
            from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength
            from System.Windows.Media import SolidColorBrush, Color

            row = Grid()
            row.Margin = Thickness(0, 0, 48, 10)
            col_av = ColumnDefinition()
            col_av.Width = GridLength.Auto
            col_msg = ColumnDefinition()
            col_msg.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col_av)
            row.ColumnDefinitions.Add(col_msg)

            # Avatar
            av = self._make_avatar("T3", (37, 99, 235), (56, 189, 248))
            Grid.SetColumn(av, 0)
            row.Children.Add(av)

            # Bubble
            bubble = Border()
            bubble.Background = SolidColorBrush(Color.FromRgb(255, 255, 255))
            bubble.CornerRadius = CornerRadius(4, 12, 12, 12)
            bubble.Padding = Thickness(12, 8, 12, 8)
            bubble.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
            bubble.BorderThickness = Thickness(1)

            msg_text = TextBlock()
            msg_text.Text = text
            msg_text.FontSize = 12
            msg_text.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
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
