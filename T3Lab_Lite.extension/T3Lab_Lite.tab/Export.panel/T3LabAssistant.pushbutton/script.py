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
from System.Windows import Visibility, WindowState
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

# ─── Tool launchers ───────────────────────────────────────────────────────────
# Each function opens the corresponding T3Lab tool.

def _get_tool_script_dir(panel, pushbutton):
    """Return the path to a pushbutton script.py given panel and pushbutton names."""
    tab_dir = os.path.dirname(os.path.dirname(__file__))  # T3Lab_Lite.tab
    return os.path.join(tab_dir, panel, pushbutton, 'script.py')


def launch_batchout():
    """Open the BatchOut export dialog."""
    try:
        script_path = _get_tool_script_dir('Export.panel', 'BatchOut.pushbutton')
        if script_path not in sys.path:
            sys.path.insert(0, os.path.dirname(script_path))
        import importlib.util
        spec = importlib.util.spec_from_file_location("batchout_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        window = mod.ExportManagerWindow()
        window.ShowDialog()
        return True
    except Exception as ex:
        logger.error("Error launching BatchOut: {}".format(ex))
        return False


def launch_parasync():
    """Open the ParaSync parameter sync tool."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'ParaSync.pushbutton')
        import importlib.util
        spec = importlib.util.spec_from_file_location("parasync_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
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
        import importlib.util
        spec = importlib.util.spec_from_file_location("projectname_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return True
    except Exception as ex:
        logger.error("Error launching ProjectName: {}".format(ex))
        return False


def launch_workset():
    """Open the Workset manager."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'Workset.pushbutton')
        import importlib.util
        spec = importlib.util.spec_from_file_location("workset_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return True
    except Exception as ex:
        logger.error("Error launching Workset: {}".format(ex))
        return False


def launch_dimtext():
    """Open the Dim Text tool."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'DimText.pushbutton')
        import importlib.util
        spec = importlib.util.spec_from_file_location("dimtext_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return True
    except Exception as ex:
        logger.error("Error launching DimText: {}".format(ex))
        return False


def launch_upperdimtext():
    """Open the Upper Dim Text tool."""
    try:
        script_path = _get_tool_script_dir('Project.panel', 'UpperDimText.pushbutton')
        import importlib.util
        spec = importlib.util.spec_from_file_location("upperdimtext_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return True
    except Exception as ex:
        logger.error("Error launching UpperDimText: {}".format(ex))
        return False


def launch_resetoverrides():
    """Run the Reset Overrides tool."""
    try:
        script_path = _get_tool_script_dir('Graphic.panel', 'Reset Overrides.pushbutton')
        import importlib.util
        spec = importlib.util.spec_from_file_location("resetoverrides_script", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return True
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

            # Hide previous status
            self.status_panel.Visibility = Visibility.Collapsed

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
                        u"Không hiểu lệnh. Thử: 'mở batchout', 'parasync', 'load family'..."
                    )

        except Exception as ex:
            logger.error("Error in _process_input: {}".format(ex))

    # ─── Execute intent ────────────────────────────────────────────────────────

    def _execute_result(self, result):
        """Execute the action described by a parsed result dict."""
        intent = result.get("intent", "unknown")
        message = result.get("message", "")

        if intent == "help":
            answer = result.get("params", {}).get("answer", message)
            self._append_bot_message(answer or message)
            return

        if intent in TOOL_LAUNCHERS:
            confirmation = message or u"Đang mở công cụ..."
            self._append_bot_message(confirmation)
            # Run the tool launcher (may open a blocking dialog)
            launcher = TOOL_LAUNCHERS[intent]

            def run_tool():
                ok = launcher()
                def update():
                    if not ok:
                        self._append_bot_message(
                            u"Không thể mở công cụ. Xem console để biết lỗi."
                        )
                self.Dispatcher.Invoke(Action(update))

            t = Thread(ThreadStart(run_tool))
            t.IsBackground = True
            t.Start()
        elif intent == "unknown":
            self._append_bot_message(
                result.get("params", {}).get("message", u"Lệnh không rõ.")
            )
        else:
            self._append_bot_message(message or u"Đã xử lý lệnh.")

    def _run_tool(self, intent, default_msg):
        """Helper: show confirmation message and run a tool launcher."""
        self._append_bot_message(default_msg)
        launcher = TOOL_LAUNCHERS.get(intent)
        if not launcher:
            return

        def run():
            ok = launcher()
            def update():
                if not ok:
                    self._append_bot_message(
                        u"Không thể mở công cụ. Xem console để biết lỗi."
                    )
            self.Dispatcher.Invoke(Action(update))

        t = Thread(ThreadStart(run))
        t.IsBackground = True
        t.Start()

    # ─── Chat UI helpers ──────────────────────────────────────────────────────

    def _append_user_message(self, text):
        """Add a user bubble to the chat history."""
        try:
            from System.Windows.Controls import Border, TextBlock, StackPanel
            from System.Windows import Thickness, CornerRadius, FontWeights, TextWrapping
            from System.Windows.Media import SolidColorBrush, Color

            panel = StackPanel()
            panel.Margin = Thickness(40, 0, 0, 8)

            label = TextBlock()
            label.Text = u"Bạn"
            label.FontSize = 10
            label.FontWeight = FontWeights.SemiBold
            label.Foreground = SolidColorBrush(Color.FromRgb(52, 152, 219))
            label.Margin = Thickness(0, 0, 0, 2)

            bubble = Border()
            bubble.Background = SolidColorBrush(Color.FromRgb(52, 152, 219))
            bubble.CornerRadius = CornerRadius(6)
            bubble.Padding = Thickness(10, 7, 10, 7)

            msg_text = TextBlock()
            msg_text.Text = text
            msg_text.FontSize = 12
            msg_text.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
            msg_text.TextWrapping = TextWrapping.Wrap
            bubble.Child = msg_text

            panel.Children.Add(label)
            panel.Children.Add(bubble)
            self.chat_history_panel.Children.Add(panel)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("Error adding user message: {}".format(ex))

    def _append_bot_message(self, text):
        """Add a bot bubble to the chat history."""
        try:
            from System.Windows.Controls import Border, TextBlock, StackPanel
            from System.Windows import Thickness, CornerRadius, TextWrapping
            from System.Windows.Media import SolidColorBrush, Color

            panel = StackPanel()
            panel.Margin = Thickness(0, 0, 40, 8)

            label = TextBlock()
            label.Text = u"T3Lab Assistant"
            label.FontSize = 10
            label.Foreground = SolidColorBrush(Color.FromRgb(41, 128, 185))
            label.Margin = Thickness(0, 0, 0, 2)

            bubble = Border()
            bubble.Background = SolidColorBrush(Color.FromRgb(238, 244, 251))
            bubble.CornerRadius = CornerRadius(6)
            bubble.Padding = Thickness(10, 7, 10, 7)

            msg_text = TextBlock()
            msg_text.Text = text
            msg_text.FontSize = 12
            msg_text.Foreground = SolidColorBrush(Color.FromRgb(44, 62, 80))
            msg_text.TextWrapping = TextWrapping.Wrap
            bubble.Child = msg_text

            panel.Children.Add(label)
            panel.Children.Add(bubble)
            self.chat_history_panel.Children.Add(panel)
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
