# -*- coding: utf-8 -*-
"""
T3Lab Assistant

Open the T3Lab AI assistant for natural language Revit commands.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "T3Lab Assistant"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==================================================
import os
import sys
import clr
import json
import re
import datetime

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility, WindowState, GridLength
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Action
from System.Threading import Thread, ThreadStart, ApartmentState

from pyrevit import revit, forms, script
from Autodesk.Revit import DB

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()
output = script.get_output()

try:
    REVIT_VERSION = int(revit.doc.Application.VersionNumber)
except Exception:
    try:
        from pyrevit import HOST_APP
        REVIT_VERSION = int(HOST_APP.version)
    except Exception:
        REVIT_VERSION = 2023  # safe fallback

# ─── Lib path setup ───────────────────────────────────────────────────────────
# __file__ → .../T3LabAssistant.pushbutton/script.py
# extension_dir → .../T3Lab_Lite.extension
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# ─── NLP module ───────────────────────────────────────────────────────────────
try:
    from Intelligence.t3lab_assistant import (parse_command, has_api_key, keyword_parse,
                                              learn_pattern, find_learned_match,
                                              has_local_llm, parse_command_local,
                                              get_local_model_name, parse_command_nlu,
                                              inject_discovered_tools,
                                              get_active_provider_name,
                                              get_provider_display_label,
                                              _RAG_SYSTEM_PREFIX)
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
    def get_active_provider_name(*a, **kw): return "claude"
    def get_provider_display_label(*a, **kw): return "AI"
    _RAG_SYSTEM_PREFIX = u""  # fallback: no RAG prefix if NLP module unavailable

# ─── Tool discovery module ────────────────────────────────────────────────────
try:
    from Services.tool_discovery import (discover_new_tools, get_registered_tools,
                                         make_generic_launcher)
    HAS_DISCOVERY = True
except Exception as e:
    logger.warning("Could not import tool_discovery: {}".format(e))
    HAS_DISCOVERY = False
    def discover_new_tools(): return []
    def get_registered_tools(): return []
    def make_generic_launcher(script_path, title): return lambda: False

# ─── Context Scout (BIM Context) ────────────────────────────────────────────────
try:
    from Selection.scout import ContextScout
    HAS_SCOUT = True
except Exception as e:
    logger.warning("Could not import ContextScout: {}".format(e))
    HAS_SCOUT = False
    class ContextScout:
        @staticmethod
        def get_context_summary_for_ai(): return ""

# ─── BatchOut executor (configure + direct export) ────────────────────────────
try:
    from Services.batchout_executor import configure_batchout_window, direct_export
    HAS_EXECUTOR = True
except Exception as e:
    logger.warning("Could not import batchout_executor: {}".format(e))
    HAS_EXECUTOR = False

# ─── RAG processor (PDF / image attachments) ──────────────────────────────────
try:
    from Intelligence.rag_processor import (is_supported, is_image, is_pdf,
                                           build_text_context, build_vision_content_blocks,
                                           has_images, summarize_attachments, SUPPORTED_EXTS)
    HAS_RAG = True
except Exception as e:
    logger.warning("Could not import rag_processor: {}".format(e))
    HAS_RAG = False
    def is_supported(p): return False
    def is_image(p): return False
    def is_pdf(p): return False
    def build_text_context(files): return ''
    def build_vision_content_blocks(text, files): return [{"type": "text", "text": text}]
    def has_images(files): return False
    def summarize_attachments(files): return ''
    SUPPORTED_EXTS = set()

# ─── Streaming message extractor (live token rendering) ───────────────────────
try:
    from Intelligence.llm_provider import StreamingJSONExtractor
except Exception as e:
    logger.warning("Could not import StreamingJSONExtractor: {}".format(e))
    class StreamingJSONExtractor(object):
        """Fallback: show whatever raw text streams in (no JSON unwrapping)."""
        def display(self, raw):
            return (raw or u"").strip()

# ─── Tool launchers ───────────────────────────────────────────────────────────
# Each function opens the corresponding T3Lab tool.

def _get_tool_script_dir(*parts):
    """Return the path to a pushbutton script.py given path parts relative to the tab.

    Usage:
        _get_tool_script_dir('Export.panel', 'BatchOut.pushbutton')
        _get_tool_script_dir('Annotation & Select.panel', 'Text.stack', 'DimText.pushbutton')
    """
    # __file__ = .../T3Lab_Lite.tab/AI Connection.panel/T3LabAssistant.pushbutton/script.py
    # dirname x1 = T3LabAssistant.pushbutton/
    # dirname x2 = AI Connection.panel/
    # dirname x3 = T3Lab_Lite.tab/
    tab_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    return os.path.join(tab_dir, *parts + ('script.py',))


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
    script_path = _get_tool_script_dir('Views & Sheets.panel', 'BatchOut.pushbutton')
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
        script_path = _get_tool_script_dir('Modeling & Datum.panel', 'ParaSync.pushbutton')
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
        from GUI.ManaFamiDialog import show_family_manager
        show_family_manager(default_tab=0)
        return True
    except Exception as ex:
        logger.error("Error launching LoadFamily: {}".format(ex))
        return False


def launch_loadfamily_cloud():
    """Open the Load Family (Cloud) dialog."""
    try:
        from GUI.ManaFamiDialog import show_family_manager
        show_family_manager(default_tab=0)
        return True
    except Exception as ex:
        logger.error("Error launching LoadFamily Cloud: {}".format(ex))
        return False


def launch_projectname():
    """Open the Project Name tool."""
    try:
        script_path = _get_tool_script_dir('Modeling & Datum.panel', 'ProjectName.pushbutton')
        mod = _load_script('projectname_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching ProjectName: {}".format(ex))
        return False


def launch_workset():
    """Open the Workset manager."""
    try:
        script_path = _get_tool_script_dir('Modeling & Datum.panel', 'Workset.pushbutton')
        mod = _load_script('workset_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching Workset: {}".format(ex))
        return False


def launch_dimtext():
    """Run the Dim Text tool on current selection."""
    try:
        script_path = _get_tool_script_dir('Annotation & Select.panel', 'Text.stack', 'TextTagTools.pulldown', 'DimText.pushbutton')
        mod = _load_script('dimtext_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching DimText: {}".format(ex))
        return False


def launch_upperall():
    """Run the Upper All Text tool on current selection."""
    try:
        script_path = _get_tool_script_dir('Annotation & Select.panel', 'Text.stack', 'TextTagTools.pulldown', 'UpperAll.pushbutton')
        mod = _load_script('upperall_script', script_path)
        if mod and hasattr(mod, 'main'):
            mod.main()
        return mod is not None
    except Exception as ex:
        logger.error("Error launching UpperAll: {}".format(ex))
        return False


def launch_resetoverrides():
    """Run the Reset Overrides tool on the active view."""
    try:
        script_path = _get_tool_script_dir('Annotation & Select.panel', 'Graphic 2.stack', 'Reset Overrides.pushbutton')
        mod = _load_script('resetoverrides_script', script_path)
        return mod is not None
    except Exception as ex:
        logger.error("Error launching Reset Overrides: {}".format(ex))
        return False


def launch_cadtobeam():
    """Open the CAD to Beam tool."""
    try:
        script_path = _get_tool_script_dir('Modeling & Datum.panel', 'Create.stack', 'Create Elements.pulldown', 'Beam.pushbutton')
        mod = _load_script('cadtobeam_script', script_path)
        if mod:
            window = mod.CADtoBeamWindow()
            window.ShowDialog()
            return True
        return False
    except Exception as ex:
        logger.error("Error launching CADtoBeam: {}".format(ex))
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
    "open_upperdimtext":     launch_upperall,
    "open_resetoverrides":   launch_resetoverrides,
    "open_cad_to_beam":      launch_cadtobeam,
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


# CLASS/FUNCTIONS
# ==================================================

class T3LabAssistantWindow(forms.WPFWindow):
    """Standalone T3Lab Assistant chatbox window."""

    # Dynamic buttons added by _bootstrap_discovered_tools
    _DYNAMIC_BTNS = []   # list of Button WPF objects (not names)

    def __init__(self, is_docked=False):
        self.is_docked = is_docked
        try:
            xaml_path = os.path.join(extension_dir, 'lib', 'GUI', 'Tools', 'T3LabAssistant.xaml')
            forms.WPFWindow.__init__(self, xaml_path)
        except Exception as ex:
            logger.error("Could not load T3LabAssistant XAML: {}".format(ex))
            raise

        self.doc = revit.doc

        # ── Session state ─────────────────────────────────────────────────────
        self._busy             = False          # concurrency guard
        self._typing_row       = None           # reference to typing indicator element
        self._conversation_history = []         # [{role, content}, ...] multi-turn context
        self._last_raw         = ''             # last user input (for learning)
        self._doc_key          = _get_doc_key() # document identifier for history
        self._persisted_msgs   = []             # flat list with timestamps, for save/load
        self._attached_files   = []             # list of file paths (images / PDFs)
        self._models_cache     = {
            "claude":    ["claude-3-5-sonnet-20241022", "claude-3-haiku-20240307", "claude-3-opus-20240229"],
            "openai":    ["gpt-4o", "gpt-4o-mini", "gpt-4-turbo", "gpt-3.5-turbo"],
            "deepseek":  ["deepseek-chat", "deepseek-coder"],
            "ollama":    ["qwen2.5:0.5b", "qwen2.5:1.5b", "llama3", "mistral"],
            "lmstudio":  ["local-model"],
        }             # {provider_name: [model_list]} — avoids HTTP on sidebar open
        
        # ── History & Typing Animation State ──────────────────────────────────
        self._input_history    = []
        self._history_index    = -1
        self._current_input_temp = ""
        self._typing_timer     = None
        self._typing_elapsed   = 0

        # ── Live streaming bubble state ───────────────────────────────────────
        self._stream_row       = None           # Grid row of the live reply bubble
        self._stream_tb        = None           # TextBlock being filled token-by-token

        # ── Easy-to-use input: multi-line with Shift+Enter ────────────────────
        try:
            from System.Windows.Controls import ScrollBarVisibility
            from System.Windows import TextWrapping
            self.chat_input.AcceptsReturn = False
            self.chat_input.TextWrapping  = TextWrapping.Wrap
            self.chat_input.MaxHeight     = 120
            self.chat_input.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            self.chat_input.ToolTip = (u"Nhập lệnh Tiếng Việt hoặc English  •  "
                                       u"Enter để gửi, Shift+Enter xuống dòng")
        except Exception:
            pass


        # ── Logo ──────────────────────────────────────────────────────────────

        # ── Restore conversation from previous session ─────────────────────────
        self._restore_history()
        self._update_welcome_greeting()

        # ── Tool discovery (background, then inject chips into UI) ─────────────
        def _discover_and_update():
            import time; time.sleep(0.3)
            self.Dispatcher.Invoke(Action(self._bootstrap_discovered_tools))

        _dt = Thread(ThreadStart(_discover_and_update))
        _dt.IsBackground = True
        _dt.SetApartmentState(ApartmentState.STA)
        _dt.Start()

        # Restore window geometry and sidebar state from last session
        self._restore_window_state()

        # Update AI badge, pre-load models cache, and warm up router status in background
        def _bg_startup_probe():
            try:
                import time
                time.sleep(0.5)   # let window render first

                # ─── 1. Auto-start Revit MCP Server & File Watcher ───
                try:
                    from Services.mcp_service import MCPService
                    # Start MCP HTTP Server if stopped
                    srv_status = MCPService.server_status()
                    if not srv_status.get('running'):
                        MCPService.start_server()
                    # Start File Task Watcher if stopped
                    wat_status = MCPService.watcher_status()
                    if not wat_status.get('running'):
                        MCPService.start_watcher()
                except Exception as ex:
                    logger.debug("Auto-start MCP/watcher failed: {}".format(ex))

                # ─── 2. Auto-start Ollama Local Engine ───
                try:
                    from Intelligence import local_llm
                    if not local_llm.is_running():
                        import subprocess
                        import os
                        user_profile = os.environ.get('USERPROFILE', '')
                        ollama_paths = [
                            "ollama",
                            os.path.join(user_profile, "AppData", "Local", "Programs", "Ollama", "ollama.exe")
                        ]
                        launched = False
                        for path in ollama_paths:
                            try:
                                # Start Ollama server in background (no console window)
                                if os.path.exists(path) or path == "ollama":
                                    subprocess.Popen([path, "serve"], 
                                                     creationflags=0x08000000) # CREATE_NO_WINDOW
                                    launched = True
                                    break
                            except Exception:
                                pass
                        
                        if launched:
                            # Wait for server to boot up
                            for _ in range(10):
                                time.sleep(0.5)
                                if local_llm.is_running():
                                    break
                except Exception as ex:
                    logger.debug("Ollama launch failed: {}".format(ex))

                # ─── 3. Auto-download Default Model if empty ───
                try:
                    from Intelligence import local_llm
                    if local_llm.is_running():
                        models = local_llm.list_models()
                        if not models:
                            # Notify user in the chat panel
                            self.Dispatcher.Invoke(Action(lambda: self._append_bot_message(
                                u"📥 Không tìm thấy mô hình AI cục bộ nào. Đang tự động tải mô hình mặc định (qwen2.5:1.5b) về máy bạn. Quá trình này chạy ngầm và có thể mất vài phút..."
                            )))
                            
                            payload = {"name": "qwen2.5:1.5b", "stream": False}
                            local_llm._post_json(local_llm.OLLAMA_HOST + "/api/pull", payload, timeout=600)
                            
                            self.Dispatcher.Invoke(Action(lambda: self._append_bot_message(
                                u"✅ Tải thành công mô hình AI qwen2.5:1.5b! Bạn có thể sử dụng T3Lab Assistant ngoại tuyến."
                            )))
                except Exception as ex:
                    logger.debug("Auto-pull model failed: {}".format(ex))

                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                active = router.get_active_name()

                # Update the badge display instantly from router settings
                self.Dispatcher.Invoke(Action(self._update_ai_badge))

                # Step 1: Probe the active provider first to fill cache ASAP
                try:
                    router.probe_provider(active)
                    provider = router.get_provider(active)
                    if provider:
                        self._models_cache[active] = provider.get_models()
                except Exception:
                    pass

                # If sidebar is open, update UI with active provider's data
                sidebar_state = [False]
                def _check_sidebar():
                    sidebar_state[0] = (self.settings_sidebar.Visibility == Visibility.Visible)
                self.Dispatcher.Invoke(Action(_check_sidebar))
                if sidebar_state[0]:
                    self.Dispatcher.Invoke(Action(self._update_sidebar))

                # Step 2: Pre-load/cache models for all other providers in the background
                for name in router.get_provider_names():
                    if name == active:
                        continue
                    try:
                        p = router.get_provider(name)
                        if p and p.check_health():
                            self._models_cache[name] = p.get_models()
                    except Exception:
                        pass

                # Step 3: Warm up LLMRouter status cache
                try:
                    router.get_status(use_cache=False)
                except Exception:
                    pass

                # Step 4: Final update to sidebar if open
                self.Dispatcher.Invoke(Action(_check_sidebar))
                if sidebar_state[0]:
                    self.Dispatcher.Invoke(Action(self._update_sidebar))
            except Exception:
                pass

        _t = Thread(ThreadStart(_bg_startup_probe))
        _t.IsBackground = True
        _t.SetApartmentState(ApartmentState.STA)
        _t.Start()

        # ── First-run onboarding (new installs only) ──────────────────────────
        try:
            from config.user_profile import UserProfile
            if UserProfile().is_first_run():
                self._show_onboarding()
        except Exception as ex:
            logger.debug("onboarding check error: {}".format(ex))

        # Hide minimize/maximize buttons if hosted inside Dockable Pane
        if self.is_docked:
            try:
                self.btn_minimize.Visibility = Visibility.Collapsed
                self.btn_maximize.Visibility = Visibility.Collapsed
            except Exception:
                pass

    def setup_icon(self):
        """Override pyRevit's setup_icon to remove the window icon from the title bar."""
        pass

    # ─── Window state persistence ─────────────────────────────────────────────

    def _restore_window_state(self):
        """Restore window position, size and sidebar visibility from settings."""
        try:
            from config.settings import T3LabAISettings
            ws = T3LabAISettings().get_window_state()

            # Restore size
            w = ws.get('width')
            h = ws.get('height')
            if w and h:
                self.Width  = float(w)
                self.Height = float(h)

            # Restore position — validate it is still on-screen
            left = ws.get('left')
            top  = ws.get('top')
            if left is not None and top is not None:
                try:
                    import System.Windows
                    sw = System.Windows.SystemParameters.PrimaryScreenWidth
                    sh = System.Windows.SystemParameters.PrimaryScreenHeight
                    left_f = float(left)
                    top_f  = float(top)
                    if 0 <= left_f <= sw - 100 and 0 <= top_f <= sh - 60:
                        self.Left = left_f
                        self.Top  = top_f
                except Exception:
                    pass

            # Restore sidebar
            if ws.get('sidebar_open'):
                self.settings_sidebar.Visibility = Visibility.Visible
                self._update_sidebar_instant()

        except Exception as ex:
            logger.debug("_restore_window_state error: {}".format(ex))

    def _save_window_state(self):
        """Persist current window geometry and sidebar state to settings."""
        try:
            from config.settings import T3LabAISettings
            sidebar_open = (self.settings_sidebar.Visibility == Visibility.Visible)
            T3LabAISettings().save_window_state(
                self.Left, self.Top,
                self.Width, self.Height,
                sidebar_open,
            )
        except Exception as ex:
            logger.debug("_save_window_state error: {}".format(ex))

    # ─── Window controls ──────────────────────────────────────────────────────

    def close_clicked(self, sender, e):
        self._save_window_state()
        if self.is_docked:
            try:
                from Autodesk.Revit.UI import DockablePaneId
                from System import Guid
                from GUI.AssistantPaneControl import ASSISTANT_PANE_GUID
                from pyrevit import HOST_APP

                pane_id = DockablePaneId(ASSISTANT_PANE_GUID)
                pane = HOST_APP.uiapp.GetDockablePane(pane_id)
                if pane and pane.IsShown():
                    pane.Hide()
                    return
            except Exception:
                pass
        self.Close()

    def minimize_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_clicked(self, sender, e):
        from System.Windows import WindowState
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def undo_clicked(self, sender, e):
        """Undo the last Revit transaction."""
        try:
            if revit.doc.CanUndo():
                revit.doc.Undo()
                self._append_bot_message(u"↺ Đã hoàn tác (Undo) hành động cuối cùng.")
            else:
                self._append_bot_message(u"Không có hành động nào để hoàn tác.")
        except Exception as ex:
            logger.debug("Undo error: {}".format(ex))

    # ─── Tool discovery bootstrap ──────────────────────────────────────────────

    def _bootstrap_discovered_tools(self):
        """
        Run on startup (UI thread):
          1. Discover new tools → register launchers → update NLP prompt.
          2. Post a chat notification for truly NEW tools.
        Must be called from the UI thread (via Dispatcher.Invoke).
        """
        try:
            if not HAS_DISCOVERY:
                return

            # ── Discover (writes registry) ────────────────────────────────────
            new_tools = discover_new_tools()

            # ── Register launchers + inject into NLP ─────────────────────────
            _register_discovered_launchers(new_tools)

            # Also register launchers for tools already in registry from previous runs
            all_tools = get_registered_tools()
            if all_tools:
                _register_discovered_launchers(all_tools)

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
            # Remove all children except the welcome greeting panel (first child)
            while self.chat_history_panel.Children.Count > 1:
                self.chat_history_panel.Children.RemoveAt(1)
            # Clear in-memory state
            self._conversation_history = []
            self._persisted_msgs = []
            # Delete saved file
            clear_chat_history(self._doc_key)
            self._update_welcome_greeting()
            # Show fresh welcome message
            self._append_bot_message(
                u"Cuộc trò chuyện đã được làm mới! 👋\n"
                u"Tôi có thể giúp gì cho bạn?"
            )
        except Exception as ex:
            logger.debug("reset_chat error: {}".format(ex))

    # ─── AI badge & provider switcher ────────────────────────────────────────

    # Provider brand colors (shared by badge + health refresh)
    _BADGE_COLORS = {
        "claude":    (217, 119,  87),  # Anthropic orange
        "openai":    ( 16, 163, 127),  # OpenAI green
        "deepseek":  ( 37,  99, 235),  # DeepSeek blue
        "ollama":    ( 59, 130, 246),  # Ollama blue
        "lmstudio":  (124,  58, 237),  # LM Studio purple
    }
    _BADGE_GRAY = (161, 161, 170)       # #A1A1AA — no provider / offline

    def _render_greeting(self, name):
        """Set the welcome greeting text for a given name (no settings read)."""
        try:
            import datetime
            hour = datetime.datetime.now().hour
            if hour < 12:
                greet = u"Good morning"
            elif hour < 18:
                greet = u"Good afternoon"
            else:
                greet = u"Good evening"
            self.welcome_greeting_text.Text = u"{}, {}".format(greet, name or u"Thạnh")
        except Exception:
            pass

    def _update_welcome_greeting(self):
        """Refresh greeting text from saved settings and toggle panel visibility."""
        try:
            from config.user_profile import UserProfile
            name = UserProfile().get_name() or u"Thạnh"
            self._render_greeting(name)

            # The welcome banner only shows on a fresh chat (no history yet).
            if self._persisted_msgs:
                self.welcome_greeting_panel.Visibility = Visibility.Collapsed
            else:
                self.welcome_greeting_panel.Visibility = Visibility.Visible
        except Exception:
            pass

    def sidebar_username_changed(self, sender, e):
        """Live-preview the greeting + profile card as the user types."""
        try:
            name = (self.sidebar_username_box.Text or u"").strip()
            if name:
                self._render_greeting(name)
                self._update_profile_card(name)
        except Exception:
            pass

    def _update_profile_card(self, name=None):
        """Refresh the Claude-style profile card in the settings sidebar."""
        try:
            from config.user_profile import UserProfile
            prof = UserProfile()
            nm = (name or u"").strip() or prof.get_name() or u"Thạnh"
            try:
                self.sidebar_profile_name.Text    = nm
                self.sidebar_profile_initial.Text = (nm[:1].upper() if nm else u"T")
            except Exception:
                pass
            try:
                from Intelligence.llm_router import LLMRouter
                self.sidebar_profile_sub.Text = LLMRouter().get_display_label()
            except Exception:
                try:
                    self.sidebar_profile_sub.Text = u"T3Lab Assistant"
                except Exception:
                    pass
        except Exception:
            pass

    def sidebar_save_username_clicked(self, sender, e):
        """Persist the user name and reflect it in the greeting immediately."""
        username = (self.sidebar_username_box.Text or u"").strip()
        if not username:
            return

        # 1) Update the greeting from the typed value FIRST — independent of disk
        #    I/O — so the UI always syncs even if persistence happens to fail.
        self._render_greeting(username)

        # 2) Persist to the user profile (which also syncs settings.username).
        saved = False
        try:
            from config.user_profile import UserProfile
            UserProfile().set_name(username)
            saved = True
        except Exception as ex:
            logger.debug("set_name error: {}".format(ex))

        # 3) Confirm visually (covers the mid-chat case where the greeting banner
        #    itself is collapsed and the text change isn't visible).
        self._update_profile_card(username)
        if saved:
            self._flash_username_saved()

    def _flash_username_saved(self):
        """Briefly tint the username box green to confirm a successful save."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan

            box = self.sidebar_username_box
            box.BorderBrush = SolidColorBrush(Color.FromRgb(16, 185, 129))   # emerald
            box.Background  = SolidColorBrush(Color.FromRgb(240, 253, 244))

            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromSeconds(1.3)

            def _revert(s, ev):
                try:
                    box.BorderBrush = SolidColorBrush(Color.FromRgb(230, 230, 234))  # #E6E6EA
                    box.Background  = SolidColorBrush(Color.FromRgb(255, 255, 255))
                finally:
                    timer.Stop()

            timer.Tick += _revert
            timer.Start()
        except Exception:
            pass

    # ─── First-run onboarding ─────────────────────────────────────────────────

    def _selected_onboarding_provider(self):
        """Return the provider tag selected in the onboarding combo."""
        try:
            item = self.onboarding_provider_combo.SelectedItem
            if item is not None and item.Tag:
                return str(item.Tag)
        except Exception:
            pass
        return "claude"

    def _sync_onboarding_key_panel(self):
        """Hide the API-key field for local providers; relabel for remote ones."""
        try:
            prov = self._selected_onboarding_provider()
            is_local = prov in ("ollama", "lmstudio")
            self.onboarding_key_panel.Visibility = (
                Visibility.Collapsed if is_local else Visibility.Visible)
            if not is_local:
                labels = {
                    "claude":   u"ANTHROPIC API KEY",
                    "openai":   u"OPENAI API KEY",
                    "deepseek": u"DEEPSEEK API KEY",
                }
                self.onboarding_key_label.Text = labels.get(prov, u"API KEY")
        except Exception:
            pass

    def _show_onboarding(self):
        """Display the first-run onboarding card."""
        try:
            from config.user_profile import UserProfile
            prof = UserProfile()
            try:
                nm = prof.get_name(fallback=False)
                if nm:
                    self.onboarding_name_box.Text = nm
            except Exception:
                pass

            import datetime
            h = datetime.datetime.now().hour
            greet = (u"Good morning" if h < 12 else
                     u"Good afternoon" if h < 18 else u"Good evening")
            try:
                self.onboarding_greeting.Text = u"{}! 👋".format(greet)
            except Exception:
                pass

            self._sync_onboarding_key_panel()
            self.onboarding_overlay.Visibility = Visibility.Visible
            try:
                self.onboarding_name_box.Focus()
            except Exception:
                pass
        except Exception as ex:
            logger.debug("_show_onboarding error: {}".format(ex))

    def _hide_onboarding(self):
        try:
            self.onboarding_overlay.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def onboarding_provider_changed(self, sender, e):
        self._sync_onboarding_key_panel()

    def onboarding_skip_clicked(self, sender, e):
        """Dismiss onboarding without saving — won't show again."""
        try:
            from config.user_profile import UserProfile
            UserProfile().mark_setup_completed()
        except Exception:
            pass
        self._hide_onboarding()

    def onboarding_save_clicked(self, sender, e):
        """Persist the new user's profile + model setup, then dismiss onboarding."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            name = (self.onboarding_name_box.Text or u"").strip()
            if not name:
                # Gently require a name.
                try:
                    self.onboarding_name_box.BorderBrush = SolidColorBrush(
                        Color.FromRgb(239, 68, 68))      # rose
                    self.onboarding_name_box.Focus()
                except Exception:
                    pass
                return

            from config.user_profile import UserProfile
            prof     = UserProfile()
            provider = self._selected_onboarding_provider()

            prof.set_name(name)
            prof.set_model_setup(provider)

            # Optional API key for remote providers.
            key = u""
            try:
                key = (self.onboarding_key_box.Text or u"").strip()
            except Exception:
                pass
            if key and provider in ("claude", "openai", "deepseek"):
                key_map = {"claude": "Claude", "openai": "OpenAI", "deepseek": "DeepSeek"}
                try:
                    from config.settings import T3LabAISettings
                    T3LabAISettings().set_api_key(key_map[provider], key)
                except Exception:
                    pass

            prof.mark_setup_completed()

            # Activate the chosen provider live.
            try:
                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                router.switch_provider(provider)
                p = router.get_active_provider()
                if p and hasattr(p, "reload_credentials"):
                    p.reload_credentials()
            except Exception:
                pass

            self._hide_onboarding()

            # Refresh greeting + badge to reflect the new profile.
            self._render_greeting(name)
            try:
                self._update_welcome_greeting()
            except Exception:
                pass
            try:
                self._update_ai_badge()
            except Exception:
                pass
            try:
                self._update_profile_card(name)
            except Exception:
                pass

            self._append_bot_message(
                u"Rất vui được gặp bạn, {}! 🎉\n"
                u"Hồ sơ của bạn đã được lưu. Hãy thử 'mở batchout' hoặc hỏi tôi bất cứ điều gì về Revit.".format(name))
        except Exception as ex:
            logger.debug("onboarding_save_clicked error: {}".format(ex))

    def _update_ai_badge(self):
        """Render the provider pill INSTANTLY (no network). Health is refined async.

        This keeps provider/model switching snappy — the old version blocked the
        UI thread on provider.check_health() and a full router.get_status() probe.
        """
        try:
            from System.Windows.Media import SolidColorBrush, Color
            from Intelligence.llm_router import LLMRouter

            # Pill is always visible so the user can always open the switcher
            self.ai_status_badge.Visibility = Visibility.Visible

            if not HAS_NLP:
                self.ai_status_text.Text = u"No AI"
                self.ai_provider_dot.Fill = SolidColorBrush(Color.FromRgb(*self._BADGE_GRAY))
                self.ai_status_badge.ToolTip = "NLP module not loaded"
                return

            router = LLMRouter()
            name   = router.get_active_name()
            # Pure string ops — no network
            label  = router.get_display_label()
            rgb    = self._BADGE_COLORS.get(name, (52, 152, 219))

            self.ai_status_text.Text     = label
            self.ai_provider_dot.Fill    = SolidColorBrush(Color.FromRgb(*rgb))
            self.ai_status_badge.ToolTip = u"{}\nClick to switch provider".format(label)

            # Refine health (dot color + tooltip) on a background thread
            self._refresh_badge_health_async()

        except Exception:
            pass

    def _refresh_badge_health_async(self):
        """Probe the active provider's health off the UI thread, then update the dot."""
        def _work():
            try:
                from Intelligence.llm_router import LLMRouter
                from System.Windows.Media import SolidColorBrush, Color
                router   = LLMRouter()
                name     = router.get_active_name()
                provider = router.get_active_provider()
                healthy  = (provider is not None and provider.check_health())
                label    = router.get_display_label()
                try:
                    model = provider.get_active_model() or u"" if provider else u""
                except Exception:
                    model = u""

                def _apply():
                    try:
                        if healthy:
                            rgb = self._BADGE_COLORS.get(name, (52, 152, 219))
                            self.ai_provider_dot.Fill = SolidColorBrush(Color.FromRgb(*rgb))
                            self.ai_status_text.Text  = label
                            self.ai_status_badge.ToolTip = u"{} — {}\nClick to switch provider".format(
                                label, model) if model else (
                                u"{}\nClick to switch provider".format(label))
                        else:
                            self.ai_provider_dot.Fill = SolidColorBrush(Color.FromRgb(*self._BADGE_GRAY))
                            self.ai_status_text.Text  = u"Set up AI"
                            self.ai_status_badge.ToolTip = (
                                u"No AI provider configured.\n"
                                u"Click to select Claude, GPT, or Local LLM.")
                    except Exception:
                        pass

                self.Dispatcher.Invoke(Action(_apply))
            except Exception:
                pass

        try:
            t = Thread(ThreadStart(_work))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception:
            pass

    def ai_badge_clicked(self, sender, args):
        """Left-click on the badge → show provider switcher context menu."""
        try:
            from System.Windows.Controls import ContextMenu, MenuItem
            import System.Windows.Controls.Primitives as Primitives
            from Intelligence.llm_router import LLMRouter

            router = LLMRouter()
            status = router.get_status()

            menu = ContextMenu()

            for name in router.get_provider_names():
                info      = status.get(name, {})
                display   = info.get("display_name", name)
                model     = info.get("model") or ""
                available = info.get("available", False)
                is_active = info.get("active", False)

                if model:
                    short = model.split(":")[0] if ":" in model else model
                    header = u"{} ({})".format(display, short)
                else:
                    header = display

                if is_active:
                    header = u"✓  " + header
                if not available:
                    header = header + u"  [no key / offline]"

                item          = MenuItem()
                item.Header   = header
                item.IsEnabled = available and not is_active

                def _make_handler(n):
                    def _handler(s, e):
                        self._switch_provider(n)
                    return _handler

                item.Click += _make_handler(name)
                menu.Items.Add(item)

            menu.PlacementTarget = sender
            menu.Placement       = Primitives.PlacementMode.Bottom
            menu.IsOpen          = True

        except Exception as ex:
            logger.debug("ai_badge_clicked error: {}".format(ex))

    def _switch_provider(self, name):
        """Hot-swap the active LLM provider — instant UI, network probes in background."""
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            ok = router.switch_provider(name)
            if not ok:
                return

            # Instant UI: badge + sidebar render from cached/saved data (no network)
            self.Dispatcher.Invoke(Action(self._update_ai_badge))
            sidebar_open = (self.settings_sidebar.Visibility == Visibility.Visible)
            if sidebar_open:
                self.Dispatcher.Invoke(Action(self._update_sidebar_instant))

            # Background: probe the newly-active provider + refresh its model list
            def _bg():
                try:
                    provider = router.get_active_provider()
                    router.probe_provider(name)
                    if provider:
                        try:
                            self._models_cache[name] = provider.get_models()
                        except Exception:
                            pass
                    if sidebar_open:
                        self.Dispatcher.Invoke(Action(self._update_sidebar))
                except Exception:
                    pass

            t = Thread(ThreadStart(_bg))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception as ex:
            logger.debug("_switch_provider error: {}".format(ex))

    # ─── Settings sidebar ─────────────────────────────────────────────────────

    def settings_btn_clicked(self, sender, e):
        """Toggle settings sidebar — renders instantly (zero HTTP), then probes in background."""
        if self.settings_sidebar.Visibility == Visibility.Visible:
            self.settings_sidebar.Visibility = Visibility.Collapsed
            return

        self.settings_sidebar.Visibility = Visibility.Visible
        # Phase 1: instant render — no network calls, uses cached/local data
        self._update_sidebar_instant()

        # Phase 2: background probe — ACTIVE provider first (fast feedback),
        # then the remaining providers for the full status list.
        def _bg_probe():
            try:
                from Intelligence.llm_router import LLMRouter
                router   = LLMRouter()
                active   = router.get_active_name()
                provider = router.get_active_provider()

                # 2a. Probe only the active provider → update model list + dot ASAP
                router.probe_provider(active)
                if provider:
                    try:
                        self._models_cache[active] = provider.get_models()
                    except Exception:
                        pass
                self.Dispatcher.Invoke(Action(self._update_sidebar))

                # 2b. Probe everything else for the full status section
                router.get_status(use_cache=False)
                self.Dispatcher.Invoke(Action(self._update_sidebar))
            except Exception:
                pass

        _pt = Thread(ThreadStart(_bg_probe))
        _pt.IsBackground = True
        _pt.SetApartmentState(ApartmentState.STA)
        _pt.Start()

    def sidebar_provider_changed(self, sender, e):
        """Handle provider ComboBox selection change."""
        try:
            item = self.sidebar_provider_combo.SelectedItem
            if item is None:
                return
            tag = item.Tag
            if tag:
                # Repopulate the MODEL list for the chosen provider IMMEDIATELY so
                # it never lingers on the previous provider's models, then hot-swap.
                self._populate_model_combo(tag)
                self._switch_provider(tag)
        except Exception as ex:
            logger.debug("sidebar_provider_changed error: {}".format(ex))

    def sidebar_model_changed(self, sender, e):
        """Persist the newly selected model for the active provider."""
        try:
            from Intelligence.llm_router import LLMRouter
            item = self.sidebar_model_combo.SelectedItem
            if item is None:
                return
            model = item.ToString()
            router = LLMRouter()
            name = router.get_active_name()
            router.set_model(name, model)
            self._update_ai_badge()
        except Exception as ex:
            logger.debug("sidebar_model_changed error: {}".format(ex))

    def sidebar_save_model_clicked(self, sender, e):
        """Explicitly save the selected model as default + show confirmation."""
        try:
            from Intelligence.llm_router import LLMRouter
            item = self.sidebar_model_combo.SelectedItem
            if item is None:
                self.model_saved_hint.Text = u"Select a model first"
                self._flash_saved_hint(clear_after=True)
                return
            model  = item.ToString()
            router = LLMRouter()
            name   = router.get_active_name()
            router.set_model(name, model)
            self._update_ai_badge()
            self.model_saved_hint.Text = u"✓ Saved"
            self._flash_saved_hint(clear_after=True)
        except Exception as ex:
            logger.debug("sidebar_save_model_clicked error: {}".format(ex))

    def _flash_saved_hint(self, clear_after=False):
        """Briefly show the 'Saved' hint next to MODEL, then fade it out."""
        if not clear_after:
            return
        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromSeconds(2.0)

            def _clear(s, ev):
                try:
                    self.model_saved_hint.Text = u""
                finally:
                    timer.Stop()

            timer.Tick += _clear
            timer.Start()
        except Exception:
            # Fallback: just clear immediately if timer is unavailable
            try:
                self.model_saved_hint.Text = u""
            except Exception:
                pass

    def sidebar_test_clicked(self, sender, e):
        """Send a minimal test message to the active provider and show the raw reply."""
        from System.Windows.Media import SolidColorBrush, Color

        # Show "testing..." state immediately
        self.sidebar_test_result_border.Visibility = Visibility.Visible
        self.sidebar_test_result_border.Background = SolidColorBrush(Color.FromRgb(249, 250, 251))
        self.sidebar_test_result_border.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
        self.sidebar_test_label.Text       = u"Testing…"
        self.sidebar_test_label.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
        self.sidebar_test_result.Text      = u""
        self.sidebar_test_btn.IsEnabled    = False

        def _do_test():
            ok    = False
            label = u"Result"
            msg   = u""
            try:
                from Intelligence.llm_router import LLMRouter
                router   = LLMRouter()
                name     = router.get_active_name()
                provider = router.get_active_provider()

                if provider is None:
                    msg = u"Provider '{}' not loaded.".format(name)
                elif not provider.check_health():
                    if name == "ollama":
                        msg = (u"Ollama not available or no models installed.\n"
                               u"1. Make sure Ollama is running.\n"
                               u"2. Run: ollama pull qwen2.5:0.5b")
                    elif name == "lmstudio":
                        msg = (u"LM Studio not available or no model loaded.\n"
                               u"1. Open LM Studio.\n"
                               u"2. Load a model in LM Studio first.")
                    else:
                        msg = u"Provider not reachable.\nCheck API key or service status."
                else:
                    # Check a model is actually available before calling chat()
                    active_model = None
                    try:
                        active_model = provider.get_active_model()
                    except Exception:
                        pass
                    if not active_model:
                        msg = u"No model selected. Choose a model from the Model dropdown."
                    else:
                        # Direct call — bypass command-parsing layer
                        resp = provider.chat(
                            [],
                            u"You are a concise assistant. Do not think. Reply in one short sentence only.",
                            u"Reply with exactly this sentence: 'Connected OK'",
                            max_tokens=120,
                        )
                        if resp and resp.strip():
                            ok    = True
                            label = u"Connected"
                            msg   = resp.strip()[:120]
                        else:
                            msg = u"Provider responded but returned an empty reply."
            except Exception as ex:
                msg = u"Error: {}".format(str(ex)[:100])

            _ok = ok; _label = label; _msg = msg

            def _update():
                from System.Windows.Media import SolidColorBrush, Color
                if _ok:
                    self.sidebar_test_result_border.Background   = SolidColorBrush(Color.FromRgb(240, 253, 244))
                    self.sidebar_test_result_border.BorderBrush  = SolidColorBrush(Color.FromRgb(187, 247, 208))
                    self.sidebar_test_label.Foreground           = SolidColorBrush(Color.FromRgb(21, 128, 61))
                else:
                    self.sidebar_test_result_border.Background   = SolidColorBrush(Color.FromRgb(254, 242, 242))
                    self.sidebar_test_result_border.BorderBrush  = SolidColorBrush(Color.FromRgb(254, 202, 202))
                    self.sidebar_test_label.Foreground           = SolidColorBrush(Color.FromRgb(185, 28, 28))
                self.sidebar_test_label.Text  = _label
                self.sidebar_test_result.Text = _msg
                self.sidebar_test_btn.IsEnabled = True

            self.Dispatcher.Invoke(Action(_update))

        _tt = Thread(ThreadStart(_do_test))
        _tt.IsBackground = True
        _tt.SetApartmentState(ApartmentState.STA)
        _tt.Start()

    def sidebar_save_key_clicked(self, sender, e):
        """Setup rule: save API key → verify connection → fetch models → enable.

        The MODEL list stays disabled until the key is confirmed working, so a
        provider can never be 'used' with an invalid/missing key (which would
        otherwise stream wrong/fallback output).
        """
        try:
            from System.Windows.Media import SolidColorBrush, Color
            from config.settings import T3LabAISettings
            from Intelligence.llm_router import LLMRouter

            key = self.sidebar_api_key_box.Text.strip()
            if not key or key.endswith("..."):
                return

            router = LLMRouter()
            name   = router.get_active_name()
            settings_key = self._KEY_NAME_MAP.get(name)
            if not settings_key:
                return   # not a key-based provider

            # 1) Save the key.
            T3LabAISettings().set_api_key(settings_key, key)

            provider = router.get_active_provider()
            if provider and hasattr(provider, "reload_credentials"):
                provider.reload_credentials()
            elif provider and hasattr(provider, "invalidate_models_cache"):
                provider.invalidate_models_cache()
            self._models_cache.pop(name, None)

            # 2) Show "checking" state; keep MODEL disabled until verified.
            self.sidebar_model_combo.IsEnabled    = False
            self.sidebar_save_model_btn.IsEnabled = False
            self.sidebar_save_key_btn.IsEnabled   = False
            self.model_saved_hint.Foreground = SolidColorBrush(Color.FromRgb(113, 113, 122))
            self.model_saved_hint.Text       = u"Đang kiểm tra kết nối…"
            self._update_ai_badge()

            # 3) Verify connection + fetch models off the UI thread.
            def _validate():
                ok = False
                models = []
                try:
                    if provider and provider.check_health():
                        models = provider.get_models() or []
                        ok = len(models) > 0
                except Exception:
                    ok = False

                def _apply():
                    from System.Windows.Media import SolidColorBrush, Color
                    try:
                        self.sidebar_save_key_btn.IsEnabled = True
                    except Exception:
                        pass
                    if ok:
                        # 4) Fetched → enable model selection.
                        self._models_cache[name] = models
                        self._populate_model_combo(name)
                        self.model_saved_hint.Foreground = SolidColorBrush(Color.FromRgb(16, 185, 129))
                        self.model_saved_hint.Text = u"✓ Đã kết nối ({} model)".format(len(models))
                    else:
                        self._models_cache.pop(name, None)
                        self._populate_model_combo(name)   # stays disabled + hint
                        self.model_saved_hint.Foreground = SolidColorBrush(Color.FromRgb(239, 68, 68))
                        self.model_saved_hint.Text = u"✗ Key sai hoặc không kết nối được"
                    self._set_status_dot(name, ok)
                    self._update_ai_badge()

                self.Dispatcher.Invoke(Action(_apply))

            _kt = Thread(ThreadStart(_validate))
            _kt.IsBackground = True
            _kt.SetApartmentState(ApartmentState.STA)
            _kt.Start()
        except Exception as ex:
            logger.debug("sidebar_save_key_clicked error: {}".format(ex))

    def sidebar_save_host_clicked(self, sender, e):
        """Save the typed local server URL for LM Studio / Ollama."""
        try:
            from config.settings import T3LabAISettings
            from Intelligence.llm_router import LLMRouter
            host = self.sidebar_host_box.Text.strip()
            if not host:
                return

            router = LLMRouter()
            name   = router.get_active_name()

            if name == "lmstudio":
                # LM Studio reads host from settings key "LMStudio_Host"
                T3LabAISettings().set_api_key("LMStudio_Host", host)
                provider = router.get_active_provider()
                if provider and hasattr(provider, "reload_credentials"):
                    provider.reload_credentials()
            elif name == "ollama":
                # Ollama provider exposes a set_host() method
                provider = router.get_active_provider()
                if provider and hasattr(provider, "set_host"):
                    provider.set_host(host)

            # Clear model cache so next probe fetches live data from new host
            self._models_cache.pop(name, None)

            # Refresh badge and sidebar instantly, then kick background probe
            self._update_ai_badge()
            self._update_sidebar_instant()

            def _probe():
                try:
                    router.get_status(use_cache=False)
                    provider = router.get_active_provider()
                    live_models = provider.get_models() if provider else []
                    if live_models:
                        self._models_cache[name] = live_models
                    self.Dispatcher.Invoke(Action(self._update_sidebar))
                except Exception:
                    pass
            _ht = Thread(ThreadStart(_probe))
            _ht.IsBackground = True
            _ht.SetApartmentState(ApartmentState.STA)
            _ht.Start()
        except Exception as ex:
            logger.debug("sidebar_save_host_clicked error: {}".format(ex))

    # ─── Command palette ──────────────────────────────────────────────────────

    # All commands exposed to the AI, grouped by category.
    # Each entry: icon (Segoe MDL2 Assets codepoint), name, example phrase (inserted into chat).
    _COMMANDS = {
        "export": [
            (u"", u"Export PDF — All Sheets",       u"xuất pdf toàn bộ sheet"),
            (u"", u"Export DWG — All Sheets",       u"xuất dwg toàn bộ sheet"),
            (u"", u"Export G Sheets → PDF",         u"xuất pdf G sheet"),
            (u"", u"Export A Sheets → PDF",         u"xuất pdf A sheet"),
            (u"", u"Export IFC",                    u"xuất ifc"),
            (u"", u"Export Image (PNG/JPEG)",       u"xuất hình ảnh sheet"),
            (u"", u"Open BatchOut",                 u"mở batchout"),
            (u"", u"BatchOut — G Sheet PDF",        u"mở batchout G sheet pdf"),
            (u"", u"BatchOut — Configured",         u"mở batchout đã cấu hình"),
        ],
        "tools": [
            (u"", u"ParaSync",                      u"mở parasync"),
            (u"", u"Load Family",                   u"mở load family"),
            (u"", u"Load Family (Cloud)",           u"mở load family cloud"),
            (u"", u"DimText",                       u"mở dimtext"),
            (u"", u"Upper DimText",                 u"mở upper dimtext"),
            (u"", u"Reset Overrides",               u"reset graphic overrides"),
            (u"", u"Workset",                       u"mở workset"),
            (u"", u"Project Name",                  u"mở project name"),
            (u"", u"Grids",                         u"mở grids"),
        ],
        "ai": [
            (u"", u"Project Info",                  u"thông tin project hiện tại"),
            (u"", u"Active View Info",              u"view hiện tại là gì?"),
            (u"", u"Selected Elements",             u"thông tin element đã chọn"),
            (u"", u"Revit Question",                u"giải thích workset trong Revit"),
            (u"", u"Analyze Attachment",            u"phân tích tài liệu đính kèm"),
            (u"", u"Help",                          u"trợ giúp"),
        ],
        # ── Revit MCP commands (revit-mcp server) ──────────────────────────
        "query": [
            (u"", u"View Info",             u"thong tin view hien tai la gi?"),
            (u"", u"Elements in View",      u"lay danh sach element trong view hien tai"),
            (u"", u"Selected Elements",     u"thong tin cac element dang duoc chon"),
            (u"", u"Filter by Category",    u"loc tat ca tuong (OST_Walls) trong view"),
            (u"", u"Available Families",    u"family types cua di co san trong project"),
            (u"", u"Material Quantities",   u"tinh khoi luong vat lieu tuong trong du an"),
            (u"", u"Model Statistics",      u"thong ke toan bo model - so element, family, view"),
            (u"", u"Export Room Data",      u"xuat danh sach toan bo phong va dien tich"),
            (u"", u"Query Stored Data",     u"truy van du lieu project da luu"),
        ],
        "create": [
            (u"", u"Create Wall",           u"tao tuong dai 5000mm tu (0,0,0) den (5000,0,0) cao 3000mm"),
            (u"", u"Create Floor",          u"tao san be tong 10x10m tai Level 1"),
            (u"", u"Create Roof",           u"tao mai phang bao quanh building footprint"),
            (u"", u"Create Room",           u"tao phong ten Phong hop so 101 tai toa do (3000,3000,0)"),
            (u"", u"Create Grid System",    u"tao luoi truc A-F cach nhau 6000mm, 1-8 cach nhau 7200mm"),
            (u"", u"Create Level",          u"tao tang Level 2 o do cao 4000mm"),
            (u"", u"Create Door/Window",    u"tao cua di tai vi tri (2500,0,0) trong tuong"),
            (u"", u"Create Furniture",      u"tao ban lam viec tai (1000,1000,0) xoay 90 do"),
            (u"", u"Create Beam System",    u"tao he dam tai Level 2, khoang cach 1200mm, vung 10x20m"),
        ],
        "modify": [
            (u"", u"Select by ID",          u"chon element co ID 123456"),
            (u"", u"Color by Parameter",    u"to mau tuong theo loai (Type) - moi type mot mau"),
            (u"", u"Color Rooms by Dept",   u"to mau phong theo department"),
            (u"", u"Highlight Elements",    u"highlight do cac element co ID 111 222 333"),
            (u"", u"Hide Elements",         u"an element co ID 444 555 trong view"),
            (u"", u"Isolate Elements",      u"isolate tat ca cot (OST_Columns) trong view"),
            (u"", u"Reset Visibility",      u"reset isolate - khoi phuc hien thi binh thuong"),
            (u"", u"Delete Elements",       u"xoa element co ID 123 456 789"),
            (u"", u"Tag All Walls",         u"gan tag tat ca tuong trong view hien tai"),
            (u"", u"Tag All Rooms",         u"gan tag tat ca phong trong view hien tai"),
            (u"", u"Run C# Code",           u"chay code C# trong Revit de lay ten project"),
            (u"", u"Store Project Data",    u"luu metadata project hien tai vao database"),
        ],
    }

    _active_cmd_cat = "export"   # current palette category

    def cmd_palette_toggle_clicked(self, sender, e):
        """Toggle the command palette panel."""
        if self.cmd_palette.Visibility == Visibility.Visible:
            self.cmd_palette.Visibility = Visibility.Collapsed
        else:
            self.cmd_palette.Visibility = Visibility.Visible
            # Build initial category if panel was never opened
            if not self.cmd_cards_panel.Children.Count:
                self._build_cmd_cards(self._active_cmd_cat)

    def cmd_palette_close_clicked(self, sender, e):
        self.cmd_palette.Visibility = Visibility.Collapsed

    def cmd_cat_clicked(self, sender, e):
        """Switch the active category tab and rebuild the cards panel."""
        try:
            tag = sender.Tag
            if tag:
                self._active_cmd_cat = tag
                self._set_active_cmd_cat(sender)
                self._build_cmd_cards(tag)
        except Exception as ex:
            logger.debug("cmd_cat_clicked error: {}".format(ex))

    def _set_active_cmd_cat(self, active_btn):
        """Update category tab backgrounds/foregrounds."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            _DARK = SolidColorBrush(Color.FromRgb(24, 24, 27))
            _LITE = SolidColorBrush(Color.FromRgb(244, 244, 246))
            _WHITE = SolidColorBrush(Color.FromRgb(255, 255, 255))
            _INK  = SolidColorBrush(Color.FromRgb(39, 39, 42))

            for btn in (self.btn_cat_export, self.btn_cat_tools,
                        self.btn_cat_query, self.btn_cat_create, self.btn_cat_modify,
                        self.btn_cat_ai):
                if btn is active_btn:
                    btn.Background = _DARK
                    btn.Foreground = _WHITE
                else:
                    btn.Background = _LITE
                    btn.Foreground = _INK
        except Exception as ex:
            logger.debug("_set_active_cmd_cat error: {}".format(ex))

    def _build_cmd_cards(self, category):
        """Populate cmd_cards_panel with command cards for the given category."""
        try:
            self.cmd_cards_panel.Children.Clear()
            for icon, name, example in self._COMMANDS.get(category, []):
                card = self._make_cmd_card(icon, name, example)
                self.cmd_cards_panel.Children.Add(card)
        except Exception as ex:
            logger.debug("_build_cmd_cards error: {}".format(ex))

    def _make_cmd_card(self, icon, name, example):
        """Create a single clickable command card."""
        import System.Windows.Controls as WC
        import System.Windows as SW
        from System.Windows.Media import SolidColorBrush, Color, FontFamily
        from System.Windows.Input import Cursors

        _BG        = SolidColorBrush(Color.FromRgb(248, 250, 252))
        _BG_HOVER  = SolidColorBrush(Color.FromRgb(239, 246, 255))
        _BORDER    = SolidColorBrush(Color.FromRgb(226, 232, 240))
        _ACCENT    = SolidColorBrush(Color.FromRgb(59, 130, 246))
        _INK       = SolidColorBrush(Color.FromRgb(39, 39, 42))
        _MUTED     = SolidColorBrush(Color.FromRgb(100, 116, 139))

        card = WC.Border()
        card.Width            = 186
        card.Margin           = SW.Thickness(0, 0, 8, 8)
        card.Padding          = SW.Thickness(11, 9, 11, 9)
        card.CornerRadius     = SW.CornerRadius(10)
        card.Background       = _BG
        card.BorderBrush      = _BORDER
        card.BorderThickness  = SW.Thickness(1)
        card.Cursor           = Cursors.Hand
        card.Tag              = example

        sp = WC.StackPanel()

        # Icon row
        icon_row = WC.StackPanel()
        icon_row.Orientation = WC.Orientation.Horizontal
        icon_row.Margin      = SW.Thickness(0, 0, 0, 5)

        icon_tb = WC.TextBlock()
        icon_tb.Text       = icon
        icon_tb.FontFamily = FontFamily(u"Segoe MDL2 Assets")
        icon_tb.FontSize   = 13
        icon_tb.Foreground = _ACCENT
        icon_tb.Margin     = SW.Thickness(0, 0, 7, 0)
        icon_tb.VerticalAlignment = SW.VerticalAlignment.Center

        name_tb = WC.TextBlock()
        name_tb.Text          = name
        name_tb.FontSize      = 11
        name_tb.FontWeight    = SW.FontWeights.SemiBold
        name_tb.Foreground    = _INK
        name_tb.FontFamily    = FontFamily(u"Hanken Grotesk")
        name_tb.TextWrapping  = SW.TextWrapping.Wrap

        icon_row.Children.Add(icon_tb)
        icon_row.Children.Add(name_tb)

        # Example phrase
        ex_tb = WC.TextBlock()
        ex_tb.Text         = u'"' + example + u'"'
        ex_tb.FontSize     = 10
        ex_tb.Foreground   = _MUTED
        ex_tb.FontFamily   = FontFamily(u"Inter")
        ex_tb.TextWrapping = SW.TextWrapping.Wrap

        sp.Children.Add(icon_row)
        sp.Children.Add(ex_tb)
        card.Child = sp

        card.MouseLeftButtonUp += self._cmd_card_clicked
        card.MouseEnter += lambda s, ev: setattr(s, 'Background', _BG_HOVER)
        card.MouseLeave += lambda s, ev: setattr(s, 'Background', _BG)

        return card

    def _cmd_card_clicked(self, sender, e):
        """Insert the example phrase into the chat input and close the palette."""
        try:
            phrase = sender.Tag
            if phrase:
                self.chat_input.Text = phrase
                self.chat_input.Focus()
                self.chat_input.CaretIndex = len(phrase)
            self.cmd_palette.Visibility = Visibility.Collapsed
        except Exception as ex:
            logger.debug("_cmd_card_clicked error: {}".format(ex))

    # ── Brand colors for provider dot (Ellipse on ComboBox) ──────────────────
    _BRAND_COLORS = {
        "claude":   (217, 119,  87),   # Anthropic orange
        "openai":   ( 16, 163, 127),   # OpenAI green
        "deepseek": ( 37,  99, 235),   # DeepSeek blue
        "ollama":   ( 59, 130, 246),   # Ollama blue
        "lmstudio": (124,  58, 237),   # LM Studio purple
    }
    _PROV_INDEX = {"claude": 0, "openai": 1, "deepseek": 2, "ollama": 3, "lmstudio": 4}

    _KEY_PROVIDERS  = ("claude", "openai", "deepseek")
    _KEY_NAME_MAP   = {"claude": "Claude", "openai": "OpenAI", "deepseek": "DeepSeek"}

    # Where to obtain an API key for each key-based provider.
    _API_KEY_URLS = {
        "claude":   "https://console.anthropic.com/settings/keys",
        "openai":   "https://platform.openai.com/api-keys",
        "deepseek": "https://platform.deepseek.com/api_keys",
    }

    def get_api_key_clicked(self, sender, e):
        """Open the active provider's API-key page in the default browser."""
        try:
            from Intelligence.llm_router import LLMRouter
            name = LLMRouter().get_active_name()
            url  = self._API_KEY_URLS.get(name)
            if not url:
                return
            import System.Diagnostics
            System.Diagnostics.Process.Start(url)
        except Exception as ex:
            logger.debug("get_api_key_clicked error: {}".format(ex))

    def _has_saved_key(self, provider):
        """True if an API key is stored for a key-based provider."""
        try:
            from config.settings import T3LabAISettings
            return bool(T3LabAISettings().get_api_key(self._KEY_NAME_MAP.get(provider, "")))
        except Exception:
            return False

    def _populate_model_combo(self, active):
        """Fill the MODEL combo for `active`, enforcing the setup rule.

        A key-based provider exposes models ONLY after its key was saved and the
        connection verified — i.e. live models are present in the cache (the probe
        only caches them when check_health() succeeds). Until then the combo is
        empty + disabled with a hint. Local providers expose models once the
        server has been probed. No network here.
        """
        try:
            try:
                self.sidebar_model_combo.SelectionChanged -= self.sidebar_model_changed
            except Exception:
                pass

            self.sidebar_model_combo.Items.Clear()
            models    = list(self._models_cache.get(active, []))   # live, validated only
            enabled   = bool(models)
            needs_key = active in self._KEY_PROVIDERS
            hint      = u""

            if enabled:
                # Pin the saved model (stored per-provider) first if present.
                saved = None
                try:
                    from config.settings import T3LabAISettings
                    saved = T3LabAISettings().get_provider_model(active)
                except Exception:
                    pass
                if saved and saved not in models:
                    models.insert(0, saved)
                for m in models:
                    self.sidebar_model_combo.Items.Add(m)
                if saved and saved in models:
                    self.sidebar_model_combo.SelectedItem = saved
                else:
                    self.sidebar_model_combo.SelectedIndex = 0
            else:
                if needs_key:
                    hint = (u"Nhập API key trước" if not self._has_saved_key(active)
                            else u"Chưa kết nối — nhấn Save để kiểm tra")
                else:
                    hint = u"Khởi động server & load model"

            self.sidebar_model_combo.IsEnabled = enabled
            try:
                self.sidebar_save_model_btn.IsEnabled = enabled
            except Exception:
                pass

            try:
                from System.Windows.Media import SolidColorBrush, Color
                self.model_saved_hint.Text = hint
                self.model_saved_hint.Foreground = SolidColorBrush(Color.FromRgb(161, 161, 170))
            except Exception:
                pass
        except Exception as ex:
            logger.debug("_populate_model_combo error: {}".format(ex))
        finally:
            try:
                self.sidebar_model_combo.SelectionChanged += self.sidebar_model_changed
            except Exception:
                pass

    def _set_status_dot(self, name, available):
        """Update a single provider's STATUS dot + label (UI thread)."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            mapping = {
                "claude":   (self.status_dot_claude,   self.status_text_claude),
                "openai":   (self.status_dot_openai,   self.status_text_openai),
                "deepseek": (self.status_dot_deepseek, self.status_text_deepseek),
                "ollama":   (self.status_dot_ollama,   self.status_text_ollama),
                "lmstudio": (self.status_dot_lmstudio, self.status_text_lmstudio),
            }
            pair = mapping.get(name)
            if not pair:
                return
            dot, txt = pair
            if available:
                dot.Fill = SolidColorBrush(Color.FromRgb(16, 185, 129))
                txt.Text = u"Ready"
                txt.Foreground = SolidColorBrush(Color.FromRgb(16, 185, 129))
            else:
                dot.Fill = SolidColorBrush(Color.FromRgb(230, 230, 234))
                txt.Text = u"Not set up"
                txt.Foreground = SolidColorBrush(Color.FromRgb(161, 161, 170))
        except Exception:
            pass

    def _update_sidebar_instant(self):
        """Phase-1 sidebar render — zero HTTP, instant.  Uses cached/local data only."""
        try:
            from System.Windows.Media import SolidColorBrush, Color

            # Username Profile Section
            try:
                from config.user_profile import UserProfile
                self.sidebar_username_box.Text = UserProfile().get_name() or u"Thạnh"
            except Exception:
                pass
            self._update_profile_card()
            from Intelligence.llm_router import LLMRouter

            router = LLMRouter()
            active = router.get_active_name()

            # Provider ComboBox — select active provider
            self.sidebar_provider_combo.SelectionChanged -= self.sidebar_provider_changed
            self.sidebar_provider_combo.SelectedIndex = self._PROV_INDEX.get(active, 0)
            self.sidebar_provider_combo.SelectionChanged += self.sidebar_provider_changed

            # Brand-color dot on the provider ComboBox
            rgb = self._BRAND_COLORS.get(active, (161, 161, 170))
            self.provider_brand_dot.Fill = SolidColorBrush(
                Color.FromRgb(rgb[0], rgb[1], rgb[2]))

            # API key section (from settings.json — no HTTP)
            _KEY_LABELS = {
                "claude":   u"Anthropic API Key (sk-ant-...)",
                "openai":   u"OpenAI API Key (sk-...)",
                "deepseek": u"DeepSeek API Key (sk-...)",
                "ollama":   u"No key needed (local)",
                "lmstudio": u"No key needed — start LM Studio first",
            }
            self.sidebar_key_label.Text = _KEY_LABELS.get(active, u"API Key")

            needs_key = active in self._KEY_PROVIDERS
            self.sidebar_api_key_box.IsEnabled  = needs_key
            self.sidebar_save_key_btn.IsEnabled = needs_key
            try:
                self.get_api_key_link.Visibility = (
                    Visibility.Visible if needs_key else Visibility.Collapsed)
            except Exception:
                pass

            if needs_key:
                try:
                    from config.settings import T3LabAISettings
                    k_map = {"claude": "Claude", "openai": "OpenAI", "deepseek": "DeepSeek"}
                    saved = T3LabAISettings().get_api_key(k_map.get(active, "")) or ""
                    self.sidebar_api_key_box.Text = (saved[:8] + u"...") if len(saved) > 8 else saved
                except Exception:
                    pass
            else:
                self.sidebar_api_key_box.Text = u""

            # Model ComboBox — instant, no HTTP (cache → fallback → saved model).
            self._populate_model_combo(active)

        except Exception as ex:
            logger.debug("_update_sidebar_instant error: {}".format(ex))

    def _update_sidebar(self):
        """Phase-2 sidebar refresh — called from background after HTTP probes complete."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            from Intelligence.llm_router import LLMRouter

            router = LLMRouter()
            status = router.get_status(use_cache=True)   # cache already warm from bg probe
            active = router.get_active_name()

            _GRAY  = Color.FromRgb(230, 230, 234)
            _MUTED = Color.FromRgb(161, 161, 170)
            _READY = Color.FromRgb( 16, 185, 129)

            # Provider ComboBox + brand dot
            self.sidebar_provider_combo.SelectionChanged -= self.sidebar_provider_changed
            self.sidebar_provider_combo.SelectedIndex = self._PROV_INDEX.get(active, 0)
            self.sidebar_provider_combo.SelectionChanged += self.sidebar_provider_changed

            rgb = self._BRAND_COLORS.get(active, (161, 161, 170))
            self.provider_brand_dot.Fill = SolidColorBrush(
                Color.FromRgb(rgb[0], rgb[1], rgb[2]))

            # API key section
            _KEY_LABELS = {
                "claude":   u"Anthropic API Key (sk-ant-...)",
                "openai":   u"OpenAI API Key (sk-...)",
                "deepseek": u"DeepSeek API Key (sk-...)",
                "ollama":   u"No key needed (local)",
                "lmstudio": u"No key needed — start LM Studio first",
            }
            self.sidebar_key_label.Text = _KEY_LABELS.get(active, u"API Key")

            needs_key = active in ("claude", "openai", "deepseek")
            self.sidebar_api_key_box.IsEnabled  = needs_key
            self.sidebar_save_key_btn.IsEnabled = needs_key
            try:
                self.get_api_key_link.Visibility = (
                    Visibility.Visible if needs_key else Visibility.Collapsed)
            except Exception:
                pass

            if needs_key:
                try:
                    from config.settings import T3LabAISettings
                    k_map = {"claude": "Claude", "openai": "OpenAI", "deepseek": "DeepSeek"}
                    saved = T3LabAISettings().get_api_key(k_map.get(active, "")) or ""
                    self.sidebar_api_key_box.Text = (saved[:8] + u"...") if len(saved) > 8 else saved
                except Exception:
                    pass
            else:
                self.sidebar_api_key_box.Text = u""

            # Model ComboBox — refresh from the freshly probed cache for `active`.
            self._populate_model_combo(active)

            # Status dots
            dot_rows = [
                ("claude",   self.status_dot_claude,   self.status_text_claude),
                ("openai",   self.status_dot_openai,   self.status_text_openai),
                ("deepseek", self.status_dot_deepseek, self.status_text_deepseek),
                ("ollama",   self.status_dot_ollama,   self.status_text_ollama),
                ("lmstudio", self.status_dot_lmstudio, self.status_text_lmstudio),
            ]
            for name, dot, txt in dot_rows:
                available  = status.get(name, {}).get("available", False)
                dot.Fill   = SolidColorBrush(_READY if available else _GRAY)
                txt.Text   = u"Ready" if available else u"Not set up"
                txt.Foreground = SolidColorBrush(_READY if available else _MUTED)

        except Exception as ex:
            logger.debug("_update_sidebar error: {}".format(ex))

    # ─── Session guard & UI state ─────────────────────────────────────────────

    def _set_busy(self, busy):
        """Lock/unlock the whole input area. Call from UI thread only."""
        self._busy = busy
        try:
            self.send_button.IsEnabled  = not busy
            self.chat_input.IsEnabled   = not busy
            # btn_attach stays locked (import file feature disabled)
        except Exception:
            pass
        for btn in self._DYNAMIC_BTNS:
            try:
                btn.IsEnabled = not busy
            except Exception:
                pass
        
        # Toggle top ProgressBar visibility
        try:
            from System.Windows import Visibility
            if busy:
                self.top_loading_bar.Visibility = Visibility.Visible
            else:
                self.top_loading_bar.Visibility = Visibility.Collapsed
        except Exception:
            pass

        if busy:
            self._show_typing_indicator()
        else:
            self._hide_typing_indicator()
    def _safe_update_typing_text(self, text):
        """Thread-safe update of the typing indicator text."""
        def action():
            try:
                if hasattr(self, "_typing_text_block") and self._typing_text_block is not None:
                    self._typing_text_block.Text = text
            except Exception:
                pass
        try:
            self.Dispatcher.Invoke(Action(action))
        except Exception:
            pass

    def _show_typing_indicator(self):
        """Add an animated '● ● ●' bubble to the chat."""
        try:
            if self._typing_row is not None:
                return  # already shown
            self._typing_row = self._make_typing_row()
            self.chat_history_panel.Children.Add(self._typing_row)
            self._scroll_to_bottom()

            # Start dynamic text timer
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            self._typing_elapsed = 0
            self._update_typing_text() # initial text
            
            self._typing_timer = DispatcherTimer()
            self._typing_timer.Interval = TimeSpan.FromSeconds(1)
            self._typing_timer.Tick += self._on_typing_timer_tick
            self._typing_timer.Start()
        except Exception:
            pass

    def _hide_typing_indicator(self):
        """Remove the typing indicator bubble."""
        try:
            if self._typing_timer is not None:
                self._typing_timer.Stop()
                self._typing_timer = None
            if self._typing_row is not None:
                self.chat_history_panel.Children.Remove(self._typing_row)
                self._typing_row = None
                self._typing_text_block = None
        except Exception:
            pass

    def _on_typing_timer_tick(self, sender, e):
        self._typing_elapsed += 1
        self._update_typing_text()

    def _update_typing_text(self):
        try:
            if not hasattr(self, "_typing_text_block") or self._typing_text_block is None:
                return
            
            # Select Vietnamese text if the input was Vietnamese, else English
            is_vn = _is_viet_text(self._last_raw)
            
            if self._typing_elapsed < 1:
                txt = u"● ● ●  Đang đọc dữ liệu Revit..." if is_vn else "● ● ●  Reading Revit data..."
            elif self._typing_elapsed < 3:
                txt = u"● ● ●  Đang phân tích yêu cầu..." if is_vn else "● ● ●  Formulating response..."
            else:
                txt = u"● ● ●  Đang phản hồi..." if is_vn else "● ● ●  Responding..."
                
            self._typing_text_block.Text = txt
        except Exception:
            pass

    def _make_typing_row(self):
        """Build the typing indicator WPF element with text description."""
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
        dots.FontSize  = 12
        dots.Foreground = SolidColorBrush(Color.FromRgb(127, 140, 141))  # #7F8C8D
        
        self._typing_text_block = dots

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

    # ─── File attachment ──────────────────────────────────────────────────────

    def attach_clicked(self, sender, e):
        """Open a file picker and add selected file to attachment list."""
        try:
            import clr
            clr.AddReference('System.Windows.Forms')
            from System.Windows.Forms import OpenFileDialog, DialogResult

            exts = "PDF và Hình ảnh|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"
            dlg = OpenFileDialog()
            dlg.Title  = u"Chọn PDF hoặc hình ảnh để đính kèm"
            dlg.Filter = exts
            dlg.Multiselect = True

            if dlg.ShowDialog() == DialogResult.OK:
                for path in dlg.FileNames:
                    if not HAS_RAG or is_supported(path):
                        if path not in self._attached_files:
                            self._attached_files.append(path)
                            self._add_attachment_chip(path)
                self._refresh_attachment_panel()
        except Exception as ex:
            logger.error("attach_clicked error: {}".format(ex))

    def clear_attachments_clicked(self, sender, e):
        """Remove all attachments."""
        self._attached_files = []
        try:
            self.attachment_chips_panel.Children.Clear()
            self._refresh_attachment_panel()
        except Exception:
            pass

    def _add_attachment_chip(self, file_path):
        """Add a small chip label for an attached file."""
        try:
            from System.Windows.Controls import Button, StackPanel, TextBlock
            from System.Windows import Thickness
            import os as _os

            name = _os.path.basename(file_path)
            ext  = _os.path.splitext(name)[1].lower()
            icon = u"🖼️" if ext in ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp') else u"📄"

            btn = Button()
            try:
                btn.Style = self.FindResource('AttachChipBtn')
            except Exception:
                pass
            btn.Margin = Thickness(0, 0, 4, 4)

            sp = StackPanel()
            sp.Orientation = System.Windows.Controls.Orientation.Horizontal

            icon_lbl = TextBlock()
            icon_lbl.Text = icon + u" "
            icon_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(icon_lbl)

            name_lbl = TextBlock()
            name_lbl.Text = name if len(name) <= 22 else name[:19] + u"..."
            name_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(name_lbl)

            x_lbl = TextBlock()
            x_lbl.Text = u"  ✕"
            x_lbl.FontSize = 9
            x_lbl.Foreground = System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(150, 150, 150))
            x_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(x_lbl)

            btn.Content = sp
            btn.ToolTip = file_path

            _path = file_path

            def _on_remove(s, ev, p=_path):
                if p in self._attached_files:
                    self._attached_files.remove(p)
                try:
                    self.attachment_chips_panel.Children.Remove(s)
                except Exception:
                    pass
                self._refresh_attachment_panel()

            btn.Click += _on_remove
            self.attachment_chips_panel.Children.Add(btn)
        except Exception as ex:
            logger.debug("_add_attachment_chip error: {}".format(ex))

    def _refresh_attachment_panel(self):
        """Show or hide the attachment preview border depending on file list."""
        try:
            self.attachment_preview_border.Visibility = (
                Visibility.Visible if self._attached_files else Visibility.Collapsed
            )
        except Exception:
            pass

    # ─── Chat input ───────────────────────────────────────────────────────────

    def send_clicked(self, sender, e):
        self._process_input()

    def input_keydown(self, sender, e):
        from System.Windows.Input import Key, Keyboard, ModifierKeys
        if e.Key == Key.Return or e.Key == Key.Enter:
            # Shift+Enter inserts a newline (multi-line input)
            if (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift:
                caret = self.chat_input.CaretIndex
                text = self.chat_input.Text or ""
                self.chat_input.Text = text[:caret] + "\n" + text[caret:]
                self.chat_input.CaretIndex = caret + 1
                e.Handled = True
                return
            
            # Enter sends.
            self._process_input()
            e.Handled = True
        elif e.Key == Key.Up:
            # Don't hijack Up while editing a multi-line draft.
            if u"\n" in (self.chat_input.Text or u""):
                return
            if self._input_history:
                if self._history_index == -1:
                    self._current_input_temp = self.chat_input.Text
                    self._history_index = len(self._input_history) - 1
                else:
                    self._history_index = max(0, self._history_index - 1)
                self.chat_input.Text = self._input_history[self._history_index]
                self.chat_input.CaretIndex = len(self.chat_input.Text)
                e.Handled = True
        elif e.Key == Key.Down:
            if self._input_history and self._history_index != -1:
                if self._history_index == len(self._input_history) - 1:
                    self.chat_input.Text = self._current_input_temp
                    self._history_index = -1
                else:
                    self._history_index += 1
                    self.chat_input.Text = self._input_history[self._history_index]
                self.chat_input.CaretIndex = len(self.chat_input.Text)
                e.Handled = True

    # Keyword groups for the DB-only fast path (no LLM round-trip)
    _FAST_CTX_PROJECT = (
        u"thông tin dự án", u"thong tin du an", u"thông tin project", u"thông tin model",
        u"thong tin model", u"dự án này", u"du an nay", u"project info",
        u"project information", u"about project", u"model info", u"project", u"dự án", u"du an",
    )
    _FAST_CTX_VIEW = (
        u"view hiện tại", u"view hien tai", u"thông tin view", u"thong tin view",
        u"current view", u"active view", u"view đang mở", u"view dang mo",
        u"activeview", u"currentview", u"view",
    )
    _FAST_CTX_SELECTION = (
        u"đang chọn", u"dang chon", u"đang được chọn", u"dang duoc chon",
        u"selected element", u"selection", u"đã chọn", u"da chon",
        u"bao nhiêu element", u"how many selected", u"sel", u"selected",
        u"đang chọn gì", u"dang chon gi",
    )
    # Action verbs and terms that prevent fast-path routing (exact word-level match)
    _FAST_CTX_EXCLUDE_WORDS = {
        u"tạo", u"create", u"mở", u"open", u"xuất", u"export", u"in", u"print",
        u"ẩn", u"hide", u"color", u"màu", u"xoá", u"delete", u"remove",
        u"vẽ", u"draw", u"tag", u"add", u"thêm", u"sửa", u"edit", u"update",
        u"chạy", u"run", u"c#", u"csharp", u"load", u"tải", u"nap",
        u"batchout", u"parasync", u"dim", u"override", u"isolate", u"highlight"
    }

    # Substrings that prevent fast-path routing
    _FAST_CTX_EXCLUDE_SUBS = (
        u"batchout", u"parasync", u"override", u"isolate", u"highlight"
    )

    # Specific Revit element type and query keywords that prevent generic fast-path hijacking
    _FAST_CTX_ELEMENT_WORDS = {
        u"tường", u"wall", u"cửa", u"door", u"sàn", u"floor", u"mái", u"roof",
        u"phòng", u"room", u"dầm", u"beam", u"cột", u"column", u"trần", u"ceiling",
        u"family", u"type", u"vật liệu", u"material", u"level", u"tầng", u"grid",
        u"lưới", u"sheet", u"bản vẽ", u"ban ve", u"parameter", u"tham số", u"tham so",
        u"khối lượng", u"khoi luong", u"quantity", u"count", u"nhiêu", u"nhieu",
        u"bao nhiêu", u"bao nhieu", u"how many", u"list", u"danh sách", u"danh sach"
    }

    _FAST_CTX_ELEMENT_MAP = {
        u"tường": ([DB.BuiltInCategory.OST_Walls], u"Tường", u"Walls"),
        u"wall": ([DB.BuiltInCategory.OST_Walls], u"Tường", u"Walls"),
        u"cửa": ([DB.BuiltInCategory.OST_Doors], u"Cửa đi", u"Doors"),
        u"door": ([DB.BuiltInCategory.OST_Doors], u"Cửa đi", u"Doors"),
        u"sàn": ([DB.BuiltInCategory.OST_Floors], u"Sàn", u"Floors"),
        u"floor": ([DB.BuiltInCategory.OST_Floors], u"Sàn", u"Floors"),
        u"mái": ([DB.BuiltInCategory.OST_Roofs], u"Mái", u"Roofs"),
        u"roof": ([DB.BuiltInCategory.OST_Roofs], u"Mái", u"Roofs"),
        u"phòng": ([DB.BuiltInCategory.OST_Rooms], u"Phòng", u"Rooms"),
        u"room": ([DB.BuiltInCategory.OST_Rooms], u"Phòng", u"Rooms"),
        u"dầm": ([DB.BuiltInCategory.OST_StructuralFraming], u"Dầm", u"Beams"),
        u"beam": ([DB.BuiltInCategory.OST_StructuralFraming], u"Dầm", u"Beams"),
        u"cột": ([DB.BuiltInCategory.OST_StructuralColumns, DB.BuiltInCategory.OST_Columns], u"Cột", u"Columns"),
        u"column": ([DB.BuiltInCategory.OST_StructuralColumns, DB.BuiltInCategory.OST_Columns], u"Cột", u"Columns"),
        u"trần": ([DB.BuiltInCategory.OST_Ceilings], u"Trần", u"Ceilings"),
        u"ceiling": ([DB.BuiltInCategory.OST_Ceilings], u"Trần", u"Ceilings"),
        u"cửa sổ": ([DB.BuiltInCategory.OST_Windows], u"Cửa sổ", u"Windows"),
        u"window": ([DB.BuiltInCategory.OST_Windows], u"Cửa sổ", u"Windows"),
        u"lưới": ([DB.BuiltInCategory.OST_Grids], u"Lưới trục", u"Grids"),
        u"grid": ([DB.BuiltInCategory.OST_Grids], u"Lưới trục", u"Grids"),
        u"level": ([DB.BuiltInCategory.OST_Levels], u"Tầng", u"Levels"),
        u"tầng": ([DB.BuiltInCategory.OST_Levels], u"Tầng", u"Levels"),
    }

    def _count_elements(self, category_list, in_active_view=False):
        try:
            from Autodesk.Revit import DB as _DB
            total = 0
            for bic in category_list:
                if in_active_view and self.doc.ActiveView:
                    collector = _DB.FilteredElementCollector(self.doc, self.doc.ActiveView.Id)
                else:
                    collector = _DB.FilteredElementCollector(self.doc)
                collector.OfCategory(bic).WhereElementIsNotElementType()
                total += collector.GetElementCount()
            return total
        except Exception as ex:
            logger.debug("_count_elements error: {}".format(ex))
            return 0

    def _try_fast_context_answer(self, raw):
        """Answer project/view/selection questions directly from Revit (no LLM).

        Returns a formatted message string if the query matches a known context
        question, else None (so the normal NLP/LLM pipeline runs).
        """
        try:
            if not raw:
                return None
            low = raw.lower().strip()

            # Check for fast element count query
            count_kws = [u"bao nhiêu", u"bao nhieu", u"đếm", u"dem", u"số lượng", u"so luong", u"count", u"how many", u"tổng số", u"tong so"]
            want_count = any(k in low for k in count_kws)
            target_cats = []
            cat_display_vn = u""
            cat_display_en = u""
            
            if want_count:
                for kw, (bics, vn_name, en_name) in self._FAST_CTX_ELEMENT_MAP.items():
                    if re.search(r'\b' + re.escape(kw) + r'\b', low):
                        target_cats = bics
                        cat_display_vn = vn_name
                        cat_display_en = en_name
                        break
            
            if want_count and target_cats:
                in_view = any(k in low for k in [u"trong view", u"view này", u"view dang mo", u"view đang mở", u"view hiện tại", u"view hien tai", u"in view", u"active view", u"current view"])
                count = self._count_elements(target_cats, in_active_view=in_view)
                
                viet = _is_viet_text(raw)
                lines = []
                lines.append(u"⚡ **Phản hồi tức thì từ Revit DB**")
                lines.append(u"")
                if viet:
                    scope_str = u"trong view hiện tại" if in_view else u"trong toàn bộ dự án"
                    lines.append(u"📋 **Thống kê cấu kiện**")
                    lines.append(u"- Số lượng {}: **{}** ({})".format(cat_display_vn, count, scope_str))
                else:
                    scope_str = u"in active view" if in_view else u"in entire project"
                    lines.append(u"📋 **Element Statistics**")
                    lines.append(u"- Number of {}: **{}** ({})".format(cat_display_en, count, scope_str))
                return u"\n".join(lines)

            # Word-level and substring-level checks to prevent hijacking commands
            words = set(low.split())
            if (any(w in words for w in self._FAST_CTX_EXCLUDE_WORDS) 
                    or any(s in low for s in self._FAST_CTX_EXCLUDE_SUBS)
                    or any(e in words for e in self._FAST_CTX_ELEMENT_WORDS)):
                return None

            want_project   = any(k in low for k in self._FAST_CTX_PROJECT)
            want_view      = any(k in low for k in self._FAST_CTX_VIEW)
            want_selection = any(k in low for k in self._FAST_CTX_SELECTION)
            if not (want_project or want_view or want_selection):
                return None

            ctx = ContextScout.get_active_context()
            if not ctx or "error" in ctx:
                return None

            viet = _is_viet_text(raw)
            lines = []

            # Add premium visual badge
            if viet:
                lines.append(u"⚡ **Phản hồi tức thì từ Revit DB**")
            else:
                lines.append(u"⚡ **Instant Revit DB Answer**")
            lines.append(u"")

            if want_project:
                p = ctx.get("project", {})
                r = ctx.get("revit", {})
                if viet:
                    lines.append(u"📋 **Thông tin dự án**")
                    lines.append(u"- Tên file: {}".format(p.get("title") or u"—"))
                    lines.append(u"- Tên dự án: {}".format(p.get("name") or u"—"))
                    lines.append(u"- Mã số: {}".format(p.get("number") or u"—"))
                    lines.append(u"- Khu vực: {}".format(p.get("region") or u"—"))
                    lines.append(u"- Revit: {}".format(r.get("version") or u"—"))
                else:
                    lines.append(u"📋 **Project information**")
                    lines.append(u"- File: {}".format(p.get("title") or u"—"))
                    lines.append(u"- Name: {}".format(p.get("name") or u"—"))
                    lines.append(u"- Number: {}".format(p.get("number") or u"—"))
                    lines.append(u"- Region: {}".format(p.get("region") or u"—"))
                    lines.append(u"- Revit: {}".format(r.get("version") or u"—"))

            if want_view:
                v = ctx.get("active_view", {})
                if len(lines) > 2:
                    lines.append(u"")
                if viet:
                    lines.append(u"🖼️ **View hiện tại**")
                    lines.append(u"- Tên: {}".format(v.get("name") or u"—"))
                    lines.append(u"- Loại: {}".format(v.get("type") or u"—"))
                    lines.append(u"- Tỷ lệ: {}".format(v.get("scale") or u"—"))
                    lines.append(u"- Bộ môn: {}".format(v.get("discipline") or u"—"))
                else:
                    lines.append(u"🖼️ **Active view**")
                    lines.append(u"- Name: {}".format(v.get("name") or u"—"))
                    lines.append(u"- Type: {}".format(v.get("type") or u"—"))
                    lines.append(u"- Scale: {}".format(v.get("scale") or u"—"))
                    lines.append(u"- Discipline: {}".format(v.get("discipline") or u"—"))

            if want_selection:
                s = ctx.get("selection", {})
                cnt = s.get("count", 0)
                details = s.get("details", [])
                if len(lines) > 2:
                    lines.append(u"")
                if viet:
                    lines.append(u"🎯 **Đối tượng đang chọn: {}**".format(cnt))
                    if details:
                        for idx, d in enumerate(details):
                            lines.append(u"  {}. {} (Category: {}) [ID: {}]".format(idx + 1, d["name"], d["category"], d["id"]))
                        if cnt > len(details):
                            lines.append(u"  *... và {} đối tượng khác.*".format(cnt - len(details)))
                else:
                    lines.append(u"🎯 **Selected elements: {}**".format(cnt))
                    if details:
                        for idx, d in enumerate(details):
                            lines.append(u"  {}. {} (Category: {}) [ID: {}]".format(idx + 1, d["name"], d["category"], d["id"]))
                        if cnt > len(details):
                            lines.append(u"  *... and {} more item(s).*".format(cnt - len(details)))

            return u"\n".join(lines) if len(lines) > 2 else None
        except Exception as ex:
            logger.debug("_try_fast_context_answer error: {}".format(ex))
            return None

    def _process_input(self):
        """Read input (+ any attachments), dispatch to NLP or keyword fallback."""
        try:
            raw = self.chat_input.Text.strip()
            attached = list(self._attached_files)   # snapshot

            # Must have text OR attachments
            if not raw and not attached:
                return

            # Record in command history
            if raw:
                if not self._input_history or self._input_history[-1] != raw:
                    self._input_history.append(raw)
                    if len(self._input_history) > 50:
                        self._input_history.pop(0)
                self._history_index = -1
                self._current_input_temp = ""

            # ── Concurrency guard ─────────────────────────────────────────────
            if self._busy:
                self._append_bot_message(
                    u"⏳ Đang xử lý lệnh trước, vui lòng chờ một chút..."
                )
                return


            self.chat_input.Text = ""
            self._last_raw = raw or u"[đính kèm tài liệu]"

            # ── Show user message in chat ──────────────────────────────────────
            display_text = raw
            if attached:
                attach_label = u"\n📎 " + summarize_attachments(attached)
                display_text = (raw + attach_label) if raw else attach_label.strip()
            self._append_user_message(display_text)
            self._add_to_history("user", display_text)

            # ── Clear attachments from UI after sending ────────────────────────
            if attached:
                self._attached_files = []
                try:
                    self.attachment_chips_panel.Children.Clear()
                    self._refresh_attachment_panel()
                except Exception:
                    pass

            # Lock UI
            self._set_busy(True)

            # ── If attachments present and no tool-like text, go straight to RAG ─
            has_attach = bool(attached) and HAS_RAG
            if has_attach and not raw:
                # No text — summarise the documents
                raw = u"Phân tích và tóm tắt nội dung tài liệu đính kèm."

            # Build context-enriched prompt for NLP / Claude
            # (PDF text is injected; images will be sent via vision API)
            rag_context = ''
            if has_attach:
                rag_context = build_text_context(attached)

            # For NLP routing we use only the raw text (no PDF dump)
            captured = raw
            if HAS_SCOUT:
                captured = ContextScout.get_context_summary_for_ai() + "\n" + raw

            history  = list(self._conversation_history[:-1])

            use_local        = HAS_NLP and has_local_llm()
            use_claude       = HAS_NLP and has_api_key()   # True for any configured provider
            _active_provider = get_active_provider_name()  # "claude" | "openai" | "ollama"

            # ── 0. Fast context answer (DB-only, no LLM) ──────────────
            if HAS_SCOUT and not has_attach:
                fast = self._try_fast_context_answer(raw)
                if fast:
                    self._hide_typing_indicator()
                    self._append_bot_message(fast)
                    self._add_to_history("assistant", fast)
                    self._set_busy(False)
                    return

            # ── 1. Learned patterns (skip if attachments present) ─────────────
            if HAS_NLP and not has_attach:
                learned = find_learned_match(raw)
                if learned:
                    self._execute_result(learned)
                    return

            # ── 2. Built-in NLU (skip for RAG / attachment queries) ───────────
            nlu_result = None
            if HAS_NLP and not has_attach:
                nlu_result = parse_command_nlu(captured, history)
                if nlu_result and nlu_result.get("intent") not in (None, "unknown"):
                    if nlu_result["intent"] not in ("chat", "help") \
                            or not (use_local or use_claude):
                        self._execute_result(nlu_result)
                        return

            if use_local or use_claude or has_attach:
                # ── 3/4. Async LLM path ────────────────────────────────────────
                self._show_typing_indicator()
                nlu_hint = nlu_result if (HAS_NLP and not has_attach) else None

                def do_nlp():
                    result = None
                    from Intelligence.llm_router import LLMRouter
                    from Intelligence.t3lab_agent import build_system_prompt
                    import json as _json

                    _router = LLMRouter()
                    _provider = _router.get_active_provider()

                    # Run up to 5 iterations of tool execution
                    max_iterations = 5
                    current_iteration = 0
                    current_history = list(history)

                    # Initial user prompt
                    current_query = (rag_context + u"\n\n" + captured) if rag_context else captured

                    while current_iteration < max_iterations:
                        current_iteration += 1

                        # Dynamically query Revit context on each iteration
                        _ctx_block = u""
                        if HAS_SCOUT:
                            try:
                                _ctx_block = ContextScout.get_context_summary_for_ai()
                            except Exception:
                                pass

                        # Retrieve registered tools from server.py
                        server_tools_str = u""
                        try:
                            from core.server import get_t3labai_server
                            srv = get_t3labai_server()
                            tools_list = srv._handle_tools_list().get('tools', [])
                            if tools_list:
                                server_tools_str = u"\n\nLocal MCP Server Tools:\n"
                                for tool in tools_list:
                                    server_tools_str += u"- `{}`: {} (Schema: {})\n".format(
                                        tool['name'],
                                        tool['description'],
                                        _json.dumps(tool['inputSchema'])
                                    )
                        except Exception as tool_err:
                            logger.debug("Failed to list server tools: {}".format(tool_err))

                        system_prompt = build_system_prompt(revit_context=_ctx_block)
                        if server_tools_str:
                            system_prompt += server_tools_str

                        if has_attach or rag_context:
                            system_prompt = _RAG_SYSTEM_PREFIX + system_prompt

                        # Perform the chat completion.
                        # Iteration 1 streams live into a growing bubble (smooth,
                        # low time-to-first-token). Subsequent tool-loop turns are
                        # internal, so they use a plain blocking call.
                        _resp = None
                        try:
                            if current_iteration == 1:
                                _resp = self._stream_llm_turn(
                                    _provider, _router, current_history,
                                    system_prompt, current_query, max_tokens=1200,
                                    response_format={"type": "json_object"}
                                )
                            elif _provider and _provider.check_health():
                                _resp = _provider.chat(current_history[-16:], system_prompt, current_query, max_tokens=1200, response_format={"type": "json_object"})
                            else:
                                _resp = _router.chat(current_history[-16:], system_prompt, current_query, max_tokens=1200, response_format={"type": "json_object"})
                        except Exception as chat_ex:
                            logger.debug("Router chat error: {}".format(chat_ex))

                        if not _resp or not _resp.strip():
                            break

                        # Parse response JSON
                        _cleaned = self._clean_bot_response(_resp)
                        _parsed = None
                        try:
                            import re as _re
                            # First try extracting JSON block using regex if model included extra text
                            _m = _re.search(r'\{[\s\S]*\}', _resp)
                            if _m:
                                _parsed = _json.loads(_m.group())
                            else:
                                # Fallback to direct parsing
                                _parsed = _json.loads(_resp)
                                
                            # Ensure intent is present
                            if "intent" not in _parsed:
                                _parsed["intent"] = "unknown"
                        except Exception:
                            # Fallback if model returns corrupted JSON/text
                            _parsed = {
                                "intent": "unknown",
                                "message": u"Không thể đọc dữ liệu từ Model. Vui lòng thử lại."
                            }

                        if _parsed and _parsed.get("intent"):
                            intent = _parsed.get("intent")
                            params = _parsed.get("params", {}) or {}
                            message = _parsed.get("message", "")

                            # Determine if it's a local MCP tool call
                            is_local_tool = False
                            if intent.startswith("revit_"):
                                is_local_tool = True
                            else:
                                try:
                                    from core.server import get_t3labai_server
                                    srv = get_t3labai_server()
                                    if intent in srv._tools:
                                        is_local_tool = True
                                except Exception:
                                    pass

                            if is_local_tool:
                                # The streamed preview (if any) is superseded by
                                # explicit tool-execution feedback below.
                                self.Dispatcher.Invoke(Action(self._remove_stream_bubble))
                                # Update typing indicator UI to show tool execution
                                is_vn = _is_viet_text(captured)
                                status_msg = u"● ● ●  Đang chạy công cụ `{}`...".format(intent) if is_vn else "● ● ●  Executing tool `{}`...".format(intent)
                                self._safe_update_typing_text(status_msg)

                                # Display temporary feedback message in chat so the user is updated
                                tool_display_msg = u"🔧 [Tool Call] `{}` (params: {})".format(intent, _json.dumps(params))
                                if message:
                                    tool_display_msg = u"🔧 [Tool Call] {}\nRevit tool: `{}`".format(message, intent)
                                self._safe_append_bot(tool_display_msg)

                                # Execute the tool in the Revit context using the external event handler
                                tool_result = None
                                try:
                                    from core.server import get_t3labai_server
                                    srv = get_t3labai_server()
                                    tool_result = srv._execute_tool(intent, params)
                                except Exception as execute_err:
                                    tool_result = {"error": str(execute_err)}

                                # Log tool call and result to current_history using portable roles
                                current_history.append({"role": "assistant", "content": _json.dumps(_parsed, ensure_ascii=False)})
                                current_history.append({"role": "user", "content": u"Tool `{}` returned: {}".format(intent, _json.dumps(tool_result, ensure_ascii=False))})

                                # Setup subsequent query
                                current_query = u"Tool `{}` successfully returned: {}. Please proceed to the next step or conclude if done.".format(intent, _json.dumps(tool_result, ensure_ascii=False))
                                continue
                            else:
                                result = _parsed
                                break
                        else:
                            result = {"intent": "chat", "message": _cleaned}
                            break

                    def finish():
                        try:
                            has_stream = (self._stream_tb is not None)
                            r_intent   = result.get("intent") if result else None
                            _conv      = ("chat", "help", "greet")

                            # Plain conversational reply that already streamed live
                            # → keep the bubble, just apply markdown + record it.
                            if has_stream and r_intent in _conv:
                                if r_intent == "help":
                                    msg = result.get("params", {}).get(
                                        "answer", result.get("message", ""))
                                else:
                                    msg = result.get("message", "")
                                msg = self._clean_bot_response(msg) if msg else u""
                                if not msg:
                                    try:
                                        msg = self._stream_tb.Text or u""
                                    except Exception:
                                        msg = u""
                                if not msg:
                                    msg = u"Có thể giúp gì thêm không?"
                                self._finalize_stream_bubble(msg)
                                self._add_to_history("assistant", msg)
                                self._clear_stream_refs()
                                self._set_busy(False)
                                return

                            # Action/tool result → discard any preview bubble and
                            # let the executor render its own message + run it.
                            if has_stream:
                                self._remove_stream_bubble()
                            self._hide_typing_indicator()
                            if result and result.get("intent") not in (None, "unknown"):
                                self._execute_result(result)
                            elif nlu_hint and nlu_hint.get("intent") not in (None, "unknown"):
                                self._execute_result(nlu_hint)
                            else:
                                if has_attach and not use_claude and not use_local:
                                    if rag_context:
                                        self._append_bot_message(
                                            u"📄 Nội dung tài liệu:\n\n" + rag_context[:2000]
                                        )
                                    else:
                                        self._append_bot_message(
                                            u"Không trích xuất được văn bản từ tài liệu. PDF có thể là dạng scan."
                                        )
                                    self._set_busy(False)
                                    return
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
                            self._clear_stream_refs()
                            self._set_busy(False)

                    self.Dispatcher.Invoke(Action(finish))

                t = Thread(ThreadStart(do_nlp))
                t.IsBackground = True
                t.SetApartmentState(ApartmentState.STA)
                t.Start()
            else:
                # ── 5. Keyword fallback ────────────────────────────────────────
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
            reply = self._clean_bot_response(reply) if reply else u""
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

        # ── MCP Revit intents ─────────────────────────────────────────────────
        try:
            from Intelligence.t3lab_agent import is_mcp_intent, get_intent_info
            if is_mcp_intent(intent):
                cat, desc = get_intent_info(intent)
                param_txt = u""
                if params:
                    param_txt = u"\n**Parameters:** " + u", ".join(
                        u"{}={}".format(k, v) for k, v in params.items() if v is not None
                    )
                reply = (
                    message + u"\n\n"
                    if message else u""
                ) + (
                    u"**Revit API:** `{}`\n"
                    u"**Action:** {}{}\n\n"
                    u"*Requires the [revit-mcp](https://github.com/mcp-servers-for-revit/revit-mcp) "
                    u"server running alongside Revit.*"
                ).format(intent, desc, param_txt)
                _learn(reply)
                _bot(reply)
                self._set_busy(False)
                return
        except Exception as _mcp_ex:
            logger.debug("MCP intent handler error: {}".format(_mcp_ex))

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

    @staticmethod
    def _clean_bot_response(text):
        """Strip chain-of-thought, meta-commentary, and excessive whitespace."""
        import re as _re
        # Remove <think>...</think> blocks (reasoning models)
        text = _re.sub(r'<think>[\s\S]*?</think>', '', text, flags=_re.IGNORECASE)
        # Remove lines that read like internal planning/meta-commentary
        _skip_prefixes = (
            u"* user", u"* role", u"* tone", u"* language", u"* option",
            u"* the user", u"* since i am", u"* as an ai",
            u"- user", u"- role", u"- tone", u"- option",
            u"i am an ai", u"as an ai assistant",
            u"let me analyze", u"let me think", u"i need to consider",
            u"i'll analyze", u"i'll consider",
        )
        lines_out = []
        for line in text.splitlines():
            low = line.strip().lower()
            if any(low.startswith(p) for p in _skip_prefixes):
                continue
            lines_out.append(line)
        # Collapse 3+ consecutive blank lines to 1
        text = _re.sub(r'\n{3,}', u'\n\n', u'\n'.join(lines_out))
        return text.strip()

    @staticmethod
    def _build_md_inlines(text_block, text):
        """
        Populate text_block.Inlines with basic markdown rendering.
        Handles: **bold**, bullet lines (* / -), plain text, line breaks.
        """
        import re as _re
        from System.Windows.Documents import Run, LineBreak
        from System.Windows import FontWeights

        lines = text.splitlines()
        for i, line in enumerate(lines):
            if i > 0:
                text_block.Inlines.Add(LineBreak())

            # Bullet lines
            stripped = line.strip()
            if stripped.startswith(u'* ') or stripped.startswith(u'- '):
                prefix_run = Run()
                prefix_run.Text = u'• '   # bullet •
                text_block.Inlines.Add(prefix_run)
                line = stripped[2:]             # remaining text after bullet marker

            # Inline **bold** spans
            parts = _re.split(r'\*\*', line)
            for idx, part in enumerate(parts):
                if not part:
                    continue
                r = Run()
                r.Text = part
                if idx % 2 == 1:               # inside ** pair = bold
                    r.FontWeight = FontWeights.SemiBold
                text_block.Inlines.Add(r)

    def _append_bot_message(self, text):
        """Add a left-aligned bot bubble with avatar and basic markdown rendering."""
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

            # Bubble
            bubble = Border()
            bubble.Background      = SolidColorBrush(Color.FromRgb(255, 255, 255))
            bubble.CornerRadius    = CornerRadius(3, 8, 8, 8)
            bubble.Padding         = Thickness(14, 10, 14, 10)
            bubble.BorderBrush     = SolidColorBrush(Color.FromRgb(189, 195, 199))
            bubble.BorderThickness = Thickness(1)

            msg_text = TextBlock()
            msg_text.FontSize     = 13
            msg_text.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            msg_text.Foreground   = SolidColorBrush(Color.FromRgb(39, 39, 42))   # #27272A
            msg_text.TextWrapping = TextWrapping.Wrap
            msg_text.LineHeight   = 20

            try:
                self._build_md_inlines(msg_text, text)
            except Exception:
                # Fallback: plain text if inline building fails
                msg_text.Text = text

            bubble.Child = msg_text
            Grid.SetColumn(bubble, 1)
            row.Children.Add(bubble)
            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("Error adding bot message: {}".format(ex))

    # ─── Live streaming bubble ────────────────────────────────────────────────

    def _begin_stream_bubble(self):
        """Create an empty bot bubble that will be filled token-by-token.

        Stores references in self._stream_row / self._stream_tb. UI thread only.
        """
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

            av = self._make_avatar("T3")
            Grid.SetColumn(av, 0)
            row.Children.Add(av)

            bubble = Border()
            bubble.Background      = SolidColorBrush(Color.FromRgb(255, 255, 255))
            bubble.CornerRadius    = CornerRadius(3, 8, 8, 8)
            bubble.Padding         = Thickness(14, 10, 14, 10)
            bubble.BorderBrush     = SolidColorBrush(Color.FromRgb(189, 195, 199))
            bubble.BorderThickness = Thickness(1)

            tb = TextBlock()
            tb.FontSize     = 13
            tb.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            tb.Foreground   = SolidColorBrush(Color.FromRgb(39, 39, 42))
            tb.TextWrapping = TextWrapping.Wrap
            tb.LineHeight   = 20

            bubble.Child = tb
            Grid.SetColumn(bubble, 1)
            row.Children.Add(bubble)
            self.chat_history_panel.Children.Add(row)
            self._stream_row = row
            self._stream_tb  = tb
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("_begin_stream_bubble error: {}".format(ex))
            self._stream_row = None
            self._stream_tb  = None

    def _finalize_stream_bubble(self, text):
        """Re-render the live bubble with full markdown once streaming is done."""
        try:
            tb = self._stream_tb
            if tb is None:
                return
            tb.Inlines.Clear()
            try:
                self._build_md_inlines(tb, text)
            except Exception:
                tb.Text = text
            self._scroll_to_bottom()
        except Exception:
            pass

    def _remove_stream_bubble(self):
        """Discard the live bubble (used when an action renders its own reply)."""
        try:
            if self._stream_row is not None:
                self.chat_history_panel.Children.Remove(self._stream_row)
        except Exception:
            pass
        self._stream_row = None
        self._stream_tb  = None

    def _clear_stream_refs(self):
        """Detach references so the bubble persists but is no longer 'live'."""
        self._stream_row = None
        self._stream_tb  = None

    def _stream_llm_turn(self, provider, router, history, system_prompt,
                         query, max_tokens=1200, **kwargs):
        """Run one streaming LLM turn, filling a live bubble. Worker thread.

        The bubble is created lazily on the first token so the animated typing
        indicator stays visible until the model actually starts replying.
        Returns the full raw response text (str) or None.
        """
        extractor = StreamingJSONExtractor()
        state = {"raw": [], "started": False}

        def _on_delta(chunk):
            if not chunk:
                return
            state["raw"].append(chunk)
            disp = extractor.display(u"".join(state["raw"]))

            def _ui():
                if not state["started"]:
                    state["started"] = True
                    self._hide_typing_indicator()
                    self._begin_stream_bubble()
                try:
                    if self._stream_tb is not None:
                        # Live text is plain (fast); markdown applied on finalize.
                        self._stream_tb.Text = disp
                        self._scroll_to_bottom()
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        try:
            if provider is not None and provider.check_health():
                return provider.chat_stream(
                    history[-16:], system_prompt, query, _on_delta, max_tokens)
            return router.chat_stream(
                history[-16:], system_prompt, query, _on_delta, max_tokens)
        except Exception as ex:
            logger.debug("_stream_llm_turn error: {}".format(ex))
            return None

    def _scroll_to_bottom(self):
        try:
            self.chat_scroll.ScrollToBottom()
        except Exception:
            pass


# MAIN SCRIPT
# ==================================================

if __name__ == '__main__':
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)

    try:
        from Autodesk.Revit.UI import DockablePaneId
        from System import Guid
        from GUI.AssistantPaneControl import ASSISTANT_PANE_GUID
        from pyrevit import HOST_APP

        pane_id = DockablePaneId(ASSISTANT_PANE_GUID)
        uiapp = HOST_APP.uiapp
        pane = uiapp.GetDockablePane(pane_id)
        if pane:
            if pane.IsShown():
                pane.Hide()
            else:
                pane.Show()
    except Exception as ex:
        logger.warning("Could not toggle dockable pane: {}. Falling back to floating window.".format(ex))
        window = T3LabAssistantWindow()
        window.ShowDialog()
