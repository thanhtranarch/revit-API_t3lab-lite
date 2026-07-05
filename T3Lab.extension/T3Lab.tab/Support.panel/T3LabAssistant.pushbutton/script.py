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
import io
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
                                              get_setup_guidance_message,
                                              _build_system_prompt,
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
    def get_setup_guidance_message(viet=True):
        return (u"Chưa hiểu yêu cầu — bạn mô tả cụ thể hơn nhé." if viet
                else "I didn't understand — could you describe it more specifically?")
    def _build_system_prompt(revit_context=u""):
        from Intelligence.t3lab_agent import build_system_prompt
        return build_system_prompt(revit_context=revit_context)
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
        # Serialize to an ASCII string FIRST, then write in one shot.
        # json.dump(..., ensure_ascii=False) on a bytes-mode file blows up
        # with UnicodeEncodeError mid-write under IronPython 2.7 as soon as
        # the chat carries Vietnamese — truncating the history file to 0
        # bytes, silently (same failure class as the tool_registry.json bug).
        data = json.dumps({"doc_key": doc_key, "messages": to_save},
                          ensure_ascii=True, indent=2)
        if isinstance(data, bytes):
            data = data.decode('ascii')
        with io.open(path, 'w', encoding='utf-8') as f:
            f.write(data)
    except Exception as ex:
        logger.debug("Could not save chat history: {}".format(ex))


def load_chat_history(doc_key):
    """Load saved messages for doc_key.  Returns [] if none / error."""
    try:
        path = _history_file(doc_key)
        if not os.path.exists(path):
            return []
        with io.open(path, 'r', encoding='utf-8') as f:
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


# ─── Minimal icon vocabulary for bot chat messages ─────────────────────────────
# Segoe MDL2 Assets glyphs already verified elsewhere in this codebase (window
# chrome buttons, command palette, tool-card status) — reused here rather than
# guessing new codepoints, per the Lumina rule that chat icons must be minimal
# monochrome glyphs, never full-color detailed emoji (🔍🤖⚠️📎🔧👋🎉 etc.).
_ICON_INFO    = u""   # Info — neutral notices ("no AI configured", RAG note)
_ICON_SEARCH  = u""   # Zoom — tool discovery
_ICON_WARNING = u""   # Warning triangle
_ICON_SUCCESS = u""   # CheckMark
_ICON_SYNC    = u""   # Sync — "in progress" (matches tool-card running glyph)
_ICON_STOP    = u""   # Stop
_ICON_REFRESH = u""   # Refresh (chat cleared)
_ICON_ATTACH  = u""   # Attach (paperclip)
_ICON_ANALYZE = u""   # Analyze — fast-context "instant DB answer" badge
_ICON_LIST    = u""   # List/reference — stats & selection section headers

_ICON_BLUE  = (59, 130, 246)     # #3B82F6 accent — info/discovery
_ICON_AMBER = (245, 158, 11)     # #F59E0B — warning
_ICON_GREEN = (16, 185, 129)     # #10B981 — success
_ICON_SLATE = (100, 116, 139)    # #64748B — neutral/muted


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
        self._switching_provider = False        # guard: _switch_provider bg probe in flight
        self._probing_sidebar    = False        # guard: settings sidebar bg probe in flight
        self._typing_row       = None           # reference to typing indicator element
        self._conversation_history = []         # [{role, content}, ...] multi-turn context
        self._last_raw         = ''             # last user input (for learning)
        self._doc_key          = _get_doc_key() # document identifier for history
        self._persisted_msgs   = []             # flat list with timestamps, for save/load
        self._attached_files   = []             # list of file paths (images / PDFs)
        # {provider_name: [model_list]} — filled ONLY by live probes against
        # each vendor's models endpoint after a connection succeeds. Never
        # pre-seeded with hardcoded names: the MODEL combo stays empty and
        # disabled (with a hint) until the provider actually reports which
        # models the account/server currently has.
        self._models_cache     = {}
        
        # ── History & Typing Animation State ──────────────────────────────────
        self._input_history    = []
        self._history_index    = -1
        self._current_input_temp = ""
        self._typing_timer     = None
        self._typing_elapsed   = 0

        # ── Live streaming bubble state ───────────────────────────────────────
        self._stream_row       = None           # Grid row of the live reply bubble
        self._stream_tb        = None           # TextBlock being filled token-by-token

        # ── Native agent loop state ───────────────────────────────────────────
        self._agent_loop        = None    # running AgentLoop (native tools path)
        self._cancel_requested  = False   # Stop pressed before the loop existed
        self._send_orig_content = None    # cached "Gửi" content of send_button

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

        # ── Register the MCP ExternalEvent on the UI thread ────────────────────
        # This MUST happen here, on Revit's main thread — ExternalEvent.Create
        # throws if called from the background startup probe below, which would
        # leave every model-editing MCP tool (create wall/floor/column, rename,
        # set_parameter, …) unable to open a transaction. Creating it up-front
        # here guarantees the server can marshal those tools onto the UI thread.
        try:
            from Services.mcp_service import MCPService as _MCPService
            _ok, _ee_err = _MCPService.ensure_external_event()
            if not _ok:
                logger.debug("MCP ExternalEvent init failed: {}".format(_ee_err))
        except Exception as _ex:
            logger.debug("MCP ExternalEvent init error: {}".format(_ex))

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
                                u"Không tìm thấy mô hình AI cục bộ nào. Đang tự động tải mô hình mặc định (qwen2.5:1.5b) về máy bạn. Quá trình này chạy ngầm và có thể mất vài phút...",
                                icon=_ICON_SYNC, icon_color=_ICON_SLATE
                            )))

                            payload = {"name": "qwen2.5:1.5b", "stream": False}
                            local_llm._post_json(local_llm.OLLAMA_HOST + "/api/pull", payload, timeout=600)

                            self.Dispatcher.Invoke(Action(lambda: self._append_bot_message(
                                u"Tải thành công mô hình AI qwen2.5:1.5b! Bạn có thể sử dụng T3Lab Assistant ngoại tuyến.",
                                icon=_ICON_SUCCESS, icon_color=_ICON_GREEN
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

                # Step 5: Proactive setup nudge — only on a fresh chat (no saved
                # history for this document yet), only after first-run onboarding
                # has already been shown/dismissed (avoids duplicating that flow),
                # and only if auto-start above didn't already find a provider.
                try:
                    from config.user_profile import UserProfile
                    already_onboarded = not UserProfile().is_first_run()
                    fresh_chat = not self._persisted_msgs
                    no_provider = not has_api_key() and not has_local_llm()
                    if already_onboarded and fresh_chat and no_provider:
                        def _nudge():
                            self._append_bot_message(get_setup_guidance_message(True),
                                                     icon=_ICON_INFO, icon_color=_ICON_SLATE)
                            self._add_to_history("assistant", get_setup_guidance_message(True))
                        self.Dispatcher.Invoke(Action(_nudge))
                except Exception:
                    pass
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

        # Persist window geometry/sidebar on close (custom X button was removed)
        try:
            self.Closing += self._on_closing
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

    def _on_closing(self, sender, e):
        """Persist window geometry/sidebar state when the window closes.

        Wired to the WPF Closing event in __init__. Previously this logic lived
        in close_clicked (the custom X button), which was removed — this keeps
        window-state saving alive for Alt+F4 / Revit pane close.
        """
        self._save_window_state()

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
                    u"Phát hiện {} công cụ mới: {}.\n"
                    u"Tôi đã tự học và có thể mở chúng bằng lệnh tự nhiên.".format(
                        len(new_tools), names),
                    icon=_ICON_SEARCH, icon_color=_ICON_BLUE
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
                u"Cuộc trò chuyện đã được làm mới!\n"
                u"Tôi có thể giúp gì cho bạn?",
                icon=_ICON_REFRESH, icon_color=_ICON_SLATE
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
                self.onboarding_greeting.Text = u"{}!".format(greet)
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
                u"Rất vui được gặp bạn, {}!\n"
                u"Hồ sơ của bạn đã được lưu. Hãy thử 'mở batchout' hoặc hỏi tôi bất cứ điều gì về Revit.".format(name),
                icon=_ICON_SUCCESS, icon_color=_ICON_GREEN)
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
            # Instant, zero-HTTP snapshot — the old get_status() call here
            # probed every provider synchronously on the UI thread whenever
            # the cache was cold, freezing the window for seconds on click.
            status = router.get_status_instant()

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
                # Switching is always allowed — the sidebar guides setup if the
                # provider isn't configured yet (disabling based on a possibly
                # stale snapshot locked users out of providers that were fine).
                item.IsEnabled = not is_active

                def _make_handler(n):
                    def _handler(s, e):
                        self._switch_provider(n)
                    return _handler

                item.Click += _make_handler(name)
                menu.Items.Add(item)

            # Warm the real status cache in the background for the NEXT open.
            def _warm():
                try:
                    LLMRouter().get_status(use_cache=True)
                except Exception:
                    pass
            _wt = Thread(ThreadStart(_warm))
            _wt.IsBackground = True
            _wt.Start()

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

            # Background: probe the newly-active provider + refresh its model list.
            # Guarded so rapid repeated provider switches don't pile up threads
            # all probing at once — a switch while one is already probing just
            # skips spawning a second thread (the in-flight one already covers
            # the freshest switch target read at its start).
            if self._switching_provider:
                return
            self._switching_provider = True

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
                finally:
                    self._switching_provider = False

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
        # then the remaining providers for the full status list. Guarded so
        # rapidly toggling the sidebar open/closed doesn't spawn a new probe
        # thread on top of one that's already running.
        if self._probing_sidebar:
            return
        self._probing_sidebar = True

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
            finally:
                self._probing_sidebar = False

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
                # Pre-select the saved model (stored per-provider) ONLY if the
                # vendor still reports it in the live list — never inject a
                # name the vendor didn't confirm (it may be retired/renamed).
                saved = None
                try:
                    from config.settings import T3LabAISettings
                    saved = T3LabAISettings().get_provider_model(active)
                except Exception:
                    pass
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
            # NEVER get_status() here — this runs on the UI thread (Dispatcher),
            # and when the 30s cache TTL has lapsed (probe_provider merges data
            # but doesn't refresh the TTL) get_status(use_cache=True) silently
            # re-probes EVERY provider synchronously: multi-second UI freeze on
            # every sidebar open / provider switch. Cached-or-cheap only.
            status = router.get_status_instant()
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
        """Lock/unlock the input area. Call from UI thread only.

        While busy the send button STAYS ENABLED and becomes a Stop button so
        the user can cancel a running agent request mid-flight.
        """
        self._busy = busy
        if busy:
            self._cancel_requested = False
        try:
            self.chat_input.IsEnabled = not busy
            self._render_send_button(busy)
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

    def _render_send_button(self, busy):
        """Swap the send button between 'Gửi ➢' and 'Dừng ⏹'. UI thread only."""
        try:
            from System.Windows.Controls import StackPanel, TextBlock, Orientation
            from System.Windows.Media import SolidColorBrush, Color, FontFamily
            from System.Windows import Thickness, VerticalAlignment

            btn = self.send_button
            if self._send_orig_content is None:
                self._send_orig_content = btn.Content

            if not busy:
                btn.Content   = self._send_orig_content
                btn.IsEnabled = True
                btn.ToolTip   = u"Gửi (Enter)"
                return

            _white = SolidColorBrush(Color.FromRgb(255, 255, 255))
            sp = StackPanel()
            sp.Orientation = Orientation.Horizontal

            t1 = TextBlock()
            t1.Text              = u"Dừng"
            t1.FontSize          = 13
            t1.FontWeight        = System.Windows.FontWeights.SemiBold
            t1.Foreground        = _white
            t1.VerticalAlignment = VerticalAlignment.Center

            t2 = TextBlock()
            t2.Text              = u""   # MDL2 Stop
            t2.FontFamily        = FontFamily(u"Segoe MDL2 Assets")
            t2.FontSize          = 11
            t2.Foreground        = _white
            t2.VerticalAlignment = VerticalAlignment.Center
            t2.Margin            = Thickness(4, 0, 0, 0)

            sp.Children.Add(t1)
            sp.Children.Add(t2)
            btn.Content   = sp
            btn.IsEnabled = True
            btn.ToolTip   = u"Dừng tác vụ đang chạy"
        except Exception as ex:
            logger.debug("_render_send_button error: {}".format(ex))

    def _request_stop(self):
        """Stop button pressed while a request is running (UI thread)."""
        try:
            self._cancel_requested = True
            loop = self._agent_loop
            if loop is not None:
                loop.cancel()
            try:
                self.send_button.IsEnabled = False   # re-enabled by _set_busy(False)
            except Exception:
                pass
            self._safe_update_typing_text(u"● ● ●  Đang dừng sau bước hiện tại…")
        except Exception as ex:
            logger.debug("_request_stop error: {}".format(ex))

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


    def _safe_append_bot(self, msg, icon=None, icon_color=None):
        """Thread-safe bot message append (can be called from background threads)."""
        try:
            self.Dispatcher.Invoke(Action(
                lambda: self._append_bot_message(msg, icon=icon, icon_color=icon_color)))
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

            btn = Button()
            try:
                btn.Style = self.FindResource('AttachChipBtn')
            except Exception:
                pass
            btn.Margin = Thickness(0, 0, 4, 4)

            sp = StackPanel()
            sp.Orientation = System.Windows.Controls.Orientation.Horizontal

            # Minimal MDL2 Attach glyph (same one used on the attach button
            # itself) — replaces the old colored 🖼️/📄 emoji pair.
            icon_lbl = TextBlock()
            icon_lbl.Text = _ICON_ATTACH + u" "
            icon_lbl.FontFamily = System.Windows.Media.FontFamily(u"Segoe MDL2 Assets")
            icon_lbl.FontSize = 11
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
        if self._busy:
            # While busy the button is a Stop button. Effective on the native
            # agent path; the legacy path finishes its current step regardless.
            self._request_stop()
            return
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
                    u"Đang xử lý lệnh trước, vui lòng chờ một chút...",
                    icon=_ICON_SYNC, icon_color=_ICON_SLATE
                )
                return


            self.chat_input.Text = ""
            self._last_raw = raw or u"[đính kèm tài liệu]"

            # ── Show user message in chat ──────────────────────────────────────
            display_text     = raw
            attachment_note  = summarize_attachments(attached) if attached else None
            self._append_user_message(display_text, attachment_note=attachment_note)
            # History/LLM context still gets plain text — no raw icon glyph.
            history_text = (u"{}\n[đính kèm: {}]".format(display_text, attachment_note)
                            if attachment_note else display_text)
            self._add_to_history("user", history_text)

            # ── Clear attachments from UI after sending ────────────────────────
            if attached:
                self._attached_files = []
                try:
                    self.attachment_chips_panel.Children.Clear()
                    self._refresh_attachment_panel()
                except Exception:
                    pass

            # Lock UI + show the typing indicator IMMEDIATELY (set_busy does
            # both). Everything further down — PDF text extraction, the
            # Ollama HTTP probe, Revit-DB fast-answer collectors, NLU catalog
            # scoring — takes tens of ms to SECONDS. It used to run right
            # here on the UI thread, so every Enter press froze the window
            # until routing finished. It now runs on a routing worker thread;
            # only the results marshal back onto the dispatcher.
            self._set_busy(True)

            def _route():
                try:
                    self._route_input(raw, attached)
                except Exception as ex:
                    logger.error("_route_input error: {}".format(ex))

                    def _fail():
                        self._hide_typing_indicator()
                        self._set_busy(False)
                    try:
                        self.Dispatcher.Invoke(Action(_fail))
                    except Exception:
                        pass

            rt = Thread(ThreadStart(_route))
            rt.IsBackground = True
            rt.SetApartmentState(ApartmentState.STA)
            rt.Start()

        except Exception as ex:
            logger.error("Error in _process_input: {}".format(ex))
            self._set_busy(False)

    def _route_input(self, raw, attached):
        """Classify + dispatch one user request. WORKER THREAD.

        Runs the whole routing ladder (RAG context build → fast Revit-DB
        answer → learned patterns → offline NLU → LLM turn) off the UI
        thread, so pressing Enter stays instant no matter how heavy the
        request is. All UI feedback and _execute_result calls marshal back
        through the Dispatcher.
        """
        try:
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

            # For NLP routing we use ONLY the raw user text. Prepending the
            # ContextScout model summary here poisoned the offline NLU and
            # keyword scoring (its words — "sheet", "view", "wall"... — leak
            # into intent triggers) and duplicated context the LLM already
            # receives via the system prompt in do_nlp().
            captured = raw

            history  = list(self._conversation_history[:-1])

            use_local        = HAS_NLP and has_local_llm()
            use_claude       = HAS_NLP and has_api_key()   # True for any configured provider
            _active_provider = get_active_provider_name()  # "claude" | "openai" | "ollama"

            # ── 0. Fast context answer (DB-only, no LLM) ──────────────
            if HAS_SCOUT and not has_attach:
                fast = self._try_fast_context_answer(raw)
                if fast:
                    def _show_fast(_fast=fast):
                        self._hide_typing_indicator()
                        self._append_bot_message(_fast)
                        self._add_to_history("assistant", _fast)
                        self._set_busy(False)
                    self.Dispatcher.Invoke(Action(_show_fast))
                    return

            # ── 1. Learned patterns (skip if attachments present) ─────────────
            if HAS_NLP and not has_attach:
                learned = find_learned_match(raw)
                if learned:
                    def _run_learned(_r=learned):
                        self._execute_result(_r)
                    self.Dispatcher.Invoke(Action(_run_learned))
                    return

            # ── 2. Built-in NLU (skip for RAG / attachment queries) ───────────
            nlu_result = None
            if HAS_NLP and not has_attach:
                nlu_result = parse_command_nlu(captured, history)
                if nlu_result and nlu_result.get("intent") not in (None, "unknown"):
                    # _authoritative = answered from the real tool catalog
                    # (capability questions, ambiguity clarifications) — the
                    # LLM must not get a chance to override it with a guess.
                    if nlu_result["intent"] not in ("chat", "help") \
                            or nlu_result.get("_authoritative") \
                            or not (use_local or use_claude):
                        def _run_nlu(_r=nlu_result):
                            self._execute_result(_r)
                        self.Dispatcher.Invoke(Action(_run_nlu))
                        return

            if use_local or use_claude or has_attach:
                # ── 3/4. LLM path (typing indicator already showing) ──────────
                nlu_hint = nlu_result if (HAS_NLP and not has_attach) else None

                def do_nlp():
                    result = None
                    # Distinguishes "the model returned nothing at all" (timeout,
                    # connection error, empty body) from "the model answered but
                    # picked an unrecognised intent" — without this, both looked
                    # identical to the user: the same generic offline fallback
                    # text, no matter what they typed.
                    llm_call_failed = False
                    from Intelligence.llm_router import LLMRouter
                    import json as _json

                    _router = LLMRouter()
                    _provider = _router.get_active_provider()

                    # ── Native function-calling agent path ─────────────────────
                    # Providers with SUPPORTS_NATIVE_TOOLS run the real agentic
                    # loop: tool schemas travel through the API `tools` param,
                    # replies are plain text (no JSON-in-prompt). If the path
                    # can't even start (registry unavailable, provider mute on
                    # turn 1 — e.g. an Ollama model without tool support), it
                    # returns False and the legacy JSON-intent loop below runs.
                    if not has_attach and not rag_context:
                        _handled = False
                        try:
                            _handled = self._run_native_agent(
                                _provider, list(history), captured)
                        except Exception as _na_ex:
                            logger.debug("native agent path error: {}".format(_na_ex))
                        if _handled:
                            return

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

                        # Retrieve registered tools from server.py — compact
                        # catalog (C4): names + one-line purpose only. The full
                        # JSON schemas of all ~75 tools used to be inlined here
                        # on EVERY iteration, which crushed small local models
                        # (token bloat + broken JSON). The model now requests a
                        # schema on demand via the describe_tool meta-intent.
                        server_tools_str = u""
                        tools_list = []
                        try:
                            from core.server import get_t3labai_server
                            srv = get_t3labai_server()
                            tools_list = srv._handle_tools_list().get('tools', [])
                            if tools_list:
                                server_tools_str = u"\n\nLocal MCP Server Tools (name: purpose):\n"
                                for tool in tools_list:
                                    desc = (tool.get('description') or u'').strip()
                                    desc = desc.splitlines()[0][:110] if desc else tool['name']
                                    server_tools_str += u"- `{}`: {}\n".format(
                                        tool['name'], desc)
                                server_tools_str += (
                                    u"\nTool calls need correct parameters. If you are "
                                    u"unsure of a tool's parameters, FIRST reply exactly "
                                    u'{"intent": "describe_tool", "params": {"name": "<tool_name>"}} '
                                    u"to receive its full JSON schema, then call the tool "
                                    u"in your next reply.\n")
                        except Exception as tool_err:
                            logger.debug("Failed to list server tools: {}".format(tool_err))

                        system_prompt = _build_system_prompt(revit_context=_ctx_block)
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
                            llm_call_failed = True
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

                            # C4 meta-intent: the model asks for one tool's full
                            # schema (the catalog above only carries one-liners).
                            # Feed the schema back and let it produce the real
                            # tool call on the next iteration.
                            if intent == "describe_tool":
                                self.Dispatcher.Invoke(Action(self._remove_stream_bubble))
                                tname = u"{}".format(
                                    (params or {}).get("name")
                                    or (params or {}).get("tool") or u"").strip()
                                match = None
                                for _t in (tools_list or []):
                                    if _t.get('name') == tname:
                                        match = _t
                                        break
                                if match:
                                    info = _json.dumps(
                                        {"name": match['name'],
                                         "description": match.get('description', ''),
                                         "inputSchema": match.get('inputSchema', {})},
                                        ensure_ascii=False)
                                else:
                                    info = _json.dumps(
                                        {"error": "unknown tool", "name": tname},
                                        ensure_ascii=False)
                                current_history.append(
                                    {"role": "assistant",
                                     "content": _json.dumps(_parsed, ensure_ascii=False)})
                                current_history.append(
                                    {"role": "user",
                                     "content": u"Tool schema: {}".format(info)})
                                current_query = (
                                    u"Schema: {}. Now respond with the actual tool call as "
                                    u'{{"intent": "<tool_name>", "params": {{...}}, '
                                    u'"message": "..."}}.'.format(info))
                                continue

                            # Determine if it's a local MCP tool call — check
                            # against the real registry, not a "revit_" prefix
                            # guess, since several real tool names don't start
                            # with "revit_" (place_wall, create_grid, etc.) and
                            # a hallucinated "revit_*" name that ISN'T
                            # registered would otherwise be sent into the
                            # External Event round-trip for nothing.
                            is_local_tool = False
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
                                tool_display_msg = u"[Tool Call] `{}` (params: {})".format(intent, _json.dumps(params))
                                if message:
                                    tool_display_msg = u"[Tool Call] {}\nRevit tool: `{}`".format(message, intent)
                                self._safe_append_bot(tool_display_msg,
                                                      icon=_ICON_SYNC, icon_color=_ICON_BLUE)

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
                                # A model that picks "unknown" but still wrote a
                                # message (small/local models do this a lot —
                                # they don't map cleanly to one of the listed
                                # intents but still try to answer) should have
                                # that answer shown, not silently swapped out
                                # for the generic offline fallback in finish().
                                if intent == "unknown" and message.strip():
                                    _parsed["intent"] = "chat"
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
                            # A generic "I didn't understand" guess from the offline
                            # NLU is only worth showing when the LLM call actually
                            # completed and ALSO had nothing better — if the LLM
                            # never responded at all, that guess would silently
                            # masquerade as "the assistant tried and failed to
                            # match your request", which is misleading; the
                            # llm_call_failed branch below gives the real reason.
                            nlu_hint_usable = (
                                nlu_hint and nlu_hint.get("intent") not in (None, "unknown")
                                and not (llm_call_failed and nlu_hint.get("_generic_fallback"))
                            )
                            if result and result.get("intent") not in (None, "unknown"):
                                self._execute_result(result)
                            elif nlu_hint_usable:
                                self._execute_result(nlu_hint)
                            else:
                                if has_attach and not use_claude and not use_local:
                                    if rag_context:
                                        self._append_bot_message(
                                            u"Nội dung tài liệu:\n\n" + rag_context[:2000],
                                            icon=_ICON_ATTACH, icon_color=_ICON_SLATE
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
                                elif llm_call_failed:
                                    # The AI genuinely never answered (timeout /
                                    # connection error / API rejection) — say so
                                    # plainly, WITH the provider's real reason
                                    # when it was recorded (get_last_error).
                                    label = get_provider_display_label()
                                    detail = u""
                                    try:
                                        _le = (_provider.get_last_error()
                                               if _provider else None)
                                        if _le:
                                            detail = (u"\nChi tiết: {}".format(_le)
                                                      if _is_viet_text(captured)
                                                      else u"\nDetail: {}".format(_le))
                                    except Exception:
                                        pass
                                    msg = (u"Model AI ({}) không phản hồi (model quá nặng, mất "
                                           u"kết nối hoặc API báo lỗi). Thử lại hoặc chọn model "
                                           u"khác trong Cài đặt.{}".format(label, detail)
                                           if _is_viet_text(captured) else
                                           u"The AI model ({}) didn't respond (too heavy, "
                                           u"disconnected, or the API returned an error). Try again "
                                           u"or pick another model in Settings.{}".format(label, detail))
                                    self._append_bot_message(msg, icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                                    self._set_busy(False)
                                else:
                                    msg = (u"Mình chưa hiểu yêu cầu này — bạn mô tả cụ thể hơn nhé."
                                           if _is_viet_text(captured) else
                                           "I didn't understand this request — could you describe it more specifically?")
                                    self._append_bot_message(msg)
                                    self._set_busy(False)
                        except Exception as finish_ex:
                            logger.error("finish error: {}".format(finish_ex))
                            self._hide_typing_indicator()
                            self._clear_stream_refs()
                            self._set_busy(False)

                    self.Dispatcher.Invoke(Action(finish))

                # Already on the routing worker thread — run the LLM turn
                # inline instead of spawning yet another thread.
                do_nlp()
            else:
                # ── 5. No provider configured at all — keyword fallback ─────────
                fb = keyword_parse(raw)

                def _finish_offline(_fb=fb):
                    if _fb:
                        self._execute_result(_fb)
                    else:
                        self._append_bot_message(
                            get_setup_guidance_message(_is_viet_text(raw)),
                            icon=_ICON_INFO, icon_color=_ICON_SLATE)
                        self._set_busy(False)
                self.Dispatcher.Invoke(Action(_finish_offline))
        except Exception:
            # Bubble to _route()'s handler in _process_input — it hides the
            # typing indicator and releases the busy lock on the UI thread.
            raise

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

        # ── Conversational-input guard (last line of defence) ─────────────────
        # Pure small talk ("morning", "thanks", "ok"...) must never launch a
        # tool, no matter which layer produced the intent — a mis-learned
        # pattern or an LLM hallucination could map it to open_*/export.
        if intent not in ("help", "chat", "greet", "unknown") and HAS_NLP:
            try:
                from Intelligence.nlu_engine import is_conversational
                _is_smalltalk = is_conversational(raw)
            except Exception:
                _is_smalltalk = False
            if _is_smalltalk:
                conv = parse_command_nlu(raw) or {}
                if conv.get("intent") in ("greet", "chat", "help") and conv.get("message"):
                    reply = conv["message"]
                elif _is_viet_text(raw):
                    reply = u"Xin chào! Tôi là T3Lab Assistant.\nBạn muốn làm gì hôm nay?"
                else:
                    reply = u"Hello! I'm T3Lab Assistant.\nWhat would you like to do today?"
                _bot(reply)
                self._set_busy(False)
                return

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

        # ── Recover hallucinated open_* intents ──────────────────────────────
        # The LLM sometimes invents a near-miss intent name ("open_mcp_control"
        # instead of "open_mcpcontrol"). Before declaring the tool missing,
        # resolve the user's own words against the full tool catalog and
        # launch the tool only if one clearly wins.
        if intent.startswith("open_") and HAS_NLP:
            _match = None
            try:
                from Intelligence.nlu_engine import resolve_tool
                _match, _cands = resolve_tool(raw)
            except Exception:
                _match = None
            if _match and _match['intent'] in TOOL_LAUNCHERS:
                label = _match.get('title', _match['intent'])
                confirm = (u"Đang mở {}...".format(label) if _is_viet_text(raw)
                           else u"Opening {}...".format(label))
                _bot(confirm)
                learn_pattern(raw, _match['intent'], {}, confirm)
                ok = TOOL_LAUNCHERS[_match['intent']]()
                if not ok:
                    self._append_bot_message(u"Không thể mở công cụ. Xem console.")
                self._set_busy(False)
                return

        # ── Unknown / fallthrough ─────────────────────────────────────────────
        # Reaching here with a non-empty, non-"unknown" intent means the model
        # returned a tool name that isn't registered anywhere (T3Lab UI tool
        # or MCP tool) — say so plainly instead of the misleading "Đã thực
        # hiện." ("Done."), since nothing was actually executed.
        if intent == "unknown":
            _bot(params.get("message", u"Yêu cầu chưa rõ — bạn mô tả cụ thể hơn nhé."))
        elif message:
            _bot(message)
        else:
            _bot(u"Công cụ `{}` không tồn tại — kiểm tra lại tên hoặc mô tả việc cần làm.".format(intent))
        self._set_busy(False)

    def _run_tool(self, intent, default_msg):
        """Helper for quick-button clicks: guard, show message, run launcher."""
        if self._busy:
            self._append_bot_message(u"Đang xử lý lệnh trước, vui lòng chờ...",
                                     icon=_ICON_SYNC, icon_color=_ICON_SLATE)
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

    # ─── Native agentic loop (function calling) ────────────────────────────────

    def _run_native_agent(self, provider, history, captured):
        """Run the native tool-calling agent loop. WORKER THREAD.

        Returns True when the request was fully handled (UI updated, busy
        released). Returns False so the legacy JSON-intent path can run —
        only when nothing was shown to the user yet.
        """
        if provider is None or not getattr(provider, "SUPPORTS_NATIVE_TOOLS", False):
            return False

        from Intelligence.agent_loop import AgentLoop, build_agent_system_prompt
        from Intelligence import tool_schema

        try:
            from core.server import get_t3labai_server
            srv = get_t3labai_server()
        except Exception:
            return False

        launcher = tool_schema.make_launcher_tool(list(TOOL_LAUNCHERS.keys()))
        tools = tool_schema.get_tools_for_provider(provider.NAME, [launcher])
        if len(tools) <= 1:          # only the launcher → MCP registry unavailable
            return False

        ctx = u""
        if HAS_SCOUT:
            try:
                ctx = ContextScout.get_context_summary_for_ai()
            except Exception:
                pass
        system_prompt = build_agent_system_prompt(ctx)

        viet = _is_viet_text(captured)

        # ── C2: vision — "look at this view" ships a snapshot of the active
        # view as an image block (Claude agent path only; other providers'
        # agent calls don't convert Claude-format blocks).
        user_content = captured
        if self._wants_view_snapshot(captured, provider):
            self._safe_update_typing_text(
                u"● ● ●  Đang chụp active view…" if viet
                else u"● ● ●  Capturing the active view…")
            shot = self._capture_active_view(srv)
            if shot:
                try:
                    from Intelligence.rag_processor import build_vision_content_blocks
                    user_content = build_vision_content_blocks(captured, [shot])
                except Exception:
                    user_content = captured

        # ── B4 + B5: tool-execution wrapper ───────────────────────────────────
        # B4: the first model-mutating tool opens ONE TransactionGroup on the
        #     Revit main thread (via the __begin_action_group pseudo-tool), so
        #     the whole request assimilates into a single Undo entry.
        # B5: destructive tools block on an in-chat Confirm/Cancel card; the
        #     first purge_unused of a request is always forced to dry_run.
        req_title = u" ".join((captured or u"AI request").split())[:60]
        group = {"open": False}
        purge = {"first_done": False}
        # Write tools that must NOT trigger the request group: they don't
        # change the model (selection / export / UI) or run arbitrary code.
        _group_exempt = frozenset((
            'say_hello', 'show_assistant_pane', 'set_active_view',
            'select_elements', 'export_sheets_pdf', 'export_dwg',
            'export_image', 'send_code_to_revit',
            '__begin_action_group', '__end_action_group',
        ))

        def _exec_tool(name, args):
            args = dict(args or {})
            if name == 'purge_unused' and not purge["first_done"]:
                purge["first_done"] = True
                if not bool(args.get('dry_run', True)):
                    args['dry_run'] = True   # first pass is ALWAYS a report
            destructive = (
                name == 'delete_element'
                or (name == 'purge_unused' and not bool(args.get('dry_run', True))))
            if destructive and not self._confirm_tool_blocking(name, args, viet):
                return {"cancelled": True,
                        "note": "User declined the '{}' action.".format(name)}
            if (not group["open"] and name in srv._WRITE_TOOLS
                    and name not in _group_exempt):
                try:
                    res = srv._execute_tool(
                        "__begin_action_group",
                        {"title": u"T3Lab AI: " + req_title})
                    group["open"] = isinstance(res, dict) and bool(res.get("success"))
                except Exception:
                    group["open"] = False
            return srv._execute_tool(name, args)

        # ── D1: cancel the loop when the user switches documents mid-request ──
        doc_key0 = _get_doc_key()

        def _guard_check():
            try:
                cur = _get_doc_key()
                # "default" = the read failed (API busy / no context) — that is
                # UNKNOWN, not "changed"; only trip on a positive mismatch.
                return cur != "default" and cur != doc_key0
            except Exception:
                return False

        import time as _time
        stream = {"text": u"", "last": 0.0, "open": False}
        card   = {"cur": None}

        # ── Throttled live-stream rendering ───────────────────────────────────
        # Deltas are batched: at most ~25 UI updates/second, pushed with
        # BeginInvoke so the worker never blocks on the dispatcher. Only
        # state-transition callbacks below use the synchronous Invoke.
        def _push_stream(force):
            now = _time.time()
            if not force and (now - stream["last"]) < 0.04:
                return
            stream["last"] = now
            snap = stream["text"]

            def _ui():
                try:
                    if not stream["open"]:
                        stream["open"] = True
                        self._hide_typing_indicator()
                        self._begin_stream_bubble()
                    if self._stream_tb is not None:
                        self._stream_tb.Text = snap
                        self._scroll_to_bottom()
                except Exception:
                    pass

            try:
                self.Dispatcher.BeginInvoke(Action(_ui))
            except Exception:
                pass

        def on_text_delta(chunk):
            stream["text"] += chunk
            _push_stream(False)

        def on_turn_text(text, is_final):
            stream["text"] = u""

            def _ui():
                try:
                    if stream["open"]:
                        stream["open"] = False
                        self._finalize_stream_bubble(text)
                        self._clear_stream_refs()
                    else:
                        self._hide_typing_indicator()
                        self._append_bot_message(text)
                    self._add_to_history("assistant", text)
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        def on_tool_start(name, args, iteration):
            def _ui():
                try:
                    self._hide_typing_indicator()
                    card["cur"] = self._append_tool_card(name, args)
                    self._show_typing_indicator()
                    if getattr(self, "_typing_text_block", None) is not None:
                        self._typing_text_block.Text = (
                            u"● ● ●  Đang chạy `{}`…".format(name))
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        def on_tool_done(name, result, ok, seconds):
            def _ui():
                try:
                    self._update_tool_card(card["cur"], ok, seconds, result)
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        loop = AgentLoop(
            provider, _exec_tool, tools,
            callbacks={
                "on_text_delta": on_text_delta,
                "on_turn_text":  on_turn_text,
                "on_tool_start": on_tool_start,
                "on_tool_done":  on_tool_done,
                "guard_check":   _guard_check,
            },
            max_iterations=10, max_tokens=1500)

        self._agent_loop = loop
        if self._cancel_requested:
            loop.cancel()
        try:
            result = loop.run(history, system_prompt, user_content)
        finally:
            self._agent_loop = None
            # B4: always close the request group — a group left open would
            # block every subsequent transaction in the session.
            if group["open"]:
                group["open"] = False
                try:
                    srv._execute_tool("__end_action_group", {})
                except Exception:
                    pass

        # Provider never answered turn 1 and nothing reached the UI →
        # hand back to the legacy path (it has its own fallbacks).
        if (result.get("status") == "failed"
                and result.get("iterations", 0) <= 1
                and not result.get("text")
                and not result.get("tool_runs")
                and not stream["open"]):
            return False

        def _finish_ui():
            try:
                # A turn interrupted mid-stream leaves an open live bubble.
                if stream["open"]:
                    stream["open"] = False
                    txt = stream["text"]
                    self._finalize_stream_bubble(txt)
                    self._clear_stream_refs()
                    if txt:
                        self._add_to_history("assistant", txt)
                self._hide_typing_indicator()

                st = result.get("status")
                if st == "cancelled":
                    self._append_bot_message(
                        u"Đã dừng theo yêu cầu." if viet else u"Stopped.",
                        icon=_ICON_STOP, icon_color=_ICON_SLATE)
                elif st == "doc_changed":
                    self._append_bot_message(
                        (u"Bạn đã chuyển sang document khác — yêu cầu bị hủy "
                         u"để tránh sửa nhầm model.") if viet else
                        (u"The active document changed — request cancelled "
                         u"to avoid editing the wrong model."),
                        icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                elif st == "failed":
                    label = get_provider_display_label()
                    detail = u""
                    try:
                        _le = provider.get_last_error()
                        if _le:
                            detail = (u"\nChi tiết: {}".format(_le) if viet
                                      else u"\nDetail: {}".format(_le))
                    except Exception:
                        pass
                    self._append_bot_message(
                        (u"Model AI ({}) bị gián đoạn giữa chừng — kết quả có "
                         u"thể chưa trọn vẹn. Thử lại nhé.{}".format(label, detail))
                        if viet else
                        (u"The AI model ({}) dropped mid-request — the result "
                         u"may be incomplete. Please retry.{}".format(label, detail)),
                        icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                elif st in ("max_iterations", "timeout"):
                    self._append_bot_message(
                        (u"Yêu cầu quá dài — đã dừng sau {} bước. Hãy chia nhỏ "
                         u"yêu cầu để tiếp tục.").format(result.get("iterations"))
                        if viet else
                        (u"Request too long — stopped after {} steps. Split it "
                         u"up to continue.").format(result.get("iterations")),
                        icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                elif (st == "done" and not result.get("text")
                        and result.get("tool_runs")):
                    self._append_bot_message(
                        u"Đã thực hiện xong {} bước công cụ.".format(
                            result.get("tool_runs")) if viet else
                        u"Completed {} tool step(s).".format(
                            result.get("tool_runs")),
                        icon=_ICON_SUCCESS, icon_color=_ICON_GREEN)

                li = result.get("launch_intent")
                if li:
                    # Terminal launcher: _execute_result opens the window on
                    # the UI thread and releases the busy state itself.
                    self._execute_result({"intent": li, "message": u"",
                                          "params": {}})
                else:
                    self._set_busy(False)
            except Exception as ex:
                logger.error("native agent finish error: {}".format(ex))
                self._set_busy(False)

        self.Dispatcher.Invoke(Action(_finish_ui))
        return True

    # ─── Tool-call cards ───────────────────────────────────────────────────────

    def _append_tool_card(self, name, args):
        """Add a tool-call status card to the chat. UI thread only.

        Returns a handle dict for _update_tool_card, or None on failure.
        """
        try:
            from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation
            from System.Windows import Thickness, CornerRadius, TextWrapping
            from System.Windows.Media import SolidColorBrush, Color, FontFamily

            card = Border()
            card.Background      = SolidColorBrush(Color.FromRgb(248, 250, 252))  # #F8FAFC
            card.BorderBrush     = SolidColorBrush(Color.FromRgb(226, 232, 240))  # #E2E8F0
            card.BorderThickness = Thickness(1)
            card.CornerRadius    = CornerRadius(8)
            card.Padding         = Thickness(12, 8, 12, 8)
            # Left margin lines the card up with bot bubbles (avatar 36 + 10).
            card.Margin          = Thickness(46, 0, 60, 8)

            panel = StackPanel()

            head = StackPanel()
            head.Orientation = Orientation.Horizontal

            status = TextBlock()
            status.Text       = u""   # MDL2 Sync — running
            status.FontFamily = FontFamily(u"Segoe MDL2 Assets")
            status.FontSize   = 12
            status.Foreground = SolidColorBrush(Color.FromRgb(59, 130, 246))      # #3B82F6
            status.Margin     = Thickness(0, 1, 8, 0)

            title = TextBlock()
            title.Text       = name
            title.FontFamily = FontFamily(u"Consolas")
            title.FontSize   = 12
            title.FontWeight = System.Windows.FontWeights.SemiBold
            title.Foreground = SolidColorBrush(Color.FromRgb(15, 23, 42))          # #0F172A

            dur = TextBlock()
            dur.Text       = u"đang chạy…"
            dur.FontSize   = 11
            dur.Foreground = SolidColorBrush(Color.FromRgb(148, 163, 184))         # #94A3B8
            dur.Margin     = Thickness(8, 1, 0, 0)

            head.Children.Add(status)
            head.Children.Add(title)
            head.Children.Add(dur)
            panel.Children.Add(head)

            try:
                args_s = json.dumps(args, ensure_ascii=False)
            except Exception:
                args_s = u"{}".format(args)
            if len(args_s) > 160:
                args_s = args_s[:160] + u"…"
            args_tb = TextBlock()
            args_tb.Text         = args_s
            args_tb.FontSize     = 11
            args_tb.Foreground   = SolidColorBrush(Color.FromRgb(100, 116, 139))   # #64748B
            args_tb.TextWrapping = TextWrapping.Wrap
            args_tb.Margin       = Thickness(20, 2, 0, 0)
            panel.Children.Add(args_tb)

            result_tb = TextBlock()
            result_tb.FontSize     = 11
            result_tb.Foreground   = SolidColorBrush(Color.FromRgb(51, 65, 85))    # #334155
            result_tb.TextWrapping = TextWrapping.Wrap
            result_tb.Margin       = Thickness(20, 3, 0, 0)
            result_tb.Visibility   = Visibility.Collapsed
            panel.Children.Add(result_tb)

            card.Child = panel
            self.chat_history_panel.Children.Add(card)
            self._scroll_to_bottom()
            return {"card": card, "status": status, "dur": dur,
                    "result": result_tb}
        except Exception as ex:
            logger.debug("_append_tool_card error: {}".format(ex))
            return None

    def _update_tool_card(self, handle, ok, seconds, result):
        """Mark a tool card done/failed and show the result summary. UI thread."""
        if not handle:
            return
        try:
            from System.Windows.Media import SolidColorBrush, Color

            status = handle["status"]
            if ok:
                status.Text       = u""   # MDL2 CheckMark
                status.Foreground = SolidColorBrush(Color.FromRgb(16, 185, 129))   # #10B981
            else:
                status.Text       = u""   # MDL2 Cancel
                status.Foreground = SolidColorBrush(Color.FromRgb(239, 68, 68))    # #EF4444

            handle["dur"].Text = u"{0:.1f}s".format(seconds)

            try:
                res_s = json.dumps(result, ensure_ascii=False)
            except Exception:
                res_s = u"{}".format(result)
            rt = handle["result"]
            rt.Text       = res_s[:240] + (u"…" if len(res_s) > 240 else u"")
            rt.ToolTip    = res_s[:4000]
            rt.Visibility = Visibility.Visible

            # C1: clickable element-id links → select & zoom in Revit.
            try:
                ids = self._extract_element_ids(result)
                if ok and ids:
                    from System.Windows.Controls import StackPanel, TextBlock, Orientation
                    from System.Windows import Thickness, TextDecorations
                    from System.Windows.Input import Cursors

                    links = StackPanel()
                    links.Orientation = Orientation.Horizontal
                    links.Margin = Thickness(20, 4, 0, 0)

                    def _mk_link(label, id_list):
                        tb = TextBlock()
                        tb.Text           = label
                        tb.FontSize       = 11
                        tb.Foreground     = SolidColorBrush(Color.FromRgb(59, 130, 246))  # #3B82F6
                        tb.TextDecorations = TextDecorations.Underline
                        tb.Cursor         = Cursors.Hand
                        tb.Margin         = Thickness(0, 0, 10, 0)
                        tb.ToolTip        = u"Chọn & zoom trong Revit"

                        def _click(s, e, _ids=list(id_list)):
                            self._select_in_revit_async(_ids)

                        tb.MouseLeftButtonUp += _click
                        return tb

                    for eid in ids[:6]:
                        links.Children.Add(_mk_link(u"#{}".format(eid), [eid]))
                    if len(ids) > 1:
                        links.Children.Add(
                            _mk_link(u"chọn cả {}".format(len(ids)), ids))
                    handle["card"].Child.Children.Add(links)
            except Exception:
                pass

            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug("_update_tool_card error: {}".format(ex))

    @staticmethod
    def _extract_element_ids(result, _limit=60):
        """Collect Revit element ids out of a tool-result dict (C1).

        Recognizes the common id-bearing shapes across the ~75 MCP tools:
        {'id': n}, {'element_id': n}, {'element_ids'|'created_ids'|'ids': [...]},
        and nested lists of {'id': n} dicts (get_current_view_elements, ...).
        Order-preserving, deduped, capped so a huge element dump stays cheap.
        """
        out  = []
        seen = set()

        def _add(v):
            try:
                n = int(v)
            except Exception:
                return
            if n > 0 and n not in seen:
                seen.add(n)
                out.append(n)

        def _walk(node, depth):
            if depth > 4 or len(out) >= _limit:
                return
            if isinstance(node, dict):
                for k, v in node.items():
                    lk = u"{}".format(k).lower()
                    if lk in ("id", "element_id", "new_element_id", "new_id",
                              "tag_id", "wall_id", "grid_id", "level_id"):
                        _add(v)
                    elif lk in ("element_ids", "created_ids", "ids", "new_ids",
                                "tag_ids", "wall_ids", "created_element_ids"):
                        if isinstance(v, (list, tuple)):
                            for item in v:
                                _add(item)
                    elif isinstance(v, (dict, list, tuple)):
                        _walk(v, depth + 1)
            elif isinstance(node, (list, tuple)):
                for item in node:
                    _walk(item, depth + 1)

        _walk(result if isinstance(result, dict) else {}, 0)
        return out

    def _select_in_revit_async(self, element_ids):
        """Select + zoom elements from an element-link click (C1).

        Spawns a WORKER thread: _execute_tool blocks on the ExternalEvent,
        and waiting for that on the UI thread would deadlock (the handler
        itself needs the UI thread to run).
        """
        ids = [i for i in (element_ids or [])]
        if not ids:
            return

        def _work():
            try:
                from core.server import get_t3labai_server
                srv = get_t3labai_server()
                srv._execute_tool('select_elements',
                                  {'element_ids': ids, 'show': True,
                                   'limit': len(ids)})
            except Exception:
                pass

        t = Thread(ThreadStart(_work))
        t.IsBackground = True
        t.Start()

    # ─── Destructive-tool confirmation (B5) ────────────────────────────────────

    def _confirm_tool_blocking(self, name, args, viet, timeout_sec=120):
        """WORKER thread: render a Confirm/Cancel card and block until the
        user decides. Returns True only on an explicit Confirm click —
        timeout, Stop, or any error all count as declined.
        """
        import threading
        state = {"decision": None, "seal": None}
        evt = threading.Event()

        def _ui():
            try:
                self._hide_typing_indicator()
                self._append_confirm_card(name, args, state, evt, viet)
            except Exception:
                state["decision"] = False
                evt.set()

        try:
            self.Dispatcher.Invoke(Action(_ui))
        except Exception:
            return False

        waited = 0.0
        while waited < timeout_sec and not evt.is_set():
            evt.wait(0.25)
            waited += 0.25
            loop = self._agent_loop
            if loop is not None and loop.is_cancelled():
                break

        if state["decision"] is None:
            # Timeout / Stop — seal the card so stale buttons can't approve
            # a request that is already over.
            def _expire():
                try:
                    if state.get("seal"):
                        state["seal"](u"⏱ Hết hạn — đã bỏ qua" if viet
                                      else u"⏱ Expired — skipped")
                except Exception:
                    pass
            try:
                self.Dispatcher.BeginInvoke(Action(_expire))
            except Exception:
                pass
        return state["decision"] is True

    def _append_confirm_card(self, name, args, state, evt, viet):
        """Confirm/Cancel card for a destructive tool call. UI thread only."""
        from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation, Button
        from System.Windows.Documents import Run
        from System.Windows import Thickness, CornerRadius, TextWrapping
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows.Input import Cursors

        card = Border()
        card.Background      = SolidColorBrush(Color.FromRgb(254, 242, 242))  # #FEF2F2
        card.BorderBrush     = SolidColorBrush(Color.FromRgb(252, 165, 165))  # #FCA5A5
        card.BorderThickness = Thickness(1)
        card.CornerRadius    = CornerRadius(8)
        card.Padding         = Thickness(12, 10, 12, 10)
        card.Margin          = Thickness(46, 0, 60, 8)

        panel = StackPanel()

        head = TextBlock()
        head.FontSize   = 12
        head.FontWeight = System.Windows.FontWeights.SemiBold
        head.Foreground = SolidColorBrush(Color.FromRgb(185, 28, 28))          # #B91C1C
        # Minimal MDL2 warning glyph — needs its own FontFamily run; the plain
        # "⚠" character rendered in the body font would show as a colored
        # emoji glyph (or tofu) instead of a flat monochrome icon.
        self._add_icon_run(head, _ICON_WARNING, (185, 28, 28), size=12)
        head.Inlines.Add(Run(u"Xác nhận hành động phá hủy" if viet
                             else u"Confirm destructive action"))
        panel.Children.Add(head)

        try:
            args_s = json.dumps(args, ensure_ascii=False)
        except Exception:
            args_s = u"{}".format(args)
        if len(args_s) > 200:
            args_s = args_s[:200] + u"…"
        body = TextBlock()
        body.Text         = u"`{}` — {}".format(name, args_s)
        body.FontSize     = 11.5
        body.TextWrapping = TextWrapping.Wrap
        body.Foreground   = SolidColorBrush(Color.FromRgb(51, 65, 85))         # #334155
        body.Margin       = Thickness(0, 4, 0, 8)
        panel.Children.Add(body)

        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal

        status_tb = TextBlock()
        status_tb.FontSize   = 11.5
        status_tb.Foreground = SolidColorBrush(Color.FromRgb(100, 116, 139))   # #64748B
        status_tb.Margin     = Thickness(10, 5, 0, 0)
        status_tb.Visibility = Visibility.Collapsed

        def _mk_btn(label, bg, fg):
            b = Button()
            b.Content         = label
            b.FontSize        = 12
            b.FontWeight      = System.Windows.FontWeights.SemiBold
            b.Padding         = Thickness(14, 5, 14, 5)
            b.Margin          = Thickness(0, 0, 8, 0)
            b.Cursor          = Cursors.Hand
            b.Background      = SolidColorBrush(bg)
            b.Foreground      = SolidColorBrush(fg)
            b.BorderThickness = Thickness(0)
            return b

        ok_btn = _mk_btn(u"Xác nhận" if viet else u"Confirm",
                         Color.FromRgb(239, 68, 68), Color.FromRgb(255, 255, 255))
        no_btn = _mk_btn(u"Hủy" if viet else u"Cancel",
                         Color.FromRgb(241, 245, 249), Color.FromRgb(15, 23, 42))

        def _seal(msg):
            try:
                ok_btn.IsEnabled     = False
                no_btn.IsEnabled     = False
                status_tb.Text       = msg
                status_tb.Visibility = Visibility.Visible
            except Exception:
                pass
        state["seal"] = _seal

        def _on_ok(s, e):
            state["decision"] = True
            _seal(u"✓ Đã xác nhận" if viet else u"✓ Confirmed")
            evt.set()

        def _on_cancel(s, e):
            state["decision"] = False
            _seal(u"✗ Đã hủy" if viet else u"✗ Cancelled")
            evt.set()

        ok_btn.Click += _on_ok
        no_btn.Click += _on_cancel

        btn_row.Children.Add(ok_btn)
        btn_row.Children.Add(no_btn)
        btn_row.Children.Add(status_tb)
        panel.Children.Add(btn_row)

        card.Child = panel
        self.chat_history_panel.Children.Add(card)
        self._scroll_to_bottom()

    # ─── Vision view capture (C2) ──────────────────────────────────────────────

    def _wants_view_snapshot(self, text, provider):
        """True when the user asks the assistant to LOOK at the current view
        and the active provider can take Claude-format image blocks (the
        agent path currently ships vision only for Claude)."""
        if not text:
            return False
        if provider is None or getattr(provider, "NAME", "") != "claude":
            return False
        try:
            if not provider.supports_vision():
                return False
        except Exception:
            return False
        import re as _re
        pat = (u"(nhìn|xem|quan sát|soi|chụp|đánh giá|kiểm tra"
               u"|look|inspect|review|check|analy)"
               u"[^\n]{0,24}"
               u"(view|màn hình|bố cục|layout|screen)")
        return _re.search(pat, text, _re.IGNORECASE | _re.UNICODE) is not None

    def _capture_active_view(self, srv):
        """WORKER thread: export the active view as a PNG (~1280px) through
        the ExternalEvent write path. Returns the PNG path or None."""
        try:
            import tempfile
            folder = os.path.join(tempfile.gettempdir(), 'T3Lab_ViewShots')
            res = srv._execute_tool('export_image',
                                    {'width': 1280, 'output_folder': folder})
            files = (res or {}).get('files') or []
            for p in files:
                if p.lower().endswith('.png') and os.path.isfile(p):
                    return p
            return files[0] if files else None
        except Exception:
            return None

    # ─── Chat UI helpers ──────────────────────────────────────────────────────

    _avatar_bitmap = None   # lazy-loaded, cached BitmapImage of icon.png (per window)
    _avatar_bitmap_failed = False

    def _get_avatar_bitmap(self):
        """Return the cached BitmapImage for the tool's own icon.png, or None.

        icon.png is the SAME icon Revit shows on the T3Lab Assistant ribbon
        button (T3LabAssistant.pushbutton/icon.png) — using it as the chat
        avatar means the bubble literally shows the tool's real icon instead
        of a generic colored "T3" initials badge.
        """
        if self._avatar_bitmap is not None:
            return self._avatar_bitmap
        if self._avatar_bitmap_failed:
            return None
        try:
            from System import Uri, UriKind
            from System.Windows.Media.Imaging import BitmapCacheOption
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.png')
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource   = Uri(icon_path, UriKind.Absolute)
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            self._avatar_bitmap = bmp
            return bmp
        except Exception as ex:
            logger.debug("_get_avatar_bitmap error: {}".format(ex))
            self._avatar_bitmap_failed = True
            return None

    def _make_avatar(self, letter, _unused_start=None, _unused_end=None):
        """Create a circular avatar showing the tool's real icon.png.

        Falls back to a plain initials badge (BatchOut blue #3498DB) only if
        the icon file can't be loaded — the app must never crash a chat
        bubble over a missing/locked icon asset.
        """
        from System.Windows.Controls import Border, TextBlock
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Media import SolidColorBrush, Color, ImageBrush, Stretch
        from System.Windows import HorizontalAlignment, VerticalAlignment

        av = Border()
        av.Width = 36
        av.Height = 36
        av.CornerRadius = CornerRadius(18)
        av.Margin = Thickness(0, 2, 10, 0)
        av.VerticalAlignment = VerticalAlignment.Top

        bmp = self._get_avatar_bitmap()
        if bmp is not None:
            brush = ImageBrush(bmp)
            brush.Stretch = Stretch.UniformToFill
            av.Background = brush
            return av

        av.Background = SolidColorBrush(Color.FromRgb(52, 152, 219))   # #3498DB fallback
        lbl = TextBlock()
        lbl.Text = letter
        lbl.FontSize = 12
        lbl.FontWeight = System.Windows.FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment = VerticalAlignment.Center
        av.Child = lbl
        return av

    @staticmethod
    def _add_icon_run(text_block, glyph, rgb=None, size=13.5):
        """Prepend a Segoe MDL2 Assets icon Run to a TextBlock, followed by a
        thin space — the ONLY way to render a real icon glyph inline: the rest
        of the bubble uses Hanken Grotesk/Inter, which has no glyph at these
        private-use codepoints, so the icon needs its own FontFamily run
        rather than living in the same string as the body text.
        """
        from System.Windows.Documents import Run
        from System.Windows.Media import FontFamily as _WpfFontFamily, SolidColorBrush, Color

        icon_run = Run(glyph)
        icon_run.FontFamily = _WpfFontFamily(u"Segoe MDL2 Assets")
        icon_run.FontSize   = size
        if rgb:
            icon_run.Foreground = SolidColorBrush(Color.FromRgb(*rgb))
        text_block.Inlines.Add(icon_run)
        text_block.Inlines.Add(Run(u"  "))

    def _append_user_message(self, text, attachment_note=None):
        """Add a right-aligned user bubble (BatchOut #3498DB).

        attachment_note, if given, renders as its own line with a minimal
        Attach glyph — replaces the old baked-in "📎 filename" text so the
        indicator is a real icon, not a colored emoji character.
        """
        try:
            from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
            from System.Windows.Documents import Run, LineBreak
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
            msg_text.FontSize    = 13
            msg_text.Foreground  = SolidColorBrush(Color.FromRgb(255, 255, 255))
            msg_text.TextWrapping = TextWrapping.Wrap
            if text:
                msg_text.Inlines.Add(Run(text))
            if attachment_note:
                if text:
                    msg_text.Inlines.Add(LineBreak())
                self._add_icon_run(msg_text, _ICON_ATTACH, size=12)
                msg_text.Inlines.Add(Run(attachment_note))
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

    # Legacy colored emoji markers (still produced by older message-builders
    # like _try_fast_context_answer) mapped to a minimal MDL2 glyph + Lumina
    # color — converted at render time so no caller needs to change its
    # markdown text, only this one renderer.
    _MD_ICON_MARKERS = [
        (u"⚡ ",            "_ICON_ANALYZE", "_ICON_BLUE"),   # instant DB answer
        (u"\U0001f4cb ",   "_ICON_LIST",    "_ICON_SLATE"),  # stats/info section
        (u"\U0001f5bc️ ", "_ICON_SEARCH", "_ICON_SLATE"),  # view section
        (u"\U0001f3af ",   "_ICON_LIST",    "_ICON_SLATE"),  # selection section
        (u"\U0001f4ca ",   "_ICON_LIST",    "_ICON_BLUE"),   # LLM chart/stats emoji
        (u"\U0001f4c8 ",   "_ICON_LIST",    "_ICON_BLUE"),   # LLM trend emoji
    ]

    @staticmethod
    def _add_inline_md(text_block, line):
        """Render ONE line's inline markdown into text_block:
        **bold** and `code` spans (code = Consolas, ink color)."""
        import re as _re
        from System.Windows.Documents import Run
        from System.Windows import FontWeights
        from System.Windows.Media import FontFamily as _WpfFontFamily, SolidColorBrush, Color

        for b_idx, b_seg in enumerate(_re.split(r'\*\*', line)):
            if not b_seg:
                continue
            bold = (b_idx % 2 == 1)
            for c_idx, c_seg in enumerate(b_seg.split(u'`')):
                if not c_seg:
                    continue
                r = Run()
                r.Text = c_seg
                if bold:
                    r.FontWeight = FontWeights.SemiBold
                if c_idx % 2 == 1:              # inside `code`
                    r.FontFamily = _WpfFontFamily(u"Consolas")
                    r.Foreground = SolidColorBrush(Color.FromRgb(15, 23, 42))
                text_block.Inlines.Add(r)

    @staticmethod
    def _build_md_inlines(text_block, text):
        """
        Populate text_block.Inlines with paragraph-level markdown:
        leading emoji→MDL2 icon markers, # headings, bullet lines (* / -),
        **bold** / `code` inline spans, line breaks. Tables are handled one
        level up by _render_md_blocks (a TextBlock cannot host a grid).
        """
        from System.Windows.Documents import Run, LineBreak
        from System.Windows import FontWeights

        module_globals = globals()
        lines = text.splitlines()
        for i, line in enumerate(lines):
            if i > 0:
                text_block.Inlines.Add(LineBreak())

            # Legacy emoji section markers → minimal MDL2 icon run
            stripped = line.strip()
            for marker, glyph_name, color_name in T3LabAssistantWindow._MD_ICON_MARKERS:
                if stripped.startswith(marker):
                    T3LabAssistantWindow._add_icon_run(
                        text_block, module_globals[glyph_name], module_globals[color_name])
                    line = stripped[len(marker):]
                    break

            # Headings: "# " / "## " / "### " → semibold, slightly larger
            stripped = line.strip()
            if stripped.startswith(u'#'):
                h = len(stripped) - len(stripped.lstrip(u'#'))
                if 1 <= h <= 4 and stripped[h:h + 1] == u' ':
                    r = Run()
                    r.Text       = stripped[h + 1:].strip()
                    r.FontWeight = FontWeights.SemiBold
                    r.FontSize   = 14.5 if h <= 2 else 13.5
                    text_block.Inlines.Add(r)
                    continue

            # Bullet lines
            if stripped.startswith(u'* ') or stripped.startswith(u'- '):
                prefix_run = Run()
                prefix_run.Text = u'• '   # bullet •
                text_block.Inlines.Add(prefix_run)
                line = stripped[2:]             # remaining text after bullet marker

            T3LabAssistantWindow._add_inline_md(text_block, line)

    def _make_md_table(self, raw_rows):
        """Render markdown pipe-rows ("| a | b |") as a bordered WPF Grid.

        Returns a Border-wrapped Grid, or None when the rows don't form a
        usable table (caller then falls back to showing the raw text).
        Header row (followed by |---|---|) gets the Lumina table header
        treatment; star-sized columns so the table always fits the bubble.
        """
        try:
            import re as _re
            from System.Windows.Controls import Grid, ColumnDefinition, RowDefinition, Border, TextBlock
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength, GridUnitType
            from System.Windows.Media import SolidColorBrush, Color

            parsed = []
            for r in raw_rows:
                inner = r.strip()
                if inner.startswith(u"|"):
                    inner = inner[1:]
                if inner.endswith(u"|"):
                    inner = inner[:-1]
                parsed.append([c.strip() for c in inner.split(u"|")])

            def _is_sep(cells):
                return bool(cells) and all(
                    _re.match(r'^:?-{2,}:?$', c or u'') for c in cells)

            header = None
            if len(parsed) >= 2 and _is_sep(parsed[1]):
                header = parsed[0]
                body = [r for r in parsed[2:] if not _is_sep(r)]
            else:
                body = [r for r in parsed if not _is_sep(r)]
            rows = ([header] if header is not None else []) + body
            if not rows:
                return None
            ncols = max(len(r) for r in rows)
            if ncols < 2:
                return None

            _line  = Color.FromRgb(226, 232, 240)   # #E2E8F0 divider
            _inkhd = Color.FromRgb(15, 23, 42)      # #0F172A header ink
            _ink   = Color.FromRgb(39, 39, 42)      # #27272A body ink

            g = Grid()
            for _c in range(ncols):
                cd = ColumnDefinition()
                cd.Width = GridLength(1, GridUnitType.Star)
                g.ColumnDefinitions.Add(cd)

            for ri, row in enumerate(rows):
                g.RowDefinitions.Add(RowDefinition())
                is_head = (header is not None and ri == 0)
                for ci in range(ncols):
                    cell = Border()
                    cell.BorderBrush = SolidColorBrush(_line)
                    cell.BorderThickness = Thickness(
                        0, 0,
                        1 if ci < ncols - 1 else 0,
                        1 if ri < len(rows) - 1 else 0)
                    cell.Padding = Thickness(8, 4, 8, 4)
                    if is_head:
                        cell.Background = SolidColorBrush(Color.FromRgb(248, 250, 252))  # #F8FAFC

                    tb = TextBlock()
                    tb.FontSize     = 12 if is_head else 12.5
                    tb.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
                    tb.TextWrapping = TextWrapping.Wrap
                    tb.Foreground   = SolidColorBrush(_inkhd if is_head else _ink)
                    if is_head:
                        tb.FontWeight = System.Windows.FontWeights.SemiBold
                    self._add_inline_md(tb, row[ci] if ci < len(row) else u"")
                    cell.Child = tb
                    Grid.SetRow(cell, ri)
                    Grid.SetColumn(cell, ci)
                    g.Children.Add(cell)

            outer = Border()
            outer.BorderBrush     = SolidColorBrush(_line)
            outer.BorderThickness = Thickness(1)
            outer.CornerRadius    = CornerRadius(6)
            outer.Margin          = Thickness(0, 6, 0, 6)
            outer.Child = g
            return outer
        except Exception as ex:
            logger.debug("_make_md_table error: {}".format(ex))
            return None

    def _render_md_blocks(self, text, icon=None, icon_color=None):
        """Build the CONTENT of a bot bubble: a StackPanel of paragraph
        TextBlocks and real table Grids.

        The old single-TextBlock renderer showed markdown tables as raw
        "| a | b |" pipe text; consecutive pipe-lines now become a bordered
        grid with a header row, so LLM answers containing tables read
        cleanly. icon/icon_color prefix the first paragraph.
        """
        from System.Windows.Controls import StackPanel, TextBlock
        from System.Windows import TextWrapping, Thickness
        from System.Windows.Media import SolidColorBrush, Color

        panel = StackPanel()
        state = {"icon": icon}

        def _new_tb():
            tb = TextBlock()
            tb.FontSize     = 13
            tb.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            tb.Foreground   = SolidColorBrush(Color.FromRgb(39, 39, 42))   # #27272A
            tb.TextWrapping = TextWrapping.Wrap
            tb.LineHeight   = 20
            return tb

        para = []

        def _flush_para():
            if not para:
                return
            chunk = u"\n".join(para).strip(u"\n")
            del para[:]
            if not chunk.strip() and not state["icon"]:
                return
            tb = _new_tb()
            if state["icon"]:
                self._add_icon_run(tb, state["icon"], icon_color)
                state["icon"] = None
            self._build_md_inlines(tb, chunk)
            tb.Margin = Thickness(0, 0, 0, 2)
            panel.Children.Add(tb)

        lines = (text or u"").splitlines() or [u""]
        i = 0
        n = len(lines)
        while i < n:
            s = lines[i].strip()
            if s.startswith(u"|") and s.count(u"|") >= 2:
                tbl = []
                while i < n:
                    s2 = lines[i].strip()
                    if s2.startswith(u"|") and s2.count(u"|") >= 2:
                        tbl.append(s2)
                        i += 1
                    else:
                        break
                _flush_para()
                table = self._make_md_table(tbl)
                if table is not None:
                    panel.Children.Add(table)
                else:
                    para.extend(tbl)     # unparseable → show as raw text
                continue
            para.append(lines[i])
            i += 1
        _flush_para()

        if panel.Children.Count == 0:
            panel.Children.Add(_new_tb())
        return panel

    def _append_bot_message(self, text, icon=None, icon_color=None):
        """Add a left-aligned bot bubble with avatar and basic markdown rendering.

        icon/icon_color: optional Segoe MDL2 glyph + Lumina RGB tuple rendered
        before the text, in its own FontFamily run — the minimal-icon
        replacement for the old baked-in colored emoji prefixes.
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

            # Block renderer: paragraphs + real table grids. Falls back to a
            # plain TextBlock if anything in the renderer throws.
            try:
                content = self._render_md_blocks(text, icon=icon, icon_color=icon_color)
            except Exception:
                content = TextBlock()
                content.FontSize     = 13
                content.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
                content.Foreground   = SolidColorBrush(Color.FromRgb(39, 39, 42))   # #27272A
                content.TextWrapping = TextWrapping.Wrap
                content.LineHeight   = 20
                content.Text         = text

            bubble.Child = content
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
        """Re-render the live bubble with full markdown once streaming is done.

        The streaming TextBlock showed plain text; swap the bubble's content
        for the block renderer so tables/headings in the finished reply get
        their real layout. Falls back to inline-only rendering on error.
        """
        try:
            tb = self._stream_tb
            if tb is None:
                return
            try:
                bubble = tb.Parent          # the bubble Border hosting the stream
                bubble.Child = self._render_md_blocks(text)
            except Exception:
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
