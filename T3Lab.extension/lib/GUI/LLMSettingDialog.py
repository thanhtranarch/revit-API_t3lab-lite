# -*- coding: utf-8 -*-
"""
LLMs Setting Dialog — the T3Lab Assistant settings hub

Tabbed settings window: General (profile, action mode, data), Models
(provider/model/API key/connection), Projects (workspaces), Knowledge
(RAG index) and Skills. All state lives in the shared LLMRouter /
T3LabAISettings / UserProfile / ProjectStore singletons, so anything
changed here is immediately visible to the T3Lab Assistant (and any
other AI-powered tool) too.
"""

from __future__ import unicode_literals

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System import Action
from System.Threading import Thread, ThreadStart, ApartmentState
from System.Windows import Visibility, WindowState
from System.Windows.Media import SolidColorBrush, Color

from pyrevit import forms, script

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'LLMSetting.xaml')

logger = script.get_logger()

try:
    from Intelligence.knowledge.knowledge_store import (get_active_store,
                                                        default_knowledge_dir)
    HAS_KNOWLEDGE = True
except Exception:
    HAS_KNOWLEDGE = False

    def get_active_store():
        return None

    def default_knowledge_dir():
        return ''


def _brush(r, g, b):
    return SolidColorBrush(Color.FromRgb(r, g, b))


_READY  = _brush(16, 185, 129)
_GRAY   = _brush(230, 230, 234)
_MUTED  = _brush(161, 161, 170)
_GREEN  = _brush(16, 185, 129)
_RED    = _brush(239, 68, 68)


class LLMSettingWindow(forms.WPFWindow):
    """Standalone dialog for LLM provider / model / API key / connection setup."""

    _BRAND_COLORS = {
        "claude":   (217, 119,  87),
        "openai":   ( 16, 163, 127),
        "deepseek": ( 37,  99, 235),
        "ollama":   ( 59, 130, 246),
        "lmstudio": (124,  58, 237),
    }
    _PROV_INDEX = {"claude": 0, "openai": 1, "deepseek": 2, "ollama": 3, "lmstudio": 4}

    _KEY_PROVIDERS = ("claude", "openai", "deepseek")
    _KEY_NAME_MAP  = {"claude": "Claude", "openai": "OpenAI", "deepseek": "DeepSeek"}
    _KEY_LABELS = {
        "claude":   u"Anthropic API Key (sk-ant-...)",
        "openai":   u"OpenAI API Key (sk-...)",
        "deepseek": u"DeepSeek API Key (sk-...)",
        "ollama":   u"No key needed (local)",
        "lmstudio": u"No key needed — start LM Studio first",
    }
    _API_KEY_URLS = {
        "claude":   "https://console.anthropic.com/settings/keys",
        "openai":   "https://platform.openai.com/api-keys",
        "deepseek": "https://platform.deepseek.com/api_keys",
    }
    _HOST_DEFAULTS = {
        "ollama":   "http://localhost:11434",
        "lmstudio": "http://localhost:1234",
    }
    _HOST_LABELS = {
        "ollama":   u"Ollama base URL",
        "lmstudio": u"LM Studio base URL",
    }

    def __init__(self):
        self._ui_ready     = False   # guards tab_changed during XAML load
        self._action_guard  = False  # guards action_mode_toggled re-entry
        self._think_guard   = False  # guards extended_thinking_toggled re-entry
        self._quality_guard = False  # guards quality_mode_toggled re-entry
        self._embed_guard   = False  # guards knowledge_embed_toggled re-entry
        self._kn_scan_busy = False

        forms.WPFWindow.__init__(self, _XAML)
        self._models_cache = {}
        self._probing = False

        self._update_instant()
        self._load_general_tab()
        self._load_projects_tab()
        self._load_knowledge_tab()
        self._load_skills_tab()
        self._ui_ready = True
        self._probe_all_async()

    # ─── Tabs ───────────────────────────────────────────────────────────────

    def tab_changed(self, sender, e):
        """Swap the visible settings panel when a tab pill is checked."""
        if not getattr(self, '_ui_ready', False):
            return
        try:
            tag = sender.Tag
            panels = {
                'general':   self.panel_general,
                'provider':  self.panel_provider,
                'projects':  self.panel_projects,
                'knowledge': self.panel_knowledge,
                'skills':    self.panel_skills,
            }
            for key in panels:
                panels[key].Visibility = (
                    Visibility.Visible if key == tag else Visibility.Collapsed)
        except Exception as ex:
            logger.debug("tab_changed error: {}".format(ex))

    # ─── Chrome ─────────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def close_button_clicked(self, sender, e):
        self.Close()

    # ─── Provider / model ───────────────────────────────────────────────────

    def provider_changed(self, sender, e):
        try:
            item = self.provider_combo.SelectedItem
            if item is None:
                return
            tag = item.Tag
            if not tag:
                return
            self._populate_model_combo(tag)
            self._apply_host_panel(tag)
            self._switch_provider(tag)
        except Exception as ex:
            logger.debug("provider_changed error: {}".format(ex))

    def _switch_provider(self, name):
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            if not router.switch_provider(name):
                return
            self._apply_provider_chrome(name)

            if self._probing:
                return
            self._probing = True

            def _bg():
                try:
                    router.probe_provider(name)
                    provider = router.get_active_provider()
                    if provider:
                        try:
                            self._models_cache[name] = provider.get_models()
                        except Exception:
                            pass
                    self.Dispatcher.Invoke(Action(self._update_probed))
                except Exception:
                    pass
                finally:
                    self._probing = False

            t = Thread(ThreadStart(_bg))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception as ex:
            logger.debug("_switch_provider error: {}".format(ex))

    def model_changed(self, sender, e):
        try:
            from Intelligence.llm_router import LLMRouter
            item = self.model_combo.SelectedItem
            if item is None:
                return
            router = LLMRouter()
            router.set_model(router.get_active_name(), item.ToString())
        except Exception as ex:
            logger.debug("model_changed error: {}".format(ex))

    def save_model_clicked(self, sender, e):
        try:
            from Intelligence.llm_router import LLMRouter
            item = self.model_combo.SelectedItem
            if item is None:
                self.model_saved_hint.Text = u"Select a model first"
                self._flash_saved_hint()
                return
            router = LLMRouter()
            router.set_model(router.get_active_name(), item.ToString())
            self.model_saved_hint.Foreground = _GREEN
            self.model_saved_hint.Text = u"✓ Saved"
            self._flash_saved_hint()
        except Exception as ex:
            logger.debug("save_model_clicked error: {}".format(ex))

    def _flash_saved_hint(self):
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
            pass

    def _populate_model_combo(self, active):
        """Fill MODEL combo for `active`. Empty+disabled until a live model list exists."""
        try:
            try:
                self.model_combo.SelectionChanged -= self.model_changed
            except Exception:
                pass

            self.model_combo.Items.Clear()
            models  = list(self._models_cache.get(active, []))
            enabled = bool(models)
            hint    = u""

            if enabled:
                saved = None
                try:
                    from config.settings import T3LabAISettings
                    saved = T3LabAISettings().get_provider_model(active)
                except Exception:
                    pass
                for m in models:
                    self.model_combo.Items.Add(m)
                if saved and saved in models:
                    self.model_combo.SelectedItem = saved
                else:
                    self.model_combo.SelectedIndex = 0
            else:
                if active in self._KEY_PROVIDERS:
                    hint = (u"Enter an API key first" if not self._has_saved_key(active)
                            else u"Not connected — press Save to test")
                else:
                    hint = u"Start the server & press Test Connection"

            self.model_combo.IsEnabled = enabled
            try:
                self.save_model_btn.IsEnabled = enabled
            except Exception:
                pass
            self.model_saved_hint.Text = hint
            self.model_saved_hint.Foreground = _MUTED
        except Exception as ex:
            logger.debug("_populate_model_combo error: {}".format(ex))
        finally:
            try:
                self.model_combo.SelectionChanged += self.model_changed
            except Exception:
                pass

    def _has_saved_key(self, provider):
        try:
            from config.settings import T3LabAISettings
            return bool(T3LabAISettings().get_api_key(self._KEY_NAME_MAP.get(provider, "")))
        except Exception:
            return False

    # ─── API key ────────────────────────────────────────────────────────────

    def get_api_key_clicked(self, sender, e):
        try:
            from Intelligence.llm_router import LLMRouter
            name = LLMRouter().get_active_name()
            url = self._API_KEY_URLS.get(name)
            if not url:
                return
            import System.Diagnostics
            System.Diagnostics.Process.Start(url)
        except Exception as ex:
            logger.debug("get_api_key_clicked error: {}".format(ex))

    def save_key_clicked(self, sender, e):
        """Save API key → verify connection → fetch models → enable MODEL combo."""
        try:
            from config.settings import T3LabAISettings
            from Intelligence.llm_router import LLMRouter

            key = self.api_key_box.Text.strip()
            if not key or key.endswith("..."):
                return

            router = LLMRouter()
            name = router.get_active_name()
            settings_key = self._KEY_NAME_MAP.get(name)
            if not settings_key:
                return

            T3LabAISettings().set_api_key(settings_key, key)

            provider = router.get_active_provider()
            if provider and hasattr(provider, "reload_credentials"):
                provider.reload_credentials()
            elif provider and hasattr(provider, "invalidate_models_cache"):
                provider.invalidate_models_cache()
            self._models_cache.pop(name, None)

            self.model_combo.IsEnabled = False
            self.save_model_btn.IsEnabled = False
            self.save_key_btn.IsEnabled = False
            self.model_saved_hint.Foreground = _MUTED
            self.model_saved_hint.Text = u"Checking connection…"

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
                    try:
                        self.save_key_btn.IsEnabled = True
                    except Exception:
                        pass
                    if ok:
                        self._models_cache[name] = models
                        self._populate_model_combo(name)
                        self.model_saved_hint.Foreground = _GREEN
                        self.model_saved_hint.Text = u"✓ Connected ({} models)".format(len(models))
                    else:
                        self._models_cache.pop(name, None)
                        self._populate_model_combo(name)
                        self.model_saved_hint.Foreground = _RED
                        self.model_saved_hint.Text = u"✗ Invalid key or connection failed"
                    self._set_status_dot(name, ok)

                self.Dispatcher.Invoke(Action(_apply))

            t = Thread(ThreadStart(_validate))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception as ex:
            logger.debug("save_key_clicked error: {}".format(ex))

    # ─── Local server host (Ollama / LM Studio) ────────────────────────────

    def save_host_clicked(self, sender, e):
        try:
            from config.settings import T3LabAISettings
            from Intelligence.llm_router import LLMRouter
            host = self.host_box.Text.strip()
            if not host:
                return

            router = LLMRouter()
            name = router.get_active_name()

            if name == "lmstudio":
                T3LabAISettings().set_api_key("LMStudio_Host", host)
                provider = router.get_active_provider()
                if provider and hasattr(provider, "reload_credentials"):
                    provider.reload_credentials()
            elif name == "ollama":
                provider = router.get_active_provider()
                if provider and hasattr(provider, "set_host"):
                    provider.set_host(host)

            self._models_cache.pop(name, None)
            self._update_instant()

            def _probe():
                try:
                    router.get_status(use_cache=False)
                    provider = router.get_active_provider()
                    live_models = provider.get_models() if provider else []
                    if live_models:
                        self._models_cache[name] = live_models
                    self.Dispatcher.Invoke(Action(self._update_probed))
                except Exception:
                    pass

            t = Thread(ThreadStart(_probe))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception as ex:
            logger.debug("save_host_clicked error: {}".format(ex))

    def _apply_host_panel(self, name):
        try:
            if name in self._HOST_DEFAULTS:
                self.host_panel.Visibility = Visibility.Visible
                self.host_label.Text = self._HOST_LABELS.get(name, u"Server URL")
                current = u""
                try:
                    if name == "lmstudio":
                        from config.settings import T3LabAISettings
                        current = T3LabAISettings().get_api_key("LMStudio_Host") or u""
                except Exception:
                    pass
                self.host_box.Text = current or self._HOST_DEFAULTS[name]
            else:
                self.host_panel.Visibility = Visibility.Collapsed
        except Exception as ex:
            logger.debug("_apply_host_panel error: {}".format(ex))

    # ─── Test connection ────────────────────────────────────────────────────

    def test_clicked(self, sender, e):
        self.test_result_border.Visibility = Visibility.Visible
        self.test_result_border.Background = _brush(249, 250, 251)
        self.test_result_border.BorderBrush = _brush(229, 231, 235)
        self.test_label.Text = u"Testing…"
        self.test_label.Foreground = _brush(107, 114, 128)
        self.test_result.Text = u""
        self.test_btn.IsEnabled = False
        self.status_text.Text = u"Testing connection…"

        def _do_test():
            ok, label, msg = False, u"Result", u""
            try:
                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                name = router.get_active_name()
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
                    active_model = None
                    try:
                        active_model = provider.get_active_model()
                    except Exception:
                        pass
                    if not active_model:
                        msg = u"No model selected. Choose a model from the Model dropdown."
                    else:
                        resp = provider.chat(
                            [],
                            u"You are a concise assistant. Do not think. Reply in one short sentence only.",
                            u"Reply with exactly this sentence: 'Connected OK'",
                            max_tokens=120,
                        )
                        if resp and resp.strip():
                            ok, label, msg = True, u"Connected", resp.strip()[:120]
                        else:
                            msg = u"Provider responded but returned an empty reply."
            except Exception as ex:
                msg = u"Error: {}".format(str(ex)[:100])

            _ok, _label, _msg = ok, label, msg

            def _update():
                if _ok:
                    self.test_result_border.Background = _brush(240, 253, 244)
                    self.test_result_border.BorderBrush = _brush(187, 247, 208)
                    self.test_label.Foreground = _brush(21, 128, 61)
                else:
                    self.test_result_border.Background = _brush(254, 242, 242)
                    self.test_result_border.BorderBrush = _brush(254, 202, 202)
                    self.test_label.Foreground = _brush(185, 28, 28)
                self.test_label.Text = _label
                self.test_result.Text = _msg
                self.test_btn.IsEnabled = True
                self.status_text.Text = u"Ready"

            self.Dispatcher.Invoke(Action(_update))

        t = Thread(ThreadStart(_do_test))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()

    # ─── Status dots ────────────────────────────────────────────────────────

    def _set_status_dot(self, name, available):
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
            dot.Fill = _READY
            txt.Text = u"Ready"
            txt.Foreground = _READY
        else:
            dot.Fill = _GRAY
            txt.Text = u"Not set up"
            txt.Foreground = _MUTED

    # ─── Full render passes ─────────────────────────────────────────────────

    def _apply_provider_chrome(self, active):
        """Provider combo selection + brand dot + API-key section, no HTTP."""
        self.provider_combo.SelectionChanged -= self.provider_changed
        self.provider_combo.SelectedIndex = self._PROV_INDEX.get(active, 0)
        self.provider_combo.SelectionChanged += self.provider_changed

        rgb = self._BRAND_COLORS.get(active, (161, 161, 170))
        self.provider_brand_dot.Fill = _brush(*rgb)

        self.key_label.Text = self._KEY_LABELS.get(active, u"API Key")
        needs_key = active in self._KEY_PROVIDERS
        self.api_key_box.IsEnabled = needs_key
        self.save_key_btn.IsEnabled = needs_key
        self.get_api_key_link.Visibility = (
            Visibility.Visible if needs_key else Visibility.Collapsed)

        if needs_key:
            try:
                from config.settings import T3LabAISettings
                saved = T3LabAISettings().get_api_key(self._KEY_NAME_MAP.get(active, "")) or ""
                self.api_key_box.Text = (saved[:8] + u"...") if len(saved) > 8 else saved
            except Exception:
                pass
        else:
            self.api_key_box.Text = u""

        self._apply_host_panel(active)

    def _update_instant(self):
        """Phase 1 — zero HTTP, cached/local data only."""
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            active = router.get_active_name()
            self._apply_provider_chrome(active)
            self._populate_model_combo(active)
        except Exception as ex:
            logger.debug("_update_instant error: {}".format(ex))

    def _update_probed(self):
        """Phase 2 — called on the UI thread after a background probe completes."""
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            status = router.get_status_instant()
            active = router.get_active_name()

            self._apply_provider_chrome(active)
            self._populate_model_combo(active)

            for name in ("claude", "openai", "deepseek", "ollama", "lmstudio"):
                info = status.get(name, {})
                self._set_status_dot(name, info.get("available", False))
        except Exception as ex:
            logger.debug("_update_probed error: {}".format(ex))

    def _probe_all_async(self):
        if self._probing:
            return
        self._probing = True

        def _bg():
            try:
                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                active = router.get_active_name()

                router.probe_provider(active)
                provider = router.get_active_provider()
                if provider:
                    try:
                        self._models_cache[active] = provider.get_models()
                    except Exception:
                        pass
                self.Dispatcher.Invoke(Action(self._update_probed))

                router.get_status(use_cache=False)
                self.Dispatcher.Invoke(Action(self._update_probed))
            except Exception:
                pass
            finally:
                self._probing = False

        t = Thread(ThreadStart(_bg))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()

    # ─── Generic hint flash ─────────────────────────────────────────────────

    def _flash_hint(self, tb, text=u"✓ Saved", seconds=2.0):
        """Show a short confirmation next to a field, then clear it."""
        try:
            tb.Foreground = _GREEN
            tb.Text = text
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromSeconds(seconds)

            def _clear(s, ev):
                try:
                    tb.Text = u""
                finally:
                    timer.Stop()

            timer.Tick += _clear
            timer.Start()
        except Exception:
            pass

    # ─── TAB: General ───────────────────────────────────────────────────────

    def _load_general_tab(self):
        """Fill display name + action-mode + reasoning toggles from settings."""
        try:
            from config.user_profile import UserProfile
            self.username_box.Text = UserProfile().get_name() or u""
        except Exception:
            pass
        try:
            from config.settings import get_settings
            self._action_guard = True
            self.action_mode_toggle.IsChecked = (
                get_settings().get_action_mode() == 'confirm')
        except Exception:
            pass
        finally:
            self._action_guard = False
        try:
            from config.settings import get_settings
            self._think_guard = True
            self.extended_thinking_toggle.IsChecked = bool(
                get_settings().get_agent_option('extended_thinking', False))
        except Exception:
            pass
        finally:
            self._think_guard = False
        try:
            from config.settings import get_settings
            self._quality_guard = True
            self.quality_mode_toggle.IsChecked = bool(
                get_settings().is_quality_mode_enabled())
        except Exception:
            pass
        finally:
            self._quality_guard = False

    def save_username_clicked(self, sender, e):
        name = (self.username_box.Text or u"").strip()
        if not name:
            return
        try:
            from config.user_profile import UserProfile
            UserProfile().set_name(name)
            self._flash_hint(self.username_saved_hint)
        except Exception as ex:
            logger.debug("save_username_clicked error: {}".format(ex))

    def action_mode_toggled(self, sender, e):
        """Persist 'ask before model edits' (confirm) vs 'auto'."""
        if getattr(self, '_action_guard', False):
            return
        try:
            from config.settings import get_settings
            mode = 'confirm' if self.action_mode_toggle.IsChecked else 'auto'
            get_settings().set_agent_option('action_mode', mode)
        except Exception as ex:
            logger.debug("action_mode_toggled error: {}".format(ex))

    def extended_thinking_toggled(self, sender, e):
        """Persist the Claude extended-thinking (deep reasoning) switch."""
        if getattr(self, '_think_guard', False):
            return
        try:
            from config.settings import get_settings
            get_settings().set_agent_option(
                'extended_thinking',
                bool(self.extended_thinking_toggle.IsChecked))
        except Exception as ex:
            logger.debug("extended_thinking_toggled error: {}".format(ex))

    def quality_mode_toggled(self, sender, e):
        """Persist the Opus-parity (maximum quality) switch.

        When on, the Claude provider defaults to the most capable model,
        forces deep reasoning on agent turns, and widens the token ceiling.
        """
        if getattr(self, '_quality_guard', False):
            return
        try:
            from config.settings import get_settings
            get_settings().set_agent_option(
                'quality_mode',
                bool(self.quality_mode_toggle.IsChecked))
        except Exception as ex:
            logger.debug("quality_mode_toggled error: {}".format(ex))

    def open_data_dir_clicked(self, sender, e):
        """Open %APPDATA%/T3LabAI in Explorer."""
        try:
            import System.Diagnostics
            d = os.path.join(os.environ.get('APPDATA', ''), 'T3LabAI')
            if not os.path.isdir(d):
                os.makedirs(d)
            System.Diagnostics.Process.Start(d)
        except Exception as ex:
            logger.debug("open_data_dir_clicked error: {}".format(ex))

    # ─── TAB: Projects ──────────────────────────────────────────────────────

    def _selected_project_id(self):
        try:
            item = self.project_combo.SelectedItem
            return getattr(item, 'Tag', None) if item is not None else None
        except Exception:
            return None

    def _load_projects_tab(self, select_pid=None):
        """Fill the project combo (edit scope only — activation happens in
        the chat composer's project chip, never here)."""
        try:
            from System.Windows.Controls import ComboBoxItem
            from config.project_store import ProjectStore
            ps = ProjectStore()
            combo = self.project_combo
            try:
                combo.SelectionChanged -= self.project_changed
            except Exception:
                pass
            combo.Items.Clear()
            projects = ps.list_projects()
            want = select_pid or ps.get_active_project_id()
            sel = -1
            for i, meta in enumerate(projects):
                item = ComboBoxItem()
                item.Content = meta['name']
                item.Tag = meta['id']
                combo.Items.Add(item)
                if meta['id'] == want:
                    sel = i
            if projects:
                combo.SelectedIndex = sel if sel >= 0 else 0
            try:
                combo.SelectionChanged += self.project_changed
            except Exception:
                pass
            self._fill_project_edit(self._selected_project_id())
        except Exception as ex:
            logger.debug("_load_projects_tab error: {}".format(ex))

    def _fill_project_edit(self, pid):
        try:
            from config.project_store import ProjectStore
            meta = ProjectStore().get_project(pid) if pid else None
            if meta:
                self.project_edit_panel.Visibility = Visibility.Visible
                self.project_name_box.Text = meta.get('name', u'')
                self.project_instructions_box.Text = meta.get('instructions', u'')
                # Default AI override (applied on activation)
                try:
                    want = meta.get('provider') or ''
                    combo = self.project_provider_combo
                    idx = 0
                    for i in range(combo.Items.Count):
                        if str(combo.Items[i].Tag or '') == want:
                            idx = i
                            break
                    combo.SelectedIndex = idx
                    self.project_model_box.Text = meta.get('model') or u''
                except Exception:
                    pass
                self._update_project_files_status(pid)
                self._render_project_sched(pid)
            else:
                self.project_edit_panel.Visibility = Visibility.Collapsed
        except Exception as ex:
            logger.debug("_fill_project_edit error: {}".format(ex))

    def project_changed(self, sender, e):
        self._fill_project_edit(self._selected_project_id())

    def new_project_clicked(self, sender, e):
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            n = len(ps.list_projects()) + 1
            meta = ps.create_project(u"Project {}".format(n))
            self._load_projects_tab(select_pid=meta['id'])
            try:
                self.project_name_box.Focus()
                self.project_name_box.SelectAll()
            except Exception:
                pass
        except Exception as ex:
            logger.debug("new_project_clicked error: {}".format(ex))

    def project_save_clicked(self, sender, e):
        try:
            pid = self._selected_project_id()
            if not pid:
                return
            prov = None
            try:
                item = self.project_provider_combo.SelectedItem
                if item is not None and item.Tag:
                    prov = str(item.Tag)
            except Exception:
                prov = None
            model = (self.project_model_box.Text or u'').strip() or None
            from config.project_store import ProjectStore
            ProjectStore().update_project(pid, {
                'name': self.project_name_box.Text.strip() or u"Project",
                'instructions': self.project_instructions_box.Text.strip(),
                'provider': prov,
                'model': model,
            })
            self._load_projects_tab(select_pid=pid)
            self._flash_hint(self.project_saved_hint)
        except Exception as ex:
            logger.debug("project_save_clicked error: {}".format(ex))

    def project_delete_clicked(self, sender, e):
        try:
            pid = self._selected_project_id()
            if not pid:
                return
            from config.project_store import ProjectStore
            ps = ProjectStore()
            meta = ps.get_project(pid) or {}
            from System.Windows import MessageBox, MessageBoxButton, MessageBoxResult
            res = MessageBox.Show(
                u"Delete project '{}' (including its index + chat history)?".format(
                    meta.get('name', pid)),
                u"LLMs Setting", MessageBoxButton.YesNo)
            if res != MessageBoxResult.Yes:
                return
            ps.delete_project(pid)
            self._load_projects_tab()
        except Exception as ex:
            logger.debug("project_delete_clicked error: {}".format(ex))

    # ── Projects: context files ──────────────────────────────────────────────

    def _project_files_dir(self, pid):
        """projects/<pid>/files — created on demand."""
        from config.project_store import ProjectStore
        d = os.path.join(ProjectStore().project_dir(pid), 'files')
        if not os.path.isdir(d):
            try:
                os.makedirs(d)
            except Exception:
                pass
        return d

    def _update_project_files_status(self, pid):
        try:
            n = 0
            for _r, _s, _fs in os.walk(self._project_files_dir(pid)):
                n += len(_fs)
                if n > 99:
                    break
            self.project_files_status.Text = u"{} file{} in the knowledge folder".format(
                u"99+" if n > 99 else n, u"" if n == 1 else u"s")
        except Exception as ex:
            logger.debug("_update_project_files_status error: {}".format(ex))

    def project_add_files_clicked(self, sender, e):
        """Copy picked documents into the project's knowledge folder."""
        try:
            pid = self._selected_project_id()
            if not pid:
                return
            clr.AddReference('System.Windows.Forms')
            from System.Windows.Forms import OpenFileDialog, DialogResult
            dlg = OpenFileDialog()
            dlg.Multiselect = True
            dlg.Title = "Add documents to this project"
            dlg.Filter = ("Documents|*.pdf;*.docx;*.txt;*.md;*.csv;*.xlsx|"
                          "All files|*.*")
            if dlg.ShowDialog() != DialogResult.OK:
                return
            import shutil
            dst = self._project_files_dir(pid)
            for f in dlg.FileNames:
                try:
                    shutil.copy2(f, os.path.join(dst, os.path.basename(f)))
                except Exception:
                    pass
            self._update_project_files_status(pid)
            # Editing the ACTIVE project → refresh its RAG index now;
            # other projects get indexed on activation.
            try:
                from config.project_store import ProjectStore
                if ProjectStore().get_active_project_id() == pid:
                    self._kick_knowledge_scan()
            except Exception:
                pass
        except Exception as ex:
            logger.debug("project_add_files_clicked error: {}".format(ex))

    def project_open_folder_clicked(self, sender, e):
        try:
            pid = self._selected_project_id()
            if not pid:
                return
            import System.Diagnostics
            System.Diagnostics.Process.Start(self._project_files_dir(pid))
        except Exception as ex:
            logger.debug("project_open_folder_clicked error: {}".format(ex))

    # ── Projects: scheduled daily prompts ────────────────────────────────────

    def _render_project_sched(self, pid):
        """Rebuild the scheduled-prompt rows for the selected project."""
        try:
            from config.project_store import ProjectStore
            from System.Windows.Controls import (Grid, ColumnDefinition,
                                                 TextBlock)
            from System.Windows import Thickness, GridLength, TextWrapping
            from System.Windows.Input import Cursors

            panel = self.project_sched_panel
            panel.Children.Clear()
            items = (ProjectStore().get_project(pid) or {}) \
                .get('scheduled') or []
            if not items:
                t = TextBlock()
                t.Text = u"No scheduled prompts yet."
                t.FontSize = 11
                t.Foreground = _MUTED
                t.Margin = Thickness(1, 0, 0, 4)
                panel.Children.Add(t)
                return
            for it in items:
                g = Grid()
                g.Margin = Thickness(1, 2, 0, 2)
                g.ColumnDefinitions.Add(ColumnDefinition())
                c1 = ColumnDefinition()
                c1.Width = GridLength.Auto
                g.ColumnDefinitions.Add(c1)
                _pv = u" ".join((it.get('prompt') or u'').split())
                if len(_pv) > 60:
                    _pv = _pv[:59] + u"…"
                lbl = TextBlock()
                lbl.Text = u"{} — {}".format(it.get('time') or u'?', _pv)
                lbl.FontSize = 12
                lbl.Foreground = _brush(82, 82, 91)
                lbl.TextWrapping = TextWrapping.Wrap
                g.Children.Add(lbl)

                x = TextBlock()
                x.Text = u""
                x.FontFamily = System.Windows.Media.FontFamily(
                    "Segoe MDL2 Assets")
                x.FontSize = 10
                x.Foreground = _MUTED
                x.Cursor = Cursors.Hand
                x.Margin = Thickness(10, 2, 2, 0)
                x.ToolTip = u"Remove"

                def _rm(s, ev, _tid=it.get('id'), _pid=pid):
                    try:
                        from config.project_store import ProjectStore as _PS
                        _items = [t2 for t2 in
                                  ((_PS().get_project(_pid) or {})
                                   .get('scheduled') or [])
                                  if t2.get('id') != _tid]
                        _PS().update_project(_pid, {'scheduled': _items})
                    except Exception:
                        pass
                    self._render_project_sched(_pid)

                x.MouseLeftButtonUp += _rm
                Grid.SetColumn(x, 1)
                g.Children.Add(x)
                panel.Children.Add(g)
        except Exception as ex:
            logger.debug("_render_project_sched error: {}".format(ex))

    def project_sched_add_clicked(self, sender, e):
        """Validate + append one scheduled daily prompt."""
        try:
            pid = self._selected_project_id()
            if not pid:
                return
            import re as _re
            import time as _time
            prompt = (self.project_sched_prompt_box.Text or u'').strip()
            m = _re.match(r'^(\d{1,2}):(\d{2})$',
                          (self.project_sched_time_box.Text or u'').strip())
            if not prompt or not m or int(m.group(1)) > 23 \
                    or int(m.group(2)) > 59:
                (self.project_sched_time_box if prompt
                 else self.project_sched_prompt_box).BorderBrush = _RED
                return
            self.project_sched_prompt_box.BorderBrush = _brush(203, 213, 225)
            self.project_sched_time_box.BorderBrush = _brush(203, 213, 225)
            from config.project_store import ProjectStore
            items = list((ProjectStore().get_project(pid) or {})
                         .get('scheduled') or [])
            items.append({
                'id': u's_{}'.format(int(_time.time() * 1000)),
                'prompt': prompt,
                'time': u"{:02d}:{}".format(int(m.group(1)), m.group(2)),
                'enabled': True,
                'last_run': u''})
            ProjectStore().update_project(pid, {'scheduled': items})
            self.project_sched_prompt_box.Text = u''
            self._render_project_sched(pid)
        except Exception as ex:
            logger.debug("project_sched_add_clicked error: {}".format(ex))

    # ─── TAB: Knowledge ─────────────────────────────────────────────────────

    def _load_knowledge_tab(self):
        """Refresh status label, dir rows and the embeddings toggle."""
        if not HAS_KNOWLEDGE:
            try:
                self.knowledge_index_status.Text = u"Knowledge module not available"
            except Exception:
                pass
            return
        try:
            store = get_active_store()
            if store is not None:
                st = store.stats()
                self.knowledge_index_status.Text = u"{} files · {} chunks".format(
                    st['files'], st['chunks'])
            self._refresh_knowledge_dirs_panel()
            try:
                from config.settings import get_settings
                want = bool(get_settings().get_knowledge_option(
                    'embeddings_enabled', True))
                self._embed_guard = True
                self.knowledge_embed_toggle.IsChecked = want
            except Exception:
                pass
            finally:
                self._embed_guard = False
        except Exception as ex:
            logger.debug("_load_knowledge_tab error: {}".format(ex))

    def _refresh_knowledge_dirs_panel(self):
        """Rebuild the knowledge directory rows. UI THREAD."""
        try:
            from System.Windows.Controls import (Border, TextBlock, Grid,
                                                 ColumnDefinition, Button)
            from System.Windows import Thickness, CornerRadius, GridLength

            panel = self.knowledge_dirs_panel
            panel.Children.Clear()

            rows = [(default_knowledge_dir(), False)]
            try:
                from config.settings import get_settings
                for d in get_settings().get_knowledge_dirs():
                    rows.append((d, True))
            except Exception:
                pass

            for path, removable in rows:
                row = Border()
                row.Background = SolidColorBrush(Color.FromRgb(255, 255, 255))
                row.BorderBrush = SolidColorBrush(Color.FromRgb(230, 230, 234))
                row.BorderThickness = Thickness(1)
                row.CornerRadius = CornerRadius(8)
                row.Padding = Thickness(10, 6, 8, 6)
                row.Margin = Thickness(0, 0, 0, 4)

                grid = Grid()
                col_txt = ColumnDefinition()
                col_txt.Width = GridLength(1, System.Windows.GridUnitType.Star)
                col_btn = ColumnDefinition()
                col_btn.Width = GridLength.Auto
                grid.ColumnDefinitions.Add(col_txt)
                grid.ColumnDefinitions.Add(col_btn)

                tb = TextBlock()
                tb.Text = os.path.basename(path.rstrip(u'\\/')) or path
                tb.ToolTip = path
                tb.FontSize = 11.5
                tb.FontFamily = System.Windows.Media.FontFamily("Hanken Grotesk")
                tb.Foreground = SolidColorBrush(Color.FromRgb(82, 82, 91))
                tb.VerticalAlignment = System.Windows.VerticalAlignment.Center
                tb.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
                Grid.SetColumn(tb, 0)
                grid.Children.Add(tb)

                if removable:
                    btn = Button()
                    btn.Content = u"✕"
                    btn.FontSize = 10
                    btn.Width = 20
                    btn.Height = 20
                    btn.Cursor = System.Windows.Input.Cursors.Hand
                    btn.Background = SolidColorBrush(Color.FromArgb(0, 0, 0, 0))
                    btn.BorderThickness = Thickness(0)
                    btn.Foreground = SolidColorBrush(Color.FromRgb(161, 161, 170))
                    btn.ToolTip = u"Remove folder from index"

                    def _make_remove(p):
                        def _remove(sender, e):
                            try:
                                from config.settings import get_settings
                                get_settings().remove_knowledge_dir(p)
                                self._refresh_knowledge_dirs_panel()
                                self._kick_knowledge_scan()
                            except Exception as rex:
                                logger.debug("remove dir error: {}".format(rex))
                        return _remove
                    btn.Click += _make_remove(path)
                    Grid.SetColumn(btn, 1)
                    grid.Children.Add(btn)

                row.Child = grid
                panel.Children.Add(row)
        except Exception as ex:
            logger.debug("_refresh_knowledge_dirs_panel error: {}".format(ex))

    def _kick_knowledge_scan(self):
        """(Re)scan the active knowledge store on a background thread."""
        if not HAS_KNOWLEDGE or self._kn_scan_busy:
            return
        self._kn_scan_busy = True
        try:
            self.knowledge_index_status.Text = u"Scanning..."
        except Exception:
            pass

        def _scan():
            try:
                store = get_active_store()
                if store is not None:
                    def _prog(name):
                        def _ui(_n=name):
                            try:
                                self.knowledge_index_status.Text = \
                                    u"Index: " + _n[:24]
                            except Exception:
                                pass
                        try:
                            self.Dispatcher.BeginInvoke(Action(_ui))
                        except Exception:
                            pass
                    store.scan(progress_cb=_prog)
                    try:
                        from Intelligence.knowledge.embeddings import (
                            get_default_embedder)
                        emb = get_default_embedder()
                        if emb is not None and emb.is_available():
                            store.embed_pending(emb, budget_sec=120)
                    except Exception:
                        pass
            except Exception as ex:
                logger.debug("knowledge scan error: {}".format(ex))
            finally:
                self._kn_scan_busy = False
                try:
                    self.Dispatcher.Invoke(Action(self._load_knowledge_tab))
                except Exception:
                    pass
        _kt = Thread(ThreadStart(_scan))
        _kt.IsBackground = True
        _kt.SetApartmentState(ApartmentState.STA)
        _kt.Start()

    def add_knowledge_dir_clicked(self, sender, e):
        """Pick a folder to add to the knowledge index. UI THREAD."""
        try:
            clr.AddReference('System.Windows.Forms')
            from System.Windows.Forms import FolderBrowserDialog, DialogResult
            dlg = FolderBrowserDialog()
            dlg.Description = "Chon thu muc tai lieu (PDF/TXT/MD) de index"
            if dlg.ShowDialog() == DialogResult.OK and dlg.SelectedPath:
                from config.settings import get_settings
                get_settings().add_knowledge_dir(dlg.SelectedPath)
                self._refresh_knowledge_dirs_panel()
                self._kick_knowledge_scan()
        except Exception as ex:
            logger.debug("add_knowledge_dir_clicked error: {}".format(ex))

    def reindex_clicked(self, sender, e):
        self._kick_knowledge_scan()

    def knowledge_embed_toggled(self, sender, e):
        """Persist the semantic-search switch; pull the embed model when
        first enabled (background — ~270 MB)."""
        if getattr(self, '_embed_guard', False):
            return
        try:
            on = bool(self.knowledge_embed_toggle.IsChecked)
            from config.settings import get_settings
            get_settings().set_knowledge_option('embeddings_enabled', on)
            if not on:
                return

            def _ensure():
                try:
                    from Intelligence.knowledge.embeddings import (
                        get_default_embedder)
                    emb = get_default_embedder()
                    if emb is None:
                        return
                    if not emb.is_available():
                        if not emb.ensure_model():
                            def _fail():
                                try:
                                    self.knowledge_index_status.Text = (
                                        u"Embed model unavailable — check Ollama")
                                except Exception:
                                    pass
                            self.Dispatcher.Invoke(Action(_fail))
                            return
                    store = get_active_store()
                    if store is not None:
                        store.embed_pending(emb, budget_sec=300)
                    self.Dispatcher.Invoke(Action(self._load_knowledge_tab))
                except Exception as ex2:
                    logger.debug("embed enable error: {}".format(ex2))
            _et = Thread(ThreadStart(_ensure))
            _et.IsBackground = True
            _et.SetApartmentState(ApartmentState.STA)
            _et.Start()
        except Exception as ex:
            logger.debug("knowledge_embed_toggled error: {}".format(ex))

    # ─── TAB: Skills ────────────────────────────────────────────────────────

    def _load_skills_tab(self):
        """Rebuild the skills rows (name + enable toggle). UI THREAD."""
        try:
            from System.Windows.Controls import (Border, TextBlock, Grid,
                                                 ColumnDefinition, CheckBox)
            from System.Windows import Thickness, CornerRadius, GridLength
            from Intelligence.skills_engine import get_skills_engine

            panel = self.skills_list_panel
            panel.Children.Clear()
            skills = get_skills_engine().all_skills()
            if not skills:
                tb = TextBlock()
                tb.Text = u"No skills yet."
                tb.FontSize = 11.5
                tb.Foreground = SolidColorBrush(Color.FromRgb(161, 161, 170))
                tb.FontFamily = System.Windows.Media.FontFamily("Hanken Grotesk")
                panel.Children.Add(tb)
                return

            for meta in skills:
                row = Border()
                row.Background = SolidColorBrush(Color.FromRgb(255, 255, 255))
                row.BorderBrush = SolidColorBrush(Color.FromRgb(230, 230, 234))
                row.BorderThickness = Thickness(1)
                row.CornerRadius = CornerRadius(8)
                row.Padding = Thickness(10, 6, 10, 6)
                row.Margin = Thickness(0, 0, 0, 4)
                row.ToolTip = meta.get('description', '')

                grid = Grid()
                col_txt = ColumnDefinition()
                col_txt.Width = GridLength(1, System.Windows.GridUnitType.Star)
                col_tgl = ColumnDefinition()
                col_tgl.Width = GridLength.Auto
                grid.ColumnDefinitions.Add(col_txt)
                grid.ColumnDefinitions.Add(col_tgl)

                tb = TextBlock()
                tb.Text = meta.get('name', meta['id'])
                tb.FontSize = 11.5
                tb.FontFamily = System.Windows.Media.FontFamily("Hanken Grotesk")
                tb.Foreground = SolidColorBrush(Color.FromRgb(82, 82, 91))
                tb.VerticalAlignment = System.Windows.VerticalAlignment.Center
                tb.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
                Grid.SetColumn(tb, 0)
                grid.Children.Add(tb)

                cb = CheckBox()
                cb.IsChecked = bool(meta.get('enabled', True))
                cb.VerticalAlignment = System.Windows.VerticalAlignment.Center
                try:
                    cb.Style = self.FindResource("T3ToggleSwitch")
                    cb.LayoutTransform = System.Windows.Media.ScaleTransform(0.7, 0.7)
                except Exception:
                    pass

                def _make_toggle(sid, box):
                    def _toggled(sender, e):
                        try:
                            from Intelligence.skills_engine import get_skills_engine
                            get_skills_engine().set_enabled(
                                sid, bool(box.IsChecked))
                        except Exception as tex:
                            logger.debug("skill toggle error: {}".format(tex))
                    return _toggled
                handler = _make_toggle(meta['id'], cb)
                cb.Checked += handler
                cb.Unchecked += handler
                Grid.SetColumn(cb, 1)
                grid.Children.Add(cb)

                row.Child = grid
                panel.Children.Add(row)
        except Exception as ex:
            logger.debug("_load_skills_tab error: {}".format(ex))

    def refresh_skills_clicked(self, sender, e):
        """Rescan skill folders and rebuild the list."""
        try:
            from Intelligence.skills_engine import get_skills_engine
            get_skills_engine().scan()
            self._load_skills_tab()
        except Exception as ex:
            logger.debug("refresh_skills_clicked error: {}".format(ex))

    def open_skills_dir_clicked(self, sender, e):
        """Open the user skills folder in Explorer."""
        try:
            from Intelligence.skills_engine import _user_skills_dir
            import System.Diagnostics
            System.Diagnostics.Process.Start(_user_skills_dir())
        except Exception as ex:
            logger.debug("open_skills_dir_clicked error: {}".format(ex))


def show_llm_setting_dialog():
    """Show the LLMs Setting dialog."""
    LLMSettingWindow().ShowDialog()
