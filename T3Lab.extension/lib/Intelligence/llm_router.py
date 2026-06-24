# -*- coding: utf-8 -*-
"""
LLM Router

Singleton that manages LLM provider selection, hot-swap, and automatic fallback.

Usage:
    from Intelligence.llm_router import LLMRouter

    router = LLMRouter()
    router.switch_provider("openai")              # hot-swap
    router.set_model("openai", "gpt-4o")          # change model
    text = router.chat(history, sys_prompt, user_content)

    status = router.get_status()
    # {"claude": {"available": True, "model": "claude-haiku-...", ...}, ...}

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "LLM Router"

import os
import sys


# ─── Router ────────────────────────────────────────────────────────────────────

class LLMRouter(object):
    """
    Singleton router for multi-provider LLM access.

    Provider priority / fallback order:
        claude → openai → ollama

    Switching is hot (no session restart required).
    The active provider and per-provider model are persisted in settings.json.
    """

    _instance = None

    # Canonical fallback order — also used as the display order in UI
    FALLBACK_CHAIN = ["claude", "openai", "deepseek", "ollama", "lmstudio"]

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(LLMRouter, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._providers     = {}
        self._active_name   = "claude"
        self._fallback_on   = True
        self._status_cache  = None   # cached result of get_status()
        self._status_ts     = 0.0    # epoch time of last cache fill
        self._initialized   = True

        self._load_providers()
        self._restore_settings()

    # ── Initialisation helpers ─────────────────────────────────────────────────

    def _load_providers(self):
        """Instantiate all provider adapters (failures are silenced)."""
        adapters = [
            ("claude",    "Intelligence.claude_provider",    "ClaudeProvider"),
            ("openai",    "Intelligence.openai_provider",    "OpenAIProvider"),
            ("deepseek",  "Intelligence.deepseek_provider",  "DeepSeekProvider"),
            ("ollama",    "Intelligence.ollama_provider",    "OllamaProvider"),
            ("lmstudio",  "Intelligence.lmstudio_provider",  "LMStudioProvider"),
        ]
        for name, module_path, class_name in adapters:
            try:
                parts  = module_path.split(".")
                mod    = __import__(module_path, fromlist=[parts[-1]])
                cls_   = getattr(mod, class_name)
                self._providers[name] = cls_()
            except Exception:
                pass

    def _restore_settings(self):
        """Load the last-used provider and per-provider model from settings."""
        try:
            _ensure_lib_in_path()
            from config.settings import T3LabAISettings
            s = T3LabAISettings()

            saved_provider = s.get_active_provider()
            if saved_provider and saved_provider in self._providers:
                self._active_name = saved_provider

            for name, provider in self._providers.items():
                saved_model = s.get_provider_model(name)
                if saved_model:
                    try:
                        provider.set_model(saved_model)
                    except Exception:
                        pass

            # Log restored startup configuration
            active_provider = self.get_active_provider()
            active_model = active_provider.get_active_model() if active_provider else None
            s.log_model_usage("STARTUP_RESTORE", self._active_name, active_model)
        except Exception:
            pass

    # ── Provider management ────────────────────────────────────────────────────

    def get_active_name(self):
        """Return the name of the currently active provider."""
        return self._active_name

    def get_active_provider(self):
        """Return the active provider instance, or None."""
        return self._providers.get(self._active_name)

    def switch_provider(self, name, model=None):
        """
        Hot-swap to a different provider.

        Args:
            name (str): provider name — "claude", "openai", or "ollama".
            model (str|None): optional model to activate on the new provider.

        Returns:
            bool: True on success, False if provider is unknown.
        """
        if name not in self._providers:
            return False

        self._active_name = name
        # Note: do NOT invalidate the status cache here. The only thing a switch
        # changes is the "active" flag, which get_status() applies live to the
        # cached snapshot. Invalidating would force a slow full re-probe of every
        # provider on the next status call (e.g. opening the badge menu).

        from config.settings import T3LabAISettings
        s = T3LabAISettings()

        if model is None:
            model = s.get_provider_model(name)

        if model:
            self._providers[name].set_model(model)
        else:
            try:
                model = self._providers[name].get_active_model()
            except Exception:
                pass

        # Persist provider selection
        try:
            s.set_active_provider(name)
        except Exception:
            pass

        # Log provider hot-swap action
        try:
            s.log_model_usage("SWITCH_PROVIDER", name, model)
        except Exception:
            pass

        return True

    def set_model(self, provider_name, model_name):
        """
        Set the model for a specific provider and persist the choice.

        Returns:
            bool: True if the provider exists and accepted the model.
        """
        provider = self._providers.get(provider_name)
        if not provider:
            return False

        ok = provider.set_model(model_name)
        if ok:
            try:
                _ensure_lib_in_path()
                from config.settings import T3LabAISettings
                s = T3LabAISettings()
                s.set_provider_model(provider_name, model_name)
                s.log_model_usage("SET_MODEL", provider_name, model_name)
            except Exception:
                pass

        return ok


    def get_provider(self, name):
        """Return a provider instance by name, or None."""
        return self._providers.get(name)

    def get_provider_names(self):
        """Return list of loaded provider names in FALLBACK_CHAIN order."""
        return [n for n in self.FALLBACK_CHAIN if n in self._providers]

    def enable_fallback(self, enabled=True):
        """Enable or disable automatic fallback when the active provider fails."""
        self._fallback_on = enabled

    # ── Chat ──────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Route a chat request to the active provider and return raw response text.

        If the active provider fails and fallback is enabled, the router tries
        each provider in FALLBACK_CHAIN order until one succeeds.

        Args:
            messages (list): conversation history [{role, content}, ...].
            system_prompt (str): system instruction.
            user_content (str|list): current user input (string or content blocks).
            max_tokens (int): maximum response tokens.

        Returns:
            str | None: raw response text, or None if all providers fail.
        """
        active = self._providers.get(self._active_name)

        # ── Active provider ────────────────────────────────────────────────────
        if active:
            try:
                result = active.chat(messages, system_prompt, user_content, max_tokens, **kwargs)
                if result is not None:
                    return result
            except Exception:
                pass

        if not self._fallback_on:
            return None

        # ── Fallback chain ─────────────────────────────────────────────────────
        for name in self.FALLBACK_CHAIN:
            if name == self._active_name:
                continue
            provider = self._providers.get(name)
            if not provider:
                continue
            try:
                if not provider.check_health():
                    continue
                result = provider.chat(messages, system_prompt, user_content, max_tokens, **kwargs)
                if result is not None:
                    return result
            except Exception:
                continue

        return None

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """
        Route a streaming chat request to the active provider, with fallback.

        on_delta(text_chunk) is invoked for each streamed piece of text. Mirrors
        chat()'s active-then-fallback strategy. Returns the full response text,
        or None if every provider fails.
        """
        active = self._providers.get(self._active_name)

        if active:
            try:
                result = active.chat_stream(
                    messages, system_prompt, user_content, on_delta, max_tokens, **kwargs)
                if result is not None:
                    return result
            except Exception:
                pass

        if not self._fallback_on:
            return None

        for name in self.FALLBACK_CHAIN:
            if name == self._active_name:
                continue
            provider = self._providers.get(name)
            if not provider:
                continue
            try:
                if not provider.check_health():
                    continue
                result = provider.chat_stream(
                    messages, system_prompt, user_content, on_delta, max_tokens, **kwargs)
                if result is not None:
                    return result
            except Exception:
                continue

        return None

    # ── Status ────────────────────────────────────────────────────────────────

    def get_status(self, use_cache=True):
        """
        Return a health snapshot of every loaded provider.

        Results are cached for 30 seconds so repeated calls (e.g. sidebar
        re-renders) are instant. Pass use_cache=False to force a fresh probe.

        Returns:
            dict: {
                "claude": {"available": bool, "model": str|None,
                           "display_name": str, "active": bool,
                           "supports_vision": bool},
                ...
            }
        """
        import time
        now = time.time()
        if use_cache and self._status_cache is not None and (now - self._status_ts) < 30:
            # Return cached snapshot but update the "active" flag live
            snap = {}
            for name, info in self._status_cache.items():
                snap[name] = dict(info)
                snap[name]["active"] = (name == self._active_name)
            return snap

        status = {}
        # Local providers first (instant, no network round-trip), then remote.
        probe_order = self.get_local_provider_names() + self.get_remote_provider_names()
        for name in probe_order:
            provider = self._providers[name]
            try:
                available = provider.check_health()
            except Exception:
                available = False
            try:
                model = provider.get_active_model()
            except Exception:
                model = None
            status[name] = {
                "available":       available,
                "model":           model,
                "display_name":    provider.DISPLAY_NAME,
                "active":          (name == self._active_name),
                "supports_vision": provider.SUPPORTS_VISION,
            }
        self._status_cache = status
        self._status_ts    = now
        return status

    def probe_provider(self, name):
        """
        Probe a SINGLE provider and merge its result into the status cache.

        Much faster than get_status() when you only need to refresh the active
        provider (e.g. right after a hot-swap). Returns the provider's status
        dict or None if the provider is unknown.
        """
        provider = self._providers.get(name)
        if not provider:
            return None
        try:
            available = provider.check_health()
        except Exception:
            available = False
        try:
            model = provider.get_active_model()
        except Exception:
            model = None
        info = {
            "available":       available,
            "model":           model,
            "display_name":    provider.DISPLAY_NAME,
            "active":          (name == self._active_name),
            "supports_vision": provider.SUPPORTS_VISION,
        }
        # Merge into cache so subsequent get_status(use_cache=True) sees it
        if self._status_cache is None:
            self._status_cache = {}
        self._status_cache[name] = dict(info)
        return info

    def get_local_provider_names(self):
        """Return loaded local (no-API-key) provider names — probed first for speed."""
        return [n for n in ("ollama", "lmstudio") if n in self._providers]

    def get_remote_provider_names(self):
        """Return loaded remote (API-key) provider names."""
        return [n for n in self.FALLBACK_CHAIN
                if n in self._providers and n not in ("ollama", "lmstudio")]

    def invalidate_status_cache(self):
        """Force the next get_status() call to do a live probe."""
        self._status_ts = 0.0

    def get_display_label(self):
        """
        Return a short label for the UI badge (e.g. "GPT-4o-mini", "LOCAL").
        Mirrors the badge logic that was previously in script.py.
        """
        provider = self._providers.get(self._active_name)
        if not provider:
            return "OFFLINE"

        name = self._active_name
        try:
            model = provider.get_active_model() or ""
        except Exception:
            model = ""

        if name == "ollama":
            return "LOCAL"
        if name == "lmstudio":
            short = model.split("/")[-1] if "/" in model else model
            return short[:14] if short else "LM Studio"
        if name == "claude":
            # Shorten the model name for display
            if "haiku" in model:
                return "Haiku"
            if "sonnet" in model:
                return "Sonnet"
            if "opus" in model:
                return "Opus"
            return "Claude"
        if name == "openai":
            if "mini" in model:
                return "GPT-4o mini"
            if "4o" in model:
                return "GPT-4o"
            if "turbo" in model:
                return "GPT-4 Turbo"
            return "GPT"
        if name == "deepseek":
            if "reasoner" in model:
                return "DS Reasoner"
            return "DeepSeek"

        return name.upper()


# ─── Module-level singleton accessor ──────────────────────────────────────────

def get_router():
    """Return the global LLMRouter singleton."""
    return LLMRouter()


# ─── Path helper ───────────────────────────────────────────────────────────────

def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
