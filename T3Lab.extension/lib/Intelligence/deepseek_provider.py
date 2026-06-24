# -*- coding: utf-8 -*-
"""
DeepSeek Provider

OpenAI-compatible adapter for DeepSeek API.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "DeepSeek Provider"

import os
import sys
import json
import re as _re

from Intelligence.llm_provider import (BaseLLMProvider, http_post, http_post_stream,
                                       http_get_auth, parse_openai_stream_line)


# DeepSeek is OpenAI-compatible. Base URL https://api.deepseek.com works for both
# the bare and /v1 paths; keep /v1 for explicit OpenAI-compat routing.
# See https://api-docs.deepseek.com/  (API docs / key management)
_BASE_URL        = "https://api.deepseek.com/v1"
# Live models from /v1/models.  Known IDs used only as offline fallback:
#   deepseek-v4-flash  → fast, non-thinking  (replaces deepseek-chat 2026/07/24)
#   deepseek-v4-pro    → thinking/reasoning  (replaces deepseek-reasoner 2026/07/24)
_FALLBACK_MODELS = ["deepseek-v4-flash", "deepseek-v4-pro",
                    "deepseek-chat", "deepseek-reasoner"]
_DEFAULT_MODEL   = "deepseek-v4-flash"


def _lib_dir():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class DeepSeekProvider(BaseLLMProvider):

    NAME            = "deepseek"
    DISPLAY_NAME    = "DeepSeek"
    SUPPORTS_VISION = False

    def __init__(self):
        self._model         = None
        self._cached_models = None   # filled on first successful models fetch
        self._load_saved_model()

    # ── Credentials ───────────────────────────────────────────────────────────

    def _get_api_key(self):
        """Read the API key fresh from settings every time.

        This ensures that keys saved via the sidebar are picked up immediately
        without requiring a provider restart (unlike the old cached pattern).
        """
        try:
            lib = _lib_dir()
            if lib not in sys.path:
                sys.path.insert(0, lib)
            from config.settings import T3LabAISettings
            return T3LabAISettings().get_api_key("DeepSeek")
        except Exception:
            return None

    def _load_saved_model(self):
        """Restore the last-used model from settings for instant fast load."""
        try:
            lib = _lib_dir()
            if lib not in sys.path:
                sys.path.insert(0, lib)
            from config.settings import T3LabAISettings
            saved = T3LabAISettings().get_provider_model("deepseek")
            if saved:
                self._model = saved
        except Exception:
            pass

    def reload_credentials(self):
        """Clear the model cache so next get_models() re-fetches live data.

        No key to reload — _get_api_key() always reads fresh from settings.
        """
        self._cached_models = None

    def invalidate_models_cache(self):
        self._cached_models = None

    # ── Health & model discovery ───────────────────────────────────────────────

    def check_health(self):
        """Return True if the API key is set and the models endpoint responds."""
        if not self._get_api_key():
            return False
        models = self.get_models()
        return len(models) > 0

    def get_models(self):
        """
        Fetch live model list from DeepSeek /v1/models.
        Returns [] if there is no key or the live fetch fails, so callers can
        treat a non-empty result as a verified connection.
        """
        if not self._get_api_key():
            return []   # no genuine live result → report unset (gates the model list)

        if self._cached_models is not None:
            return list(self._cached_models)

        try:
            text = http_get_auth(
                _BASE_URL + "/models",
                {"Authorization": "Bearer " + self._get_api_key()},
            )
            if text:
                data = json.loads(text)
                ids = [m.get("id", "") for m in data.get("data", []) if m.get("id")]
                if ids:
                    self._cached_models = ids
                    return list(ids)
        except Exception:
            pass

        return []   # no genuine live result → report unset (gates the model list)

    def get_active_model(self):
        return self._model or _DEFAULT_MODEL

    def set_model(self, model_name):
        self._model = model_name
        return True

    # ── Chat ─────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        api_key = self._get_api_key()
        if not api_key:
            return None

        if isinstance(user_content, list):
            text = self.blocks_to_text(user_content)
        else:
            text = user_content or ""

        model = self.get_active_model()
        msgs  = []
        if system_prompt:
            msgs.append({"role": "system", "content": system_prompt})
        for h in (messages or [])[-8:]:
            role    = h.get("role", "user")
            content = h.get("content", "")
            if role not in ("user", "assistant"):
                continue
            if isinstance(content, list):
                content = self.blocks_to_text(content)
            if content:
                msgs.append({"role": role, "content": content})
        msgs.append({"role": "user", "content": text})

        payload = {
            "model":       model,
            "messages":    msgs,
            "max_tokens":  max_tokens,
            "temperature": 0.3,
            "stream":      False,
        }

        try:
            resp_text = http_post(
                _BASE_URL + "/chat/completions",
                payload,
                {"Authorization": "Bearer " + api_key},
            )
            data    = json.loads(resp_text)
            msg     = data.get("choices", [{}])[0].get("message", {})
            content = msg.get("content") or msg.get("reasoning_content") or ""
            content = _re.sub(r"<think>[\s\S]*?</think>", "", content).strip()
            return content if content else None
        except Exception:
            return None

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """Stream a DeepSeek response token-by-token (OpenAI-compatible SSE)."""
        api_key = self._get_api_key()
        if not api_key:
            return None

        if isinstance(user_content, list):
            text = self.blocks_to_text(user_content)
        else:
            text = user_content or ""

        model = self.get_active_model()
        msgs  = []
        if system_prompt:
            msgs.append({"role": "system", "content": system_prompt})
        for h in (messages or [])[-8:]:
            role    = h.get("role", "user")
            content = h.get("content", "")
            if role not in ("user", "assistant"):
                continue
            if isinstance(content, list):
                content = self.blocks_to_text(content)
            if content:
                msgs.append({"role": role, "content": content})
        msgs.append({"role": "user", "content": text})

        payload = {
            "model":       model,
            "messages":    msgs,
            "max_tokens":  max_tokens,
            "temperature": 0.3,
            "stream":      True,
        }

        chunks = []

        def _on_line(line):
            delta = parse_openai_stream_line(line)
            if delta:
                chunks.append(delta)
                if on_delta:
                    try:
                        on_delta(delta)
                    except Exception:
                        pass

        try:
            http_post_stream(
                _BASE_URL + "/chat/completions",
                payload,
                {"Authorization": "Bearer " + api_key},
                _on_line,
            )
            full = _re.sub(r"<think>[\s\S]*?</think>", "", u"".join(chunks)).strip()
            return full if full else None
        except Exception:
            return self.chat(messages, system_prompt, user_content, max_tokens, **kwargs)
