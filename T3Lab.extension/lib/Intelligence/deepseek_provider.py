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

from Intelligence.llm_provider import BaseLLMProvider, http_post, http_get_auth


# DeepSeek is OpenAI-compatible. Base URL https://api.deepseek.com works for both
# the bare and /v1 paths; keep /v1 for explicit OpenAI-compat routing.
# See https://platform.deepseek.com/  (API docs / key management)
_BASE_URL        = "https://api.deepseek.com/v1"
# Live models from /v1/models. Known IDs used only as offline fallback:
#   deepseek-chat     → DeepSeek-V3 (non-thinking)
#   deepseek-reasoner → DeepSeek-R1 (thinking)
_FALLBACK_MODELS = ["deepseek-chat", "deepseek-reasoner"]
_DEFAULT_MODEL   = "deepseek-chat"


def _lib_dir():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class DeepSeekProvider(BaseLLMProvider):

    NAME            = "deepseek"
    DISPLAY_NAME    = "DeepSeek"
    SUPPORTS_VISION = False

    def __init__(self):
        self._model         = None
        self._api_key       = None
        self._cached_models = None   # filled on first successful models fetch
        self._load_key()
        self._load_saved_model()

    def _load_key(self):
        try:
            lib = _lib_dir()
            if lib not in sys.path:
                sys.path.insert(0, lib)
            from config.settings import T3LabAISettings
            self._api_key = T3LabAISettings().get_api_key("DeepSeek")
        except Exception:
            pass

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

    # ── Credentials ───────────────────────────────────────────────────────────

    def reload_credentials(self):
        """Re-read API key from settings and clear the model cache."""
        self._load_key()
        self._cached_models = None

    def invalidate_models_cache(self):
        self._cached_models = None

    # ── Health & model discovery ───────────────────────────────────────────────

    def check_health(self):
        """Return True if the API key is set and the models endpoint responds."""
        if not self._api_key:
            return False
        models = self.get_models()
        return len(models) > 0

    def get_models(self):
        """
        Fetch live model list from DeepSeek /v1/models.
        Caches result; falls back to known model IDs on network error.
        """
        if not self._api_key:
            return list(_FALLBACK_MODELS)

        if self._cached_models is not None:
            return list(self._cached_models)

        try:
            text = http_get_auth(
                _BASE_URL + "/models",
                {"Authorization": "Bearer " + self._api_key},
            )
            if text:
                data = json.loads(text)
                ids = [m.get("id", "") for m in data.get("data", []) if m.get("id")]
                if ids:
                    self._cached_models = ids
                    return list(ids)
        except Exception:
            pass

        return list(_FALLBACK_MODELS)

    def get_active_model(self):
        return self._model or _DEFAULT_MODEL

    def set_model(self, model_name):
        self._model = model_name
        return True

    # ── Chat ─────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400):
        if not self._api_key:
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
                {"Authorization": "Bearer " + self._api_key},
            )
            data    = json.loads(resp_text)
            msg     = data.get("choices", [{}])[0].get("message", {})
            content = msg.get("content") or msg.get("reasoning_content") or ""
            content = _re.sub(r"<think>[\s\S]*?</think>", "", content).strip()
            return content if content else None
        except Exception:
            return None
