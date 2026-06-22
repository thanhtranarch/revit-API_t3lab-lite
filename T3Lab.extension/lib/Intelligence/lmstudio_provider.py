# -*- coding: utf-8 -*-
"""
LM Studio Provider

Adapter for LM Studio local server (OpenAI-compatible API at localhost:1234).
No API key required — just start LM Studio and load a model.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "LM Studio Provider"

import json

from Intelligence.llm_provider import BaseLLMProvider, http_post, http_get

DEFAULT_HOST = "http://localhost:1234"


class LMStudioProvider(BaseLLMProvider):
    """Adapter for LM Studio (OpenAI-compatible local server)."""

    NAME            = "lmstudio"
    DISPLAY_NAME    = "LM Studio"
    SUPPORTS_VISION = False

    def __init__(self):
        self._model = None   # None → use whatever is loaded in LM Studio

    def _get_host(self):
        try:
            from config.settings import T3LabAISettings
            host = T3LabAISettings().get_api_key("LMStudio_Host")
            return host.rstrip("/") if host else DEFAULT_HOST
        except Exception:
            return DEFAULT_HOST

    # ── Health & discovery ─────────────────────────────────────────────────────

    def check_health(self):
        """Return True if LM Studio is running and has at least one model loaded."""
        try:
            resp = http_get(self._get_host() + "/v1/models")
            if not resp:
                return False
            data = json.loads(resp)
            return len(data.get("data", [])) > 0
        except Exception:
            return False

    def get_models(self):
        """Return list of model IDs currently loaded in LM Studio."""
        try:
            resp = http_get(self._get_host() + "/v1/models")
            if not resp:
                return []
            data = json.loads(resp)
            return [m.get("id", "") for m in data.get("data", []) if m.get("id")]
        except Exception:
            return []

    def get_active_model(self):
        if self._model:
            return self._model
        models = self.get_models()
        return models[0] if models else "local-model"

    def set_model(self, model_name):
        self._model = model_name
        return True

    def reload_credentials(self):
        """No-op — LM Studio needs no credentials. Clears nothing."""
        pass

    def invalidate_models_cache(self):
        """No-op — LM Studio always fetches live from /v1/models."""
        pass

    # ── Chat ──────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400):
        """POST to /v1/chat/completions (OpenAI format)."""
        if isinstance(user_content, list):
            text = self.blocks_to_text(user_content)
        else:
            text = user_content or ""

        model = self._model or self.get_active_model() or "local-model"

        msgs = []
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
            resp_text = http_post(self._get_host() + "/v1/chat/completions", payload)
            data = json.loads(resp_text)
            msg = data.get("choices", [{}])[0].get("message", {})

            # Standard content field
            content = msg.get("content") or ""

            # Thinking models (Qwen3, DeepSeek-R1, etc.) may return the actual
            # answer in "reasoning_content" when "content" is empty, or wrap
            # thinking in <think>...</think> tags inside "content".
            if not content.strip():
                content = msg.get("reasoning_content") or ""

            # Strip <think>...</think> blocks — keep only the final answer
            import re as _re
            content = _re.sub(r'<think>[\s\S]*?</think>', '', content).strip()

            return content if content else None
        except Exception:
            return None
