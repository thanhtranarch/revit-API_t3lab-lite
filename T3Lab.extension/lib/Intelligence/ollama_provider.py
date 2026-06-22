# -*- coding: utf-8 -*-
"""
Ollama Provider

Local Ollama LLM adapter for the T3Lab LLM router.
Reuses local_llm.py for model discovery and the HTTP helpers in llm_provider.py.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Ollama Provider"

import json
import os
import sys

from Intelligence.llm_provider import BaseLLMProvider, http_post, http_get


# ─── Provider ──────────────────────────────────────────────────────────────────

class OllamaProvider(BaseLLMProvider):
    """Adapter for a locally-running Ollama LLM server."""

    NAME            = "ollama"
    DISPLAY_NAME    = "Local LLM (Ollama)"
    SUPPORTS_VISION = False   # most small models don't support vision

    def __init__(self):
        self._model = None   # None → auto-select best installed model
        self._host  = None   # None → read from local_llm.OLLAMA_HOST

    # ── Internal helpers ───────────────────────────────────────────────────────

    def _local_llm(self):
        """Lazy-import local_llm module."""
        try:
            from Intelligence import local_llm
            return local_llm
        except Exception:
            return None

    def _get_host(self):
        mod = self._local_llm()
        return self._host or (mod.OLLAMA_HOST if mod else "http://localhost:11434")

    def _get_timeout(self):
        mod = self._local_llm()
        return mod.TIMEOUT_GEN if mod else 60

    # ── BaseLLMProvider interface ──────────────────────────────────────────────

    def check_health(self):
        """Return True if Ollama is running AND has at least one model installed."""
        try:
            tags = http_get(self._get_host() + "/api/tags")
            if not tags:
                return False
            data = json.loads(tags)
            return len(data.get("models", [])) > 0
        except Exception:
            return False

    def get_models(self):
        try:
            mod = self._local_llm()
            if mod:
                return mod.list_models()
            tags_text = http_get(self._get_host() + "/api/tags")
            if tags_text:
                data = json.loads(tags_text)
                return [m.get("name", "") for m in data.get("models", [])]
        except Exception:
            pass
        return []

    def get_active_model(self):
        if self._model:
            return self._model
        mod = self._local_llm()
        if mod:
            try:
                return mod.get_best_model()
            except Exception:
                pass
        return None

    def set_model(self, model_name):
        self._model = model_name
        return True

    def reload_credentials(self):
        """No-op — Ollama needs no credentials. Clears nothing."""
        pass

    def invalidate_models_cache(self):
        """No-op — Ollama always fetches live from /api/tags."""
        pass

    def set_host(self, host):
        """Override the Ollama server URL (e.g. 'http://192.168.1.10:11434')."""
        self._host = host

    def chat(self, messages, system_prompt, user_content, max_tokens=400):
        """
        Send a chat request to the local Ollama server.

        Vision is not supported — image blocks are stripped and only the text
        portions of user_content are sent.
        """
        # Flatten vision/multi-modal content to plain text
        if isinstance(user_content, list):
            text = self.blocks_to_text(user_content)
        else:
            text = user_content or ""

        model = self.get_active_model()
        if not model:
            return None

        msgs = [{"role": "system", "content": system_prompt}]
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
            "model":    model,
            "messages": msgs,
            "stream":   False,
            "format":   "json",
            "options":  {
                "temperature": 0.0,
                "num_predict": max_tokens,
            },
        }

        try:
            resp_text = http_post(
                self._get_host() + "/api/chat",
                payload,
            )
            data = json.loads(resp_text)
            return data.get("message", {}).get("content", "")
        except Exception:
            return None
