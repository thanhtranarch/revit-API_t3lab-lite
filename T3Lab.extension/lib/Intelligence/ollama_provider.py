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
        self._model = None        # None → auto-select best installed model
        self._host  = None        # None → read from local_llm.OLLAMA_HOST
        self._active_host = None  # last host that actually responded

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
        return self._active_host or self._host or (
            mod.OLLAMA_HOST if mod else "http://localhost:11434")

    def _candidate_hosts(self):
        """Hosts to try, in order: explicit/configured → 127.0.0.1 → localhost."""
        mod = self._local_llm()
        cfg = self._host or (mod.OLLAMA_HOST if mod else None)
        out = []
        for h in (cfg, "http://127.0.0.1:11434", "http://localhost:11434"):
            if h:
                h = h.rstrip("/")
                if h not in out:
                    out.append(h)
        return out

    def _get_timeout(self):
        mod = self._local_llm()
        return mod.TIMEOUT_GEN if mod else 60

    # ── BaseLLMProvider interface ──────────────────────────────────────────────

    def _probe_tags(self):
        """Try each candidate host; return (host, [model_names]) for the first
        reachable Ollama with installed models, else (None, [])."""
        for host in self._candidate_hosts():
            try:
                tags = http_get(host + "/api/tags", timeout_ms=800)
                if not tags:
                    continue
                data = json.loads(tags)
                names = [m.get("name", "") for m in data.get("models", []) if m.get("name")]
                if names:
                    self._active_host = host   # remember for chat/generate
                    return host, names
            except Exception:
                continue
        return None, []

    def check_health(self):
        """Return True if a reachable Ollama has at least one model installed."""
        host, names = self._probe_tags()
        return bool(names)

    def get_models(self):
        host, names = self._probe_tags()
        if names:
            return names
        # Fall back to local_llm's own discovery if direct probing found nothing.
        try:
            mod = self._local_llm()
            if mod:
                return mod.list_models()
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
        self._active_host = None   # re-probe with the new host on next check

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
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
