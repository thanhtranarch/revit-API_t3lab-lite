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

from Intelligence.llm_provider import BaseLLMProvider, http_post, http_get, http_get_auth

DEFAULT_HOST = "http://localhost:1234"


class LMStudioProvider(BaseLLMProvider):
    """Adapter for LM Studio (OpenAI-compatible local server)."""

    NAME            = "lmstudio"
    DISPLAY_NAME    = "LM Studio"
    SUPPORTS_VISION = False

    def __init__(self):
        self._model = None        # None → use whatever is loaded in LM Studio
        self._active_host = None  # last host that actually responded
        self._api_prefix = "/v1" # discovered API prefix ("/v1" or "/api/v1")

    def _configured_host(self):
        """Return the normalized host from settings (or the default)."""
        try:
            from config.settings import T3LabAISettings
            host = T3LabAISettings().get_api_key("LMStudio_Host")
        except Exception:
            host = None
        return self._normalize_host(host)

    def _get_host(self):
        """Host used for chat — prefer the one that last responded to a probe."""
        return self._active_host or self._configured_host()

    def _get_api_key(self):
        """Return the stored LM Studio API token, or None."""
        try:
            from config.settings import T3LabAISettings
            key = T3LabAISettings().get_api_key("LMStudio")
            return key if key else None
        except Exception:
            return None

    def _auth_headers(self):
        """Build Authorization header dict.

        LM Studio sometimes requires an Authorization header even for local
        requests when its security settings are enabled. We provide 'lm-studio'
        as a default fallback token if none is configured.
        """
        key = self._get_api_key() or "lm-studio"
        return {"Authorization": "Bearer " + key}

    def _candidate_hosts(self):
        """Hosts to try, in order: configured → 127.0.0.1 → localhost (deduped).
        LM Studio commonly serves on localhost AND a LAN IP; trying a couple of
        defaults means a running server is detected even if the host field is
        blank or slightly off.
        """
        out = []
        for h in (self._configured_host(), "http://127.0.0.1:1234", DEFAULT_HOST):
            if h and h not in out:
                out.append(h)
        return out

    @staticmethod
    def _normalize_host(host):
        """Return the server ROOT, tolerant of how the user typed the host.

        We always append "/v1/..." or "/api/v1/..." ourselves, so the stored
        host must NOT already end in those — otherwise requests double up.
        Accept "http://localhost:1234", "…/v1", and "…/api/v1".
        """
        if not host:
            return DEFAULT_HOST
        host = host.strip().rstrip("/")
        low = host.lower()
        if low.endswith("/api/v1"):
            host = host[:-7].rstrip("/")
        elif low.endswith("/v1"):
            host = host[:-3].rstrip("/")
        return host or DEFAULT_HOST

    # ── Health & discovery ─────────────────────────────────────────────────────

    def _probe_models(self):
        """Try each candidate host; return (host, [model_ids]) for the first that
        responds with a valid model list, else (None, []).

        LM Studio exposes two API flavours depending on configuration:
          • OpenAI-compatible:  /v1/models, /v1/chat/completions
          • LM Studio native:   /api/v1/models, /api/v1/chat
        We probe both paths on each candidate host so the health check
        succeeds regardless of which mode the user has enabled.
        """
        _PATHS = ("/v1/models", "/api/v1/models")
        headers = self._auth_headers()
        for host in self._candidate_hosts():
            for path in _PATHS:
                try:
                    url = host + path
                    # LM Studio requires authenticated GET if security is enabled
                    resp = http_get_auth(url, headers=headers, timeout_ms=800)
                    if not resp:
                        continue
                    data = json.loads(resp)
                    ids = [m.get("id", "") for m in data.get("data", []) if m.get("id")]
                    if ids:
                        self._active_host = host   # remember for chat()
                        # Remember which API prefix worked so chat() uses the
                        # correct completions endpoint later.
                        self._api_prefix = path.rsplit("/models", 1)[0]  # "/v1" or "/api/v1"
                        return host, ids
                except Exception:
                    continue
        return None, []

    def check_health(self):
        """Return True if a reachable LM Studio has at least one model loaded."""
        host, ids = self._probe_models()
        return bool(ids)

    def get_models(self):
        """Return list of model IDs currently loaded in LM Studio."""
        host, ids = self._probe_models()
        return ids

    def get_active_model(self):
        """User's choice, else the first model the server actually reports.

        Returns None when the server is unreachable or has nothing loaded —
        never a made-up placeholder name.
        """
        if self._model:
            return self._model
        models = self.get_models()
        return models[0] if models else None

    def set_model(self, model_name):
        self._model = model_name
        return True

    def reload_credentials(self):
        """No key needed, but a host may have changed — re-probe on next call."""
        self._active_host = None
        self._api_prefix = "/v1"

    def invalidate_models_cache(self):
        """LM Studio always fetches live; just forget the remembered host."""
        self._active_host = None
        self._api_prefix = "/v1"

    # ── Chat ──────────────────────────────────────────────────────────────────

    def _get_chat_endpoint(self):
        """Return the chat completions URL using the discovered API prefix.

        /v1       → /v1/chat/completions  (OpenAI-compatible)
        /api/v1   → /api/v1/chat          (LM Studio native)
        """
        prefix = getattr(self, "_api_prefix", "/v1")
        if prefix == "/api/v1":
            return self._get_host() + "/api/v1/chat"
        return self._get_host() + "/v1/chat/completions"

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """POST to the chat endpoint (auto-detected during probe)."""
        if isinstance(user_content, list):
            text = self.blocks_to_text(user_content)
        else:
            text = user_content or ""

        model = self._model or self.get_active_model()
        if not model:
            return None   # server unreachable / no model loaded

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
        
        tools = kwargs.get("tools")
        if tools:
            payload["tools"] = tools
            
        tool_choice = kwargs.get("tool_choice")
        if tool_choice:
            payload["tool_choice"] = tool_choice

        response_format = kwargs.get("response_format")
        if response_format:
            payload["response_format"] = response_format

        try:
            # Local CPU/GPU inference on a multi-billion-parameter model can
            # legitimately take minutes — the shared http_post() default
            # (60s, tuned for cloud APIs) was silently killing every slower
            # local generation, indistinguishable from "the model failed".
            resp_text = http_post(self._get_chat_endpoint(), payload,
                                  headers=self._auth_headers(), timeout_ms=180000)
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
            
            tool_calls = msg.get("tool_calls")
            if tool_calls:
                return tool_calls

            return content if content else None
        except Exception:
            return None
