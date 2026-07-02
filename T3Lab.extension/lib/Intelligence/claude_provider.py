# -*- coding: utf-8 -*-
"""
Claude Provider

Anthropic Claude API adapter for the T3Lab LLM router.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Claude Provider"

import json
import os
import sys

from Intelligence.llm_provider import (BaseLLMProvider, http_post, http_post_stream,
                                       http_get_auth, parse_anthropic_stream_line,
                                       HAS_HTTP)


# ─── Constants ─────────────────────────────────────────────────────────────────

CLAUDE_API_URL    = "https://api.anthropic.com/v1/messages"
CLAUDE_MODELS_URL = "https://api.anthropic.com/v1/models"
ANTHROPIC_API_VER = "2023-06-01"

MODEL_VISION  = "claude-sonnet-4-6"
MODEL_TEXT    = "claude-haiku-4-5-20251001"

# Shown when the API is unreachable or key not set
FALLBACK_MODELS = [
    "claude-haiku-4-5-20251001",
    "claude-sonnet-4-6",
    "claude-opus-4-8",
]


# ─── Provider ──────────────────────────────────────────────────────────────────

class ClaudeProvider(BaseLLMProvider):
    """Adapter for the Anthropic Claude API."""

    NAME            = "claude"
    DISPLAY_NAME    = "Claude (Anthropic)"
    SUPPORTS_VISION = True

    def __init__(self):
        self._model         = None   # None → auto-select text vs vision model
        self._cached_models = None   # filled on first successful /v1/models fetch

    # ── Credentials ───────────────────────────────────────────────────────────

    def _get_api_key(self):
        try:
            _ensure_lib_in_path()
            from config.settings import T3LabAISettings
            return T3LabAISettings().get_api_key("Claude")
        except Exception:
            return None

    def reload_credentials(self):
        """Clear model cache so next get_models() re-fetches live data."""
        self._cached_models = None

    def invalidate_models_cache(self):
        self._cached_models = None

    # ── Health & model discovery ───────────────────────────────────────────────

    def check_health(self):
        """Return True if the API key is set and the models endpoint responds."""
        if not HAS_HTTP:
            return False
        if not self._get_api_key():
            return False
        models = self.get_models()
        return len(models) > 0

    def get_models(self):
        """
        Fetch live model list from Anthropic /v1/models.
        Returns [] if there is no key or the live fetch fails, so callers can
        treat a non-empty result as a verified connection.
        """
        if self._cached_models is not None:
            return list(self._cached_models)

        api_key = self._get_api_key()
        if not api_key:
            return []   # no genuine live result → report unset (gates the model list)

        try:
            text = http_get_auth(
                CLAUDE_MODELS_URL,
                {
                    "x-api-key":         api_key,
                    "anthropic-version": ANTHROPIC_API_VER,
                },
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
        return self._model or MODEL_TEXT

    def set_model(self, model_name):
        self._model = model_name
        return True

    # ── Chat ─────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Post a chat request to the Claude API.

        user_content may be a plain string or a list of Claude-format content
        blocks (text + image) for vision requests.
        """
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None

        has_vision = self.has_image_blocks(user_content)

        if self._model:
            model = self._model
        else:
            model = MODEL_VISION if has_vision else MODEL_TEXT

        msgs = list(messages or [])
        msgs.append({"role": "user", "content": user_content})

        payload = {
            "model":      model,
            "max_tokens": max_tokens,
            "system":     system_prompt,
            "messages":   msgs,
        }

        headers = {
            "x-api-key":         api_key,
            "anthropic-version": ANTHROPIC_API_VER,
        }

        try:
            resp_text  = http_post(CLAUDE_API_URL, payload, headers)
            api_result = json.loads(resp_text)
            return api_result["content"][0]["text"].strip()
        except Exception as ex:
            self._debug_log("chat() failed: {}".format(ex))
            return None

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """Stream a Claude response token-by-token via the Messages SSE API."""
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None

        has_vision = self.has_image_blocks(user_content)
        model = self._model if self._model else (MODEL_VISION if has_vision else MODEL_TEXT)

        msgs = list(messages or [])
        msgs.append({"role": "user", "content": user_content})

        payload = {
            "model":      model,
            "max_tokens": max_tokens,
            "system":     system_prompt,
            "messages":   msgs,
            "stream":     True,
        }
        headers = {
            "x-api-key":         api_key,
            "anthropic-version": ANTHROPIC_API_VER,
        }

        chunks = []

        def _on_line(line):
            delta = parse_anthropic_stream_line(line)
            if delta:
                chunks.append(delta)
                if on_delta:
                    try:
                        on_delta(delta)
                    except Exception:
                        pass

        try:
            http_post_stream(CLAUDE_API_URL, payload, headers, _on_line)
            full = u"".join(chunks)
            return full.strip() if full else None
        except Exception:
            # Any streaming/transport error → fall back to a single blocking call.
            return self.chat(messages, system_prompt, user_content, max_tokens, **kwargs)


# ─── Path helper ───────────────────────────────────────────────────────────────

def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
