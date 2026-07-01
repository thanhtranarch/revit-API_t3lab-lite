# -*- coding: utf-8 -*-
"""
OpenAI Provider

OpenAI GPT API adapter for the T3Lab LLM router.
Supports GPT-4o, GPT-4o-mini, GPT-4-turbo, and vision input.

Vision format conversion:
  Claude block  → {"type":"image","source":{"type":"base64","media_type":"...","data":"..."}}
  OpenAI block  → {"type":"image_url","image_url":{"url":"data:<mime>;base64,..."}}

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "OpenAI Provider"

import json
import os
import sys

from Intelligence.llm_provider import (BaseLLMProvider, http_post, http_post_stream,
                                       http_get_auth, parse_openai_stream_line,
                                       HAS_HTTP)


# ─── Constants ─────────────────────────────────────────────────────────────────

OPENAI_CHAT_URL   = "https://api.openai.com/v1/chat/completions"
OPENAI_MODELS_URL = "https://api.openai.com/v1/models"

MODEL_DEFAULT = "gpt-4o-mini"
MODEL_VISION  = "gpt-4o"   # upgrade target when image blocks are present

# Shown when the API is unreachable or key not set
FALLBACK_MODELS = [
    "gpt-4o-mini",
    "gpt-4o",
    "gpt-4-turbo",
    "gpt-3.5-turbo",
]

# Substrings that identify non-chat models to exclude from the list
_EXCLUDE_SUBSTRINGS = (
    "embedding", "whisper", "tts", "dall-e", "davinci",
    "babbage", "ada", "curie", "moderation", "realtime",
    "audio", "instruct", "search", "similarity", "edit",
)

# Models that support image input
VISION_CAPABLE = {"gpt-4o", "gpt-4o-mini", "gpt-4-turbo"}


# ─── Provider ──────────────────────────────────────────────────────────────────

class OpenAIProvider(BaseLLMProvider):
    """Adapter for the OpenAI Chat Completions API."""

    NAME            = "openai"
    DISPLAY_NAME    = "GPT (OpenAI)"
    SUPPORTS_VISION = True

    def __init__(self):
        self._model         = MODEL_DEFAULT
        self._cached_models = None   # filled on first successful /v1/models fetch

    # ── Credentials ───────────────────────────────────────────────────────────

    def _get_api_key(self):
        try:
            _ensure_lib_in_path()
            from config.settings import T3LabAISettings
            return T3LabAISettings().get_api_key("OpenAI")
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
        Fetch live model list from OpenAI /v1/models, filtered to chat models.
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
                OPENAI_MODELS_URL,
                {"Authorization": "Bearer " + api_key},
            )
            if text:
                data = json.loads(text)
                ids = [
                    m.get("id", "")
                    for m in data.get("data", [])
                    if m.get("id") and _is_chat_model(m.get("id", ""))
                ]
                # Sort: newest models first via lexicographic sort on IDs
                ids.sort(reverse=True)
                if ids:
                    self._cached_models = ids
                    return list(ids)
        except Exception:
            pass

        return []   # no genuine live result → report unset (gates the model list)

    def get_active_model(self):
        return self._model

    def set_model(self, model_name):
        self._model = model_name
        return True

    # ── Chat ─────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Post a chat request to the OpenAI Chat Completions API.

        user_content may be a string or a list of Claude-format content blocks.
        Image blocks are automatically converted to the OpenAI image_url format.
        If the selected model is not vision-capable but images are present,
        the provider upgrades to MODEL_VISION automatically.
        """
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None

        has_vision = self.has_image_blocks(user_content)

        model = self._model
        if has_vision and model not in VISION_CAPABLE:
            model = MODEL_VISION

        openai_content = self._to_openai_content(user_content)

        msgs = [{"role": "system", "content": system_prompt}]
        for h in (messages or []):
            role    = h.get("role", "user")
            content = h.get("content", "")
            if role not in ("user", "assistant"):
                continue
            if isinstance(content, list):
                content = self.blocks_to_text(content)
            if content:
                msgs.append({"role": role, "content": content})

        msgs.append({"role": "user", "content": openai_content})

        payload = {
            "model":      model,
            "max_tokens": max_tokens,
            "messages":   msgs,
        }

        headers = {"Authorization": "Bearer {}".format(api_key)}

        try:
            resp_text  = http_post(OPENAI_CHAT_URL, payload, headers)
            api_result = json.loads(resp_text)
            return api_result["choices"][0]["message"]["content"].strip()
        except Exception as ex:
            self._debug_log("chat() failed: {}".format(ex))
            return None

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """Stream a GPT response token-by-token via the Chat Completions SSE API."""
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None

        has_vision = self.has_image_blocks(user_content)
        model = self._model
        if has_vision and model not in VISION_CAPABLE:
            model = MODEL_VISION

        openai_content = self._to_openai_content(user_content)

        msgs = [{"role": "system", "content": system_prompt}]
        for h in (messages or []):
            role    = h.get("role", "user")
            content = h.get("content", "")
            if role not in ("user", "assistant"):
                continue
            if isinstance(content, list):
                content = self.blocks_to_text(content)
            if content:
                msgs.append({"role": role, "content": content})
        msgs.append({"role": "user", "content": openai_content})

        payload = {
            "model":      model,
            "max_tokens": max_tokens,
            "messages":   msgs,
            "stream":     True,
        }
        headers = {"Authorization": "Bearer {}".format(api_key)}

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
            http_post_stream(OPENAI_CHAT_URL, payload, headers, _on_line)
            full = u"".join(chunks)
            return full.strip() if full else None
        except Exception as ex:
            # Transport/streaming error — fall back to a blocking call. If
            # that ALSO fails, chat()'s own except-block above logs it, so
            # only the streaming-specific failure needs logging here.
            self._debug_log("chat_stream() failed, falling back to chat(): {}".format(ex))
            return self.chat(messages, system_prompt, user_content, max_tokens, **kwargs)

    # ── Conversion helpers ─────────────────────────────────────────────────────

    @staticmethod
    def _to_openai_content(user_content):
        """Convert Claude-format content blocks to OpenAI format."""
        if not isinstance(user_content, list):
            return user_content

        converted = []
        for block in user_content:
            if not isinstance(block, dict):
                converted.append(block)
                continue
            if block.get("type") == "image":
                source = block.get("source", {})
                if source.get("type") == "base64":
                    mime = source.get("media_type", "image/jpeg")
                    data = source.get("data", "")
                    url  = "data:{};base64,{}".format(mime, data)
                    converted.append({
                        "type":      "image_url",
                        "image_url": {"url": url},
                    })
            else:
                converted.append(block)

        return converted


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _is_chat_model(model_id):
    """Return True if the model ID looks like a chat-capable model."""
    lower = model_id.lower()
    return not any(x in lower for x in _EXCLUDE_SUBSTRINGS)


def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
