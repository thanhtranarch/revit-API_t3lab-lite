# -*- coding: utf-8 -*-
"""
OpenAI Provider

OpenAI GPT API adapter for the T3Lab LLM router.
Models are discovered live from /v1/models after the key verifies; vision input
is supported.

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
                                       openai_chat_agent, openai_agent_tool_results,
                                       HAS_HTTP)


# ─── Constants ─────────────────────────────────────────────────────────────────

OPENAI_CHAT_URL   = "https://api.openai.com/v1/chat/completions"
OPENAI_MODELS_URL = "https://api.openai.com/v1/models"

# NO hardcoded default model: what the account can use comes exclusively from
# the live /v1/models endpoint after the key is verified. The names below are
# only capability/preference HINTS matched against that live list.

# Substrings that identify non-chat models to exclude from the list
_EXCLUDE_SUBSTRINGS = (
    "embedding", "whisper", "tts", "dall-e", "davinci",
    "babbage", "ada", "curie", "moderation", "realtime",
    "audio", "instruct", "search", "similarity", "edit",
)

# Substring hints (capable-first) used ONLY to auto-pick a vision default from
# the live list — never an exact-name whitelist to validate against.
_PREF_VISION = ("gpt-4o", "gpt-4.1", "gpt-4-turbo", "chatgpt")


# ─── Provider ──────────────────────────────────────────────────────────────────

class OpenAIProvider(BaseLLMProvider):
    """Adapter for the OpenAI Chat Completions API."""

    NAME                  = "openai"
    DISPLAY_NAME          = "GPT (OpenAI)"
    SUPPORTS_VISION       = True
    SUPPORTS_NATIVE_TOOLS = True

    def __init__(self):
        self._model         = None   # None → resolve from the live model list
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
        """User's saved choice, else the default from the CACHED live list.

        Deliberately does no HTTP (badge/sidebar call this) — returns None
        until a live model list has been fetched at least once.
        """
        if self._model:
            return self._model
        return self._pick_model(self._cached_models or [], prefer_vision=False)

    def set_model(self, model_name):
        self._model = model_name
        return True

    @staticmethod
    def _pick_model(models, prefer_vision):
        """Pick a model from a LIVE list only — never invent a name."""
        if not models:
            return None
        if prefer_vision:
            for hint in _PREF_VISION:
                for m in models:
                    if hint in m:
                        return m
        for m in models:
            if "mini" in m:          # cheap-first default for plain chat
                return m
        return models[0]

    def _resolve_model(self, has_vision=False):
        """Model to call: the user's explicit choice ALWAYS wins — it came from
        the vendor's own live list, so it is never second-guessed against a
        capability whitelist. Only when no model was picked do we auto-select
        from the LIVE list (vision-preferred if images are present; may fetch)."""
        if self._model:
            return self._model
        return self._pick_model(self.get_models(), prefer_vision=has_vision)

    # ── Chat ─────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Post a chat request to the OpenAI Chat Completions API.

        user_content may be a string or a list of Claude-format content blocks.
        Image blocks are automatically converted to the OpenAI image_url format.
        When no model is selected and images are present, a vision-preferred
        default is auto-picked from the LIVE model list.
        """
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None

        has_vision = self.has_image_blocks(user_content)

        model = self._resolve_model(has_vision)
        if not model:
            return None   # key not verified / vendor reported no models

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
            self._record_error(u"chat() failed: {}".format(ex))
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
        model = self._resolve_model(has_vision)
        if not model:
            return None   # key not verified / vendor reported no models

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

    # ── Agentic chat (native tool calling, blocking) ──────────────────────────

    def chat_agent(self, system_prompt, messages, tools,
                   on_delta=None, max_tokens=1500, **kwargs):
        """One agentic turn via the `tools` parameter. Blocking (no SSE) —
        the agent loop surfaces the text through on_text_delta itself."""
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None
        model = self._resolve_model(False)
        if not model:
            return None
        try:
            self._clear_error()
            return openai_chat_agent(
                OPENAI_CHAT_URL,
                {"Authorization": "Bearer {}".format(api_key)},
                model,
                system_prompt, messages, tools, max_tokens)
        except Exception as ex:
            self._record_error(u"chat_agent() failed: {}".format(ex))
            return None

    agent_tool_results = staticmethod(openai_agent_tool_results)

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
