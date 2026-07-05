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

# NO hardcoded model names: the models an account can use come exclusively
# from the live /v1/models endpoint after the key is verified. Defaults are
# picked from that live list by substring preference only.
_PREF_TEXT   = ("haiku", "sonnet", "opus")   # cheap-first for plain chat
_PREF_VISION = ("sonnet", "opus", "haiku")   # capable-first for images/agent


# ─── Provider ──────────────────────────────────────────────────────────────────

class ClaudeProvider(BaseLLMProvider):
    """Adapter for the Anthropic Claude API."""

    NAME                  = "claude"
    DISPLAY_NAME          = "Claude (Anthropic)"
    SUPPORTS_VISION       = True
    SUPPORTS_NATIVE_TOOLS = True

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
        """User's saved choice, else the default from the CACHED live list.

        Deliberately does no HTTP (badge/sidebar call this) — returns None
        until a live model list has been fetched at least once.
        """
        if self._model:
            return self._model
        return self._pick_model(self._cached_models or [], _PREF_TEXT)

    def set_model(self, model_name):
        self._model = model_name
        return True

    @staticmethod
    def _pick_model(models, prefs):
        """Pick a model from a LIVE list by substring preference order."""
        if not models:
            return None
        for p in prefs:
            for m in models:
                if p in m:
                    return m
        return models[0]

    def _default_model(self, prefer_vision=False):
        """Default model resolved against the live /v1/models list (may fetch)."""
        return self._pick_model(
            self.get_models(), _PREF_VISION if prefer_vision else _PREF_TEXT)

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

        model = self._model or self._default_model(prefer_vision=has_vision)
        if not model:
            return None   # key not verified / vendor reported no models

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
            self._record_error(u"chat() failed: {}".format(ex))
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
        model = self._model or self._default_model(prefer_vision=has_vision)
        if not model:
            return None   # key not verified / vendor reported no models

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

    # ── Agentic chat (native tool calling) ────────────────────────────────────

    def _agent_payload(self, model, system_prompt, messages, tools, max_tokens, stream):
        """Build a Messages API payload with prompt caching on system + tools.

        cache_control markers go on the system block and the LAST tool — the
        whole (tools + system) prefix is then cached server-side, so the ~75
        tool schemas cost input tokens once per 5-minute window instead of on
        every iteration of the agent loop.
        """
        cached_tools = list(tools or [])
        if cached_tools:
            last = dict(cached_tools[-1])          # copy — tool list is shared/cached
            last["cache_control"] = {"type": "ephemeral"}
            cached_tools = cached_tools[:-1] + [last]

        payload = {
            "model":      model,
            "max_tokens": max_tokens,
            "system":     [{"type": "text", "text": system_prompt,
                            "cache_control": {"type": "ephemeral"}}],
            "messages":   list(messages or []),
        }
        if cached_tools:
            payload["tools"] = cached_tools
        if stream:
            payload["stream"] = True
        return payload

    def chat_agent(self, system_prompt, messages, tools,
                   on_delta=None, max_tokens=1500, **kwargs):
        """One agentic turn: streams text deltas, collects tool_use blocks.

        Returns {"text", "tool_calls", "assistant_msg", "stop_reason"} or None.
        `messages` must already end with the latest user / tool_result turn.
        """
        if not HAS_HTTP:
            return None
        api_key = self._get_api_key()
        if not api_key:
            return None

        # Agent turns favour a more capable model when the user hasn't picked
        # one — resolved from the LIVE model list, never a hardcoded name.
        model = self._model or self._default_model(prefer_vision=True)
        if not model:
            return None

        headers = {
            "x-api-key":         api_key,
            "anthropic-version": ANTHROPIC_API_VER,
        }

        # ── Streaming attempt ────────────────────────────────────────────────
        state = {
            "blocks":      [],     # finalized content blocks, in order
            "cur":         None,   # block being streamed
            "stop_reason": None,
        }

        def _on_line(line):
            if not line:
                return
            line = line.strip()
            if not line.startswith("data:"):
                return
            data = line[5:].strip()
            if not data or data == "[DONE]":
                return
            try:
                obj = json.loads(data)
            except Exception:
                return
            etype = obj.get("type")

            if etype == "content_block_start":
                cb = obj.get("content_block", {}) or {}
                if cb.get("type") == "tool_use":
                    state["cur"] = {"type": "tool_use", "id": cb.get("id", ""),
                                    "name": cb.get("name", ""), "parts": []}
                else:
                    state["cur"] = {"type": "text", "parts": []}

            elif etype == "content_block_delta":
                delta = obj.get("delta", {}) or {}
                cur   = state["cur"]
                if cur is None:
                    return
                if delta.get("type") == "text_delta":
                    txt = delta.get("text") or u""
                    cur["parts"].append(txt)
                    if on_delta and txt:
                        try:
                            on_delta(txt)
                        except Exception:
                            pass
                elif delta.get("type") == "input_json_delta":
                    cur["parts"].append(delta.get("partial_json") or u"")

            elif etype == "content_block_stop":
                cur = state["cur"]
                if cur is not None:
                    state["blocks"].append(cur)
                    state["cur"] = None

            elif etype == "message_delta":
                d = obj.get("delta", {}) or {}
                if d.get("stop_reason"):
                    state["stop_reason"] = d["stop_reason"]

        streamed_ok = False
        try:
            self._clear_error()
            payload = self._agent_payload(model, system_prompt, messages, tools,
                                          max_tokens, stream=True)
            http_post_stream(CLAUDE_API_URL, payload, headers, _on_line,
                             timeout_ms=180000)
            streamed_ok = bool(state["blocks"]) or state["stop_reason"] is not None
        except Exception as ex:
            self._record_error(u"chat_agent stream failed: {}".format(ex))

        if streamed_ok:
            return self._agent_result_from_blocks(state["blocks"],
                                                  state["stop_reason"])

        # ── Blocking fallback ────────────────────────────────────────────────
        try:
            payload   = self._agent_payload(model, system_prompt, messages, tools,
                                            max_tokens, stream=False)
            resp_text = http_post(CLAUDE_API_URL, payload, headers,
                                  timeout_ms=180000)
            api_result = json.loads(resp_text)
            blocks = []
            for b in api_result.get("content", []) or []:
                if b.get("type") == "text":
                    blocks.append({"type": "text", "parts": [b.get("text", "")]})
                elif b.get("type") == "tool_use":
                    blocks.append({"type": "tool_use", "id": b.get("id", ""),
                                   "name": b.get("name", ""),
                                   "parts": [json.dumps(b.get("input") or {})]})
            return self._agent_result_from_blocks(
                blocks, api_result.get("stop_reason"))
        except Exception as ex:
            self._record_error(u"chat_agent() failed: {}".format(ex))
            return None

    @staticmethod
    def _agent_result_from_blocks(blocks, stop_reason):
        """Assemble the uniform chat_agent result from streamed blocks."""
        texts      = []
        tool_calls = []
        content    = []
        for b in blocks:
            joined = u"".join(b.get("parts", []))
            if b["type"] == "text":
                if joined:
                    texts.append(joined)
                    content.append({"type": "text", "text": joined})
            else:  # tool_use
                try:
                    args = json.loads(joined) if joined.strip() else {}
                except Exception:
                    args = {}
                tool_calls.append({"id": b.get("id", ""),
                                   "name": b.get("name", ""), "args": args})
                content.append({"type": "tool_use", "id": b.get("id", ""),
                                "name": b.get("name", ""), "input": args})
        if not content:
            content = [{"type": "text", "text": u""}]
        return {
            "text":          u"\n".join(texts),
            "tool_calls":    tool_calls,
            "assistant_msg": {"role": "assistant", "content": content},
            "stop_reason":   stop_reason or ("tool_use" if tool_calls else "end_turn"),
        }

    @staticmethod
    def agent_tool_results(tool_calls, result_strs):
        """Anthropic format: ONE user message holding all tool_result blocks.

        `result_strs` are pre-serialized JSON strings (agent_loop truncates
        them). Every tool_use id must be answered — missing results
        (cancelled run) are padded so the transcript stays valid.
        """
        blocks = []
        for i, tc in enumerate(tool_calls):
            res = result_strs[i] if i < len(result_strs) else u'{"cancelled": true}'
            blocks.append({
                "type":        "tool_result",
                "tool_use_id": tc.get("id", ""),
                "content":     res,
            })
        return [{"role": "user", "content": blocks}]


# ─── Path helper ───────────────────────────────────────────────────────────────

def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
