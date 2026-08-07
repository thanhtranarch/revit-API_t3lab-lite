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
_PREF_TEXT    = ("haiku", "sonnet", "opus")   # cheap-first for plain chat
_PREF_VISION  = ("sonnet", "opus", "haiku")   # capable-first for images/agent
_PREF_FAST    = ("haiku", "sonnet")           # latency-first for classify calls
_PREF_QUALITY = ("opus", "sonnet", "haiku")   # Opus-parity mode: most capable first

# Extended-thinking config for agent turns (agents.extended_thinking).
#
# Adaptive thinking ("type": "adaptive") is the current-generation format —
# required (not just preferred) on Claude Opus 4.7/4.8, Sonnet 5 and Fable 5,
# which reject the older "enabled"/budget_tokens shape with a 400. It is also
# accepted on Opus 4.6 / Sonnet 4.6. Some older-but-still-served models (e.g.
# Opus 4.5, Sonnet 4.5) only understand the legacy "enabled"+budget_tokens
# shape and 400 on "adaptive" instead. Since models are resolved live from
# /v1/models with no hardcoded whitelist (see module docstring), we can't
# know in advance which shape a given account's selected model wants — try
# adaptive first (the modern default for every current-generation model) and
# fall back to the legacy shape only if the API rejects it. The outcome is
# cached per-model so the fallback round-trip only happens once.
_THINK_BUDGET = 3000
_THINK_EFFORT = "high"
_THINK_BETA   = "interleaved-thinking-2025-05-14"   # legacy-mode only; GA (no header needed) on 4.6+ adaptive thinking

# Opus-parity mode: agent turns get a wider answer ceiling so rich replies
# (multi-step summaries, tables, code) never truncate. Utility calls are
# unaffected — they pin a fast model + tiny token cap via model_override.
_QUALITY_AGENT_TOKENS = 4000


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
        self._thinking_mode = {}     # model name -> "adaptive" | "legacy", learned on first 400

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

    def pick_fast_model(self):
        """Fastest model from the CACHED live list — used for tiny utility
        calls (classification). No HTTP; None until models were fetched."""
        return self._pick_model(self._cached_models or [], _PREF_FAST)

    def pick_quality_model(self):
        """Most capable model (Opus > Sonnet > Haiku) from the LIVE list.

        Used by the self-study teacher enricher to distill gold answers from
        Opus regardless of the user's saved default. May fetch /v1/models.
        Returns None when no key / no live list.
        """
        return self._pick_model(self.get_models(), _PREF_QUALITY)

    @staticmethod
    def _quality_mode():
        """Opus-parity switch (agents.quality_mode). Off by default."""
        try:
            _ensure_lib_in_path()
            from config.settings import get_settings
            return bool(get_settings().is_quality_mode_enabled())
        except Exception:
            return False

    def _default_model(self, prefer_vision=False):
        """Default model resolved against the live /v1/models list (may fetch).

        Opus-parity mode overrides the cheap-first/capable-first order with
        most-capable-first (_PREF_QUALITY) so the user-facing default becomes
        Opus. Utility calls bypass this via model_override, so classification
        stays on the fast model.
        """
        if self._quality_mode():
            prefs = _PREF_QUALITY
        else:
            prefs = _PREF_VISION if prefer_vision else _PREF_TEXT
        return self._pick_model(self.get_models(), prefs)

    # ── Chat ─────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Post a chat request to the Claude API.

        user_content may be a plain string or a list of Claude-format content
        blocks (text + image) for vision requests.
        """
        if not HAS_HTTP:
            return self._fail_no_http("chat()")
        api_key = self._get_api_key()
        if not api_key:
            return self._fail_no_key("chat()")

        has_vision = self.has_image_blocks(user_content)

        # model_override: callers doing tiny utility calls (dispatcher
        # classification) can pin the fastest model without touching the
        # user's saved choice.
        model = (kwargs.get("model_override") or self._model
                 or self._default_model(prefer_vision=has_vision))
        if not model:
            return self._fail_no_model("chat()")

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
            resp_text  = http_post(CLAUDE_API_URL, payload, headers,
                                   timeout_ms=int(kwargs.get("timeout_ms")
                                                  or 60000))
            api_result = json.loads(resp_text)
            return api_result["content"][0]["text"].strip()
        except Exception as ex:
            self._record_error(u"chat() failed: {}".format(ex))
            return None

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """Stream a Claude response token-by-token via the Messages SSE API."""
        self._clear_error()
        if not HAS_HTTP:
            return self._fail_no_http("chat_stream()")
        api_key = self._get_api_key()
        if not api_key:
            return self._fail_no_key("chat_stream()")

        has_vision = self.has_image_blocks(user_content)
        model = self._model or self._default_model(prefer_vision=has_vision)
        if not model:
            return self._fail_no_model("chat_stream()")

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
        except Exception as ex:
            # Streaming/transport error. Record the reason first — chat() clears
            # the error state on entry, so without this the chat window would
            # show "the model didn't respond" with no Details line.
            stream_err = u"chat_stream() failed: {}".format(ex)
            self._record_error(stream_err)
            partial = u"".join(chunks)
            if not partial.strip():
                # Nothing reached the live bubble yet, so a fresh blocking retry
                # cannot visibly swap an answer — keep the original safety net.
                result = self.chat(messages, system_prompt, user_content,
                                   max_tokens, **kwargs)
                if result is None:
                    self._record_error(u"{} | blocking retry: {}".format(
                        stream_err, self.get_last_error() or u"no response"))
                return result
            # Partial text ALREADY streamed into the bubble. Regenerating a whole
            # new answer here is exactly what made the reply visibly swap under
            # the user (and cost double the tokens). Instead prefill the partial
            # as the assistant turn and let the model CONTINUE the same reply.
            # Any failure keeps just the partial — never a different answer.
            prefill = partial.rstrip()   # Anthropic rejects trailing whitespace
            cont = self._continue_after_drop(
                msgs, system_prompt, model, headers, max_tokens, prefill,
                on_delta, int(kwargs.get("timeout_ms") or 60000))
            if cont:
                return (prefill + cont).strip()
            return prefill.strip()

    def _continue_after_drop(self, msgs, system_prompt, model, headers,
                             max_tokens, prefill, on_delta, timeout_ms):
        """Finish a stream that dropped mid-answer with ONE blocking call.

        Prefills `prefill` as the assistant turn so the model continues the same
        reply instead of starting a different one — the text the user already
        watched stream in stays put. Best-effort: any failure returns None and
        the caller keeps just the partial. Never raises.
        """
        try:
            cont_msgs = list(msgs)
            cont_msgs.append({"role": "assistant", "content": prefill})
            payload = {
                "model":      model,
                "max_tokens": max_tokens,
                "system":     system_prompt,
                "messages":   cont_msgs,
            }
            resp_text  = http_post(CLAUDE_API_URL, payload, headers,
                                   timeout_ms=timeout_ms)
            api_result = json.loads(resp_text)
            cont = api_result["content"][0]["text"]
            if cont and on_delta:
                try:
                    on_delta(cont)
                except Exception:
                    pass
            return cont
        except Exception as ex:
            self._record_error(u"stream continuation failed: {}".format(ex))
            return None

    # ── Agentic chat (native tool calling) ────────────────────────────────────

    @staticmethod
    def _thinking_enabled():
        """Extended-thinking switch for agent turns.

        On when EITHER the explicit agents.extended_thinking option is set OR
        Opus-parity mode is active (quality_mode implies deep reasoning
        between tool calls). Off by default — thinking costs latency.
        """
        try:
            _ensure_lib_in_path()
            from config.settings import get_settings
            s = get_settings()
            if s.is_quality_mode_enabled():
                return True
            return bool(s.get_agent_option("extended_thinking", False))
        except Exception:
            return False

    def _agent_payload(self, model, system_prompt, messages, tools, max_tokens,
                       stream, thinking_mode=None):
        """Build a Messages API payload with prompt caching on system + tools
        AND on the conversation prefix.

        Three ephemeral cache breakpoints (limit is 4):
          1. system block, 2. LAST tool  — the (tools + system) prefix is
             cached server-side, so the ~75 tool schemas cost input tokens
             once per cache window instead of on every iteration;
          3. the last content block of the LAST message — the growing agent
             transcript (user turn + tool results) is re-read from cache on
             every subsequent iteration of the loop instead of being
             re-processed from scratch. On multi-tool requests this cuts
             both latency and input cost substantially.

        thinking_mode: None (thinking off), "adaptive" (current-generation
        models), or "legacy" (older enabled+budget_tokens shape — see the
        module-level comment above _THINK_BUDGET for why both exist).
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
            "messages":   self._cache_marked_messages(messages),
        }
        if cached_tools:
            payload["tools"] = cached_tools
        if thinking_mode == "adaptive":
            payload["thinking"]      = {"type": "adaptive"}
            payload["output_config"] = {"effort": _THINK_EFFORT}
        elif thinking_mode == "legacy":
            # budget must stay below max_tokens — raise the cap so the final
            # answer never gets squeezed out by the reasoning budget.
            payload["thinking"]   = {"type": "enabled",
                                     "budget_tokens": _THINK_BUDGET}
            payload["max_tokens"] = max(max_tokens, _THINK_BUDGET + 2000)
        if stream:
            payload["stream"] = True
        return payload

    @staticmethod
    def _thinking_rejected(ex):
        """True if `ex` looks like the API rejecting the `thinking` param
        shape (wrong type for this model generation) rather than some other
        failure (network, auth, rate limit) that a retry wouldn't fix."""
        return "thinking" in u"{}".format(ex).lower()

    @staticmethod
    def _cache_marked_messages(messages):
        """Copy `messages` with a cache_control marker on the final content
        block of the final message. Shallow-copies only the path it touches —
        the shared history/tool-result structures are never mutated."""
        msgs = list(messages or [])
        if not msgs:
            return msgs
        last = dict(msgs[-1])
        content = last.get("content")
        if isinstance(content, list) and content:
            blk = content[-1]
            if isinstance(blk, dict) and blk.get("type") in (
                    "text", "tool_result", "tool_use", "image", "document"):
                blk = dict(blk)
                blk["cache_control"] = {"type": "ephemeral"}
                last["content"] = content[:-1] + [blk]
        elif isinstance(content, ("".__class__, u"".__class__)) and content:
            # plain string content (sanitized history turn)
            last["content"] = [{"type": "text", "text": content,
                                "cache_control": {"type": "ephemeral"}}]
        else:
            return msgs
        return msgs[:-1] + [last]

    def chat_agent(self, system_prompt, messages, tools,
                   on_delta=None, max_tokens=1500, **kwargs):
        """One agentic turn: streams text deltas, collects tool_use blocks.

        Returns {"text", "tool_calls", "assistant_msg", "stop_reason"} or None.
        `messages` must already end with the latest user / tool_result turn.
        """
        if not HAS_HTTP:
            return self._fail_no_http("chat_agent()")
        api_key = self._get_api_key()
        if not api_key:
            return self._fail_no_key("chat_agent()")

        # Agent turns favour a more capable model when the user hasn't picked
        # one — resolved from the LIVE model list, never a hardcoded name.
        model = self._model or self._default_model(prefer_vision=True)
        if not model:
            return self._fail_no_model("chat_agent()")

        # Opus-parity mode widens the answer ceiling so rich multi-step
        # replies never truncate. Never lower a caller's explicit higher cap.
        if self._quality_mode():
            max_tokens = max(max_tokens, _QUALITY_AGENT_TOKENS)

        # Deep reasoning (opt-in): extended thinking between tool calls.
        # Default to adaptive (current-generation shape); a cached "legacy"
        # verdict from a prior 400 on this model skips straight past it.
        thinking_mode = (self._thinking_mode.get(model, "adaptive")
                         if self._thinking_enabled() else None)

        def _headers(mode):
            h = {"x-api-key": api_key, "anthropic-version": ANTHROPIC_API_VER}
            if mode == "legacy":
                h["anthropic-beta"] = _THINK_BETA
            return h

        # ── Streaming attempt ────────────────────────────────────────────────
        state = {
            "blocks":      [],     # finalized content blocks, in order
            "cur":         None,   # block being streamed
            "stop_reason": None,
            # Token accounting. `message_start` carries the whole input side
            # (including cache_read_input_tokens / cache_creation_input_tokens,
            # the only way to tell whether prompt caching is actually working);
            # `message_delta` carries the final output_tokens.
            "usage":       {},
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

            if etype == "message_start":
                try:
                    u = (obj.get("message", {}) or {}).get("usage", {}) or {}
                    if isinstance(u, dict):
                        state["usage"].update(u)
                except Exception:
                    pass

            elif etype == "content_block_start":
                cb = obj.get("content_block", {}) or {}
                if cb.get("type") == "tool_use":
                    state["cur"] = {"type": "tool_use", "id": cb.get("id", ""),
                                    "name": cb.get("name", ""), "parts": []}
                elif cb.get("type") in ("thinking", "redacted_thinking"):
                    # Reasoning blocks are captured (they MUST be echoed back
                    # in the transcript when thinking is on) but never shown.
                    state["cur"] = {"type": cb.get("type"), "parts": [],
                                    "sig": cb.get("signature"),
                                    "data": cb.get("data")}
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
                elif delta.get("type") == "thinking_delta":
                    cur["parts"].append(delta.get("thinking") or u"")
                elif delta.get("type") == "signature_delta":
                    cur["sig"] = delta.get("signature")

            elif etype == "content_block_stop":
                cur = state["cur"]
                if cur is not None:
                    state["blocks"].append(cur)
                    state["cur"] = None

            elif etype == "message_delta":
                d = obj.get("delta", {}) or {}
                if d.get("stop_reason"):
                    state["stop_reason"] = d["stop_reason"]
                try:
                    u = obj.get("usage", {}) or {}
                    if isinstance(u, dict):
                        state["usage"].update(u)
                except Exception:
                    pass

        streamed_ok = False
        try:
            self._clear_error()
            payload = self._agent_payload(model, system_prompt, messages, tools,
                                          max_tokens, stream=True,
                                          thinking_mode=thinking_mode)
            http_post_stream(CLAUDE_API_URL, payload, _headers(thinking_mode),
                             _on_line, timeout_ms=180000)
            streamed_ok = bool(state["blocks"]) or state["stop_reason"] is not None
        except Exception as ex:
            if thinking_mode == "adaptive" and self._thinking_rejected(ex):
                # This model generation doesn't accept adaptive thinking —
                # fall back to the legacy shape and remember the verdict so
                # future calls on this model skip straight to it.
                self._thinking_mode[model] = thinking_mode = "legacy"
                state["blocks"], state["cur"], state["stop_reason"] = [], None, None
                try:
                    payload = self._agent_payload(model, system_prompt, messages,
                                                  tools, max_tokens, stream=True,
                                                  thinking_mode=thinking_mode)
                    http_post_stream(CLAUDE_API_URL, payload, _headers(thinking_mode),
                                     _on_line, timeout_ms=180000)
                    streamed_ok = (bool(state["blocks"])
                                  or state["stop_reason"] is not None)
                except Exception as ex2:
                    self._record_error(u"chat_agent stream failed: {}".format(ex2))
            else:
                self._record_error(u"chat_agent stream failed: {}".format(ex))

        if streamed_ok:
            return self._agent_result_from_blocks(state["blocks"],
                                                  state["stop_reason"],
                                                  state["usage"])

        # ── Blocking fallback ────────────────────────────────────────────────
        def _do_blocking(mode):
            payload    = self._agent_payload(model, system_prompt, messages, tools,
                                             max_tokens, stream=False,
                                             thinking_mode=mode)
            resp_text  = http_post(CLAUDE_API_URL, payload, _headers(mode),
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
                elif b.get("type") == "thinking":
                    blocks.append({"type": "thinking",
                                   "parts": [b.get("thinking", "")],
                                   "sig": b.get("signature")})
                elif b.get("type") == "redacted_thinking":
                    blocks.append({"type": "redacted_thinking", "parts": [],
                                   "data": b.get("data")})
            return (blocks, api_result.get("stop_reason"),
                    api_result.get("usage") or {})

        try:
            blocks, stop_reason, usage = _do_blocking(thinking_mode)
            return self._agent_result_from_blocks(blocks, stop_reason, usage)
        except Exception as ex:
            if thinking_mode == "adaptive" and self._thinking_rejected(ex):
                self._thinking_mode[model] = thinking_mode = "legacy"
                try:
                    blocks, stop_reason, usage = _do_blocking(thinking_mode)
                    return self._agent_result_from_blocks(blocks, stop_reason,
                                                          usage)
                except Exception as ex2:
                    self._record_error(u"chat_agent() failed: {}".format(ex2))
                    return None
            self._record_error(u"chat_agent() failed: {}".format(ex))
            return None

    @staticmethod
    def _agent_result_from_blocks(blocks, stop_reason, usage=None):
        """Assemble the uniform chat_agent result from streamed blocks.

        Thinking blocks are preserved in assistant_msg (the API requires the
        exact blocks + signatures back in the transcript when extended
        thinking is enabled) but never surface in the user-visible text.

        `usage` is the raw API token accounting for THIS call, passed through
        untouched for Intelligence/telemetry.py to accumulate. Providers that
        report nothing simply yield {}.
        """
        texts      = []
        tool_calls = []
        content    = []
        for b in blocks:
            joined = u"".join(b.get("parts", []))
            if b["type"] == "text":
                if joined:
                    texts.append(joined)
                    content.append({"type": "text", "text": joined})
            elif b["type"] == "thinking":
                blk = {"type": "thinking", "thinking": joined}
                if b.get("sig"):
                    blk["signature"] = b["sig"]
                content.append(blk)
            elif b["type"] == "redacted_thinking":
                content.append({"type": "redacted_thinking",
                                "data": b.get("data")})
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
            "usage":         dict(usage or {}),
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
