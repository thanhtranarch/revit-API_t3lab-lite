# -*- coding: utf-8 -*-
"""
Agent Loop

Provider-agnostic agentic tool-calling loop for the T3Lab Assistant.

Runs on a BACKGROUND thread. Providers that set SUPPORTS_NATIVE_TOOLS expose:

    chat_agent(system_prompt, messages, tools, on_delta=None, max_tokens=...)
        -> {"text":        unicode,          # user-visible text ("" if none)
            "tool_calls":  [{"id","name","args"}],
            "assistant_msg": <provider-native message to append to transcript>,
            "stop_reason": str}
        or None on transport/parse failure.

    agent_tool_results(tool_calls, results)
        -> list of provider-native messages carrying the tool results.

The loop itself never touches WPF or the Revit API directly:
- UI feedback flows through the `callbacks` dict (caller marshals to UI thread);
- tool execution goes through the injected `execute_tool(name, args)` callable
  (script.py passes core.server's _execute_tool, which already marshals write
  tools onto Revit's main thread via ExternalEvent).

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Agent Loop"

import json
import time

from Intelligence.tool_schema import LAUNCHER_TOOL_NAME


# ─── Result truncation ─────────────────────────────────────────────────────────
# Tool results are fed back to the model; huge payloads (element dumps) blow the
# context. Keep the JSON fed to the model bounded — the UI card shows the full
# result separately.
_MAX_RESULT_CHARS = 4000


def _result_to_json(result):
    try:
        s = json.dumps(result, ensure_ascii=False)
    except Exception:
        s = u"{}".format(result)
    if len(s) > _MAX_RESULT_CHARS:
        s = s[:_MAX_RESULT_CHARS] + u"... [truncated {} chars]".format(len(s) - _MAX_RESULT_CHARS)
    return s


def _sanitize_history(history, limit=16):
    """Reduce persisted chat history to plain-text user/assistant messages."""
    out = []
    for h in (history or [])[-limit:]:
        role    = h.get("role", "")
        content = h.get("content", "")
        if role not in ("user", "assistant"):
            continue
        if isinstance(content, list):
            parts = [b.get("text", "") for b in content
                     if isinstance(b, dict) and b.get("type") == "text"]
            content = u"\n".join(parts)
        if content:
            out.append({"role": role, "content": content})
    # Provider APIs reject a transcript that starts with an assistant turn.
    while out and out[0]["role"] == "assistant":
        out.pop(0)
    return out


class AgentLoop(object):
    """One user request = one AgentLoop.run(). Cancellable between steps."""

    def __init__(self, provider, execute_tool, tools, callbacks=None,
                 max_iterations=10, max_tokens=1500, time_budget_sec=240):
        self._provider       = provider
        self._execute_tool   = execute_tool
        self._tools          = tools
        self._cb             = callbacks or {}
        self._max_iterations = max_iterations
        self._max_tokens     = max_tokens
        self._time_budget    = time_budget_sec
        self._cancelled      = False

    # ── Cancellation ──────────────────────────────────────────────────────────

    def cancel(self):
        """Request a stop. The current model turn / tool finishes, then the
        loop ends — a Transaction mid-flight cannot be safely aborted."""
        self._cancelled = True

    def is_cancelled(self):
        return self._cancelled

    def _guard_tripped(self):
        """True when the optional guard_check callback reports the request
        context is no longer valid (e.g. the user switched Revit documents
        mid-request — writing on would edit the wrong model)."""
        guard = self._cb.get("guard_check")
        if guard is None:
            return False
        try:
            return bool(guard())
        except Exception:
            return False

    # ── Callback helpers (never raise into the loop) ──────────────────────────

    def _emit(self, name, *args):
        fn = self._cb.get(name)
        if fn:
            try:
                fn(*args)
            except Exception:
                pass

    # ── Text tool-call rescue ─────────────────────────────────────────────────

    def _registered_tool_names(self):
        """Tool names from the provider-native schema list (Anthropic uses
        {"name": ...}, OpenAI-style wraps it in {"function": {"name": ...}})."""
        names = set()
        for t in (self._tools or []):
            try:
                n = t.get("name") or (t.get("function") or {}).get("name")
                if n:
                    names.add(n)
            except Exception:
                pass
        return names

    def _rescue_text_tool_call(self, text):
        """Salvage a tool call the model wrote as plain text instead of a
        native tool_call block — local models drift into JSON-in-prose like
        {"name": "get_material_quantities", "parameters": {...}} after
        seeing a tool error. Returns a tool_calls-shaped list, or None."""
        if not text or u"{" not in text:
            return None
        try:
            import re
            m = re.search(r'\{[\s\S]*\}', text)
            obj = json.loads(m.group()) if m else None
        except Exception:
            return None
        if not isinstance(obj, dict):
            return None
        name = obj.get("name") or obj.get("tool") or obj.get("intent")
        args = (obj.get("parameters") or obj.get("arguments")
                or obj.get("params") or {})
        if not name or not isinstance(args, dict):
            return None
        if name not in self._registered_tool_names():
            return None
        return [{"id": "text_rescue", "name": u"{}".format(name),
                 "args": args}]

    # ── Main loop ─────────────────────────────────────────────────────────────

    def run(self, history, system_prompt, user_content):
        """Execute the agentic loop. Blocking — call from a worker thread.

        Returns:
            {"status": "done"|"cancelled"|"failed"|"max_iterations"|"timeout"
                       |"doc_changed",
             "text": <last user-visible text>,
             "launch_intent": <str|None>,   # open_t3lab_tool target, terminal
             "iterations": int, "tool_runs": int}
        """
        messages = _sanitize_history(history)
        messages.append({"role": "user", "content": user_content})

        started    = time.time()
        last_text  = u""
        tool_runs  = 0
        iteration  = 0

        while iteration < self._max_iterations:
            iteration += 1

            if self._cancelled:
                return self._finish("cancelled", last_text, None, iteration, tool_runs)
            if self._guard_tripped():
                return self._finish("doc_changed", last_text, None, iteration, tool_runs)
            if (time.time() - started) > self._time_budget:
                return self._finish("timeout", last_text, None, iteration, tool_runs)

            # Track whether this turn actually streamed deltas so blocking
            # providers still surface their text through on_text_delta once.
            turn = {"streamed": False}

            def _on_delta(chunk, _turn=turn):
                _turn["streamed"] = True
                self._emit("on_text_delta", chunk)

            try:
                resp = self._provider.chat_agent(
                    system_prompt, messages, self._tools,
                    on_delta=_on_delta, max_tokens=self._max_tokens)
            except Exception:
                resp = None

            if resp is None:
                # First turn failing = provider never answered; later turns
                # failing still leave earlier text/tool work worth reporting.
                status = "failed" if iteration == 1 else "done"
                return self._finish(status, last_text, None, iteration, tool_runs)

            text  = (resp.get("text") or u"").strip()
            calls = resp.get("tool_calls") or []

            # Model wrote the tool call as JSON text instead of a native
            # tool_call block → execute it anyway instead of ending the turn
            # with a dead JSON blob on screen.
            rescued = False
            if not calls and text:
                _rc = self._rescue_text_tool_call(text)
                if _rc:
                    calls = _rc
                    rescued = True

            if text:
                last_text = text
                if not turn["streamed"]:
                    self._emit("on_text_delta", text)
                self._emit("on_turn_text", text, not calls)

            messages.append(resp["assistant_msg"])

            if not calls:
                return self._finish("done", last_text, None, iteration, tool_runs)

            # ── Execute tool calls sequentially (Revit is single-threaded) ──
            results       = []
            launch_intent = None
            doc_changed   = False
            for tc in calls:
                name = tc.get("name", "")
                args = tc.get("args") or {}

                if self._guard_tripped():
                    doc_changed = True
                    results.append({"cancelled": True,
                                    "note": "Active document changed — request aborted."})
                    break

                if name == LAUNCHER_TOOL_NAME:
                    # Terminal: the window opens on the UI thread AFTER the
                    # loop ends (ShowDialog would otherwise block the loop).
                    launch_intent = args.get("tool_intent") or ""
                    results.append({"success": True,
                                    "note": "T3Lab tool window will open now."})
                    break

                if self._cancelled:
                    results.append({"cancelled": True,
                                    "note": "User stopped the request."})
                    break

                self._emit("on_tool_start", name, args, iteration)
                t0 = time.time()
                try:
                    res = self._execute_tool(name, args)
                except Exception as ex:
                    res = {"error": u"{}".format(ex), "tool": name}
                if not isinstance(res, dict):
                    res = {"result": res}
                dt = time.time() - t0
                ok = "error" not in res
                tool_runs += 1
                self._emit("on_tool_done", name, res, ok, dt)
                results.append(res)

            if launch_intent:
                return self._finish("done", last_text, launch_intent,
                                    iteration, tool_runs)
            if doc_changed:
                return self._finish("doc_changed", last_text, None,
                                    iteration, tool_runs)
            if self._cancelled:
                return self._finish("cancelled", last_text, None,
                                    iteration, tool_runs)

            # Feed results back in the provider's native format — serialized
            # and truncated here so no provider ever pushes a huge element
            # dump into the model context.
            try:
                res_strs = [_result_to_json(r) for r in results]
                if rescued:
                    # A rescued call has no native tool_call block in the
                    # transcript, so provider tool-result messages would
                    # reference an unknown id — feed the result back as plain
                    # user text instead (accepted by every provider).
                    messages.append({
                        "role": "user",
                        "content": u"Tool `{}` returned: {}\nContinue the "
                                   u"task. Use REAL tool calls, never JSON "
                                   u"as text.".format(calls[0]["name"],
                                                      res_strs[0])})
                else:
                    messages.extend(
                        self._provider.agent_tool_results(calls, res_strs))
            except Exception:
                return self._finish("failed", last_text, None, iteration, tool_runs)

        return self._finish("max_iterations", last_text, None,
                            self._max_iterations, tool_runs)

    def _finish(self, status, text, launch_intent, iterations, tool_runs):
        result = {
            "status":        status,
            "text":          text,
            "launch_intent": launch_intent,
            "iterations":    iterations,
            "tool_runs":     tool_runs,
        }
        self._emit("on_finish", result)
        return result


# ─── System prompt (native tool-calling mode) ──────────────────────────────────
# Unlike the legacy JSON-intent prompt, this one carries NO tool schemas (they
# travel through the API `tools` parameter) and does NOT force JSON output.

_AGENT_PROMPT = u"""You are T3Lab Assistant, an AI agent embedded in Autodesk Revit via the T3Lab pyRevit extension. You can read and modify the live Revit model through the tools provided.

## Language & formatting
Always reply in English, regardless of the language the user writes in. Keep replies short and practical — one or two sentences between tool calls, a compact summary at the end.
Use markdown when it helps: **bold**, `code`, bullet lists, and pipe tables (| a | b |) for numeric summaries — the chat renders them natively. Do NOT use emoji.

## Units
All tool coordinates and dimensions are in METERS. Convert user input: 5000mm = 5.0, 3m = 3.0. Element ids are integers.

## Plan, then act
For a multi-step request, decide the full tool sequence BEFORE the first call and state it in one short sentence. Chain the steps yourself — never ask the user to run intermediate steps you can do with tools. For a simple request, skip the plan and just act.

## Efficiency (each model turn is expensive — minimize turns)
- BATCH independent read-only calls into ONE turn: when you need e.g. levels + selection + view info, emit all those tool calls together instead of one per turn.
- Prefer ONE bulk tool over many single calls: `bulk_set_parameter` over repeated `set_parameter`, `ai_element_filter` over per-element inspection, `tag_all_*` over per-element tags.
- Do not re-query data that is already in this conversation (earlier tool results, the context block below) unless the model may have changed since.
- Never call a tool "just to check" after a success result already told you the outcome; verify with a read tool only after LARGE modifications (20+ elements) or when a result looks suspicious.

## Working rules
1. Query before you modify: when element ids or names are unknown, use read tools (get_current_view_elements, ai_element_filter, list_levels, ...) first.
2. After finishing, summarize WHAT changed (counts + element ids) in English.
3. Destructive actions (delete_element, purge_unused, and anything removing model data): unless the user's current message already explicitly requested the deletion, ask for confirmation in text FIRST and stop — do not call the tool in the same turn.
4. If a tool returns an error, diagnose from the error text and retry ONCE with corrected arguments; if it fails again, report exactly what is missing — never loop on the same failing call.
5. If the request is ambiguous about WHICH elements to change, ask ONE precise clarifying question instead of guessing — a wrong bulk edit is worse than a question.
6. `open_t3lab_tool` opens a T3Lab window and ENDS your turn — only ever call it last, and never together with other tools.
7. When the user refers to the current selection ("these elements", "the selected walls", "các element này", "đang chọn"), call `revit_get_selected_elements` FIRST and operate on those element ids — never guess ids.
8. Trust tool results over assumptions: counts, names and ids come from the model, not from memory of typical projects.
9. Scope resolution when the request names no explicit target: for CHECKING / auditing / statistics / spell-check requests the default scope is the ENTIRE PROJECT — never silently limit to the current view (limit only when the user says so, or ask ONE question when unsure). For MODIFY actions, use the "Current Revit context" block below — the user's selection first, then the active view. Never report "nothing found" from one filtered query: re-check with corrected arguments or wider scope first.
10. NUMBERS MUST COME FROM TOOLS, NEVER FROM YOU. Tool results may be truncated before you see them, so counting or summing rows/elements yourself gives wrong answers. Use the exact fields the tools compute: `total_count` (ai_element_filter), `row_count` / `column_totals` (get_schedule_data), `element_counts` (analyze_model_statistics). If a statistic you need has no tool-computed field, say so — do not estimate.
11. Color / highlight requests — "tô màu", "tô đỏ X", "bôi xanh X", "đổi màu X", "highlight X", "color X red": the color word is the OVERRIDE color to APPLY, never a parameter value to filter by (Revit elements have no "Color" or "Fill Color" parameter). Resolve the target ids first (the user's selection, or `ai_element_filter` with category only — no parameter filter), then call `revit_override_color` with that color and those ids. Use `color_elements` only when the user wants elements colored BY a parameter's values (one distinct color per value, e.g. "tô màu tường theo Type") — and its parameter_name must be a REAL parameter, never empty, never a color name.
12. NEVER fabricate tool activity. Every action happens ONLY through a real tool call in this conversation — if you say you will use a tool, emit that tool call in the same turn, and only report results that came back from an actual tool result. Writing an invented "Result:" for a call you never made is a critical failure.

## Multiple open models
Several documents can be open in this Revit session; every tool operates on the ACTIVE one. To work across models: `list_open_documents` → `switch_active_document` (exact title) → read/modify there → switch again as needed, then combine everything into ONE final answer stating which model each number comes from. Element ids are only valid inside their own document — never reuse ids across models. When the user names a model that is not in `list_open_documents`, it may live in a different Revit window (separate process): say so and point them to the assistant in that window — do not guess.

## Current Revit context
{context}
"""


def build_agent_system_prompt(revit_context=u""):
    """System prompt for the native tool-calling agent path."""
    ctx = revit_context.strip() if revit_context else u"(no context snapshot available)"
    return _AGENT_PROMPT.format(context=ctx)
