# -*- coding: utf-8 -*-
"""
Tool Schema

Converts the T3Lab MCP tool registry (core/server.py) into the native
function-calling formats of each LLM provider, so tool schemas are sent
through the API's `tools` parameter instead of being dumped as text into
the system prompt on every turn.

Formats:
    Anthropic:  {"name", "description", "input_schema"}
    OpenAI:     {"type": "function", "function": {"name", "description", "parameters"}}
    Ollama:     same wire shape as OpenAI (/api/chat `tools`)

Results are cached per session — the registry only changes on pyRevit reload.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Tool Schema"

import copy


# Synthetic tool exposed ONLY to the in-chat agent (not the MCP HTTP surface):
# lets the model open a T3Lab pushbutton window. The agent loop treats it as
# terminal — the launch happens on the UI thread after the loop finishes.
LAUNCHER_TOOL_NAME = "open_t3lab_tool"


def make_launcher_tool(intents):
    """Build the open_t3lab_tool schema in server-registry shape.

    Args:
        intents: list of launcher intent strings (e.g. ["open_batchout", ...]).
    """
    return {
        "name": LAUNCHER_TOOL_NAME,
        "description": (
            "Open a T3Lab tool window inside Revit by its intent name. "
            "This ends the conversation turn — call it as your FINAL action, "
            "never followed by other tool calls."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "tool_intent": {
                    "type": "string",
                    "description": "Which T3Lab tool to open.",
                    "enum": sorted(intents or []),
                },
            },
            "required": ["tool_intent"],
        },
    }


# ─── Registry access ───────────────────────────────────────────────────────────

_cache = {}   # {"raw": [...], "anthropic": [...], "openai": [...]}


def invalidate_cache():
    _cache.clear()


def get_server_tools():
    """Return the raw tool list from the local MCP server registry.

    Each item: {'name', 'description', 'inputSchema'}. Returns [] on failure
    (server module unavailable) so callers can fall back gracefully.
    """
    if "raw" in _cache:
        return _cache["raw"]
    try:
        from core.server import get_t3labai_server
        srv   = get_t3labai_server()
        tools = srv._handle_tools_list().get("tools", []) or []
        # Keep only well-formed entries; never let one bad schema kill the list.
        clean = [t for t in tools
                 if isinstance(t, dict) and t.get("name") and t.get("inputSchema")]
        _cache["raw"] = clean
        return clean
    except Exception:
        return []


# ─── Converters ────────────────────────────────────────────────────────────────

def to_anthropic_tools(extra_tools=None):
    """Convert registry (+ optional extra tools in registry shape) to Anthropic format."""
    key = "anthropic" if not extra_tools else None
    if key and key in _cache:
        return _cache[key]
    out = []
    for t in get_server_tools() + list(extra_tools or []):
        out.append({
            "name":         t["name"],
            "description":  t.get("description", "") or t["name"],
            "input_schema": copy.deepcopy(t["inputSchema"]),
        })
    if key:
        _cache[key] = out
    return out


def to_openai_tools(extra_tools=None):
    """Convert registry (+ optional extras) to OpenAI/Ollama function format."""
    key = "openai" if not extra_tools else None
    if key and key in _cache:
        return _cache[key]
    out = []
    for t in get_server_tools() + list(extra_tools or []):
        out.append({
            "type": "function",
            "function": {
                "name":        t["name"],
                "description": t.get("description", "") or t["name"],
                "parameters":  copy.deepcopy(t["inputSchema"]),
            },
        })
    if key:
        _cache[key] = out
    return out


def get_tools_for_provider(provider_name, extra_tools=None):
    """Return the tool list in the right wire format for a provider name."""
    if provider_name == "claude":
        return to_anthropic_tools(extra_tools)
    # openai / deepseek / ollama / lmstudio all use the OpenAI function shape
    return to_openai_tools(extra_tools)


def is_registered_tool(name):
    """True if `name` is a real MCP tool in the server registry."""
    for t in get_server_tools():
        if t["name"] == name:
            return True
    return False
