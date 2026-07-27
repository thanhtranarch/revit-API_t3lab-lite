# -*- coding: utf-8 -*-
"""
Local-model tool-calling benchmark for the T3Lab Assistant.

Answers ONE question on YOUR hardware: which installed local model is worth
using for the assistant? A general leaderboard doesn't — what matters here is
whether the model reliably picks the RIGHT tool with valid arguments on the
T3Lab agentic path, in Vietnamese and English, without hallucinating tools or
firing a tool at plain small talk.

It reuses the REAL providers (OllamaProvider / LMStudioProvider), so it
measures the exact path the assistant uses — the same sampling, num_ctx,
`think` control and reasoning-model handling shipped in lib/Intelligence.

Run (CPython 3, with Ollama and/or LM Studio running locally):
    python3 dev/bench_local_models.py                 # all Ollama models
    python3 dev/bench_local_models.py --provider lmstudio
    python3 dev/bench_local_models.py --models qwen3:14b,qwen3:8b
    python3 dev/bench_local_models.py --json out.json  # also dump raw results

No external test framework, no Revit needed. The tool schemas below mirror the
real ESSENTIAL_TOOL_NAMES subset (lib/Intelligence/tool_schema.py) so scores
track the shipping behaviour, but the script stays self-contained because the
live MCP registry only exists inside Revit.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import argparse
import json
import os
import sys
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

# Sandbox %APPDATA% before any settings import (LM Studio host/key lookups).
import tempfile
os.environ.setdefault('APPDATA', tempfile.mkdtemp(prefix='t3lab_bench_'))


# ─── Tool schemas (mirror of the essential subset, OpenAI function shape) ──────

def _fn(name, description, properties, required):
    return {"type": "function", "function": {
        "name": name, "description": description,
        "parameters": {"type": "object", "properties": properties,
                       "required": required}}}


BENCH_TOOLS = [
    _fn("ai_element_filter",
        "Query/count elements by category (Walls, Doors, Windows, Columns, "
        "Floors, Rooms, ...). Returns total_count. Use for 'how many X', "
        "'count X', 'liệt kê X', 'có bao nhiêu X'.",
        {"category": {"type": "string",
                      "description": "Element category, e.g. Walls, Doors"},
         "parameter_name": {"type": "string"},
         "parameter_value": {"type": "string"},
         "limit": {"type": "integer"}},
        ["category"]),
    _fn("revit_override_color",
        "Apply ONE color override to elements in the active view — for "
        "'tô/color/highlight X red/green/blue'. Pass category for a whole "
        "category (server collects all, no ids). The color word is what to "
        "APPLY, never a filter value.",
        {"color": {"type": "string",
                   "description": "Hex (#FF0000) or name (red, green, blue)"},
         "category": {"type": "string"},
         "element_ids": {"type": "array", "items": {"type": "integer"}}},
        ["color"]),
    _fn("revit_get_selected_elements",
        "Return the elements the user currently has selected in Revit. Use "
        "when the user refers to 'the selection', 'đang chọn', 'these elements'.",
        {}, []),
    _fn("rename_element",
        "Rename a view, sheet, level or any named element.",
        {"element_id": {"type": "integer"},
         "new_name": {"type": "string"}},
        ["element_id", "new_name"]),
    _fn("set_parameter",
        "Set a named parameter's value on one element.",
        {"element_id": {"type": "integer"},
         "parameter_name": {"type": "string"},
         "value": {"type": "string"}},
        ["element_id", "parameter_name", "value"]),
    _fn("list_levels",
        "List all levels in the model (name + elevation).",
        {}, []),
    _fn("revit_list_sheets",
        "List drawing sheets (number + name).",
        {"filter": {"type": "string"}}, []),
    _fn("create_text_note",
        "Place a text note in the active view.",
        {"text": {"type": "string"},
         "x": {"type": "number"}, "y": {"type": "number"}},
        ["text"]),
    _fn("analyze_model_statistics",
        "Element counts by category + warnings summary for the whole model.",
        {}, []),
    _fn("open_t3lab_tool",
        "Open a T3Lab tool window. Terminal action — call LAST and alone. "
        "Use for 'mở/open <tool>' (batchout, parasync, load family, ...).",
        {"tool_intent": {"type": "string",
                         "description": "e.g. open_batchout, open_parasync"}},
        ["tool_intent"]),
]

BENCH_TOOL_NAMES = frozenset(t["function"]["name"] for t in BENCH_TOOLS)


# ─── Test cases: (prompt, expected_tool_or_None, expected_arg_key_or_None) ─────
# expected_tool None  → the model should answer in TEXT with NO tool call
#                       (small-talk trap: firing a tool here is a real failure).
# expected_arg        → a param that the right call must include (soft check).

CASES = [
    # counting / listing (read)
    ("có bao nhiêu bức tường trong dự án?", "ai_element_filter", "category"),
    ("how many doors are there?",           "ai_element_filter", "category"),
    ("đếm số cửa sổ",                        "ai_element_filter", "category"),
    ("liệt kê tất cả level",                 "list_levels",       None),
    ("list all sheets",                      "revit_list_sheets", None),
    ("thống kê tổng quan model",             "analyze_model_statistics", None),
    # color (the classic "color is what to apply, not a filter" trap)
    ("tô đỏ tường",                          "revit_override_color", "color"),
    ("color all floors green",               "revit_override_color", "color"),
    ("bôi xanh dương các cột",               "revit_override_color", "color"),
    # selection
    ("cho tôi biết đang chọn bao nhiêu element",
                                             "revit_get_selected_elements", None),
    # modify with explicit ids
    ("đổi tên phần tử id 12345 thành A-102", "rename_element", "new_name"),
    ("set parameter Comments = OK for element 555",
                                             "set_parameter", "parameter_name"),
    # launcher (terminal)
    ("mở batchout",                          "open_t3lab_tool", "tool_intent"),
    ("mở parasync giúp mình",                "open_t3lab_tool", "tool_intent"),
    # small-talk traps → NO tool
    ("chào bạn, khỏe không?",                None, None),
    ("cảm ơn nhé",                           None, None),
    ("ok tốt rồi",                           None, None),
]


# ─── Provider access ───────────────────────────────────────────────────────────

def _get_provider(name):
    if name == "ollama":
        from Intelligence.ollama_provider import OllamaProvider
        return OllamaProvider()
    if name == "lmstudio":
        from Intelligence.lmstudio_provider import LMStudioProvider
        return LMStudioProvider()
    raise ValueError("unknown provider: " + name)


def _system_prompt():
    try:
        from Intelligence.agent_loop import build_agent_system_prompt
        return build_agent_system_prompt(
            "(benchmark — no live model; assume a typical project)", local=True)
    except Exception:
        return ("You are the T3Lab Assistant agent in Revit. Use the provided "
                "tools. Reply in English. Never invent a tool.")


# ─── Scoring ───────────────────────────────────────────────────────────────────

def _first_call(result):
    calls = (result or {}).get("tool_calls") or []
    return calls[0] if calls else None


def score_case(result, expected_tool, expected_arg):
    """Return (passed, detail) for one case."""
    call = _first_call(result)
    called = call["name"] if call else None

    # hallucinated tool (not in our schema) is always a failure
    if called is not None and called not in BENCH_TOOL_NAMES:
        return False, "hallucinated:" + called

    if expected_tool is None:
        # small-talk trap: any tool call is wrong
        if call is None:
            return True, "text-ok"
        return False, "fired:" + (called or "?")

    if call is None:
        return False, "no-call(text)"
    if called != expected_tool:
        return False, "wrong:" + (called or "?")
    # right tool → check args are a dict and (soft) the expected key exists
    args = call.get("args")
    if not isinstance(args, dict):
        return False, "bad-args"
    if expected_arg and expected_arg not in args:
        return True, "ok(missing-arg:%s)" % expected_arg   # tool right, arg soft
    return True, "ok"


def bench_model(provider, model, system_prompt, warmup=True):
    provider.set_model(model)
    rows = []
    if warmup:
        # First call also loads the model into VRAM — don't time it.
        try:
            provider.chat_agent(system_prompt,
                                [{"role": "user", "content": "hi"}],
                                BENCH_TOOLS, max_tokens=64)
        except Exception:
            pass
    for prompt, exp_tool, exp_arg in CASES:
        t0 = time.time()
        try:
            res = provider.chat_agent(
                system_prompt, [{"role": "user", "content": prompt}],
                BENCH_TOOLS, max_tokens=1024)
        except Exception as ex:
            res = None
            err = u"{}".format(ex)
        else:
            err = None
        dt = time.time() - t0
        passed, detail = score_case(res, exp_tool, exp_arg)
        if res is None and err:
            detail = "provider-none"
        rows.append({
            "prompt": prompt, "expected": exp_tool,
            "passed": passed, "detail": detail,
            "seconds": round(dt, 2),
            "muted": res is None,
        })
    return rows


def summarize(rows):
    n = len(rows)
    passed = sum(1 for r in rows if r["passed"])
    tool_cases = [r for r in rows if r["expected"] is not None]
    talk_cases = [r for r in rows if r["expected"] is None]
    tool_ok = sum(1 for r in tool_cases if r["passed"])
    talk_ok = sum(1 for r in talk_cases if r["passed"])
    halluc = sum(1 for r in rows if r["detail"].startswith("hallucinated"))
    muted = sum(1 for r in rows if r["muted"])
    timed = [r["seconds"] for r in rows if not r["muted"]]
    avg_s = round(sum(timed) / len(timed), 2) if timed else 0.0
    return {
        "n": n, "passed": passed, "acc": round(100.0 * passed / n, 1) if n else 0,
        "tool_ok": tool_ok, "tool_n": len(tool_cases),
        "talk_ok": talk_ok, "talk_n": len(talk_cases),
        "hallucinations": halluc, "muted": muted, "avg_s": avg_s,
    }


# ─── Reporting ─────────────────────────────────────────────────────────────────

def print_report(results):
    print("")
    print("=" * 78)
    print("T3Lab local-model tool-calling benchmark")
    print("=" * 78)
    # leaderboard sorted by accuracy, then speed
    board = sorted(results.items(),
                   key=lambda kv: (-kv[1]["summary"]["acc"],
                                   kv[1]["summary"]["avg_s"]))
    print("")
    print("{:<26} {:>7} {:>10} {:>9} {:>7} {:>8}".format(
        "MODEL", "SCORE", "TOOLS", "SMALLTALK", "HALLU", "AVG s"))
    print("-" * 78)
    for model, data in board:
        s = data["summary"]
        print("{:<26} {:>6}% {:>7}/{:<2} {:>7}/{:<2} {:>7} {:>8}".format(
            model[:26], s["acc"],
            s["tool_ok"], s["tool_n"],
            s["talk_ok"], s["talk_n"],
            s["hallucinations"], s["avg_s"]))
    print("-" * 78)
    print("SCORE = overall pass rate · TOOLS = right tool+args · "
          "SMALLTALK = correctly NO tool")
    print("HALLU = invented tools (lower better) · AVG s = latency/call "
          "(warm, excludes load)")

    # per-model failure detail (only the misses, so the table stays readable)
    for model, data in board:
        misses = [r for r in data["rows"] if not r["passed"]]
        if not misses:
            print("\n{}  — clean sweep, no misses".format(model))
            continue
        print("\n{}  misses:".format(model))
        for r in misses:
            print("   - {:<44} → {}".format(r["prompt"][:44], r["detail"]))

    if board:
        best = board[0][0]
        print("\nTop pick on this hardware: {}".format(best))
    print("")


# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Benchmark local models for T3Lab tool-calling.")
    ap.add_argument("--provider", default="ollama",
                    choices=["ollama", "lmstudio"],
                    help="which local server to benchmark (default: ollama)")
    ap.add_argument("--models", default="",
                    help="comma-separated model names (default: all installed)")
    ap.add_argument("--json", default="",
                    help="also write raw results to this JSON file")
    ap.add_argument("--no-warmup", action="store_true",
                    help="don't do an untimed warm-up call per model")
    args = ap.parse_args()

    try:
        provider = _get_provider(args.provider)
    except Exception as ex:
        print("Could not init provider {}: {}".format(args.provider, ex))
        return 2

    if args.models.strip():
        models = [m.strip() for m in args.models.split(",") if m.strip()]
    else:
        try:
            models = provider.get_models()
        except Exception as ex:
            models = []
            print("Model discovery failed: {}".format(ex))
    if not models:
        print("No models found for provider '{}'. Is the server running and "
              "a model installed/loaded?".format(args.provider))
        return 2

    system_prompt = _system_prompt()
    print("Benchmarking {} model(s) on '{}': {}".format(
        len(models), args.provider, ", ".join(models)))
    print("{} cases/model (VI+EN, tool + small-talk). This can take a while on "
          "CPU.".format(len(CASES)))

    results = {}
    for model in models:
        print("\n>>> {} ...".format(model))
        rows = bench_model(provider, model, system_prompt,
                           warmup=not args.no_warmup)
        summary = summarize(rows)
        results[model] = {"rows": rows, "summary": summary}
        print("    score {}%  tools {}/{}  small-talk {}/{}  hallu {}  {}s/call".format(
            summary["acc"], summary["tool_ok"], summary["tool_n"],
            summary["talk_ok"], summary["talk_n"], summary["hallucinations"],
            summary["avg_s"]))

    print_report(results)

    if args.json.strip():
        try:
            with open(args.json, "w") as f:
                json.dump({"provider": args.provider, "cases": len(CASES),
                           "results": results}, f, indent=2, ensure_ascii=False)
            print("Raw results written to {}".format(args.json))
        except Exception as ex:
            print("Could not write JSON: {}".format(ex))

    return 0


if __name__ == "__main__":
    sys.exit(main())
