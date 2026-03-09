# -*- coding: utf-8 -*-
"""Local LLM integration via Ollama for T3Lab Assistant.

Connects to a locally-running Ollama server (http://localhost:11434)
so T3Lab Assistant works completely offline with no API key and no
rate limits.

Quick start
-----------
1. Install Ollama:  https://ollama.ai
2. Pull a model:
       ollama pull qwen2.5:0.5b      # ~400 MB  fastest
       ollama pull qwen2.5:1.5b      # ~1.0 GB  balanced
       ollama pull llama3.2:1b       # ~1.3 GB  good
       ollama pull phi3:mini         # ~2.3 GB  best quality
3. Start Ollama (it auto-starts on most systems after install).
4. Open T3Lab Assistant — the "LOCAL" badge will appear in the header.

Environment variable
--------------------
  OLLAMA_HOST  override the server address (default http://localhost:11434)
"""
from __future__ import unicode_literals

import json
import os
import re

# ─── HTTP back-end selection ───────────────────────────────────────────────────
# Try .NET WebClient first (available in IronPython / pyRevit)
# then fall back to standard-library urllib.

_USE_NET = False
try:
    import clr
    clr.AddReference('System.Net')
    from System.Net import WebClient, WebException, ServicePointManager
    from System.Text import Encoding as _NetEncoding
    _USE_NET = True
except Exception:
    pass

_HAS_URLLIB = False
if not _USE_NET:
    try:
        from urllib2 import urlopen, Request  # Python 2 / IronPython
        _HAS_URLLIB = True
    except ImportError:
        try:
            from urllib.request import urlopen, Request  # Python 3
            _HAS_URLLIB = True
        except Exception:
            pass


# ─── Configuration ─────────────────────────────────────────────────────────────

OLLAMA_HOST    = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
TIMEOUT_GEN    = 60   # seconds — generation call
TIMEOUT_PROBE  = 3    # seconds — availability ping

# Preferred models smallest/fastest first.
# T3Lab Assistant will pick the first one that is installed.
PREFERRED_MODELS = [
    "qwen2.5:0.5b",
    "qwen2.5:1.5b",
    "llama3.2:1b",
    "phi3:mini",
    "gemma2:2b",
    "qwen2.5:3b",
    "llama3.2:3b",
    "mistral:7b",
    "qwen2.5:7b",
    "llama3:8b",
]

# ─── System prompt (tuned for small models: concise, structured) ───────────────

SYSTEM_PROMPT = u"""\
You are T3Lab Assistant for Autodesk Revit. Your job: read user input and \
return a single JSON object. No explanation, no markdown, only JSON.

INTENTS (pick the best one):
  export_direct           - export/print sheets without opening UI
  open_batchout_configured- open BatchOut pre-configured
  open_batchout           - open BatchOut (no config)
  open_parasync           - open ParaSync
  open_loadfamily         - open Load Family
  open_loadfamily_cloud   - open Load Family (Cloud)
  open_projectname        - open Project Name
  open_workset            - open Workset manager
  open_dimtext            - edit dimension text
  open_upperdimtext       - edit upper dimension text
  open_resetoverrides     - reset graphic overrides
  open_grids              - open Grid tool
  help                    - answer a question about T3Lab
  greet                   - reply to a greeting
  chat                    - general conversation
  unknown                 - cannot understand

PARAMS for export_direct / open_batchout_configured:
  format: "pdf"|"dwg"|"dwf"|"dgn"|"ifc"|"nwd"|"img"  (default "pdf")
  filter: sheet-prefix letter like "G", "A", "S" — or "" for all sheets
  combine: false

RULES:
- "xuất/export + format" with no "mở/open" → export_direct
- "mở batchout" + extra info → open_batchout_configured
- "mở batchout" alone → open_batchout
- Reply in the SAME language as the user (Vietnamese or English).

OUTPUT (JSON only, no other text):
{"intent":"<intent>","params":{<params>},"message":"<short friendly reply>"}

EXAMPLES:
input:  xuất pdf G sheet
output: {"intent":"export_direct","params":{"format":"pdf","filter":"G","combine":false},"message":"Đang xuất G sheet sang PDF..."}

input:  export all sheets to dwg
output: {"intent":"export_direct","params":{"format":"dwg","filter":"","combine":false},"message":"Exporting all sheets to DWG..."}

input:  mở batchout
output: {"intent":"open_batchout","params":{},"message":"Đang mở BatchOut..."}

input:  mở batchout G sheet pdf
output: {"intent":"open_batchout_configured","params":{"format":"pdf","filter":"G"},"message":"Mở BatchOut với G sheet..."}

input:  parasync
output: {"intent":"open_parasync","params":{},"message":"Đang mở ParaSync..."}

input:  load family
output: {"intent":"open_loadfamily","params":{},"message":"Đang mở Load Family..."}

input:  hello
output: {"intent":"greet","params":{},"message":"Hello! I'm T3Lab Assistant. How can I help?"}

input:  batchout là gì
output: {"intent":"help","params":{"answer":"BatchOut xuất hàng loạt sheets sang PDF, DWG, DWF..."},"message":"BatchOut là công cụ xuất sheets hàng loạt."}

input:  cảm ơn
output: {"intent":"chat","params":{},"message":"Không có gì! Cần gì cứ hỏi nhé."}
"""


# ─── HTTP helpers ──────────────────────────────────────────────────────────────

def _post_json(url, payload, timeout=TIMEOUT_GEN):
    """POST a JSON-serialisable payload; return response string."""
    body = json.dumps(payload, ensure_ascii=False)
    if isinstance(body, type(u"")):
        body_bytes = body.encode("utf-8")
    else:
        body_bytes = body

    if _USE_NET:
        client = WebClient()
        client.Encoding = _NetEncoding.UTF8
        client.Headers.Add("Content-Type", "application/json; charset=utf-8")
        resp = client.UploadData(url, "POST", body_bytes)
        return _NetEncoding.UTF8.GetString(resp)

    if _HAS_URLLIB:
        req = Request(url, body_bytes,
                      {"Content-Type": "application/json; charset=utf-8"})
        resp = urlopen(req, timeout=timeout)
        raw = resp.read()
        return raw.decode("utf-8") if isinstance(raw, bytes) else raw

    raise RuntimeError("No HTTP client available")


def _get_text(url, timeout=TIMEOUT_PROBE):
    """GET url; return response string or None."""
    try:
        if _USE_NET:
            client = WebClient()
            client.Encoding = _NetEncoding.UTF8
            return client.DownloadString(url)
        if _HAS_URLLIB:
            resp = urlopen(url, timeout=timeout)
            raw = resp.read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
    except Exception:
        return None


# ─── Public API ───────────────────────────────────────────────────────────────

def is_running():
    """Return True if Ollama server is reachable."""
    text = _get_text(OLLAMA_HOST + "/api/tags")
    return text is not None


def list_models():
    """Return list of installed model name strings (or empty list)."""
    text = _get_text(OLLAMA_HOST + "/api/tags")
    if not text:
        return []
    try:
        data = json.loads(text)
        return [m.get("name", "") for m in data.get("models", [])]
    except Exception:
        return []


def get_best_model():
    """Return the best available small-model name, or None if none installed."""
    installed = list_models()
    if not installed:
        return None
    # Try preferred list first
    for pref in PREFERRED_MODELS:
        pref_base = pref.split(":")[0]
        for inst in installed:
            if inst == pref or inst.startswith(pref_base + ":"):
                return inst
    return installed[0]


def parse_command(user_input, history=None, model=None):
    """Ask the local Ollama LLM to parse a natural-language command.

    Args:
        user_input: raw text from user.
        history: list of {role, content} dicts (conversation context).
        model: explicit model name; auto-selects best if None.

    Returns:
        dict with keys {intent, params, message}, or None on failure.
    """
    if model is None:
        model = get_best_model()
    if not model:
        return None

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]

    if history:
        for h in history[-8:]:
            role    = h.get("role", "user")
            content = h.get("content", "")
            if role in ("user", "assistant") and content:
                messages.append({"role": role, "content": content})

    messages.append({"role": "user", "content": user_input})

    try:
        resp_text = _post_json(
            OLLAMA_HOST + "/api/chat",
            {
                "model":   model,
                "messages": messages,
                "stream":  False,
                "format":  "json",          # force JSON output mode
                "options": {
                    "temperature": 0.0,     # deterministic
                    "num_predict": 300,
                },
            },
            timeout=TIMEOUT_GEN,
        )
        data    = json.loads(resp_text)
        content = data.get("message", {}).get("content", "")
        return _extract_json(content)
    except Exception:
        return None


# ─── JSON extraction ──────────────────────────────────────────────────────────

def _extract_json(text):
    """Parse the first JSON object from LLM response text."""
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    m = re.search(r'\{[\s\S]*\}', text)
    if m:
        try:
            return json.loads(m.group())
        except Exception:
            pass
    return None
