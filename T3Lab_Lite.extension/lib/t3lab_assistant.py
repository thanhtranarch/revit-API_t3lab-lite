# -*- coding: utf-8 -*-
"""T3Lab Assistant NLP module.

Uses Claude API to understand natural-language commands and map them
to T3Lab tool actions. Falls back to keyword matching when API is
not configured.

Supports Tiếng Việt, English, or mixed input.
"""
from __future__ import unicode_literals

import json
import re
import os
import sys

try:
    import clr
    clr.AddReference('System.Net')
    from System.Net import WebClient, WebException
    from System.Text import Encoding
    HAS_HTTP = True
except Exception:
    HAS_HTTP = False


# ─── System prompt ────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are T3Lab Assistant, an AI helper built into Autodesk Revit's T3Lab plugin.
Your job is to parse a user's free-form command and map it to one of the T3Lab tool actions below.

AVAILABLE INTENTS (choose exactly one):
  open_batchout       – open the BatchOut batch export dialog. params: {}
  open_parasync       – open the ParaSync parameter sync tool. params: {}
  open_loadfamily     – open the Load Family dialog. params: {}
  open_loadfamily_cloud – open the Load Family (Cloud) dialog. params: {}
  open_projectname    – open the Project Name tool. params: {}
  open_workset        – open the Workset manager. params: {}
  open_dimtext        – open the Dim Text tool (edit dimension text). params: {}
  open_upperdimtext   – open the Upper Dim Text tool. params: {}
  open_resetoverrides – open the Reset Overrides tool. params: {}
  open_grids          – open the Grids tool. params: {}
  help                – answer a question about T3Lab tools. params: {"answer": "<concise answer>"}
  unknown             – command is unclear. params: {"message": "<brief explanation why>"}

RULES:
- Accept commands in Vietnamese, English, or mixed.
- "batch", "batchout", "xuất sheet", "export sheet", "xuất hàng loạt" → open_batchout
- "parasync", "đồng bộ", "sync param", "parameter sync" → open_parasync
- "load family", "tải family", "nạp family" → open_loadfamily
- "load family cloud", "tải family cloud" → open_loadfamily_cloud
- "project name", "tên project", "đặt tên" → open_projectname
- "workset", "quản lý workset" → open_workset
- "dim text", "dimtext", "sửa dimension", "edit dim" → open_dimtext
- "upper dim", "upperdimtext" → open_upperdimtext
- "reset override", "xóa override", "bỏ override" → open_resetoverrides
- "grids", "lưới", "trục" → open_grids
- Questions about tools → help

RESPONSE FORMAT (JSON only, no markdown, no extra text):
{
  "intent": "<intent_name>",
  "params": { ... },
  "message": "<friendly short confirmation in the SAME language as the user's input>"
}

Examples:
  "mở batchout"         → {"intent":"open_batchout","params":{},"message":"Đang mở BatchOut..."}
  "xuất sheet PDF"      → {"intent":"open_batchout","params":{},"message":"Mở BatchOut để xuất PDF"}
  "parasync"            → {"intent":"open_parasync","params":{},"message":"Đang mở ParaSync..."}
  "load family"         → {"intent":"open_loadfamily","params":{},"message":"Đang mở Load Family..."}
  "open batchout"       → {"intent":"open_batchout","params":{},"message":"Opening BatchOut..."}
  "what is parasync?"   → {"intent":"help","params":{"answer":"ParaSync syncs parameters across linked models in your project."},"message":"ParaSync syncs parameters across linked models."}
"""


# ─── API helpers ──────────────────────────────────────────────────────────────

def _get_api_key():
    """Retrieve Claude API key from T3LabAI settings."""
    try:
        lib_dir = os.path.dirname(os.path.abspath(__file__))
        if lib_dir not in sys.path:
            sys.path.insert(0, lib_dir)
        from config.settings import T3LabAISettings
        return T3LabAISettings().get_api_key("Claude")
    except Exception:
        return None


def _extract_json(text):
    """Extract the first JSON object from text, tolerating surrounding noise."""
    try:
        return json.loads(text.strip())
    except Exception:
        pass
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    return None


def parse_command(user_input):
    """Call Claude API to parse a natural-language command.

    Args:
        user_input: The raw text the user typed.

    Returns:
        dict with keys {intent, params, message} on success, or None on failure.
    """
    if not HAS_HTTP:
        return None

    api_key = _get_api_key()
    if not api_key:
        return None

    try:
        url = "https://api.anthropic.com/v1/messages"

        body_data = {
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 300,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": user_input}],
        }
        body_json = json.dumps(body_data, ensure_ascii=False)
        body_bytes = Encoding.UTF8.GetBytes(body_json)

        client = WebClient()
        client.Encoding = Encoding.UTF8
        client.Headers.Add("Content-Type", "application/json; charset=utf-8")
        client.Headers.Add("x-api-key", api_key)
        client.Headers.Add("anthropic-version", "2023-06-01")

        response_bytes = client.UploadData(url, "POST", body_bytes)
        response_text = Encoding.UTF8.GetString(response_bytes)

        api_result = json.loads(response_text)
        content_text = api_result["content"][0]["text"].strip()

        return _extract_json(content_text)

    except Exception:
        return None


def has_api_key():
    """Return True if a Claude API key is configured."""
    return bool(_get_api_key())


# ─── Keyword fallback ─────────────────────────────────────────────────────────

def keyword_parse(raw):
    """Keyword-based fallback parser.

    Returns dict with {intent, params, message} or None.
    """
    cmd = raw.lower().strip()

    # BatchOut
    if any(k in cmd for k in ["batchout", "batch out", "xuat sheet", "xuất sheet",
                                "export sheet", "xuat hang loat", "xuất hàng loạt",
                                "xuat pdf", "export pdf", "bat xuat"]):
        return {"intent": "open_batchout", "params": {},
                "message": "Đang mở BatchOut..." if _is_viet(cmd) else "Opening BatchOut..."}

    # ParaSync
    if any(k in cmd for k in ["parasync", "para sync", "dong bo", "đồng bộ",
                                "sync param", "parameter sync"]):
        return {"intent": "open_parasync", "params": {},
                "message": "Đang mở ParaSync..." if _is_viet(cmd) else "Opening ParaSync..."}

    # Load Family Cloud
    if any(k in cmd for k in ["load family cloud", "tai family cloud", "tải family cloud",
                                "nap family cloud"]):
        return {"intent": "open_loadfamily_cloud", "params": {},
                "message": "Đang mở Load Family (Cloud)..." if _is_viet(cmd) else "Opening Load Family (Cloud)..."}

    # Load Family
    if any(k in cmd for k in ["load family", "tai family", "tải family",
                                "nap family", "nạp family"]):
        return {"intent": "open_loadfamily", "params": {},
                "message": "Đang mở Load Family..." if _is_viet(cmd) else "Opening Load Family..."}

    # Project Name
    if any(k in cmd for k in ["project name", "ten project", "tên project",
                                "dat ten", "đặt tên"]):
        return {"intent": "open_projectname", "params": {},
                "message": "Đang mở Project Name..." if _is_viet(cmd) else "Opening Project Name..."}

    # Workset
    if any(k in cmd for k in ["workset", "quan ly workset", "quản lý workset"]):
        return {"intent": "open_workset", "params": {},
                "message": "Đang mở Workset..." if _is_viet(cmd) else "Opening Workset..."}

    # Upper Dim Text
    if any(k in cmd for k in ["upper dim", "upperdimtext", "upper dimension"]):
        return {"intent": "open_upperdimtext", "params": {},
                "message": "Đang mở Upper Dim Text..." if _is_viet(cmd) else "Opening Upper Dim Text..."}

    # Dim Text
    if any(k in cmd for k in ["dim text", "dimtext", "sua dimension", "sửa dimension",
                                "edit dim"]):
        return {"intent": "open_dimtext", "params": {},
                "message": "Đang mở Dim Text..." if _is_viet(cmd) else "Opening Dim Text..."}

    # Reset Overrides
    if any(k in cmd for k in ["reset override", "xoa override", "xóa override",
                                "bo override", "bỏ override", "reset graphic"]):
        return {"intent": "open_resetoverrides", "params": {},
                "message": "Đang mở Reset Overrides..." if _is_viet(cmd) else "Opening Reset Overrides..."}

    # Grids
    if any(k in cmd for k in ["grids", "luoi", "lưới", "truc", "trục", "grid tool"]):
        return {"intent": "open_grids", "params": {},
                "message": "Đang mở Grids..." if _is_viet(cmd) else "Opening Grids..."}

    return None


def _is_viet(text):
    """Heuristic: return True if text appears to be Vietnamese."""
    viet_chars = u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ"
    viet_words = ["cua", "cua", "toi", "ban", "mo", "dung", "xuat", "chon",
                  "tat", "bat", "tab", "sheet", "sang", "loc", "bo", "them"]
    low = text.lower()
    for c in low:
        if c in viet_chars:
            return True
    for w in viet_words:
        if w in low.split():
            return True
    return False
