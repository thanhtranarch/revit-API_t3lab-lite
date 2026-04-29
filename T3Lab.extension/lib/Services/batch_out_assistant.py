# -*- coding: utf-8 -*-
"""
Batch Out Assistant

Natural language command parser for batch export operations.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Batch Out Assistant"

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
You are BatchOut Assistant, an AI helper built into Autodesk Revit's batch export tool.
Your job is to parse a user's free-form command and map it to one of the actions below.

AVAILABLE INTENTS (choose exactly one):
  select_sheets    – select sheets/views. params: {"keyword": "<text>"}  (use "all" to select everything)
  deselect_sheets  – deselect sheets/views. params: {"keyword": "<text>"}  (use "all" to clear all)
  filter_sheets    – filter/search the list. params: {"keyword": "<text>"}
  clear_filter     – clear the search filter. params: {}
  enable_format    – enable an export format. params: {"format": "pdf|dwg|dwf|dgn|img|all"}
  disable_format   – disable an export format. params: {"format": "pdf|dwg|dwf|dgn|img|all"}
  navigate_tab     – switch to a tab. params: {"tab": "selection|format|create"}
  unknown          – command is unclear. params: {"message": "<brief explanation why>"}

RULES:
- Accept commands in Vietnamese, English, or mixed.
- "chon", "chọn", "select" → select_sheets
- "bo chon", "bỏ chọn", "deselect", "uncheck" → deselect_sheets
- "loc", "lọc", "tim", "tìm", "filter", "search", "show" → filter_sheets
- "xoa loc", "xóa lọc", "reset", "tat ca", "tất cả filter", "clear filter" → clear_filter
- "bat", "bật", "enable", "dung", "dùng", "xuat", "xuất", "export" + format name → enable_format
- "tat", "tắt", "disable", "khong", "không", "remove" + format name → disable_format
- Format names: pdf, dwg, dwf, dgn, img/image/png
- "sang tab", "chuyen tab", "go to", "go", "tab", "open" + tab name → navigate_tab
- Tab names: "selection"/"chon"→selection, "format"/"dinh dang"→format, "create"/"tao"→create
- For select/deselect with a keyword, extract the keyword (e.g. "chon sheet tang 1" → keyword="tang 1")
- "chon tat ca" or "select all" → keyword="all"
- "bo chon tat ca" or "deselect all" → keyword="all"

RESPONSE FORMAT (JSON only, no markdown, no extra text):
{
  "intent": "<intent_name>",
  "params": { ... },
  "message": "<friendly short confirmation in the SAME language as the user's input>"
}

Examples:
  "chon tat ca sheet"  → {"intent":"select_sheets","params":{"keyword":"all"},"message":"Đã chọn tất cả sheet"}
  "bo chon tang 1"     → {"intent":"deselect_sheets","params":{"keyword":"tang 1"},"message":"Đã bỏ chọn sheet tầng 1"}
  "loc sheet A"        → {"intent":"filter_sheets","params":{"keyword":"A"},"message":"Đang lọc sheet chứa 'A'"}
  "xoa loc"            → {"intent":"clear_filter","params":{},"message":"Đã xóa bộ lọc"}
  "bat xuat PDF"       → {"intent":"enable_format","params":{"format":"pdf"},"message":"Đã bật xuất PDF"}
  "tat DWG"            → {"intent":"disable_format","params":{"format":"dwg"},"message":"Đã tắt DWG"}
  "xuat tat ca dinh dang" → {"intent":"enable_format","params":{"format":"all"},"message":"Đã bật tất cả định dạng"}
  "sang tab Format"    → {"intent":"navigate_tab","params":{"tab":"format"},"message":"Chuyển sang tab Format"}
  "go to create"       → {"intent":"navigate_tab","params":{"tab":"create"},"message":"Switched to Create tab"}
  "select floor plan"  → {"intent":"select_sheets","params":{"keyword":"floor plan"},"message":"Selected sheets matching 'floor plan'"}
  "enable pdf and dwg" → {"intent":"enable_format","params":{"format":"pdf"},"message":"Enabled PDF (run again for DWG)"}
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
    # Try direct parse first
    try:
        return json.loads(text.strip())
    except Exception:
        pass
    # Find {...} block
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except Exception:
            pass
    return None


def parse_command(user_input, context=""):
    """Call Claude API to parse a natural-language command.

    Args:
        user_input: The raw text the user typed.
        context:    Optional context string (e.g. "mode=sheets, visible=45, selected=12").

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

        user_content = "Command: {}\nContext: {}".format(user_input, context)
        body_data = {
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 300,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": user_content}],
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
