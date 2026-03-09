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

  ── Export (smart) ────────────────────────────────────────────────────────────
  export_direct           – export sheets directly WITHOUT opening any UI.
    params: {
      "format": "pdf"|"dwg"|"dwf"|"dgn"|"ifc"|"nwd"|"img",
      "filter": "<UPPERCASE sheet-number prefix e.g. G, A, S, or empty string for all>",
      "combine": false
    }

  open_batchout_configured – open BatchOut pre-configured (sheets selected, format set,
                             navigate to Create tab ready to export).
    params: {
      "format": "pdf"|"dwg"|"dwf"|"dgn"|"ifc"|"nwd"|"img",
      "filter": "<UPPERCASE sheet-number prefix e.g. G, A, S, or empty string for all>"
    }

  open_batchout           – open BatchOut with no pre-configuration. params: {}

  ── Other tools ───────────────────────────────────────────────────────────────
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

EXPORT RULES:
- If the user says "xuất/export + [format] + [filter] + sheet/tờ" WITHOUT "mở"/"open" → export_direct
- If the user says "mở batchout" + filter/format details → open_batchout_configured
- If the user says "mở batchout" with NO details → open_batchout
- Extract the sheet prefix letter (A, G, S, M, E, P...) from patterns like "G sheet", "tờ G", "G-sheet", "sheet G"
- "toàn bộ", "tất cả", "all sheets", "hết" → filter = "" (all sheets)
- Default format is "pdf" when only sheets are mentioned without explicit format

EXPORT EXAMPLES:
  "xuất pdf toàn bộ G sheet"   → export_direct  filter="G"  format="pdf"
  "export pdf all G sheets"    → export_direct  filter="G"  format="pdf"
  "xuất tất cả sheet ra pdf"   → export_direct  filter=""   format="pdf"
  "export all sheets DWG"      → export_direct  filter=""   format="dwg"
  "xuất sheet A ra DWG"        → export_direct  filter="A"  format="dwg"
  "mở batchout G sheet pdf"    → open_batchout_configured  filter="G"  format="pdf"
  "open batchout select A sheets" → open_batchout_configured filter="A" format="pdf"
  "mở batchout"                → open_batchout

OTHER RULES:
- "parasync", "đồng bộ", "sync param" → open_parasync
- "load family", "tải family" → open_loadfamily
- "load family cloud" → open_loadfamily_cloud
- "project name", "tên project" → open_projectname
- "workset" → open_workset
- "dim text", "dimtext" → open_dimtext
- "upper dim" → open_upperdimtext
- "reset override" → open_resetoverrides
- "grids", "lưới" → open_grids
- Questions about tools → help

RESPONSE FORMAT (JSON only, no markdown, no extra text):
{
  "intent": "<intent_name>",
  "params": { ... },
  "message": "<friendly short confirmation in the SAME language as user input>"
}

RESPONSE EXAMPLES:
  "xuất pdf toàn bộ G sheet"
    → {"intent":"export_direct","params":{"format":"pdf","filter":"G","combine":false},"message":"Đang xuất tất cả G sheet sang PDF..."}

  "export all A sheets to DWG"
    → {"intent":"export_direct","params":{"format":"dwg","filter":"A","combine":false},"message":"Exporting all A sheets to DWG..."}

  "mở batchout chọn G sheet pdf"
    → {"intent":"open_batchout_configured","params":{"format":"pdf","filter":"G"},"message":"Mở BatchOut với G sheet đã chọn và PDF đã bật..."}

  "mở batchout"
    → {"intent":"open_batchout","params":{},"message":"Đang mở BatchOut..."}

  "parasync"
    → {"intent":"open_parasync","params":{},"message":"Đang mở ParaSync..."}
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

    Returns dict with {intent, params, message} or None on failure.
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
            "max_tokens": 400,
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
    viet = _is_viet(cmd)

    # ── Export commands (detect before generic batchout) ──────────────────────
    export_kws = [
        "xuat", "xuất", "export", "in ra", "in ", "print"
    ]
    is_export_cmd = any(k in cmd for k in export_kws)

    open_kws = ["mo ", "mở ", "open ", "launch "]
    is_open_cmd = any(k in cmd for k in open_kws) or "batchout" in cmd and not is_export_cmd

    if is_export_cmd and not is_open_cmd:
        # Direct export
        params = _parse_export_params(raw, cmd)
        filt = params['filter']
        fmt  = params['format'].upper()
        filt_label = u" {} sheet".format(filt) if filt else u" tất cả sheet"
        if viet:
            msg = u"Đang xuất{} sang {}...".format(filt_label, fmt)
        else:
            msg = u"Exporting{} to {}...".format(filt_label, fmt)
        return {"intent": "export_direct", "params": params, "message": msg}

    if is_open_cmd and "batchout" in cmd:
        # Check if has filter/format details → configured
        params = _parse_export_params(raw, cmd)
        if params.get('filter') or params.get('format') != 'pdf':
            filt = params['filter']
            filt_label = u" {} sheet".format(filt) if filt else u" tất cả sheet"
            if viet:
                msg = u"Mở BatchOut với{} đã chọn...".format(filt_label)
            else:
                msg = u"Opening BatchOut pre-configured{}...".format(filt_label)
            return {"intent": "open_batchout_configured", "params": params, "message": msg}

    # ── Generic batchout ──────────────────────────────────────────────────────
    if any(k in cmd for k in ["batchout", "batch out"]):
        return {"intent": "open_batchout", "params": {},
                "message": u"Đang mở BatchOut..." if viet else "Opening BatchOut..."}

    # ── Other tools ───────────────────────────────────────────────────────────
    if any(k in cmd for k in ["parasync", "para sync", "dong bo", "đồng bộ",
                               "sync param", "parameter sync"]):
        return {"intent": "open_parasync", "params": {},
                "message": u"Đang mở ParaSync..." if viet else "Opening ParaSync..."}

    if any(k in cmd for k in ["load family cloud", "tai family cloud", "tải family cloud"]):
        return {"intent": "open_loadfamily_cloud", "params": {},
                "message": u"Đang mở Load Family (Cloud)..." if viet else "Opening Load Family (Cloud)..."}

    if any(k in cmd for k in ["load family", "tai family", "tải family", "nap family"]):
        return {"intent": "open_loadfamily", "params": {},
                "message": u"Đang mở Load Family..." if viet else "Opening Load Family..."}

    if any(k in cmd for k in ["project name", "ten project", "tên project", "dat ten"]):
        return {"intent": "open_projectname", "params": {},
                "message": u"Đang mở Project Name..." if viet else "Opening Project Name..."}

    if any(k in cmd for k in ["workset", "quan ly workset"]):
        return {"intent": "open_workset", "params": {},
                "message": u"Đang mở Workset..." if viet else "Opening Workset..."}

    if any(k in cmd for k in ["upper dim", "upperdimtext"]):
        return {"intent": "open_upperdimtext", "params": {},
                "message": u"Đang mở Upper Dim Text..." if viet else "Opening Upper Dim Text..."}

    if any(k in cmd for k in ["dim text", "dimtext", "sua dimension", "edit dim"]):
        return {"intent": "open_dimtext", "params": {},
                "message": u"Đang mở Dim Text..." if viet else "Opening Dim Text..."}

    if any(k in cmd for k in ["reset override", "xoa override", "bo override", "reset graphic"]):
        return {"intent": "open_resetoverrides", "params": {},
                "message": u"Đang mở Reset Overrides..." if viet else "Opening Reset Overrides..."}

    if any(k in cmd for k in ["grids", "luoi", "lưới", "truc", "grid tool"]):
        return {"intent": "open_grids", "params": {},
                "message": u"Đang mở Grids..." if viet else "Opening Grids..."}

    return None


# ─── Export param extraction ──────────────────────────────────────────────────

def _parse_export_params(raw, cmd=None):
    """Extract format and filter from a raw export command string."""
    if cmd is None:
        cmd = raw.lower()

    # Detect format
    fmt = 'pdf'  # default
    for f in ['dwg', 'dwf', 'dgn', 'ifc', 'nwd', 'img', 'image', 'pdf']:
        if f in cmd:
            fmt = f
            break

    # Detect sheet prefix filter
    # Pattern: uppercase letter + "sheet"/"tờ"/"bản vẽ" in original text
    m = re.search(r'\b([A-Z])\s*[-–]?\s*(?:sheet|tờ|bản\s*vẽ)', raw, re.IGNORECASE)
    if m:
        return {'format': fmt, 'filter': m.group(1).upper(), 'combine': False}

    # "sheet" + uppercase letter
    m = re.search(r'(?:sheet|tờ)\s+([A-Z])\b', raw, re.IGNORECASE)
    if m:
        return {'format': fmt, 'filter': m.group(1).upper(), 'combine': False}

    # Standalone uppercase word that is NOT a format keyword
    _ignore = {'PDF', 'DWG', 'DWF', 'DGN', 'IFC', 'NWD', 'IMG'}
    for token in raw.split():
        if re.match(r'^[A-Z]$', token) and token not in _ignore:
            return {'format': fmt, 'filter': token, 'combine': False}

    # "toàn bộ" / "tất cả" / "all" → no filter
    return {'format': fmt, 'filter': '', 'combine': False}


# ─── Language heuristic ───────────────────────────────────────────────────────

def _is_viet(text):
    """Return True if text appears to be Vietnamese."""
    viet_chars = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
                  u"ụủứừửữựỳỵỷỹ")
    for c in text.lower():
        if c in viet_chars:
            return True
    return False
