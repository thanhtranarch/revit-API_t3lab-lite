# -*- coding: utf-8 -*-
"""T3Lab Assistant NLP module.

Uses Claude API to understand natural-language commands and map them
to T3Lab tool actions. Falls back to keyword matching when API is
not configured.

Features:
- Multi-turn conversation history (natural back-and-forth)
- Self-learning: saves successful patterns to learned_patterns.json
- Keyword fallback with learned-pattern priority
- Supports Vietnamese, English, or mixed input
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
You are T3Lab Assistant, a friendly AI helper built into Autodesk Revit's T3Lab plugin.
You help users both with T3Lab tool commands AND with casual conversation.

AVAILABLE TOOL INTENTS:

  ── Export (smart) ────────────────────────────────────────────────────────────
  export_direct           – export sheets directly WITHOUT opening any UI.
    params: {
      "format": "pdf"|"dwg"|"dwf"|"dgn"|"ifc"|"nwd"|"img",
      "filter": "<UPPERCASE sheet prefix e.g. G, A, S — or empty for all>",
      "combine": false
    }

  open_batchout_configured – open BatchOut pre-configured (sheets selected, format set, Create tab).
    params: { "format": "pdf"|"dwg"|..., "filter": "<prefix or empty>" }

  open_batchout           – open BatchOut with no pre-configuration. params: {}

  ── Other tools ───────────────────────────────────────────────────────────────
  open_parasync       params: {}
  open_loadfamily     params: {}
  open_loadfamily_cloud params: {}
  open_projectname    params: {}
  open_workset        params: {}
  open_dimtext        params: {}
  open_upperdimtext   params: {}
  open_resetoverrides params: {}
  open_grids          params: {}

  ── Conversation ──────────────────────────────────────────────────────────────
  help   – answer a question about T3Lab tools.
    params: {"answer": "<concise answer>"}

  greet  – respond to a greeting (chào, hello, hi, xin chào, hey...).
    params: {}
    message: a warm, short greeting

  chat   – respond naturally to anything that is NOT a tool command
    (questions about Revit, small talk, thanks, follow-ups, etc.)
    params: {}
    message: a helpful, conversational reply in the SAME language as user input

EXPORT RULES:
- "xuất/export + format + filter + sheet" WITHOUT "mở/open" → export_direct
- "mở batchout + filter/format" → open_batchout_configured
- "mở batchout" alone → open_batchout
- Extract sheet prefix: "G sheet"→G, "tờ G"→G, "A sheet"→A etc.
- "toàn bộ"/"tất cả"/"all" → filter = "" (all sheets), default format = pdf

CONVERSATION RULES:
- Use conversation history to understand follow-up questions.
  e.g., user asks "batchout là gì?" then "nó xuất được những gì?" → use context.
- Be concise, friendly, professional. Reply in the same language as the user.
- If unsure between tool and chat → prefer tool if there is a clear keyword.

RESPONSE FORMAT (JSON only, no markdown, no extra text):
{
  "intent": "<intent_name>",
  "params": { ... },
  "message": "<friendly short message in user's language>"
}

EXAMPLES:
  "chào bạn"  → {"intent":"greet","params":{},"message":"Xin chào! Tôi là T3Lab Assistant. Cần giúp gì không?"}
  "hello"     → {"intent":"greet","params":{},"message":"Hello! I'm T3Lab Assistant. How can I help?"}
  "cảm ơn"    → {"intent":"chat","params":{},"message":"Không có gì! Nếu cần gì cứ hỏi nhé 😊"}
  "batchout làm gì?" → {"intent":"help","params":{"answer":"BatchOut giúp xuất hàng loạt sheets sang PDF, DWG, DWF... với nhiều tùy chọn nâng cao."},"message":"BatchOut là công cụ xuất sheets hàng loạt."}
  "xuất pdf toàn bộ G sheet" → {"intent":"export_direct","params":{"format":"pdf","filter":"G","combine":false},"message":"Đang xuất tất cả G sheet sang PDF..."}
  "mở batchout G sheet pdf"  → {"intent":"open_batchout_configured","params":{"format":"pdf","filter":"G"},"message":"Mở BatchOut với G sheet đã chọn..."}
  "mở batchout"              → {"intent":"open_batchout","params":{},"message":"Đang mở BatchOut..."}
  "parasync"                 → {"intent":"open_parasync","params":{},"message":"Đang mở ParaSync..."}
"""


# ─── Learned patterns ─────────────────────────────────────────────────────────

def _patterns_file():
    lib_dir = os.path.dirname(os.path.abspath(__file__))
    config_dir = os.path.join(lib_dir, 'config')
    if not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir)
        except Exception:
            pass
    return os.path.join(config_dir, 'learned_patterns.json')


def load_learned_patterns():
    """Load learned patterns from disk. Returns dict {key: data}."""
    try:
        path = _patterns_file()
        if os.path.exists(path):
            with open(path, 'r') as f:
                data = json.load(f)
            return data.get('patterns', {})
    except Exception:
        pass
    return {}


def save_learned_patterns(patterns):
    """Persist learned patterns to disk."""
    try:
        path = _patterns_file()
        with open(path, 'w') as f:
            json.dump({'patterns': patterns}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def learn_pattern(raw, intent, params, message=''):
    """Record a successful command→intent mapping.

    Only learns tool intents (not greet/chat/help).
    """
    _skip = {'help', 'chat', 'greet', 'unknown', None}
    if intent in _skip:
        return
    try:
        key = _normalize_key(raw)
        if not key or len(key.split()) < 1:
            return
        patterns = load_learned_patterns()
        if key in patterns:
            patterns[key]['hits'] = patterns[key].get('hits', 0) + 1
            patterns[key]['last_raw'] = raw
        else:
            patterns[key] = {
                'intent': intent,
                'params': params or {},
                'hits': 1,
                'last_raw': raw,
                'last_message': message,
            }
        # Keep only top 200 patterns (prune least-used)
        if len(patterns) > 200:
            sorted_keys = sorted(patterns.keys(), key=lambda k: patterns[k].get('hits', 0))
            for k in sorted_keys[:50]:
                del patterns[k]
        save_learned_patterns(patterns)
    except Exception:
        pass


def find_learned_match(raw):
    """Check learned patterns for a fuzzy match.

    Returns result dict {intent, params, message} or None.
    Uses Jaccard similarity on normalized word sets (threshold 0.65).
    """
    try:
        patterns = load_learned_patterns()
        if not patterns:
            return None
        key = _normalize_key(raw)
        key_words = set(key.split()) if key else set()
        if not key_words:
            return None

        best_score = 0.0
        best_data  = None

        for stored_key, data in patterns.items():
            stored_words = set(stored_key.split()) if stored_key else set()
            if not stored_words:
                continue
            inter = len(key_words & stored_words)
            union = len(key_words | stored_words)
            score = inter / union if union else 0.0
            if score > best_score:
                best_score = score
                best_data  = data

        if best_score >= 0.65 and best_data:
            return {
                'intent':  best_data['intent'],
                'params':  best_data.get('params', {}),
                'message': best_data.get('last_message', ''),
                '_learned': True,
            }
    except Exception:
        pass
    return None


def _normalize_key(text):
    """Normalize to a lookup key: lowercase, no diacritics, meaningful words sorted."""
    try:
        # Remove Vietnamese diacritics via unicode normalization
        import unicodedata
        nfd = unicodedata.normalize('NFD', text)
        ascii_text = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
    except Exception:
        ascii_text = text
    # Lowercase, keep alphanumeric
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', ' ', ascii_text.lower())
    # Keep words longer than 2 chars, sort for order-independence
    words = sorted(w for w in cleaned.split() if len(w) > 2)
    return ' '.join(words)


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


def parse_command(user_input, history=None):
    """Call Claude API to parse a natural-language command.

    Args:
        user_input: The raw text the user typed.
        history: optional list of previous {"role", "content"} dicts (max 8 exchanges).

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

        # Build messages with history for multi-turn context
        messages = []
        if history:
            # Include last 8 exchanges (16 messages) to keep tokens low
            messages.extend(history[-16:])
        messages.append({"role": "user", "content": user_input})

        body_data = {
            "model": "claude-haiku-4-5-20251001",
            "max_tokens": 400,
            "system": SYSTEM_PROMPT,
            "messages": messages,
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

    Checks learned patterns first (priority), then hardcoded keywords.
    Returns dict {intent, params, message} or None.
    """
    cmd = raw.lower().strip()
    viet = _is_viet(cmd)

    # ── Learned patterns (highest priority) ──────────────────────────────────
    learned = find_learned_match(raw)
    if learned:
        # Generate a fresh message so it sounds natural
        intent = learned['intent']
        if intent in _TOOL_LABELS:
            label = _TOOL_LABELS[intent]
            learned['message'] = (u"Đang mở {}...".format(label) if viet
                                  else u"Opening {}...".format(label))
        return learned

    # ── Greetings ─────────────────────────────────────────────────────────────
    greet_kws = ['chao', 'chào', 'hello', 'hi ', 'hey ', 'xin chao', 'good morning',
                 'good afternoon', 'howdy']
    if any(k in cmd for k in greet_kws) or cmd.strip() in ('hi', 'hello', 'hey'):
        if viet:
            msg = u"Xin chào! Tôi là T3Lab Assistant. Cần giúp gì không?"
        else:
            msg = u"Hello! I'm T3Lab Assistant. How can I help?"
        return {"intent": "greet", "params": {}, "message": msg}

    # ── Thanks ────────────────────────────────────────────────────────────────
    if any(k in cmd for k in ['cam on', 'cảm ơn', 'thank', 'thanks', 'cảm ơn bạn']):
        msg = (u"Không có gì! Cần gì cứ hỏi nhé." if viet
               else u"You're welcome! Let me know if you need anything.")
        return {"intent": "chat", "params": {}, "message": msg}

    # ── Export commands ───────────────────────────────────────────────────────
    export_kws = ['xuat', 'xuất', 'export', 'in ra', 'print']
    is_export = any(k in cmd for k in export_kws)
    is_open   = any(k in cmd for k in ['mo ', 'mở ', 'open ', 'launch '])

    if is_export and not is_open:
        params = _parse_export_params(raw, cmd)
        filt  = params['filter']
        fmt   = params['format'].upper()
        filt_label = u" {} sheet".format(filt) if filt else u" tất cả sheet"
        msg = (u"Đang xuất{} sang {}...".format(filt_label, fmt) if viet
               else u"Exporting{} to {}...".format(filt_label, fmt))
        return {"intent": "export_direct", "params": params, "message": msg}

    if is_open and "batchout" in cmd:
        params = _parse_export_params(raw, cmd)
        if params.get('filter') or params.get('format', 'pdf') != 'pdf':
            filt  = params['filter']
            label = u" {} sheet".format(filt) if filt else u" tất cả sheet"
            msg = (u"Mở BatchOut với{} đã chọn...".format(label) if viet
                   else u"Opening BatchOut pre-configured{}...".format(label))
            return {"intent": "open_batchout_configured", "params": params, "message": msg}

    # ── Tool keywords ─────────────────────────────────────────────────────────
    if any(k in cmd for k in ["batchout", "batch out"]):
        return {"intent": "open_batchout", "params": {},
                "message": u"Đang mở BatchOut..." if viet else "Opening BatchOut..."}

    if any(k in cmd for k in ["parasync", "para sync", "dong bo", "đồng bộ",
                               "sync param", "parameter sync"]):
        return {"intent": "open_parasync", "params": {},
                "message": u"Đang mở ParaSync..." if viet else "Opening ParaSync..."}

    if any(k in cmd for k in ["load family cloud", "tai family cloud"]):
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


# ─── Tool labels for auto-generated messages ─────────────────────────────────

_TOOL_LABELS = {
    'open_batchout':          'BatchOut',
    'open_parasync':          'ParaSync',
    'open_loadfamily':        'Load Family',
    'open_loadfamily_cloud':  'Load Family (Cloud)',
    'open_projectname':       'Project Name',
    'open_workset':           'Workset',
    'open_dimtext':           'Dim Text',
    'open_upperdimtext':      'Upper Dim Text',
    'open_resetoverrides':    'Reset Overrides',
    'open_grids':             'Grids',
}


# ─── Export param extraction ──────────────────────────────────────────────────

def _parse_export_params(raw, cmd=None):
    """Extract format and filter from a raw export command string."""
    if cmd is None:
        cmd = raw.lower()

    fmt = 'pdf'
    for f in ['dwg', 'dwf', 'dgn', 'ifc', 'nwd', 'img', 'image', 'pdf']:
        if f in cmd:
            fmt = f
            break

    m = re.search(r'\b([A-Z])\s*[-–]?\s*(?:sheet|tờ|bản\s*vẽ)', raw, re.IGNORECASE)
    if m:
        return {'format': fmt, 'filter': m.group(1).upper(), 'combine': False}

    m = re.search(r'(?:sheet|tờ)\s+([A-Z])\b', raw, re.IGNORECASE)
    if m:
        return {'format': fmt, 'filter': m.group(1).upper(), 'combine': False}

    _ignore = {'PDF', 'DWG', 'DWF', 'DGN', 'IFC', 'NWD', 'IMG'}
    for token in raw.split():
        if re.match(r'^[A-Z]$', token) and token not in _ignore:
            return {'format': fmt, 'filter': token, 'combine': False}

    return {'format': fmt, 'filter': '', 'combine': False}


# ─── Language heuristic ───────────────────────────────────────────────────────

def _is_viet(text):
    viet_chars = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
                  u"ụủứừửữựỳỵỷỹ")
    for c in text.lower():
        if c in viet_chars:
            return True
    return False
