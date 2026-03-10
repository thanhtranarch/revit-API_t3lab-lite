# -*- coding: utf-8 -*-
"""Built-in NLU (Natural Language Understanding) engine for T3Lab Assistant.

Fully self-contained — no external services, APIs, models, or downloads.
Works offline in IronPython 2.7+, CPython 3.x, and standard pyRevit.

Algorithm
---------
1. Preprocess : strip diacritics → lowercase → expand abbreviations/synonyms
2. Tokenise   : unigrams + bigrams from the normalised text
3. Classify   : weighted feature scoring per intent; pick max above threshold
4. Disambiguate: rule-based tie-breaking (export vs open_batchout_configured)
5. Slot extract: format, filter-prefix, combine flag
6. Context    : pronoun / reference resolution from conversation history
7. Respond    : build a friendly, language-matched message

Supported intents
-----------------
  export_direct · open_batchout_configured · open_batchout
  open_parasync · open_loadfamily · open_loadfamily_cloud
  open_projectname · open_workset · open_dimtext · open_upperdimtext
  open_resetoverrides · open_grids
  greet · chat · help
"""
from __future__ import unicode_literals, division

import re
import unicodedata


# ─── Diacritics helper ────────────────────────────────────────────────────────

def _strip_diacritics(text):
    """Remove all combining diacritic marks (works for Vietnamese, etc.)."""
    try:
        nfd = unicodedata.normalize('NFD', text)
        return ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
    except Exception:
        return text


def _norm(text):
    """Normalise to ASCII-lower for matching: strip diacritics + lowercase."""
    return _strip_diacritics(text).lower()


# ─── Abbreviation / synonym map ───────────────────────────────────────────────
# Applied AFTER _norm().  Order matters: longer keys first in the replacement
# loop so "batch out" doesn't eat "batch" before "batch out" is matched.

_ABBREVS = [
    # BatchOut
    ("batch out",       "batchout"),
    ("batcho",          "batchout"),
    ("b.out",           "batchout"),
    ("bo ",             "batchout "),   # trailing space avoids "bo override"
    # ParaSync
    ("para sync",       "parasync"),
    ("dong bo tham so", "parasync"),
    ("dong bo",         "parasync"),
    ("sync param",      "parasync"),
    (" ps ",            " parasync "),
    # Load Family Cloud
    ("load fam cloud",  "loadfamilycloud"),
    ("tai family cloud","loadfamilycloud"),
    ("load cloud",      "loadfamilycloud"),
    # Load Family
    ("load fam",        "loadfamily"),
    ("tai family",      "loadfamily"),
    ("nap family",      "loadfamily"),
    (" lf ",            " loadfamily "),
    # Project Name
    ("project name",    "projectname"),
    ("ten project",     "projectname"),
    ("dat ten",         "projectname"),
    ("ten du an",       "projectname"),
    (" pn ",            " projectname "),
    # Workset
    ("quan ly workset", "workset"),
    (" ws ",            " workset "),
    # Upper Dim Text
    ("upper dim text",  "upperdimtext"),
    ("upper dim",       "upperdimtext"),
    (" udt ",           " upperdimtext "),
    # Dim Text
    ("dim text",        "dimtext"),
    ("kich thuoc",      "dimtext"),
    ("sua dim",         "dimtext"),
    ("edit dim",        "dimtext"),
    (" dt ",            " dimtext "),
    # Reset Overrides
    ("reset graphic override","resetoverrides"),
    ("reset override",  "resetoverrides"),
    ("xoa override",    "resetoverrides"),
    ("bo override",     "resetoverrides"),
    # Grids
    ("luoi truc",       "grids"),
    # Export verbs (unify to "export")
    ("in ra",           "export"),
    ("xuat ra",         "export"),
    ("xuat",            "export"),
    # Open verbs (unify to "open")
    ("mo len",          "open"),
    ("bat len",         "open"),
    ("chay len",        "open"),
    ("khoi dong",       "open"),
    # "all / every sheet"
    ("tat ca sheet",    "all sheet"),
    ("toan bo sheet",   "all sheet"),
    ("all sheets",      "all sheet"),
    ("every sheet",     "all sheet"),
    ("toan bo",         "all"),
    ("tat ca",          "all"),
    # Formats
    ("hinh anh",        "img"),
    ("image",           "img"),
    # Vietnamese "open" (without trailing space — catch end of string too)
    ("mo ",             "open "),
    ("mo\n",            "open\n"),
]


def _expand(text):
    """Apply abbreviation / synonym substitutions to normalised text."""
    # Pad with spaces to allow boundary matching
    t = " " + text + " "
    for src, dst in _ABBREVS:
        t = t.replace(src, dst)
    return t.strip()


# ─── Tokeniser ────────────────────────────────────────────────────────────────

_STOPWORDS = {
    "va", "de", "cho", "cai", "len", "ra", "vao", "di",
    "toi", "ban", "minh", "ho", "no", "cua", "voi", "trong",
    "la", "duoc", "co", "khong", "nhe", "nha", "nao", "gi",
    "a", "o", "the", "an", "in", "to", "for", "of", "at",
    "me", "my", "i", "you", "it", "is", "are", "was", "be",
    "and", "or", "not", "on", "up",
}


def _tokenise(text):
    """Return (unigrams, bigrams) as frozensets of normalised strings."""
    # Keep only letters/digits/spaces
    clean = re.sub(r'[^a-z0-9\s]', ' ', text)
    tokens = [w for w in clean.split() if len(w) >= 2 and w not in _STOPWORDS]
    unigrams = frozenset(tokens)
    bigrams  = frozenset(
        tokens[i] + " " + tokens[i + 1] for i in range(len(tokens) - 1)
    )
    return unigrams, bigrams


# ─── Intent trigger tables ────────────────────────────────────────────────────
# Each entry: (feature_string, weight)
# Unigrams and bigrams are checked against the same tables.

_TRIGGERS = {
    "open_batchout": [
        ("batchout",           20),
        ("open batchout",      30),
        ("mo batchout",        30),
        ("open",                3),
    ],

    "export_direct": [
        ("export",             15),
        ("export sheet",       20),
        ("export all",         20),
        ("pdf",                 5),
        ("dwg",                 5),
        ("dwf",                 5),
        ("all sheet",          10),
        ("print",               8),
    ],

    "open_batchout_configured": [
        ("batchout",           15),
        ("open batchout",      20),
        ("mo batchout",        20),
        ("batchout pdf",       25),
        ("batchout dwg",       25),
        ("batchout dwf",       25),
        ("batchout sheet",     20),
        ("batchout all",       15),
        ("open",                3),
    ],

    "open_parasync": [
        ("parasync",           30),
        ("open parasync",      35),
        ("mo parasync",        35),
        ("open",                2),
    ],

    "open_loadfamilycloud": [
        ("loadfamilycloud",    35),
        ("open loadfamilycloud", 40),
        ("mo loadfamilycloud", 40),
    ],

    "open_loadfamily": [
        ("loadfamily",         30),
        ("open loadfamily",    35),
        ("mo loadfamily",      35),
        ("family",              8),
        ("open",                2),
    ],

    "open_projectname": [
        ("projectname",        30),
        ("open projectname",   35),
        ("mo projectname",     35),
        ("project",             8),
        ("open",                2),
    ],

    "open_workset": [
        ("workset",            30),
        ("open workset",       35),
        ("mo workset",         35),
        ("open",                2),
    ],

    "open_upperdimtext": [
        ("upperdimtext",       30),
        ("open upperdimtext",  35),
        ("mo upperdimtext",    35),
        ("upper",               8),
    ],

    "open_dimtext": [
        ("dimtext",            28),
        ("open dimtext",       35),
        ("mo dimtext",         35),
        ("dim",                10),
    ],

    "open_resetoverrides": [
        ("resetoverrides",     30),
        ("open resetoverrides",35),
        ("mo resetoverrides",  35),
        ("reset",              12),
        ("override",           12),
    ],

    "open_grids": [
        ("grids",              30),
        ("open grids",         35),
        ("mo grids",           35),
        ("grid",               12),
        ("luoi",               12),
    ],

    "greet": [
        ("chao",               25),
        ("hello",              25),
        ("xin chao",           30),
        ("hey",                15),
        ("hi",                 20),
        ("howdy",              15),
        ("good morning",       20),
    ],

    "chat": [
        ("cam on",             25),
        ("thank",              20),
        ("thanks",             20),
        ("ok",                  8),
        ("oke",                 8),
        ("got it",             15),
    ],

    "help": [
        ("la gi",              25),
        ("what is",            25),
        ("how to",             20),
        ("lam gi",             20),
        ("lam nhu the nao",    25),
        ("huong dan",          20),
        ("guide",              15),
        ("explain",            15),
        ("giai thich",         20),
        ("help",               10),
    ],
}

# Features that penalise an intent when present
_PENALTIES = {
    # Explicit "open" kills a naked export_direct
    "export_direct":             [("open", -15), ("mo", -15), ("batchout", -5)],
    # open_batchout loses if there's no "open" at all (pure export wins)
    "open_batchout":             [("export", -8)],
    "open_batchout_configured":  [("export", -8)],
    # Avoid lower-precedence dimtext when upper is present
    "open_dimtext":              [("upper", -20)],
    # Avoid loadfamily if cloud is there
    "open_loadfamily":           [("cloud", -20), ("loadfamilycloud", -30)],
}

# Minimum score required to accept an intent
_THRESHOLDS = {
    "open_batchout":             18,
    "export_direct":             18,
    "open_batchout_configured":  25,   # needs both open+batchout+params
    "open_parasync":             18,
    "open_loadfamilycloud":      25,
    "open_loadfamily":           18,
    "open_projectname":          18,
    "open_workset":              18,
    "open_upperdimtext":         22,
    "open_dimtext":              18,
    "open_resetoverrides":       18,
    "open_grids":                18,
    "greet":                     18,
    "chat":                      18,
    "help":                      18,
}


# ─── Slot extraction ──────────────────────────────────────────────────────────

_FORMATS = ["dwg", "dwf", "dgn", "ifc", "nwd", "img", "pdf"]  # pdf last = default

# Uppercase letters that are NOT format abbreviations
_FORMAT_LETTERS = {"PDF", "DWG", "DWF", "DGN", "IFC", "NWD", "IMG"}


def _extract_slots(raw):
    """Extract format, filter, and combine from the original (unicode) text.

    Returns dict: {format: str, filter: str, combine: bool}
    """
    normed = _norm(raw)

    # ── format ──────────────────────────────────────────────────────────────
    fmt = "pdf"
    for f in _FORMATS:
        if f in normed.split() or (" " + f) in normed or (f + " ") in normed:
            fmt = f
            break

    # ── sheet-prefix filter (single uppercase letter, not a format name) ────
    # Priority 1: explicit patterns  "G sheet" / "tờ G" / "G-sheet"
    sheet_filt = ""
    m = re.search(
        r'\b([A-Z])\s*[-–]?\s*(?:sheet|to|to\s|ban\s*ve|sheets)\b',
        raw, re.IGNORECASE
    )
    if m and m.group(1).upper() not in _FORMAT_LETTERS:
        sheet_filt = m.group(1).upper()

    if not sheet_filt:
        m = re.search(r'(?:sheet|to|sheets)\s+([A-Z])\b', raw, re.IGNORECASE)
        if m and m.group(1).upper() not in _FORMAT_LETTERS:
            sheet_filt = m.group(1).upper()

    # Priority 2: lone uppercase token that isn't a format name
    if not sheet_filt:
        for token in raw.split():
            tok = re.sub(r'[^A-Z]', '', token.upper())
            if len(tok) == 1 and tok not in _FORMAT_LETTERS:
                sheet_filt = tok
                break

    # ── combine flag ─────────────────────────────────────────────────────────
    combine_kws = ["combine", "merge", "gop", "ghep", "1 file", "mot file"]
    combine = any(k in normed for k in combine_kws)

    return {"format": fmt, "filter": sheet_filt, "combine": combine}


# ─── Context / pronoun resolution ─────────────────────────────────────────────

# Pronouns that refer to the most-recently-mentioned tool
_PRONOUNS = {"no", "no ay", "cai do", "cai nay", "tool do", "it", "that", "this"}

# Maps intent → tool label (for pronoun resolution messages)
_TOOL_LABELS = {
    "open_batchout":          "BatchOut",
    "export_direct":          "BatchOut",
    "open_batchout_configured":"BatchOut",
    "open_parasync":          "ParaSync",
    "open_loadfamily":        "Load Family",
    "open_loadfamilycloud":   "Load Family Cloud",
    "open_projectname":       "Project Name",
    "open_workset":           "Workset",
    "open_dimtext":           "Dim Text",
    "open_upperdimtext":      "Upper Dim Text",
    "open_resetoverrides":    "Reset Overrides",
    "open_grids":             "Grids",
}

# Tool keywords used to detect last-mentioned tool in history
_TOOL_KEYWORDS = {
    "batchout":       "open_batchout",
    "parasync":       "open_parasync",
    "loadfamilycloud":"open_loadfamilycloud",
    "loadfamily":     "open_loadfamily",
    "projectname":    "open_projectname",
    "workset":        "open_workset",
    "upperdimtext":   "open_upperdimtext",
    "dimtext":        "open_dimtext",
    "resetoverrides": "open_resetoverrides",
    "grids":          "open_grids",
}


def _last_tool_from_history(history):
    """Scan recent conversation history and return the last tool intent mentioned."""
    if not history:
        return None
    for entry in reversed(history[-6:]):
        content = _norm(_expand(_norm(entry.get("content", ""))))
        for kw, intent in _TOOL_KEYWORDS.items():
            if kw in content:
                return intent
    return None


def _is_pronoun_query(normed_expanded):
    """Return True if the input looks like a pronoun reference (e.g., 'nó là gì?')."""
    tokens = set(normed_expanded.split())
    return bool(tokens & _PRONOUNS)


# ─── Scoring ──────────────────────────────────────────────────────────────────

def _score(intent, unigrams, bigrams):
    """Compute weighted match score for one intent."""
    score = 0
    all_features = unigrams | bigrams
    for feature, weight in _TRIGGERS.get(intent, []):
        if feature in all_features:
            score += weight
    for feature, penalty in _PENALTIES.get(intent, []):
        if feature in all_features:
            score += penalty   # penalty is already negative
    return score


# ─── Disambiguation rules ─────────────────────────────────────────────────────

def _disambiguate(scores, unigrams, bigrams, slots):
    """Apply domain rules to resolve ambiguous intent scores.

    Returns the final intent string or None.
    """
    all_features = unigrams | bigrams

    # ── open_batchout vs open_batchout_configured ────────────────────────────
    # open_batchout_configured requires batchout + at least one config param
    has_batchout = "batchout" in unigrams
    has_config   = bool(slots["filter"]) or (slots["format"] != "pdf")

    if has_batchout and "open" in all_features and has_config:
        # Promote open_batchout_configured
        scores["open_batchout_configured"] = max(
            scores.get("open_batchout_configured", 0),
            scores.get("open_batchout", 0) + 10
        )
        scores["open_batchout"] -= 10

    # ── export_direct vs open_batchout_configured ────────────────────────────
    # If there is NO open/mo keyword but there IS an export keyword + batchout,
    # prefer export_direct over opening batchout configured.
    has_open   = "open" in all_features or "mo" in all_features
    has_export = "export" in all_features

    if has_export and not has_open and has_batchout:
        scores["export_direct"] = max(
            scores.get("export_direct", 0),
            scores.get("open_batchout_configured", 0) + 5
        )

    # ── Pick winner ──────────────────────────────────────────────────────────
    if not scores:
        return None
    best_intent = max(scores, key=lambda k: scores[k])
    best_score  = scores[best_intent]

    threshold = _THRESHOLDS.get(best_intent, 18)
    if best_score < threshold:
        return None
    return best_intent


# ─── Message builder ──────────────────────────────────────────────────────────

_MESSAGES_VI = {
    "open_batchout":          u"Đang mở BatchOut...",
    "open_batchout_configured": u"Mở BatchOut đã cấu hình...",
    "open_parasync":          u"Đang mở ParaSync...",
    "open_loadfamily":        u"Đang mở Load Family...",
    "open_loadfamilycloud":   u"Đang mở Load Family (Cloud)...",
    "open_projectname":       u"Đang mở Project Name...",
    "open_workset":           u"Đang mở Workset...",
    "open_dimtext":           u"Đang mở Dim Text...",
    "open_upperdimtext":      u"Đang mở Upper Dim Text...",
    "open_resetoverrides":    u"Đang mở Reset Overrides...",
    "open_grids":             u"Đang mở Grids...",
    "greet":                  u"Xin chào! Tôi là T3Lab Assistant. Cần giúp gì không?",
    "chat":                   u"Không có gì! Cần gì cứ hỏi nhé.",
    "help":                   u"Bạn có thể hỏi tôi về BatchOut, ParaSync, LoadFamily và các công cụ T3Lab khác.",
}

_MESSAGES_EN = {
    "open_batchout":          "Opening BatchOut...",
    "open_batchout_configured": "Opening BatchOut (pre-configured)...",
    "open_parasync":          "Opening ParaSync...",
    "open_loadfamily":        "Opening Load Family...",
    "open_loadfamilycloud":   "Opening Load Family (Cloud)...",
    "open_projectname":       "Opening Project Name...",
    "open_workset":           "Opening Workset...",
    "open_dimtext":           "Opening Dim Text...",
    "open_upperdimtext":      "Opening Upper Dim Text...",
    "open_resetoverrides":    "Opening Reset Overrides...",
    "open_grids":             "Opening Grids...",
    "greet":                  "Hello! I'm T3Lab Assistant. How can I help?",
    "chat":                   "You're welcome! Let me know if you need anything.",
    "help":                   "Ask me about BatchOut, ParaSync, LoadFamily and other T3Lab tools.",
}


def _is_viet(raw):
    viet_chars = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
                  u"ụủứừửữựỳỵỷỹ")
    return any(c in viet_chars for c in raw.lower())


def _build_message(intent, slots, viet):
    """Build a friendly message for the given intent and extracted slots."""
    if intent == "export_direct":
        fmt  = slots.get("format", "pdf").upper()
        filt = slots.get("filter", "")
        if viet:
            part = u" {} sheet".format(filt) if filt else u" tất cả sheet"
            return u"Đang xuất{} sang {}...".format(part, fmt)
        else:
            part = " {} sheet".format(filt) if filt else " all sheets"
            return "Exporting{} to {}...".format(part, fmt)

    if intent == "open_batchout_configured":
        fmt  = slots.get("format", "pdf").upper()
        filt = slots.get("filter", "")
        if viet:
            part = u" {} sheet".format(filt) if filt else u""
            return u"Mở BatchOut{}  ({})...".format(part, fmt)
        else:
            part = " {} sheet".format(filt) if filt else ""
            return "Opening BatchOut{} ({})...".format(part, fmt)

    if viet:
        return _MESSAGES_VI.get(intent, u"Đang xử lý...")
    return _MESSAGES_EN.get(intent, "Processing...")


# ─── Public API ───────────────────────────────────────────────────────────────

def classify(user_input, history=None):
    """Classify user_input and return a result dict, or None if uncertain.

    Args:
        user_input : raw unicode string from user.
        history    : list of {role, content} dicts (conversation context).

    Returns:
        dict  {intent, params, message, _nlu: True} on success.
        None  if confidence is below threshold.
    """
    if not user_input or not user_input.strip():
        return None

    viet = _is_viet(user_input)

    # ── Preprocess ───────────────────────────────────────────────────────────
    normed   = _norm(user_input)
    expanded = _expand(normed)

    # ── Pronoun resolution ───────────────────────────────────────────────────
    if _is_pronoun_query(expanded) and history:
        last_tool = _last_tool_from_history(history)
        if last_tool:
            slots = _extract_slots(user_input)
            label = _TOOL_LABELS.get(last_tool, last_tool)
            msg = (u"Đang mở {}...".format(label) if viet
                   else "Opening {}...".format(label))
            return {"intent": last_tool, "params": {}, "message": msg,
                    "_nlu": True}

    # ── Tokenise ─────────────────────────────────────────────────────────────
    unigrams, bigrams = _tokenise(expanded)

    # ── Score every intent ───────────────────────────────────────────────────
    scores = {intent: _score(intent, unigrams, bigrams)
              for intent in _TRIGGERS}

    # ── Extract slots (needed for disambiguation) ────────────────────────────
    slots = _extract_slots(user_input)

    # ── Disambiguate ─────────────────────────────────────────────────────────
    best = _disambiguate(dict(scores), unigrams, bigrams, slots)
    if best is None:
        return None

    # ── Build result ─────────────────────────────────────────────────────────
    params = {}
    if best in ("export_direct", "open_batchout_configured"):
        params = slots
    elif best == "open_batchout":
        params = {}

    message = _build_message(best, slots, viet)
    return {"intent": best, "params": params, "message": message, "_nlu": True}
