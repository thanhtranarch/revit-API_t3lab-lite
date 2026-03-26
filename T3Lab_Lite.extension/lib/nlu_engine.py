# -*- coding: utf-8 -*-
"""
NLU Engine

Natural Language Understanding engine for parsing Revit commands.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

from __future__ import unicode_literals, division

__author__  = "Tran Tien Thanh"
__title__   = "NLU Engine"

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
    # ── Tool name normalisations ───────────────────────────────────────────────
    # BatchOut
    ("batch out",       "batchout"),
    ("batcho",          "batchout"),
    ("b.out",           "batchout"),
    ("bo ",             "batchout "),   # trailing space avoids "bo override"
    # ParaSync
    ("para sync",       "parasync"),
    ("dong bo tham so", "parasync"),
    ("dong bo thong so","parasync"),
    ("dong bo",         "parasync"),
    ("sync param",      "parasync"),
    ("parameter sync",  "parasync"),
    (" ps ",            " parasync "),
    # Load Family Cloud
    ("load fam cloud",  "loadfamilycloud"),
    ("tai family cloud","loadfamilycloud"),
    ("load cloud",      "loadfamilycloud"),
    (" lfc ",           " loadfamilycloud "),
    # Load Family
    ("load fam",        "loadfamily"),
    ("tai family",      "loadfamily"),
    ("nap family",      "loadfamily"),
    ("keo family",      "loadfamily"),
    ("import family",   "loadfamily"),
    (" lf ",            " loadfamily "),
    # Project Name
    ("project name",    "projectname"),
    ("ten project",     "projectname"),
    ("ten du an",       "projectname"),
    ("dat ten project", "projectname"),
    ("thay ten",        "projectname"),
    (" pn ",            " projectname "),
    # Workset
    ("quan ly workset", "workset"),
    ("manage workset",  "workset"),
    (" ws ",            " workset "),
    # Upper Dim Text
    ("upper dim text",  "upperdimtext"),
    ("upper dim",       "upperdimtext"),
    (" udt ",           " upperdimtext "),
    # Dim Text
    ("dim text",        "dimtext"),
    ("dimension text",  "dimtext"),
    ("kich thuoc",      "dimtext"),
    ("sua dim",         "dimtext"),
    ("chinh dim",       "dimtext"),
    ("edit dim",        "dimtext"),
    ("edit dimension",  "dimtext"),
    (" dt ",            " dimtext "),
    # Reset Overrides
    ("reset graphic override","resetoverrides"),
    ("reset graphic",   "resetoverrides"),
    ("reset override",  "resetoverrides"),
    ("xoa override",    "resetoverrides"),
    ("bo override",     "resetoverrides"),
    ("xoa do ghi de",   "resetoverrides"),
    # Grids
    ("luoi truc",       "grids"),
    ("quan ly luoi",    "grids"),
    ("manage grid",     "grids"),

    # ── Intent-prefixed command patterns ──────────────────────────────────────
    # "muốn / cần / hãy / nhờ / làm ơn" before a verb → just keep the verb
    ("muon xuat",       "export"),
    ("muon in ",        "export "),
    ("muon mo ",        "open "),
    ("muon bat ",       "open "),
    ("muon chay ",      "open "),
    ("can xuat",        "export"),
    ("can in ",         "export "),
    ("can mo ",         "open "),
    ("hay xuat",        "export"),
    ("hay mo ",         "open "),
    ("lam on xuat",     "export"),
    ("lam on mo ",      "open "),
    ("lam on in ",      "export "),
    ("nho xuat",        "export"),
    ("nho mo ",         "open "),
    ("nho in ",         "export "),
    ("giup toi xuat",   "export"),
    ("giup toi mo ",    "open "),
    ("cho toi xuat",    "export"),
    ("cho toi mo ",     "open "),
    ("thu mo ",         "open "),
    ("thu xuat",        "export"),

    # ── Export verbs ──────────────────────────────────────────────────────────
    # Specific bigrams first (highest priority)
    ("in pdf",          "export pdf"),
    ("in dwg",          "export dwg"),
    ("in dwf",          "export dwf"),
    ("in dgn",          "export dgn"),
    ("in ifc",          "export ifc"),
    ("in nwd",          "export nwd"),
    ("in img",          "export img"),
    ("in sheet",        "export sheet"),
    ("in to ",          "export sheet "),  # "in tờ …"
    ("in het",          "export all"),
    ("in toan bo",      "export all"),
    ("in tat ca",       "export all"),
    ("in ra",           "export"),
    ("xuat ra",         "export"),
    ("xuat",            "export"),
    # Extra export synonyms
    ("ket xuat",        "export"),          # formal "render/export"
    ("save as pdf",     "export pdf"),
    ("save as dwg",     "export dwg"),
    ("convert to pdf",  "export pdf"),
    ("convert to dwg",  "export dwg"),
    ("chuyen sang pdf", "export pdf"),
    ("chuyen sang dwg", "export dwg"),
    ("luu thanh pdf",   "export pdf"),
    ("luu thanh dwg",   "export dwg"),
    ("xuat sang pdf",   "export pdf"),
    ("xuat sang dwg",   "export dwg"),
    ("out pdf",         "export pdf"),
    ("out dwg",         "export dwg"),

    # ── Open verbs ────────────────────────────────────────────────────────────
    ("mo len",          "open"),
    ("bat len",         "open"),
    ("chay len",        "open"),
    ("khoi dong",       "open"),
    ("khoi chay",       "open"),

    # ── Quantity / scope ──────────────────────────────────────────────────────
    ("tat ca sheet",    "all sheet"),
    ("toan bo sheet",   "all sheet"),
    ("all sheets",      "all sheet"),
    ("every sheet",     "all sheet"),
    ("all sheet",       "all sheet"),
    ("toan bo",         "all"),
    ("tat ca",          "all"),
    ("het sheets",      "all sheet"),
    ("het tat ca",      "all"),
    ("toan phan",       "all"),

    # ── Format synonyms ───────────────────────────────────────────────────────
    ("hinh anh",        "img"),
    ("image",           "img"),
    ("picture",         "img"),

    # ── Greeting shortcuts ────────────────────────────────────────────────────
    ("chao buoi sang",  "chao"),
    ("chao buoi chieu", "chao"),
    ("chao buoi toi",   "chao"),
    ("good morning",    "chao"),
    ("good afternoon",  "chao"),
    ("good evening",    "chao"),
    ("good night",      "chao"),
    ("xin chao",        "chao"),
    ("alo",             "chao"),
    ("a lo",            "chao"),
    ("hi ban",          "chao"),
    ("hey ban",         "chao"),
    ("chao ban",        "chao"),

    # ── Thanks shortcuts ──────────────────────────────────────────────────────
    ("cam on nhieu",    "cam on"),
    ("cam on ban nhieu","cam on"),
    ("cam on ban rat nhieu", "cam on"),
    ("rat cam on",      "cam on"),
    ("tran trong",      "cam on"),
    ("biet on",         "cam on"),
    ("thank you so much","cam on"),
    ("thanks a lot",    "cam on"),
    ("thanks so much",  "cam on"),
    ("thank u",         "cam on"),
    (" tks ",           " cam on "),
    (" thks ",          " cam on "),
    (" thx ",           " cam on "),
    (" ty ",            " cam on "),

    # ── Acknowledgement shortcuts ─────────────────────────────────────────────
    ("ok roi",          "ok roi"),    # keep for trigger matching
    ("oke roi",         "ok roi"),
    ("duoc roi",        "ok roi"),
    ("chay roi",        "ok roi"),
    ("xong roi",        "ok roi"),
    ("hieu roi",        "hieu roi"),
    ("biet roi",        "hieu roi"),
    ("ra roi",          "hieu roi"),
    ("toi hieu",        "hieu roi"),
    ("ro rang",         "hieu roi"),
    ("dung roi",        "dung roi"),
    ("chinh xac",       "dung roi"),
    ("hop ly",          "dung roi"),

    # ── Vietnamese "open" at word boundary ───────────────────────────────────
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
    # ── Vietnamese function / filler words ────────────────────────────────────
    "va", "de", "cho", "cai", "len", "ra", "vao", "di",
    "toi", "ban", "minh", "ho", "no", "cua", "voi", "trong",
    "la", "duoc", "co", "khong", "nhe", "nha", "nao", "gi",
    "a", "o", "roi", "se", "da", "dang", "rat", "khi",
    "neu", "sau", "truoc", "thi", "ma", "vay", "ay", "oi",
    "rang", "vi", "kia", "nhau", "hon",
    "mot", "hai", "ba", "bon", "nam",        # numbers (sheet count)
    "muon", "can", "hay", "thu", "giup",     # intent-prefix words (stripped by ABBREVS)
    "lam", "nho",                             # politeness words
    # NOTE: "in" is intentionally EXCLUDED — it means "print/export" in Vietnamese
    # ── English function words ─────────────────────────────────────────────────
    "the", "an", "to", "for", "of", "at",
    "me", "my", "i", "you", "it", "is", "are", "was", "be",
    "and", "or", "not", "on", "up", "do", "as", "so",
    "can", "will", "would", "could", "should",
    "with", "from", "by", "all", "any", "some",
    "please", "just", "now", "here",
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
        ("export pdf",         20),  # bigram from "in pdf" expansion
        ("export dwg",         20),
        ("pdf",                 5),
        ("dwg",                 5),
        ("dwf",                 5),
        ("all sheet",          10),
        ("print",               8),
        # "in" as Vietnamese print verb (after stopword fix, it survives tokenisation)
        ("in",                 10),
        # "sheet" alone contributes a small boost when present with other cues
        ("sheet",               4),
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
        # Core greeting words (survive after diacritic strip)
        ("chao",               25),
        ("hello",              25),
        ("hey",                18),
        ("hi",                 20),
        ("howdy",              18),
        ("alo",                25),    # phone-style "alo" (mapped via ABBREVS)
        # Bigrams / phrases (matched after ABBREVS expand "xin chao" → "chao")
        ("chao ban",           28),
        ("hi ban",             25),
        ("hey ban",            25),
        ("good morning",       20),    # also handled by ABBREVS
        ("morning",            12),
        ("chao buoi",          25),
        # Farewell (treat as greet-class conversational)
        ("tam biet",           22),
        ("bye",                20),
        ("bai",                18),
        ("see you",            20),
        ("hen gap",            22),
        ("hen gap lai",        25),
        ("good bye",           22),
        ("goodbye",            22),
        ("tam biet ban",       25),
    ],

    "chat": [
        # ── Thanks ────────────────────────────────────────────────────────────
        ("cam on",             25),
        ("thank",              20),
        ("thanks",             20),
        # ── Simple affirmations ───────────────────────────────────────────────
        ("ok",                  8),
        ("oke",                 8),
        ("ok roi",             22),    # mapped from oke roi / duoc roi / xong roi
        ("hieu roi",           22),    # mapped from biet roi / ra roi / toi hieu
        ("dung roi",           22),    # mapped from chinh xac / hop ly
        ("got it",             18),
        ("vang",               20),    # yes (polite Vietnamese)
        ("yep",                15),
        ("yeah",               15),
        ("alright",            18),
        ("sure",               15),
        ("noted",              20),
        ("understood",         20),
        ("copy",               15),
        ("roger",              15),
        # ── Positive reactions ────────────────────────────────────────────────
        ("tuyet",              20),
        ("tuyet voi",          22),
        ("tot lam",            20),
        ("hay qua",            20),
        ("ngon",               18),
        ("xin",                12),    # (only bigram use: "xin loi", "xin chao" filtered elsewhere)
        ("perfect",            20),
        ("great",              18),
        ("nice",               15),
        ("awesome",            18),
        ("wow",                15),
        ("uu viet",            18),    # excellent
        # ── State / emotion ───────────────────────────────────────────────────
        ("ban khoe",           22),    # how are you
        ("khoe khong",         22),
        ("ban oi",             15),
        ("met qua",            20),
        ("buon qua",           20),
        ("chan qua",           18),
        ("stress qua",         18),
        ("kho qua",            18),
        ("sao vay",            18),
        ("sao the",            18),
        ("tai sao",            15),
        # ── Complaints / errors ───────────────────────────────────────────────
        ("loi roi",            20),
        ("bi loi",             20),
        ("gap loi",            20),
        ("khong chay duoc",    22),
        ("sao khong chay",     22),
        ("bi hong",            20),
        ("khong hoat dong",    22),
        ("sao vay ban",        22),
        ("help me",            15),
        # ── Polite filler that doesn't map elsewhere ──────────────────────────
        ("xin loi",            20),    # sorry/excuse me
        ("sorry",              15),
        ("pardon",             12),
    ],

    "help": [
        # ── "What is X?" ──────────────────────────────────────────────────────
        ("la gi",              25),
        ("la cai gi",          25),
        ("nghia la gi",        25),
        ("what is",            25),
        ("what are",           22),
        ("what does",          22),
        ("batchout la gi",     30),
        ("parasync la gi",     28),
        ("loadfamily la gi",   28),
        # ── "How to do X?" ────────────────────────────────────────────────────
        ("lam gi",             22),
        ("lam nhu the nao",    25),
        ("lam the nao",        25),
        ("lam sao",            22),
        ("lam sao de",         25),
        ("bang cach nao",      25),
        ("cach nao",           20),
        ("how to",             22),
        ("how do",             20),
        ("how does",           20),
        ("how can",            18),
        ("xuat bang cach nao", 28),
        ("lam sao de xuat",    28),
        ("lam sao de mo",      28),
        # ── "What can you do?" ────────────────────────────────────────────────
        ("lam duoc gi",        25),
        ("dung duoc gi",       25),
        ("ho tro gi",          25),
        ("co tinh nang gi",    25),
        ("ban co the",         20),
        ("ban biet lam gi",    25),
        ("ban giup duoc gi",   25),
        ("ban lam gi",         22),
        ("ban la ai",          25),
        ("ban ten gi",         25),
        ("may la ai",          22),    # informal "who are you"
        ("tool nay",           18),
        # ── "Help/guide me" ───────────────────────────────────────────────────
        ("huong dan",          22),
        ("chi dan",            22),
        ("chi toi",            20),
        ("chi cach",           22),
        ("chi cho toi",        25),
        ("giai thich",         22),
        ("explain",            20),
        ("guide",              18),
        ("huong dan su dung",  28),
        ("cach su dung",       25),
        ("cach dung",          22),
        ("muon biet",          20),
        ("can biet",           20),
        ("cho toi biet",       25),
        ("noi cho toi",        22),
        ("tell me",            18),
        ("giup toi voi",       22),
        ("giup voi",           18),
        # ── Features / documentation ──────────────────────────────────────────
        ("tinh nang",          20),
        ("chuc nang",          20),
        ("su dung",            15),
        ("ho tro",             18),
        ("khai niem",          18),
        ("mo ta",              15),
        ("tai lieu",           18),
        ("document",           12),
        ("help",               12),
        ("info",               12),
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

    # ── Boost export_direct when a sheet-filter letter is present ────────────
    # "in G sheet" → "in"(10)+"sheet"(4)=14 which is just below threshold.
    # If the slot extractor found a filter letter AND there's any export/print
    # word, force export_direct to at least meet its threshold.
    _export_words = {"export", "in", "print", "pdf", "dwg", "dwf", "dgn",
                     "ifc", "nwd", "img", "sheet"}
    if slots.get("filter") and (_export_words & (unigrams | bigrams)):
        scores["export_direct"] = max(
            scores.get("export_direct", 0),
            _THRESHOLDS["export_direct"]
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
    "greet":  u"Xin chào! Tôi là T3Lab Assistant 👋\nBạn muốn làm gì hôm nay?",
    "farewell": u"Tạm biệt! Gặp lại bạn sau nhé 👋",
    "chat":   u"Không có gì! Cần gì cứ hỏi tôi nhé.",
    "help":   (u"Tôi có thể giúp bạn:\n"
               u"• Xuất sheet: 'xuất pdf G sheet', 'in tất cả sang dwg'\n"
               u"• Mở tool: 'mở batchout', 'parasync', 'load family'\n"
               u"• Cấu hình nhanh: 'mở batchout G sheet pdf'\n"
               u"Gõ tên tool hoặc mô tả điều bạn muốn làm!"),
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
    "greet":   "Hello! I'm T3Lab Assistant 👋\nWhat would you like to do today?",
    "farewell": "Goodbye! See you later 👋",
    "chat":    "You're welcome! Let me know if you need anything.",
    "help":    ("I can help you:\n"
                "• Export sheets: 'export pdf G sheet', 'print all to dwg'\n"
                "• Open tools: 'open batchout', 'parasync', 'load family'\n"
                "• Quick config: 'open batchout G sheet pdf'\n"
                "Type a tool name or describe what you want to do!"),
}


def _is_viet(raw):
    viet_chars = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
                  u"ụủứừửữựỳỵỷỹ")
    return any(c in viet_chars for c in raw.lower())


def _build_message(intent, slots, viet, raw_input=""):
    """Build a friendly message for the given intent and extracted slots."""
    normed_raw = _norm(raw_input)

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
            return u"Mở BatchOut{} ({})...".format(part, fmt)
        else:
            part = " {} sheet".format(filt) if filt else ""
            return "Opening BatchOut{} ({})...".format(part, fmt)

    # ── Contextual chat responses ─────────────────────────────────────────────
    if intent == "chat":
        # Farewell
        farewell_kws = ["tam biet", "bye", "bai ", "see you", "hen gap", "goodbye"]
        if any(k in normed_raw for k in farewell_kws):
            return (_MESSAGES_VI if viet else _MESSAGES_EN).get("farewell",
                    u"Tạm biệt! 👋" if viet else "Goodbye! 👋")
        # Error/complaint
        error_kws = ["loi", "bi hong", "khong chay", "khong hoat dong", "error",
                     "broken", "not working"]
        if any(k in normed_raw for k in error_kws):
            if viet:
                return (u"Xin lỗi bạn gặp vấn đề! Bạn có thể thử:\n"
                        u"• Đóng và mở lại tool\n"
                        u"• Kiểm tra Revit console để xem lỗi chi tiết")
            return ("Sorry you're having issues! You can try:\n"
                    "• Close and reopen the tool\n"
                    "• Check the Revit console for error details")
        # Positive reaction
        positive_kws = ["tuyet", "tot", "ngon", "perfect", "great", "awesome", "nice"]
        if any(k in normed_raw for k in positive_kws):
            return u"Cảm ơn bạn! 😊 Cần gì cứ hỏi nhé." if viet else "Thank you! 😊 Let me know if you need anything."
        # Thanks
        thanks_kws = ["cam on", "thank", "tks", "thks", "ty"]
        if any(k in normed_raw for k in thanks_kws):
            return u"Không có gì! Cần gì cứ hỏi tôi nhé." if viet else "You're welcome! Let me know if you need anything."
        # State question
        state_kws = ["khoe", "met", "buon", "chan", "sao vay", "stress"]
        if any(k in normed_raw for k in state_kws):
            return (u"Cảm ơn bạn hỏi thăm! Tôi ổn 😊 Bạn cần tôi giúp gì không?"
                    if viet else "Thanks for asking! I'm fine 😊 How can I help?")

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

    # ── Soft fallback for conversational input that scored nothing ────────────
    # If classification failed but the input looks conversational (no tool
    # keywords), return a gentle "I don't understand" as a chat response rather
    # than None, so the UI can show something helpful instead of silently failing.
    if best is None:
        _tool_words = {
            "batchout", "parasync", "loadfamily", "projectname",
            "workset", "dimtext", "upperdimtext", "resetoverrides", "grids",
            "export", "open", "in", "print",
        }
        if not (unigrams & _tool_words):
            # Purely conversational / unknown
            if viet:
                msg = (u"Xin lỗi, tôi chưa hiểu ý bạn. Bạn có thể thử:\n"
                       u"• 'mở batchout' / 'xuất pdf G sheet'\n"
                       u"• 'parasync', 'load family', 'workset'...")
            else:
                msg = ("Sorry, I didn't understand. You can try:\n"
                       "• 'open batchout' / 'export pdf G sheet'\n"
                       "• 'parasync', 'load family', 'workset'...")
            return {"intent": "chat", "params": {}, "message": msg, "_nlu": True}
        return None

    # ── Build result ─────────────────────────────────────────────────────────
    params = {}
    if best in ("export_direct", "open_batchout_configured"):
        params = slots
    elif best == "open_batchout":
        params = {}

    message = _build_message(best, slots, viet, raw_input=user_input)
    return {"intent": best, "params": params, "message": message, "_nlu": True}
