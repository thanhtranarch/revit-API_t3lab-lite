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
    """Remove all combining diacritic marks (works for Vietnamese, etc.).

    NFD does NOT decompose đ/Đ (letter D with stroke — a standalone letter,
    not base+mark), so they are folded to d/D explicitly. Without this,
    "được" normalises to "đuoc" and later regex cleaning eats the đ,
    silently breaking every rule containing "duoc", "dong bo", "doc"...
    """
    try:
        nfd = unicodedata.normalize('NFD', text)
        out = ''.join(c for c in nfd if unicodedata.category(c) != 'Mn')
        return out.replace(u'đ', 'd').replace(u'Đ', 'D')
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
    (" batcho ",        " batchout "),  # word-boundary padded: "batcho" is a
                                         # substring PREFIX of "batchout" itself,
                                         # so an unpadded rule here self-mangles
                                         # the correct word into "batchoutut"
    ("b.out",           "batchout"),
    # NOTE: the risky bare "bo " -> "batchout " shorthand is intentionally
    # NOT here — see the end of this list for why.
    # ParaSync
    ("para sync",       "parasync"),
    ("dong bo tham so", "parasync"),
    ("dong bo thong so","parasync"),
    ("dong bo",         "parasync"),
    ("sync param",      "parasync"),
    ("parameter sync",  "parasync"),
    (" ps ",            " parasync "),
    # Load Family Cloud
    ("load family cloud","loadfamilycloud"),
    ("load fam cloud",  "loadfamilycloud"),
    ("tai family cloud","loadfamilycloud"),
    ("load cloud",      "loadfamilycloud"),
    (" lfc ",           " loadfamilycloud "),
    # Load Family
    ("load family",     "loadfamily"),
    (" load fam ",      " loadfamily "),  # word-boundary padded: unpadded "load
                                           # fam" is a substring PREFIX of "load
                                           # family" and would mangle it into
                                           # "loadfamilyily"
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

    # ── "What can you do" capability-query normalisation ─────────────────────
    # English capability questions are made of nothing but stopwords
    # ("what", "can", "you", "do") once tokenised, so without this they score
    # zero on every intent and fall through to the generic "didn't understand"
    # reply. Collapse the whole phrase to one non-stopword marker up front.
    ("what can you do",         "capabilities query"),
    ("what can u do",           "capabilities query"),
    ("what do you do",          "capabilities query"),
    ("what could you do",      "capabilities query"),
    ("what are you capable of","capabilities query"),
    ("what can this do",        "capabilities query"),
    ("what can this tool do",   "capabilities query"),
    ("what features do you have","capabilities query"),
    ("what tools do you have",  "capabilities query"),
    ("how can you help me",     "capabilities query"),
    ("how can you help",        "capabilities query"),
    ("who are you",             "capabilities query"),
    ("what are you",            "capabilities query"),

    # ── Greeting shortcuts ────────────────────────────────────────────────────
    ("chao buoi sang",  "chao"),
    ("chao buoi chieu", "chao"),
    ("chao buoi toi",   "chao"),
    ("good morning",    "chao"),
    ("good afternoon",  "chao"),
    ("good evening",    "chao"),
    ("good night",      "chao"),
    ("xin chao",        "chao"),
    # Word-boundary padded: unpadded "alo"/"a lo" mangled every text that
    # merely CONTAINS the substring — "kiem tra loi" → "kiem trchaoi",
    # "tra loi", "dialog", "catalog"... (same class of bug as " batcho ")
    (" alo ",           " chao "),
    (" a lo ",          " chao "),
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

    # ── Bare "bo" -> BatchOut shorthand (LAST on purpose) ─────────────────────
    # "bo" is dangerously short and collides with real Vietnamese words that
    # end in "bo" after diacritics are stripped, most importantly "toàn bộ"
    # ("all") -> "toan bo". Every rule above this point already consumes its
    # own "bo" first (dong bo/parasync, bo override/resetoverrides, toan bo/
    # all, tat ca/all, etc.), so by the time this runs, any leftover
    # standalone "bo" can only be the actual BatchOut shorthand.
    ("bo ",             "batchout "),
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

    "open_loadfamily": [
        ("loadfamily",         30),
        ("open loadfamily",    35),
        ("mo loadfamily",      35),
        ("family",              8),
        ("open",                2),
    ],

    # NOTE — 2026-07-28: trigger blocks for open_parasync,
    # open_loadfamily_cloud, open_projectname, open_workset,
    # open_upperdimtext, open_dimtext, open_resetoverrides and open_grids were
    # removed. Their pushbuttons no longer exist, and because these triggers
    # scored high they beat resolve_tool() for the tools that replaced them —
    # "mo workset" matched open_workset (35) and never reached ManaWorkset.
    # Only tools with a dedicated launcher belong here; auto-discovered tools
    # are resolved from the registry by resolve_tool().

    "check_spelling": [
        # Proofread all Text Notes in the model (diacritics stripped by _norm:
        # "chính tả" → "chinh ta", "lỗi tiếng anh" → "loi tieng anh")
        ("spelling",           30),
        ("spellcheck",         35),
        ("spell check",        35),
        ("chinh ta",           35),
        ("loi chinh ta",       40),
        ("loi tieng anh",      40),
        ("tieng anh",          15),
        ("kiem tra loi",       22),
        ("check loi",          22),
        ("check english",      30),
        ("grammar",            28),
        ("ngu phap",           28),
        ("typo",               30),
        ("typos",              30),
        ("text note",          12),
        ("textnote",           12),
        ("review text",        15),
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
        # Bare time-of-day greetings must clear the 18 threshold on their own:
        # at 12, a lone "morning" fell through to the LLM, which once
        # hallucinated a tool intent that then got permanently mis-learned.
        ("morning",            20),
        ("afternoon",          20),
        ("evening",            20),
        ("yo",                 18),
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
        # ── Frustration / insult directed at the assistant ────────────────────
        # Not tool-related — give an honest, de-escalating reply instead of the
        # generic "didn't understand" message (see _build_message).
        ("stupid",             22),
        ("dumb",               20),
        ("useless",            22),
        ("garbage",            20),
        ("trash",              18),
        ("suck",               18),
        ("sucks",              18),
        ("ngu",                20),
        ("vo dung",            22),
        ("te qua",             20),
        ("qua te",             20),
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
        ("capabilities query", 30),   # normalised from EN phrasing via _ABBREVS
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
    # Create/delete verbs mean the user wants to WRITE notes, not proofread
    "check_spelling":            [("tao", -25), ("create", -25), ("them", -20),
                                  ("add", -20), ("xoa", -25), ("delete", -25),
                                  ("draw", -20), ("ve", -20)],
}

# Minimum score required to accept an intent
_THRESHOLDS = {
    "open_batchout":             18,
    "export_direct":             18,
    "open_batchout_configured":  25,   # needs both open+batchout+params
    "open_loadfamily":           18,
    # ≥ 28 so "text note"(12) + "tieng anh"(15) alone can't fire it —
    # needs a real proofread cue (spelling / chinh ta / loi tieng anh ...)
    "check_spelling":            28,
    "greet":                     18,
    "chat":                      18,
    "help":                      18,
}


# ─── Conversational-input gate ────────────────────────────────────────────────
# Distinguishes pure small talk (greeting / thanks / acknowledgement / emotion)
# from tool commands. This is the safety gate that keeps the higher layers
# honest: learned patterns must never be recorded for — or matched against —
# conversational input, and no tool may be launched from it, regardless of
# what an LLM hallucinated.

# Any of these words present → definitely a command, never small talk.
_COMMAND_WORDS = {
    "batchout", "parasync", "loadfamily", "loadfamilycloud", "projectname",
    "workset", "dimtext", "upperdimtext", "resetoverrides", "grids", "grid",
    "export", "open", "print", "sheet", "sheets", "family", "luoi", "truc",
    "pdf", "dwg", "dwf", "dgn", "ifc", "nwd", "img", "in",
    "capabilities", "query",   # normalised capability-question marker
    "spelling", "spellcheck", "grammar", "typo", "typos", "textnote",
}


def _build_conversational_words():
    """Lexicon of small-talk words, derived from the greet/chat trigger tables
    so it stays in sync, plus bare words those tables only contain inside
    multi-word phrases."""
    words = set()
    for intent in ("greet", "chat"):
        for feat, _w in _TRIGGERS[intent]:
            words.update(feat.split())
    words.update({
        "morning", "afternoon", "evening", "night", "yo", "sup",
        "gm", "gn", "haha", "hihi", "lol", "ban", "buoi",
        "please", "welcome", "greetings",
    })
    return words - _COMMAND_WORDS


_CONVERSATIONAL_WORDS = _build_conversational_words()


def is_conversational(user_input):
    """Return True when the input is pure small talk and contains no
    tool/command keyword (e.g. "morning", "cảm ơn nhé", "ok roi").

    Conservative by design: any command word, any unknown meaningful word,
    or anything longer than 6 tokens → False (treat as a possible command).
    """
    if not user_input or not user_input.strip():
        return False
    expanded = _expand(_norm(user_input))
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    tokens = clean.split()
    if not tokens or len(tokens) > 6:
        return False
    for t in tokens:
        if t in _COMMAND_WORDS:
            return False
    meaningful = [t for t in tokens if t not in _STOPWORDS and len(t) >= 2]
    if not meaningful:
        return True   # nothing but filler words → chit-chat
    return all(t in _CONVERSATIONAL_WORDS for t in meaningful)


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

    # Priority 2: lone single-letter token or pure uppercase sheet prefix
    if not sheet_filt:
        for token in raw.split():
            clean_tok = "".join(c for c in token if c.isalnum())
            if len(clean_tok) == 1 and clean_tok.isalpha():
                tok = clean_tok.upper()
                if tok not in _FORMAT_LETTERS:
                    sheet_filt = tok
                    break

    # ── combine flag ─────────────────────────────────────────────────────────
    combine_kws = ["combine", "merge", "gop", "ghep", "1 file", "mot file"]
    combine = any(k in normed for k in combine_kws)

    return {"format": fmt, "filter": sheet_filt, "combine": combine}


# ─── Context / pronoun resolution ─────────────────────────────────────────────

# Pronouns that refer to the most-recently-mentioned tool
_PRONOUNS = {"no", "no ay", "cai do", "cai nay", "tool do", "it", "that", "this"}

# Intent → tool label, for pronoun-resolution messages ("mở nó" → "BatchOut").
# Only the intents with a dedicated launcher are fixed; every other label comes
# from the live registry via _tool_label(). The hardcoded rest used to name
# eight tools that had been deleted from the extension.
_TOOL_LABELS = {
    "open_batchout":            "BatchOut",
    "export_direct":            "BatchOut",
    "open_batchout_configured": "BatchOut",
    "open_loadfamily":          "Family Loader",
}


def _tool_label(intent):
    """Display label for an intent — fixed map first, then the live registry."""
    if intent in _TOOL_LABELS:
        return _TOOL_LABELS[intent]
    for t in _tool_catalog():
        if t.get('intent') == intent:
            return t.get('title') or intent
    return intent


def _tool_keywords():
    """keyword → intent, for spotting the last-mentioned tool in history.

    Built from the live catalog (joined names such as "manaworkset",
    "cadtoelements") instead of a hardcoded table, so a renamed or newly added
    pushbutton is picked up without editing this file.
    """
    out = {
        "batchout":   "open_batchout",
        "loadfamily": "open_loadfamily",
        "manafami":   "open_loadfamily",
    }
    for t in _tool_catalog():
        intent = t.get('intent')
        if not intent or intent in ('open_batchout', 'open_loadfamily'):
            continue
        for joined in (t.get('joined') or ()):
            if len(joined) >= 4:
                out.setdefault(joined, intent)
    return out


# ─── Deterministic tool resolver ─────────────────────────────────────────────
# Ranks the user's text against EVERY known tool (builtin + auto-discovered)
# and only answers when one tool clearly wins. This replaces the old
# first-keyword-substring match, which returned whichever registry entry
# happened to be iterated first (e.g. "mở mcp control" → CAD to Elements,
# because a generic keyword matched earlier in the dict).

# Builtin tools (already covered by _TRIGGERS, listed here so the resolver
# sees one unified catalog):
#   (intent, title, joined-name, [name aliases], function description)
# Aliases cover every name a tool goes by: button folder, XAML file, old
# names — so "mở export manager" (BatchOut's XAML) opens the right tool.
# Tools that are NOT auto-discovered because script.py launches them through a
# dedicated entry point (_SPECIAL_LAUNCHERS). Everything else reaches the
# resolver via the tool_discovery registry in _tool_catalog() below.
#
# 2026-07-28: eight entries were removed from this table (ParaSync,
# LoadFamilyCloud, ProjectName, Workset, DimText, UpperDimText, ResetOverrides,
# Grids). Their pushbuttons had been deleted or renamed, but the resolver kept
# offering them, so the assistant confidently answered "Opening ParaSync..." for
# a tool that no longer shipped. Do not add a tool here unless it genuinely has
# a special launcher — an auto-discovered tool needs no entry.
_BUILTIN_TOOLS = [
    ("open_batchout", "BatchOut", "batchout",
     ["Batch Out", "Export Manager"],
     u"Xuất sheet hàng loạt sang PDF / DWG / DWF / IFC (batch export sheets)"),
    ("open_loadfamily", "Family Loader", "loadfamily",
     ["Load Family", "Family Manager", "ManaFami"],
     u"Tải family từ thư viện vào project (load family from library)"),
]

# Verbs that signal "open this tool" (post-_expand, so "mở/bật" → open).
# Vietnamese verbs that CANNOT be globally rewritten to "open" by _ABBREVS
# (e.g. "chạy" would corrupt the "không chạy được" complaint trigger) are
# stripped here instead, so "chạy autojoin" still resolves to the tool.
# NOTE: "dung" is intentionally ABSENT: after diacritics folding it is both
# "dùng" (use → open verb) and "dựng" (build/model) — keeping it hijacked
# every modeling request ("dựng model" → "which tool do you want to open?").
# "dùng X" now falls through to the LLM path, which handles it fine.
_OPEN_VERBS = {"open", "launch", "start", "run", "show",
               "chay", "xai"}


def _singularise(w):
    """Fold trivial English plurals so 'views' matches 'view' etc."""
    if len(w) >= 4 and w.endswith("s") and not w.endswith("ss"):
        return w[:-1]
    return w


def _camel_split(text):
    """'DWGManagement' → 'DWG Management' — word boundaries for camel names."""
    return re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', text)


def _name_variants(text):
    """Return (joined_names, word_set) for ONE tool name, covering the raw,
    camel-split, and _expand()-normalised forms (so "ImageToDrafting" yields
    words {image, drafting} and also matches input ABBREVS rewrote to
    "img to drafting")."""
    joined, words = set(), set()
    for base in (text, _camel_split(text)):
        for form in (_norm(base), _expand(_norm(base))):
            j = re.sub(r'[^a-z0-9]', '', form)
            if j:
                joined.add(j)
            for w in re.findall(r'[a-z0-9]+', form):
                if w not in _STOPWORDS and len(w) >= 2:
                    words.add(_singularise(w))
    return joined, words


def _desc_words(desc):
    """Normalised word set from a function description (for capability Q&A)."""
    out = set()
    for w in re.findall(r'[a-z0-9]+', _norm(_camel_split(desc or ''))):
        if w not in _STOPWORDS and len(w) >= 2:
            out.add(_singularise(w))
    return out


def _tool_entry(intent, title, names, desc='', panel='', extra_words=None):
    """Build one catalog entry. `names` is EVERY name the tool goes by
    (title, button folder, XAML basenames, aliases) — kept as separate
    variants so exact matching works per-name, plus a word union for fuzzy."""
    joined_all, variants, union = set(), [], set()
    for name in names:
        if not name:
            continue
        j, w = _name_variants(name)
        joined_all |= j
        if w and frozenset(w) not in variants:
            variants.append(frozenset(w))
        union |= w
    dwords = _desc_words(desc)
    if extra_words:
        dwords |= set(extra_words)
    return {'intent': intent, 'title': title, 'desc': (desc or '').strip(),
            'panel': panel, 'joined': joined_all, 'variants': variants,
            'words': union, 'desc_words': dwords}


def _tool_catalog():
    """Return every known tool — builtin + auto-discovered — with all its
    names (title / button folder / XAML / aliases) unified per tool."""
    catalog = []
    for intent, title, joined, aliases, desc in _BUILTIN_TOOLS:
        e = _tool_entry(intent, title, [title] + list(aliases), desc,
                        panel=u"Core")
        e['joined'].add(joined)
        catalog.append(e)
    try:
        from Services.tool_discovery import get_registered_tools
        tools = get_registered_tools()
    except Exception:
        tools = []
    for t in tools:
        title = ((t.get('title') or '')
                 .replace('&amp;', ' ').replace('&', ' ').strip())
        btn   = (t.get('button') or '').replace('.pushbutton', '')
        names = [title, btn] + list(t.get('xaml') or [])
        kw_words = set()
        for kw in (t.get('keywords') or []):
            for w in re.findall(r'[a-z0-9]+', _norm(kw)):
                if w not in _STOPWORDS and len(w) >= 2:
                    kw_words.add(_singularise(w))
        e = _tool_entry(t.get('intent'), title or btn, names,
                        t.get('doc') or '',
                        panel=(t.get('panel') or '').replace('.panel', ''),
                        extra_words=kw_words)
        if e['words'] or e['joined']:
            catalog.append(e)
    return catalog


def resolve_tool(user_input, exact_only=False):
    """Deterministically resolve a tool-open request against the full catalog.

    Returns (match, candidates):
      match      – {'intent','title',...} when exactly one tool clearly wins
      candidates – up to 3 plausible tools when the request is ambiguous

    Exact wins are judged on ALL tokens minus open-verbs, KEEPING stopwords —
    so "workset manager" is exactly Workset Manager, while "mcp control la gi"
    is NOT exact (the question words survive and block it, letting the help
    intent handle it). Exact = whole query joins to a tool's joined name
    ("mcpcontrol", "cadtoelements"), or the token set equals the tool's word
    set in any order ("manager dwg"). Fuzzy wins need score ≥ 0.75 AND a
    ≥ 0.2 lead over the runner-up — otherwise the request is reported
    ambiguous instead of guessed.
    """
    if not user_input or not user_input.strip():
        return None, []
    expanded = _expand(_norm(user_input))
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    tokens_all = [w for w in clean.split() if w not in _OPEN_VERBS]
    if not tokens_all or len(tokens_all) > 8:
        return None, []
    qjoined_all = "".join(tokens_all)
    qset_all    = set(_singularise(w) for w in tokens_all)
    # Stopword-free set for fuzzy scoring only
    qset = set(_singularise(w) for w in tokens_all if w not in _STOPWORDS)

    scored = []
    for tool in _tool_catalog():
        exact = (qjoined_all in tool['joined']
                 or any(qset_all == v for v in tool['variants'])
                 or (len(tokens_all) == 1 and tokens_all[0] in tool['joined']))
        if exact:
            scored.append((1.0, True, tool))
            continue
        inter = qset & tool['words']
        if inter:
            cov_tool  = len(inter) / float(len(tool['words']))
            cov_query = len(inter) / float(len(qset))
            # Fuzzy can reach 1.0 on perfect two-way coverage (e.g. "batchout
            # là gì" covers the word "batchout" fully) — that is NOT an exact
            # name hit, so it must stay distinguishable from exact=True.
            scored.append((0.6 * cov_tool + 0.4 * cov_query, False, tool))

    if not scored:
        return None, []
    scored.sort(key=lambda x: (-x[0], not x[1]))
    top_score, top_exact, top = scored[0]
    runner_up = scored[1][0] if len(scored) > 1 else 0.0

    if top_exact:
        # Two DIFFERENT tools both exact (e.g. shared XAML alias) → ambiguous
        exact_intents = set(t['intent'] for s, e, t in scored if e)
        if len(exact_intents) > 1:
            return None, [t for s, e, t in scored if e][:3]
        return top, []
    if exact_only:
        return None, []
    if top_score >= 0.75 and (top_score - runner_up) >= 0.2:
        return top, []
    candidates = [t for s, e, t in scored[:3] if s >= 0.45]
    return None, candidates


def _has_open_verb(expanded):
    """True if the expanded text contains an explicit open verb."""
    return bool(set(expanded.split()) & _OPEN_VERBS)


# ─── Capability questions ─────────────────────────────────────────────────────
# "Có tool nào để X không?" / "Do you have a tool for X?" must be answered
# from the REAL tool catalog — name the matching tools, or say honestly that
# none exists. Never left to the LLM to invent an answer.

_CAP_RES = [
    re.compile(r'\bco\s+(?:tool|cong cu|chuc nang|tinh nang|lenh)\b'),
    re.compile(r'\b(?:tool|cong cu|lenh)\s+nao\b'),
    re.compile(r'\b(?:tool|cong cu)\s+(?:de|giup|ho tro|cho)\b'),
    re.compile(r'\btool\s+(?:for|to)\b'),
    re.compile(r'\bdo\s+you\s+have\b'),
    re.compile(r'\bis\s+there\s+(?:a|an|any)\b'),
    re.compile(r'\b(?:which|what)\s+tool\b'),
    re.compile(r'\bany\s+tool\b'),
    re.compile(r'\bhave\s+a\s+tool\b'),
    # Generic "what can you do" (EN forms were collapsed by _ABBREVS)
    re.compile(r'\bcapabilities query\b'),
    re.compile(r'\blam duoc gi\b'),
    re.compile(r'\bdung duoc gi\b'),
    re.compile(r'\bho tro gi\b'),
    re.compile(r'\bbiet lam gi\b'),
    re.compile(r'\bgiup duoc gi\b'),
]

# Question boilerplate stripped before matching the FUNCTION words
# (many of these are already stopwords; listed for safety)
_CAP_BOILERPLATE = {
    "tool", "cong", "cu", "chuc", "nang", "tinh", "lenh", "nao",
    "have", "there", "which", "what", "any", "capabilities", "query",
    "biet", "ung", "dung", "ban", "assistant", "t3lab", "co", "khong",
    "gi", "giup", "ho", "tro", "lam", "duoc", "thuc", "hien", "does",
    "function", "feature", "help", "the",
}


def is_capability_question(expanded):
    """True if the (expanded) input asks whether a tool/feature exists."""
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    padded = u" " + u" ".join(clean.split()) + u" "
    return any(r.search(padded) for r in _CAP_RES)


# ─── Practical capability overview (grouped by workflow, verified live) ───────
# The assistant's real power is acting DIRECTLY on the open model through the
# MCP tool registry (core/server.py). "What can you do?" is answered from a
# CURATED, workflow-grouped map with concrete notes — the practical style users
# expect — but every line is VALIDATED against the live registry at render time:
#   • a line whose tools are ALL gone is dropped (no advertising dead names —
#     the drift trap documented in Intelligence/t3lab_agent.py); and
#   • any registered tool the map doesn't mention is surfaced under "Khác/Other"
#     so a newly-added tool is never silently hidden.
# Curated notes were checked against the tool schemas in core/server.py, e.g.
# purge_unused/delete_element dry_run, export_sheets_pdf `combined`, create_grid
# auto x_labels/y_labels, split_* geometry preservation, create_point_based host.

def _live_tool_names():
    """Set of MCP tool names currently registered, or an empty set when the
    registry is unreachable (running outside Revit / offline tests)."""
    try:
        from Intelligence.t3lab_agent import _get_mcp_tools
        return set(t.get("name", "") for t in _get_mcp_tools())
    except Exception:
        return set()


def _live_skills():
    """[(name, description)] of enabled assistant skills, or [] if none/error."""
    try:
        from Intelligence.skills_engine import get_skills_engine
        return [(s.get("name") or s.get("id"), s.get("description") or u"")
                for s in get_skills_engine().get_catalog(enabled_only=True)]
    except Exception:
        return []


def _user_display_name():
    """User's set display name, or '' (no fallback — used only to personalise)."""
    try:
        from config.user_profile import get_profile
        return (get_profile().get_name(fallback=False) or u"").strip()
    except Exception:
        return u""


# Internal/plumbing tools deliberately kept out of the user-facing overview
# (and out of the "Other" catch-all) — they aren't things a user asks for.
_CAP_HIDE = frozenset([
    u"say_hello", u"show_assistant_pane", u"file_watcher_status",
])

# Curated capability map. Each section: (emoji, vi_title, en_title, lines).
# Each line: (names_tuple, vi_markdown, en_markdown). The line renders when the
# registry is unknown (offline) OR at least one of its names is registered;
# every name feeds the "covered" set that drives the anti-drift catch-all.
_CAP_MAP = [
    (u"📖", u"ĐỌC / QUERY", u"READ / QUERY", [
        ((u"ai_element_filter",),
         u"`ai_element_filter` — lọc element theo category + parameter (Walls, Doors, Rooms, Grids, Pipes…)",
         u"`ai_element_filter` — filter elements by category + parameter (Walls, Doors, Rooms, Grids, Pipes…)"),
        ((u"get_all_parameters", u"get_parameter", u"revit_get_element_info"),
         u"`get_all_parameters` / `get_parameter` / `revit_get_element_info` — đọc mọi param & thông tin của 1 element",
         u"`get_all_parameters` / `get_parameter` / `revit_get_element_info` — read every parameter & info of an element"),
        ((u"get_elements_by_level", u"get_current_view_elements", u"get_current_view_info", u"revit_get_active_view", u"get_element_bounding_box"),
         u"`get_elements_by_level`, `get_current_view_elements`, `get_current_view_info` / `revit_get_active_view`, `get_element_bounding_box`",
         u"`get_elements_by_level`, `get_current_view_elements`, `get_current_view_info` / `revit_get_active_view`, `get_element_bounding_box`"),
        ((u"revit_get_selected_elements", u"select_elements", u"get_available_family_types"),
         u"`revit_get_selected_elements`, `select_elements` (chọn theo category+filter, zoom tới được), `get_available_family_types`",
         u"`revit_get_selected_elements`, `select_elements` (select by category+filter, can zoom to them), `get_available_family_types`"),
        ((u"get_schedule_data", u"get_material_quantities"),
         u"`get_schedule_data` — đọc schedule ra JSON (row_count + column_totals chính xác), `get_material_quantities`",
         u"`get_schedule_data` — read a schedule as JSON (exact row_count + column_totals), `get_material_quantities`"),
        ((u"list_levels", u"revit_list_views", u"revit_list_sheets", u"list_worksets"),
         u"`list_levels`, `revit_list_views`, `revit_list_sheets`, `list_worksets`",
         u"`list_levels`, `revit_list_views`, `revit_list_sheets`, `list_worksets`"),
        ((u"get_revit_context", u"revit_get_project_info", u"analyze_model_statistics"),
         u"`get_revit_context`, `revit_get_project_info`, `analyze_model_statistics` — tổng quan model",
         u"`get_revit_context`, `revit_get_project_info`, `analyze_model_statistics` — model overview"),
    ]),
    (u"✏️", u"TẠO / SỬA", u"CREATE / MODIFY", [
        ((u"place_wall", u"create_surface_based_element"),
         u"**Walls / Floors / Ceilings / Roofs** — `place_wall`, `create_surface_based_element`",
         u"**Walls / Floors / Ceilings / Roofs** — `place_wall`, `create_surface_based_element`"),
        ((u"create_point_based_element",),
         u"**Doors / Windows / Furniture** — `create_point_based_element` (host được vào wall)",
         u"**Doors / Windows / Furniture** — `create_point_based_element` (can be hosted into a wall)"),
        ((u"create_line_based_element", u"create_structural_framing_system"),
         u"**Beams / Pipes / Ducts** — `create_line_based_element`, `create_structural_framing_system`",
         u"**Beams / Pipes / Ducts** — `create_line_based_element`, `create_structural_framing_system`"),
        ((u"create_grid", u"create_level", u"create_room", u"load_family", u"room_to_floor"),
         u"**Grids / Levels / Rooms** — `create_grid` (tự đặt label A/B/C · 1/2/3), `create_level`, `create_room`, `load_family`, `room_to_floor`",
         u"**Grids / Levels / Rooms** — `create_grid` (auto labels A/B/C · 1/2/3), `create_level`, `create_room`, `load_family`, `room_to_floor`"),
        ((u"move_elements", u"copy_elements", u"rotate_element", u"split_element", u"split_curve"),
         u"`move_elements`, `copy_elements`, `rotate_element`, `split_element` / `split_curve` (giữ đúng arc/spline, không flatten)",
         u"`move_elements`, `copy_elements`, `rotate_element`, `split_element` / `split_curve` (keeps arcs/splines exact, no flattening)"),
        ((u"join_geometry", u"delete_element", u"operate_element"),
         u"`join_geometry` (join/unjoin), `delete_element` (có `dry_run` xem trước), `operate_element` (hide / isolate / unhide / reset màu)",
         u"`join_geometry` (join/unjoin), `delete_element` (`dry_run` preview), `operate_element` (hide / isolate / unhide / reset color)"),
        ((u"bulk_set_parameter", u"set_parameter", u"rename_element"),
         u"`bulk_set_parameter` — set 1 param cho hàng trăm element cùng lúc (filter được); `set_parameter`, `rename_element` cho từng element",
         u"`bulk_set_parameter` — set one parameter across hundreds of elements at once (filterable); `set_parameter`, `rename_element` for a single element"),
        ((u"create_project_parameter", u"create_workset", u"set_element_workset"),
         u"`create_project_parameter`, `create_workset`, `set_element_workset`",
         u"`create_project_parameter`, `create_workset`, `set_element_workset`"),
    ]),
    (u"📁", u"VIEW / SHEET / ANNOTATION", u"VIEW / SHEET / ANNOTATION", [
        ((u"create_view", u"duplicate_view", u"apply_view_template", u"create_view_filter", u"set_active_view"),
         u"`create_view`, `duplicate_view` (plain / with detailing / dependent), `apply_view_template`, `create_view_filter`, `set_active_view`",
         u"`create_view`, `duplicate_view` (plain / with detailing / dependent), `apply_view_template`, `create_view_filter`, `set_active_view`"),
        ((u"create_schedule", u"create_sheet", u"place_views_on_sheets", u"add_view_to_sheet"),
         u"`create_schedule`, `create_sheet`, `place_views_on_sheets`, `add_view_to_sheet`",
         u"`create_schedule`, `create_sheet`, `place_views_on_sheets`, `add_view_to_sheet`"),
        ((u"create_dimension", u"create_text_note", u"tag_elements", u"tag_all_rooms", u"tag_all_walls"),
         u"`create_dimension`, `create_text_note`, `tag_elements` / `tag_all_rooms` / `tag_all_walls`",
         u"`create_dimension`, `create_text_note`, `tag_elements` / `tag_all_rooms` / `tag_all_walls`"),
        ((u"color_elements", u"revit_override_color"),
         u"`color_elements` — tô màu theo giá trị parameter (rất hay để QA); `revit_override_color`",
         u"`color_elements` — color-code by parameter value (great for QA); `revit_override_color`"),
    ]),
    (u"📤", u"EXPORT", u"EXPORT", [
        ((u"export_sheets_pdf", u"export_dwg", u"export_image", u"export_room_data"),
         u"`export_sheets_pdf` (gộp thành 1 file được), `export_dwg`, `export_image` (PNG), `export_room_data`",
         u"`export_sheets_pdf` (can combine into one file), `export_dwg`, `export_image` (PNG), `export_room_data`"),
    ]),
    (u"🔍", u"QA / MODEL HEALTH", u"QA / MODEL HEALTH", [
        ((u"audit_model", u"get_model_warnings", u"get_model_health", u"analyze_model_statistics"),
         u"`audit_model`, `get_model_warnings`, `get_model_health`, `analyze_model_statistics`",
         u"`audit_model`, `get_model_warnings`, `get_model_health`, `analyze_model_statistics`"),
        ((u"purge_unused",),
         u"`purge_unused` — mặc định dry-run (chỉ báo cáo); set `dry_run=false` mới thật sự xoá",
         u"`purge_unused` — dry-run by default (report only); set `dry_run=false` to actually purge"),
    ]),
    (u"⚡", u"ADVANCED", u"ADVANCED", [
        ((u"send_code_to_revit",),
         u"`send_code_to_revit` — chạy **IronPython trực tiếp** trong Revit context (full API) → mạnh nhất, làm được cả những gì tool có sẵn chưa cover",
         u"`send_code_to_revit` — run **IronPython directly** in the Revit context (full API) → the most powerful path, does what the built-in tools don't cover yet"),
        ((u"store_project_data", u"store_room_data", u"query_stored_data"),
         u"Lưu / truy vấn dữ liệu bền vững trong model: `store_project_data`, `store_room_data`, `query_stored_data`",
         u"Persist / query data inside the model: `store_project_data`, `store_room_data`, `query_stored_data`"),
        ((u"switch_active_document", u"open_document", u"close_document", u"list_open_documents", u"list_recent_documents"),
         u"Làm việc trên nhiều model đang mở: `switch_active_document`, `open_document`, `close_document`, `list_open_documents`, `list_recent_documents`",
         u"Work across open models: `switch_active_document`, `open_document`, `close_document`, `list_open_documents`, `list_recent_documents`"),
    ]),
]


def _capabilities_overview(viet):
    """Practical, workflow-grouped overview of what the assistant can DO on the
    live model — curated notes validated against the MCP registry, plus live
    skills and a compact pointer to the openable ribbon tools. See _CAP_MAP."""
    live = _live_tool_names()   # empty set when offline → show every line

    if viet:
        out = [u"Đây là những gì mình làm được **trực tiếp trên model Revit** "
               u"(qua T3Lab connector):"]
    else:
        out = [u"Here's what I can do **directly on the Revit model** "
               u"(via the T3Lab connector):"]

    covered = set()
    for emoji, vi_title, en_title, lines in _CAP_MAP:
        rendered = []
        for names, vi_line, en_line in lines:
            covered.update(names)
            if live and not (set(names) & live):
                continue   # every referenced tool is gone → drop the line
            rendered.append(u"• " + (vi_line if viet else en_line))
        if rendered:
            out.append(u"\n{} **{}**\n{}".format(
                emoji, vi_title if viet else en_title, u"\n".join(rendered)))

    # ── Anti-drift safety net: registered tools the curated map never named ──
    if live:
        extra = sorted(live - covered - _CAP_HIDE)
        if extra:
            chips = u", ".join(u"`{}`".format(n) for n in extra)
            out.append((u"\n🧩 **Khác** (tool mới, chưa xếp nhóm): {}" if viet
                        else u"\n🧩 **Other** (newer tools, not yet grouped): {}"
                        ).format(chips))

    # ── Skills (live from SkillsEngine), personalised if a name is set ───────
    skills = _live_skills()
    if skills:
        name = _user_display_name()
        if viet:
            hdr = (u"\n🎯 **SKILLS chuyên biệt đã cài cho {}**:".format(name)
                   if name else u"\n🎯 **SKILLS chuyên biệt đã cài**:")
        else:
            hdr = (u"\n🎯 **Custom skills set up for {}**:".format(name)
                   if name else u"\n🎯 **Custom skills installed**:")
        s_lines = [hdr]
        for sname, sdesc in skills[:12]:
            s_lines.append(u"• **{}** — {}".format(sname, sdesc) if sdesc
                           else u"• **{}**".format(sname))
        rest = skills[12:]
        if rest:
            names = u", ".join(sn for sn, _ in rest)
            s_lines.append((u"• …+{} skill nữa: {}" if viet
                            else u"• …+{} more: {}").format(len(rest), names))
        out.append(u"\n".join(s_lines))

    # ── Compact pointer to the openable ribbon tools + quick affordances ─────
    total_open = len(_tool_catalog())
    if viet:
        out.append(u"\n📂 Ngoài ra mình mở được **{} tool trên ribbon** — gõ "
                   u"*mở <tên tool>*, hỏi *có tool nào để … không?*, hoặc xuất "
                   u"nhanh *xuất pdf G sheet*.".format(total_open))
    else:
        out.append(u"\n📂 I can also open **{} ribbon tools** — type "
                   u"*open <tool name>*, ask *is there a tool for …?*, or quick-"
                   u"export *export pdf G sheets*.".format(total_open))

    return u"\n".join(out)


def answer_capability_question(user_input, viet):
    """Answer 'do you have a tool for X?' from the real catalog.

    Returns a result dict {intent, params, message, _nlu, _authoritative}.
    _authoritative tells the pipeline NOT to let an LLM override this —
    the catalog is the ground truth for what tools exist.
    """
    expanded = _expand(_norm(user_input))
    clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    func = set()
    for w in clean.split():
        if (w in _STOPWORDS or w in _CAP_BOILERPLATE or w in _OPEN_VERBS
                or len(w) < 2):
            continue
        func.add(_singularise(w))

    # No function words left → generic capability question → full overview
    if not func:
        msg = _capabilities_overview(viet)
        return {"intent": "help", "params": {"answer": msg}, "message": msg,
                "_nlu": True, "_authoritative": True}

    catalog = _tool_catalog()
    # Document frequency — words appearing in ≤2 tools are distinctive
    df = {}
    vocabs = []
    for tool in catalog:
        vocab = tool['words'] | tool['desc_words']
        vocabs.append(vocab)
        for w in vocab:
            df[w] = df.get(w, 0) + 1

    matches, near = [], []
    for tool, vocab in zip(catalog, vocabs):
        inter = func & vocab
        if not inter:
            continue
        score  = len(inter) / float(len(func))
        strong = any(df.get(w, 99) <= 2 and len(w) >= 3 for w in inter)
        if score >= 0.5 or strong:
            matches.append((score + (0.5 if strong else 0.0), tool))
        else:
            near.append((score, tool))
    matches.sort(key=lambda x: -x[0])
    near.sort(key=lambda x: -x[0])

    if matches:
        # "bạn có thể mở X không?" — exact tool named + open verb → just open
        if _has_open_verb(expanded) and matches[0][0] >= 1.4:
            top = matches[0][1]
            msg = (u"Đang mở {}...".format(top['title']) if viet
                   else u"Opening {}...".format(top['title']))
            return {"intent": top['intent'], "params": {}, "message": msg,
                    "_nlu": True, "_authoritative": True}
        lines = []
        for s, t in matches[:3]:
            d = (t.get('desc') or u'').strip()
            lines.append(u"• **{}**{}".format(t['title'],
                                              u" — " + d if d else u""))
        if viet:
            msg = (u"✅ Có! Tool phù hợp:\n{}\n\n"
                   u"Gõ 'mở <tên tool>' để mở nhé.").format(u"\n".join(lines))
        else:
            msg = (u"✅ Yes! Matching tools:\n{}\n\n"
                   u"Type 'open <tool name>' to launch.").format(u"\n".join(lines))
    else:
        near_txt = u", ".join(t['title'] for s, t in near[:3])
        if viet:
            msg = u"❌ Hiện T3Lab chưa có tool riêng cho chức năng đó."
            if near_txt:
                msg += u"\nGần nhất có thể là: {}.".format(near_txt)
            msg += u"\nGõ 'bạn làm được gì' để xem toàn bộ danh sách tool."
        else:
            msg = u"❌ T3Lab doesn't have a dedicated tool for that yet."
            if near_txt:
                msg += u"\nClosest options: {}.".format(near_txt)
            msg += u"\nType 'what can you do' to see the full tool list."
    return {"intent": "help", "params": {"answer": msg}, "message": msg,
            "_nlu": True, "_authoritative": True}


def _last_tool_from_history(history):
    """Scan recent conversation history and return the last tool intent mentioned."""
    if not history:
        return None
    keywords = _tool_keywords()
    for entry in reversed(history[-6:]):
        content = _norm(_expand(_norm(entry.get("content", ""))))
        for kw, intent in keywords.items():
            if kw in content:
                return intent
    return None


def _is_pronoun_query(normed_expanded):
    """Return True if the input looks like a pronoun reference (e.g., 'nó là gì?')."""
    tokens = set(normed_expanded.split())
    return bool(tokens & _PRONOUNS)


# ─── Scoring ──────────────────────────────────────────────────────────────────

def _score(intent, unigrams, bigrams, padded_expanded):
    """Compute weighted match score for one intent.

    Single-word features are matched against the stopword-filtered unigram
    set (keeps the existing noise reduction for generic bag-of-words scoring).
    Multi-word features (2+ words, joined by a space) are matched as a literal
    phrase against the padded expanded text instead of via the bigram set:
    many curated phrases here (e.g. "lam duoc gi", "huong dan su dung") are
    built entirely from short Vietnamese filler words that _tokenise() strips
    as stopwords, so a bigram/trigram for them could never form otherwise —
    they would silently never match.
    """
    score = 0
    all_features = unigrams | bigrams
    for feature, weight in _TRIGGERS.get(intent, []):
        if " " in feature:
            if (" " + feature + " ") in padded_expanded:
                score += weight
        elif feature in all_features:
            score += weight
    for feature, penalty in _PENALTIES.get(intent, []):
        if " " in feature:
            matched = (" " + feature + " ") in padded_expanded
        else:
            matched = feature in all_features
        if matched:
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

# Per-intent canned messages. Only intents WITHOUT a registry entry need one —
# an auto-discovered tool gets "Đang mở <title>..." / "Opening <title>..."
# built from its live title in _build_message().
_MESSAGES_VI = {
    "open_batchout":          u"Đang mở BatchOut...",
    "open_batchout_configured": u"Mở BatchOut đã cấu hình...",
    "open_loadfamily":        u"Đang mở Family Loader...",
    "check_spelling":         u"Đang quét Text Note trong model để kiểm tra chính tả tiếng Anh...",
    "greet":  u"Xin chào! Tôi là T3Lab Assistant 👋\nBạn muốn làm gì hôm nay?",
    "farewell": u"Tạm biệt! Gặp lại bạn sau nhé 👋",
    "chat":   u"Không có gì! Cần gì cứ hỏi tôi nhé.",
    "help":   (u"Tôi có thể giúp bạn:\n"
               u"• Xuất sheet: 'xuất pdf G sheet', 'in tất cả sang dwg'\n"
               u"• Mở tool: 'mở batchout', 'load family', 'mở manaviews'\n"
               u"• Cấu hình nhanh: 'mở batchout G sheet pdf'\n"
               u"Gõ tên tool hoặc mô tả điều bạn muốn làm!"),
}

_MESSAGES_EN = {
    "open_batchout":          "Opening BatchOut...",
    "open_batchout_configured": "Opening BatchOut (pre-configured)...",
    "open_loadfamily":        "Opening Family Loader...",
    "check_spelling":         "Scanning model Text Notes for English spelling errors...",
    "greet":   "Hello! I'm T3Lab Assistant 👋\nWhat would you like to do today?",
    "farewell": "Goodbye! See you later 👋",
    "chat":    "You're welcome! Let me know if you need anything.",
    "help":    ("I can help you:\n"
                "• Export sheets: 'export pdf G sheet', 'print all to dwg'\n"
                "• Open tools: 'open batchout', 'load family', 'open manaviews'\n"
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
        # Frustration / insult directed at the assistant itself — acknowledge
        # honestly instead of the generic "didn't understand" reply, and point
        # at a concrete next step (works with or without an LLM connected).
        insult_kws = ["stupid", "dumb", "useless", "garbage", "trash", "suck",
                      "ngu", "vo dung", "te qua", "qua te"]
        if any(k in normed_raw for k in insult_kws):
            if viet:
                return (u"Xin lỗi vì trải nghiệm chưa tốt! Ở chế độ offline khả năng "
                        u"của tôi hạn chế — kết nối AI trong phần Cài đặt để trả "
                        u"lời tự nhiên hơn.")
            return ("Sorry that reply wasn't good enough! Offline mode is limited — "
                    "connect an AI provider in Settings for smarter answers.")
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

    table = _MESSAGES_VI if viet else _MESSAGES_EN
    if intent in table:
        return table[intent]

    # Auto-discovered tools have no canned message — build it from the live
    # registry title so a newly added pushbutton reads correctly without a
    # matching entry in the tables above.
    if intent and intent.startswith("open_"):
        label = _tool_label(intent)
        return (u"Đang mở {}...".format(label) if viet
                else u"Opening {}...".format(label))

    return u"Đang xử lý..." if viet else "Processing..."


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
            label = _tool_label(last_tool)
            msg = (u"Đang mở {}...".format(label) if viet
                   else "Opening {}...".format(label))
            return {"intent": last_tool, "params": {}, "message": msg,
                    "_nlu": True}

    # ── Exact tool-name hit → deterministic resolver first ───────────────────
    # A message that IS a tool name (with or without "mở/open") must always
    # open exactly that tool — the trigger tables below only know the builtin
    # tools and the LLM guesses. Exactness consumes every token, so commands
    # with extra words ("mở batchout G sheet pdf" → configured export,
    # "mcp control là gì" → help) still fall through to normal scoring.
    _tool, _ = resolve_tool(user_input, exact_only=True)
    if _tool:
        msg = (u"Đang mở {}...".format(_tool['title']) if viet
               else u"Opening {}...".format(_tool['title']))
        return {"intent": _tool['intent'], "params": {}, "message": msg,
                "_nlu": True}

    # ── Capability questions → answered from the real catalog ────────────────
    # "Có tool nào để X không?" gets a truthful yes (with the matching tools)
    # or a truthful no — never an LLM guess.
    if is_capability_question(expanded):
        return answer_capability_question(user_input, viet)

    # ── Tokenise ─────────────────────────────────────────────────────────────
    unigrams, bigrams = _tokenise(expanded)
    # Punctuation-stripped, whitespace-normalised text for the multi-word
    # phrase matcher in _score() — must match what _tokenise() sees, or a
    # trailing "?"/"." blocks the closing word-boundary space and the phrase
    # never matches (e.g. "what can you do?").
    _clean = re.sub(r'[^a-z0-9\s]', ' ', expanded)
    padded_expanded = u" " + u" ".join(_clean.split()) + u" "

    # ── Score every intent ───────────────────────────────────────────────────
    scores = {intent: _score(intent, unigrams, bigrams, padded_expanded)
              for intent in _TRIGGERS}

    # ── Extract slots (needed for disambiguation) ────────────────────────────
    slots = _extract_slots(user_input)

    # ── Disambiguate ─────────────────────────────────────────────────────────
    best = _disambiguate(dict(scores), unigrams, bigrams, slots)

    # ── Ranked tool resolver (offline, no LLM round-trip needed) ─────────────
    # Nothing in the trigger tables matched — try the full tool catalog with
    # confidence + margin rules. A clear winner opens; an ambiguous "open X"
    # asks the user to pick instead of guessing (or letting the LLM guess).
    if best is None:
        _tool, _cands = resolve_tool(user_input)
        if _tool:
            msg = (u"Đang mở {}...".format(_tool['title']) if viet
                   else u"Opening {}...".format(_tool['title']))
            return {"intent": _tool['intent'], "params": {}, "message": msg,
                    "_nlu": True}
        if _cands and _has_open_verb(expanded):
            names = u"\n".join(u"• {}".format(c['title']) for c in _cands)
            if viet:
                msg = (u"Bạn muốn mở tool nào? Tôi tìm thấy các tool gần giống:\n"
                       u"{}\nGõ đúng tên tool để mở chính xác nhé!".format(names))
            else:
                msg = ("Which tool do you mean? Closest matches:\n"
                       "{}\nType the exact tool name to open it!".format(names))
            return {"intent": "chat", "params": {}, "message": msg,
                    "_nlu": True, "_authoritative": True}

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
            # _generic_fallback marks this as a placeholder guess, not a real
            # answer — script.py uses it to avoid masking a clearer "the AI
            # didn't respond" message behind this same canned text whenever
            # the LLM call itself times out or errors.
            return {"intent": "chat", "params": {}, "message": msg, "_nlu": True,
                    "_generic_fallback": True}
        return None

    # ── Build result ─────────────────────────────────────────────────────────
    params = {}
    if best in ("export_direct", "open_batchout_configured"):
        params = slots
    elif best == "open_batchout":
        params = {}

    message = _build_message(best, slots, viet, raw_input=user_input)
    return {"intent": best, "params": params, "message": message, "_nlu": True}
