# -*- coding: utf-8 -*-
"""
Deterministic English spell-checker — dictionary + edit-distance engine.

This is the "intelligence" behind /english-spellcheck that does NOT depend on
a local LLM. The previous pipeline handed each batch of Text Notes to the
connected model (`router.chat`) and trusted it to spot typos; a small local
model (llama3.1:8b, qwen, ...) routinely answered NO_ERRORS on obvious
mistakes (CEMTITIOUS, ACYLIC, WATERPROFFING...), so the report came back clean
while Claude-via-MCP flagged them all.

Instead we use the same algorithm class as pyspellchecker / Peter Norvig's
corrector:
  * a bundled English word-frequency dictionary (data/en_freq.json.gz), merged
    with construction/architecture domain terms, decides what is a real word;
  * unknown tokens get edit-distance-1/2 candidates, ranked by corpus
    frequency, to produce a suggested correction;
  * a small, high-precision confusables list (data/confusables.json) catches
    valid-word/wrong-word cases the dictionary cannot (e.g. "to advice" ->
    "to advise").

Pure logic, IronPython 2.7 safe — no Revit, no WPF, no network. The bundled
asset is built by dev/build_spell_dictionary.py.

Author: Tran Tien Thanh
"""

import os
import re
import gzip
import json

# ─── Data loading (lazy, cached) ────────────────────────────────────────────

_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
_FREQ_FILE = os.path.join(_DATA_DIR, "en_freq.json.gz")
_CONFUSABLES_FILE = os.path.join(_DATA_DIR, "confusables.json")
_TERMS_FILE = os.path.join(_DATA_DIR, "construction_terms.txt")
_CUSTOM_FILE = os.path.join(_DATA_DIR, "custom_words.txt")

_WORDS = None            # {word(lower): frequency}
_DOMAIN = None           # set(single-word lowercase construction terms)
_CONFUSABLES = None      # [{"find": compiled_re, "sub": str, "reason": str}]
_LOAD_ERROR = None       # remembered so callers can surface a real failure


def _load_words():
    """Load and cache the frequency dictionary. Returns {} on failure (the
    caller treats an empty dictionary as 'engine unavailable').

    The word list bundled in the .gz is merged at load time with the plain-text
    term files (construction_terms.txt + custom_words.txt) so a word added
    there takes effect on the next run — no rebuild of the .gz asset needed.
    """
    global _WORDS, _LOAD_ERROR
    if _WORDS is not None:
        return _WORDS
    try:
        fh = gzip.open(_FREQ_FILE, "rb")
        try:
            raw = fh.read()
        finally:
            fh.close()
        _WORDS = json.loads(raw.decode("utf-8"))
    except Exception as ex:
        _LOAD_ERROR = u"{}".format(ex)
        _WORDS = {}
    if _WORDS:
        for term in _load_domain():
            if _WORDS.get(term, 0) < _DOMAIN_FLOOR:
                _WORDS[term] = _DOMAIN_FLOOR
    return _WORDS


def _read_terms(path):
    """Lowercase terms from a '# comment' + one-word-per-line text file.
    Hyphen/slash compounds contribute both the whole term and its parts."""
    terms = set()
    try:
        fh = open(path, "rb")
        try:
            raw = fh.read()
        finally:
            fh.close()
        for line in raw.decode("utf-8").splitlines():
            w = line.strip().lower()
            if not w or w.startswith(u"#"):
                continue
            if len(w) >= _MIN_LEN and w.replace(u"-", u"").replace(u"/", u"").isalpha():
                terms.add(w)
            for part in re.split(u"[-/ ]", w):        # split compounds
                if len(part) >= _MIN_LEN and part.isalpha():
                    terms.add(part)
    except Exception:
        pass
    return terms


def _load_domain():
    """Single-word construction/architecture terms (bundled list + the user's
    own custom_words.txt). They are accepted as correctly spelled AND used to
    break ties in the correction ranking: on a construction drawing a typo one
    edit from a trade term (SETING -> SETTING) should win over a more frequent
    general word (SEEING), which raw corpus frequency would otherwise pick."""
    global _DOMAIN
    if _DOMAIN is not None:
        return _DOMAIN
    _DOMAIN = _read_terms(_TERMS_FILE) | _read_terms(_CUSTOM_FILE)
    return _DOMAIN


def _load_confusables():
    global _CONFUSABLES
    if _CONFUSABLES is not None:
        return _CONFUSABLES
    rules = []
    try:
        fh = open(_CONFUSABLES_FILE, "rb")
        try:
            raw = fh.read()
        finally:
            fh.close()
        for r in json.loads(raw.decode("utf-8")):
            try:
                rules.append({
                    "find": re.compile(r["find"], re.IGNORECASE | re.UNICODE),
                    "sub":  r["sub"],
                    "reason": r.get("reason", u""),
                })
            except Exception:
                continue
    except Exception:
        rules = []
    _CONFUSABLES = rules
    return _CONFUSABLES


def is_available():
    """True when the dictionary asset loaded — i.e. the deterministic engine
    can run without any AI provider."""
    return bool(_load_words())


def load_error():
    _load_words()
    return _LOAD_ERROR


# ─── Tokenising & classification ────────────────────────────────────────────

# A token is a run of letters with optional internal apostrophes. Digits,
# hyphens and slashes act as separators, so "SETTING-OUT" -> SETTING, OUT and
# "100THK" -> THK (the digits never enter a token).
# The class is Unicode-aware on purpose: with a plain [A-Za-z]+ an accented or
# Vietnamese letter acts as a separator and shreds the word into fake tokens
# ("Façades" -> "Fa" + "ades" -> flagged as "ades" -> "ages"). Non-ASCII tokens
# are then skipped wholesale by _ignorable() — this is an English checker.
_TOKEN_RE = re.compile(u"[^\\W\\d_]+(?:['’][^\\W\\d_]+)*", re.UNICODE)
_ROMAN_RE = re.compile(u"^m{0,4}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$",
                       re.IGNORECASE)
_ALPHABET = u"abcdefghijklmnopqrstuvwxyz"

# Tokens shorter than this are never checked (single letters, most codes).
_MIN_LEN = 3
# Unknown ALL-CAPS tokens up to this length are treated as un-whitelisted
# drawing acronyms (ELV, SWBD, ...) and skipped. The shortest real typo we
# must still catch is ACYLIC (6), so 5 is safe.
_ACRONYM_MAX = 5
# Above this length we skip edit-distance-2 (too expensive / almost always a
# product code); above the hard cap we skip the token entirely.
_EDITS2_MAX = 15
_HARD_MAX = 28
# Frequency floor given to bundled domain terms, and the minimum frequency each
# half of a compound must have before the compound is accepted (see
# _is_compound). Tuned on 2085 synthetic single-edit typos: at this floor the
# rule accepts every realistic compound in the test list (COUNTERTOP,
# BACKSPLASH, STORMWATER, WORKTOP, ...) while wrongly accepting 1.2% of typos.
_DOMAIN_FLOOR = 8000
_COMPOUND_MIN_FREQ = 8000
_COMPOUND_MIN_PART = 3
# Productive prefixes: PREFINISHED, NON-SLIP written NONSLIP, SUBFRAME...
# The remainder must be at least _PREFIX_MIN_REST letters so that a genuine
# typo whose tail happens to be a short word (PRESURE -> pre + "sure") is
# still reported.
_PREFIXES = (u"pre", u"re", u"un", u"non", u"anti", u"semi", u"sub", u"super",
             u"over", u"under", u"multi", u"inter", u"intra", u"micro",
             u"mini", u"mid", u"co", u"de", u"dis", u"mis", u"out", u"up",
             u"post", u"self", u"cross", u"counter", u"back", u"down", u"off",
             u"bi", u"tri")
_PREFIX_MIN_REST = 5


def _is_compound(word, vocab):
    """True when the word is a closed compound of two common words, e.g.
    COUNTERTOP = counter + top, BACKSPLASH = back + splash, STORMWATER.
    Such compounds are absent from a general corpus but are perfectly correct
    on a drawing, and edit-distance would 'fix' them into nonsense
    (COUNTERTOP -> COUNTERTYPE, BACKSPLASH -> BACKSLASH)."""
    n = len(word)
    for i in range(_COMPOUND_MIN_PART, n - _COMPOUND_MIN_PART + 1):
        if (vocab.get(word[:i], 0) >= _COMPOUND_MIN_FREQ
                and vocab.get(word[i:], 0) >= _COMPOUND_MIN_FREQ):
            return True
    return False


def _is_prefixed(word, vocab):
    """True when the word is a productive prefix + a known word
    (PREFINISHED, REGLAZING, INTERLAYER, DOWNLIGHT...)."""
    for p in _PREFIXES:
        if word.startswith(p) and len(word) - len(p) >= _PREFIX_MIN_REST:
            if word[len(p):] in vocab:
                return True
    return False


def _is_wellformed(word, vocab, _depth=0):
    """True when the word is spelled correctly: in the dictionary, or built by
    ordinary English word-formation from dictionary words. Everything the
    checker reports has failed this test."""
    if word in vocab:
        return True
    if _is_compound(word, vocab) or _is_prefixed(word, vocab):
        return True
    # simple plural of an otherwise well-formed form (COUNTERTOPS, DOWNLIGHTS)
    if _depth == 0 and len(word) > 4 and word.endswith(u"s"):
        return _is_wellformed(word[:-1], vocab, 1)
    return False


def _looks_like_code(token):
    """Heuristic: a mixed-case/alpha token that is really a drawing code, not a
    word — e.g. one with no vowels, or CamelCase/part-number shapes."""
    low = token.lower()
    # no vowels at all -> not an English word (RWP already handled by acronym
    # rule; this catches things like 'gljmp')
    if not any(v in low for v in u"aeiouy"):
        return True
    return False


def _has_non_ascii(token):
    for ch in token:
        if ord(ch) > 127:
            return True
    return False


def _ignorable(token):
    """True when the token should never be treated as a spelling error,
    regardless of whether it is in the dictionary."""
    if len(token) < _MIN_LEN or len(token) > _HARD_MAX:
        return True
    if _has_non_ascii(token):
        return True                       # foreign / accented word, not English
    if _ROMAN_RE.match(token):
        return True
    if token.isupper() and len(token) <= _ACRONYM_MAX:
        return True
    if _looks_like_code(token):
        return True
    return False


def _apply_case(model, replacement):
    """Give `replacement` the capitalisation pattern of the original `model`
    token (ALL-CAPS drawing text -> ALL-CAPS suggestion, Titlecase -> Title)."""
    if model.isupper():
        return replacement.upper()
    if model[:1].isupper() and model[1:].islower():
        return replacement[:1].upper() + replacement[1:]
    return replacement


# ─── Edit-distance correction (Norvig) ──────────────────────────────────────

def _edits1(word):
    splits = [(word[:i], word[i:]) for i in range(len(word) + 1)]
    deletes = [a + b[1:] for a, b in splits if b]
    transposes = [a + b[1] + b[0] + b[2:] for a, b in splits if len(b) > 1]
    replaces = [a + c + b[1:] for a, b in splits if b for c in _ALPHABET]
    inserts = [a + c + b for a, b in splits for c in _ALPHABET]
    return set(deletes + transposes + replaces + inserts)


def _known(words, vocab):
    return set(w for w in words if w in vocab)


def _best(cands, vocab, domain):
    """Pick the winning correction from a candidate set. Construction terms
    are preferred as a group; frequency ranks within the chosen group."""
    domain_cands = [w for w in cands if w in domain]
    pool = domain_cands if domain_cands else list(cands)
    return max(pool, key=lambda w: vocab.get(w, 0))


def correction(word):
    """Best correction for a lowercase word, or the word itself when the engine
    has nothing confident to offer."""
    vocab = _load_words()
    if not vocab or _is_wellformed(word, vocab):
        return word
    domain = _load_domain()
    e1 = _known(_edits1(word), vocab)
    if e1:
        return _best(e1, vocab, domain)
    if len(word) <= _EDITS2_MAX:
        e2 = set()
        for w1 in _edits1(word):
            e2 |= _edits1(w1)
        e2 = _known(e2, vocab)
        if e2:
            return _best(e2, vocab, domain)
    return word


# ─── Public API ─────────────────────────────────────────────────────────────

def check_text(text):
    """Proofread one string. Returns a list of findings:
        [{"wrong": u"CEMTITIOUS", "right": u"CEMENTITIOUS",
          "kind": "spelling"|"wording", "reason": u""}]
    Spelling findings come from the dictionary; wording findings from the
    confusables rules. Order follows appearance in the text; duplicate
    wrong-tokens within one string are reported once.
    """
    vocab = _load_words()
    findings, seen = [], set()
    if text and vocab:
        for m in _TOKEN_RE.finditer(text):
            token = m.group(0)
            if _ignorable(token):
                continue
            low = token.lower().replace(u"’", u"'")
            if _is_wellformed(low, vocab):
                continue
            best = correction(low)
            if best == low:
                continue                       # no confident suggestion
            key = token.lower()
            if key in seen:
                continue
            seen.add(key)
            findings.append({
                "wrong": token,
                "right": _apply_case(token, best),
                "kind": "spelling",
                "reason": u"",
            })

    for rule in _load_confusables():
        for m in rule["find"].finditer(text or u""):
            frag = m.group(0)
            fixed = rule["find"].sub(rule["sub"], frag)
            if frag.isupper():
                fixed = fixed.upper()
            key = frag.lower()
            if key in seen:
                continue
            seen.add(key)
            findings.append({
                "wrong": frag,
                "right": fixed,
                "kind": "wording",
                "reason": rule.get("reason", u""),
            })
    return findings


def format_issue(findings):
    """Render check_text() findings into the '"wrong" -> "right"' string used
    by the report bubble and the interactive card (matches the LLM format so
    the downstream fix agent parses it identically)."""
    parts = []
    for f in findings:
        seg = u'"{}" -> "{}"'.format(f["wrong"], f["right"])
        if f.get("reason"):
            seg += u" ({})".format(f["reason"])
        parts.append(seg)
    return u"; ".join(parts)
