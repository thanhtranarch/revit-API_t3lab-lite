# -*- coding: utf-8 -*-
"""
Input normalisation — what the user typed → what they meant.

Every stage downstream (dispatcher keywords, NLU triggers, planner splitting,
BM25 retrieval) matches literal words. So a message only has to be typed the
way people actually type in a Vietnamese office to miss every table:

    "ktra warning ko"      → no keyword matched → routed to `general`
    "xoas het text note"   → typo → the action verb was invisible
    "xuat pdf dc ko a"     → teencode tail → export keywords survived, but
                              the question shape did not
    "OKKKKK"               → not an affirmation, because of the K's

This module repairs the input ONCE, and everything downstream matches against
the repaired text. It is deliberately conservative: it only rewrites tokens it
has an explicit rule for, plus single-edit typos against a closed domain
lexicon. It never invents words, and it never touches text inside quotes,
backticks, file paths, sheet numbers, or parameter names — those are data.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re

from Intelligence.knowledge import vi_text

# ─── Teencode / chat shorthand → full words ──────────────────────────────────
# Vietnamese chat shorthand. Written in full diacritics on the right so the
# expansion also RESTORES tone marks the user dropped — the folded form then
# lands in the same bucket as properly typed input.
_SHORTHAND_VI = {
    'ko': u'không', 'k': u'không', 'kh': u'không', 'khg': u'không',
    'hok': u'không', 'hong': u'không', 'khong': u'không',
    'dc': u'được', 'đc': u'được', 'duoc': u'được', 'dk': u'được',
    'j': u'gì', 'ji': u'gì', 'gi': u'gì',
    'z': u'vậy', 'v': u'vậy', 'vay': u'vậy',
    'ntn': u'như thế nào', 'nth': u'như thế nào',
    'tn': u'thế nào', 'sn': u'sao nữa',
    'bn': u'bao nhiêu', 'bnh': u'bao nhiêu',
    'mn': u'mọi người', 'ae': u'anh em',
    'tks': u'cảm ơn', 'thanks': u'cảm ơn', 'tk': u'cảm ơn',
    'hnay': u'hôm nay', 'hqua': u'hôm qua',
    'bjo': u'bao giờ', 'bgio': u'bao giờ',
    'cx': u'cũng', 'nhg': u'nhưng', 'nhug': u'nhưng',
    'vs': u'với', 'vơi': u'với',
    'ttcả': u'tất cả', 'tatca': u'tất cả', 'tc': u'tất cả',
    'tb': u'toàn bộ', 'tbo': u'toàn bộ',
    'ktra': u'kiểm tra', 'kt': u'kiểm tra', 'ktr': u'kiểm tra',
    'tke': u'thống kê', 'tket': u'thống kê',
    'sl': u'số lượng', 'slg': u'số lượng',
    'ht': u'hiện tại', 'htai': u'hiện tại',
    'hthanh': u'hoàn thành', 'ht.': u'hiện tại',
    'lm': u'làm', 'lam': u'làm',
    'đag': u'đang', 'dag': u'đang',
    'ng': u'người', 'nguoi': u'người',
    'trog': u'trong', 'trg': u'trong',
    'mún': u'muốn', 'mun': u'muốn',
    'giup': u'giúp', 'gium': u'giúp', 'dum': u'giúp', 'giùm': u'giúp',
    'oke': u'ok', 'okie': u'ok', 'okla': u'ok', 'uk': u'ừ', 'um': u'ừ',
    'ah': u'à', 'ak': u'à', 'aj': u'à',
    'nhé': u'nhé', 'nhek': u'nhé', 'nhá': u'nhé',
    'r': u'rồi', 'ro': u'rồi', 'roy': u'rồi',
    'cta': u'chúng ta', 'ctoi': u'chúng tôi',
    'pj': u'project', 'prj': u'project',
    'đm': u'định mức',
}

# Domain shorthand that means the same thing in both languages.
_SHORTHAND_DOMAIN = {
    'lv': u'level', 'lvl': u'level', 'lvls': u'levels',
    'sh': u'sheet', 'shs': u'sheets', 'shts': u'sheets',
    'vw': u'view', 'vws': u'views',
    'cat': u'category', 'cats': u'categories',
    'param': u'parameter', 'params': u'parameters', 'prm': u'parameter',
    'fam': u'family', 'fams': u'families',
    'vt': u'view template', 'vts': u'view templates',
    'wt': u'wall type', 'ft': u'floor type',
    'ws': u'workset', 'wss': u'worksets',
    'qty': u'quantity', 'qtys': u'quantities',
    'elem': u'element', 'elems': u'elements', 'el': u'element',
    'dim': u'dimension', 'dims': u'dimensions',
    'annot': u'annotation', 'annots': u'annotations',
    'temp': u'template', 'tmpl': u'template',
    'cfg': u'config', 'config': u'config',
    'doc': u'document', 'docs': u'documents',
    'mep': u'mep', 'arch': u'architecture',
    'bo': u'batchout', 'batch out': u'batchout',
    'plz': u'please', 'pls': u'please', 'ur': u'your',
    'thx': u'thanks', 'asap': u'urgent',
}

# ─── Domain lexicon for one-edit typo repair ─────────────────────────────────
# Folded, lowercase, length >= _MIN_TYPO_LEN. Only words that CHANGE ROUTING
# are listed: verbs, Revit categories, tool names. Repairing a typo in a noun
# the router never reads buys nothing and risks corrupting a parameter name.
#
# SAFETY NOTE — the repaired text is used for ROUTING SIGNALS ONLY. The message
# sent to the model, saved to history, and shown in the transcript is always
# the user's original wording. A wrong repair can therefore nudge a route; it
# can never put words in the user's mouth.
_TYPO_LEXICON = frozenset([
    # Vietnamese verbs / nouns that decide a route
    'chuyen', 'kiem', 'thong', 'liet', 'xuat', 'nhap', 'khoa',
    'tuong', 'thang', 'tran', 'vach', 'luoi', 'truc', 'khung', 'thich',
    'thuoc', 'nhan', 'huong', 'giai', 'sanh', 'nhieu', 'phong', 'mong',
    # English verbs
    'delete', 'remove', 'create', 'rename', 'select', 'export', 'import',
    'update', 'change', 'modify', 'count', 'check', 'audit', 'purge',
    'filter', 'place', 'color', 'highlight', 'isolate', 'schedule',
    # English nouns. Singular AND plural are both listed: membership means
    # "already correct, never repair", so listing only "sheets" made the
    # perfectly good "sheet" its own one-edit typo and rewrote it.
    'wall', 'walls', 'floor', 'floors', 'column', 'columns',
    'window', 'windows', 'room', 'rooms', 'stair', 'stairs',
    'railing', 'railings', 'ceiling', 'ceilings', 'sheet', 'sheets',
    'view', 'views', 'level', 'levels', 'grid', 'grids',
    'door', 'doors', 'beam', 'beams', 'roof', 'roofs',
    'warning', 'warnings', 'family', 'families', 'parameter',
    'parameters', 'category', 'element', 'elements', 'workset', 'worksets',
    'schedules', 'template', 'templates', 'material', 'materials',
    # Tool names
    'batchout', 'parasync', 'loadfamily', 'projectname',
    'dimtext', 'autodimension', 'tilelayout',
])

# Words that sit one edit from a lexicon entry but are perfectly ordinary
# English. Without this guard "sweet" became "sheet" and "wells" became
# "walls". Listed explicitly because there is no dictionary to consult under
# IronPython — the alternative is silently rewriting real words.
_NEVER_REPAIR = frozenset([
    'sweet', 'sheer', 'shoot', 'sheep', 'shelf', 'wells', 'bells', 'balls',
    'calls', 'falls', 'halls', 'malls', 'tells', 'sells', 'small', 'smell',
    'shell', 'steel', 'still', 'stall', 'skill', 'spill', 'flour', 'floors',
    'flood', 'blood', 'color', 'colour', 'coder', 'cover', 'count', 'court',
    'clear', 'clean', 'close', 'chance', 'change', 'charge', 'create',
    'crate', 'grade', 'grids', 'girls', 'goods', 'weeks', 'walks', 'wants',
    'lever', 'levels', 'lovely', 'lively', 'family', 'famous', 'filter',
    'filler', 'folder', 'former', 'faster', 'window', 'widow', 'winter',
    'render', 'reader', 'header', 'leader', 'loader', 'better', 'letter',
])

# Tokens shorter than this are never typo-repaired: below 5 characters most
# lexicon words are one edit from several others, so repair is a coin flip.
_MIN_TYPO_LEN = 5

# ─── Telex leftovers ─────────────────────────────────────────────────────────
# Telex encodes tones as a TRAILING LETTER (s = sắc, f = huyền, r = hỏi,
# x = ngã, j = nặng): "xoas" → "xóa", "tuowngf" → "tường". When the IME is off,
# or the user types into a field that swallowed it, the tone letter is left
# stranded on the end of an otherwise correct word. Stripping it is safe as
# long as the remainder is a word we actually route on.
_TELEX_TONES = frozenset(['s', 'f', 'r', 'x', 'j'])

_TELEX_BASES = frozenset([
    'xoa', 'sua', 'tao', 'dat', 'gan', 'doi', 'them', 'bot', 'chuyen',
    'kiem', 'tra', 'thong', 'ke', 'liet', 'dem', 'tim', 'loc', 'xuat',
    'nhap', 'mo', 'dong', 'luu', 'chon', 'an', 'hien', 'khoa', 'ghim',
    'to', 'boi', 'danh', 'dung', 've', 'sao', 'chep', 'nhan', 'xoay',
    'tuong', 'san', 'cot', 'dam', 'cua', 'mai', 'phong', 'tang', 'cau',
    'thang', 'tran', 'vach', 'luoi', 'truc', 'khung', 'thuoc', 'nhin',
    'bao', 'nhieu', 'huong', 'dan', 'giai', 'thich', 'sanh', 'canh',
    'mot', 'hai', 'ba', 'bon', 'nam', 'tat', 'toan', 'moi', 'tung',
])


def _strip_telex_tone(folded_token):
    """'xoas' → 'xoa'. None when the token is not a stranded-tone word."""
    if len(folded_token) < 3 or folded_token[-1] not in _TELEX_TONES:
        return None
    base = folded_token[:-1]
    if base in _TELEX_BASES and folded_token not in _TELEX_BASES:
        return base
    return None

# ─── Protected spans ──────────────────────────────────────────────────────────
# Data, not language. Masked out before rewriting and restored afterwards, so a
# sheet number ("A-101"), a path, a quoted type name or a parameter reference
# survives the normaliser untouched.
_PROTECT_RE = re.compile(
    r'"[^"]*"'                       # "double quoted"
    r"|'[^']{2,}'"                   # 'single quoted' (2+ chars: not an apostrophe)
    r'|`[^`]*`'                      # `backticked`
    r'|\b[A-Za-z]:\\[^\s]+'          # C:\path\file.rvt
    r'|\\\\[^\s]+'                   # \\server\share
    r'|https?://[^\s]+'              # url
    r'|\b[A-Za-z]{1,3}-?\d{2,4}[A-Za-z]?\b'   # A-101, A101, EL204b
    r'|\b\d+(?:[.,]\d+)?\s?(?:mm|cm|m|km|mm2|m2|m3|%)\b',  # 3000mm, 2.5 m
    re.UNICODE)

_WS_RE = re.compile(r'\s+', re.UNICODE)
_REPEAT_RE = re.compile(r'([a-zA-ZÀ-ỹ])\1{2,}', re.UNICODE)
_TOKEN_RE = re.compile(r'[^\W\d_]+', re.UNICODE)

# Unicode punctuation people paste in from Word / Zalo.
_PUNCT_MAP = {
    u'\u2018': u"'", u'\u2019': u"'", u'\u201c': u'"', u'\u201d': u'"',
    u'\u2013': u'-', u'\u2014': u'-', u'\u2026': u'...', u'\u00a0': u' ',
    u'\uff0c': u',', u'\uff1a': u':', u'\uff1f': u'?', u'\uff01': u'!',
}


def fold(text):
    """Diacritic-folded, whitespace-collapsed lowercase — the match key."""
    folded = vi_text.fold_diacritics(text or u'').lower()
    return _WS_RE.sub(u' ', folded).strip()


def _clean_punctuation(text):
    out = []
    for ch in text:
        out.append(_PUNCT_MAP.get(ch, ch))
    return u''.join(out)


def _protect(text):
    """Replace protected spans with placeholders. Returns (masked, spans)."""
    spans = []

    def _sub(m):
        spans.append(m.group(0))
        # The placeholder must survive tokenisation as ONE opaque token.
        return u'\x00{}\x00'.format(len(spans) - 1)

    return _PROTECT_RE.sub(_sub, text), spans


def _restore(text, spans):
    for i, original in enumerate(spans):
        text = text.replace(u'\x00{}\x00'.format(i), original)
    return text


def _collapse_repeats(text):
    """okkkkk → okk, khoongggg → khoong.

    Two characters are kept rather than one: Vietnamese has real doubled
    letters ("xuống" has none, but "ss"/"nn" appear in loanwords and in
    English "process"), and the shorthand table already handles the rest.
    """
    return _REPEAT_RE.sub(lambda m: m.group(1) * 2, text)


def _edit_distance_le1(a, b):
    """True when `a` and `b` are one edit apart (Damerau: transposition counts).

    Adjacent transposition is included because it is the single most common
    typing error — "deleet"/"delete", "tuwong"/"tuowng" — and plain Levenshtein
    scores it as two edits, which put every one of them out of reach.
    """
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if a == b:
        return True
    if la == lb:
        diffs = [i for i in range(la) if a[i] != b[i]]
        if len(diffs) == 1:                       # one substitution
            return True
        if len(diffs) == 2:                       # adjacent transposition?
            i, j = diffs
            return j == i + 1 and a[i] == b[j] and a[j] == b[i]
        return False
    # one insert/delete — walk the shorter against the longer
    short, long_ = (a, b) if la < lb else (b, a)
    i = j = 0
    skipped = False
    while i < len(short) and j < len(long_):
        if short[i] == long_[j]:
            i += 1
            j += 1
        elif skipped:
            return False
        else:
            skipped = True
            j += 1
    return True


def _repair_typo(folded_token):
    """Closest lexicon word within one edit, or None when ambiguous/absent."""
    if len(folded_token) < _MIN_TYPO_LEN:
        return None
    if folded_token in _TYPO_LEXICON or folded_token in _NEVER_REPAIR:
        return None
    matches = [w for w in _TYPO_LEXICON
               if abs(len(w) - len(folded_token)) <= 1
               and _edit_distance_le1(folded_token, w)]
    # Exactly one candidate. Two means we would be guessing which word the
    # user meant, and guessing is what this module exists to avoid.
    return matches[0] if len(matches) == 1 else None


def expand_shorthand(text, domain=True):
    """Expand teencode / abbreviations token by token.

    `domain=False` skips the Revit abbreviation table — used when the caller
    only wants conversational repair (e.g. deciding whether a reply is an
    affirmation) and must not turn a user's "v" into "view".
    """
    masked, spans = _protect(text or u'')

    def _sub(m):
        word = m.group(0)
        low = word.lower()
        rep = _SHORTHAND_VI.get(low)
        if rep is None and domain:
            rep = _SHORTHAND_DOMAIN.get(low)
        if rep is None:
            return word
        # Keep a leading capital so sentence case survives.
        if word[:1].isupper() and len(word) > 1:
            return rep[:1].upper() + rep[1:]
        return rep

    return _restore(_TOKEN_RE.sub(_sub, masked), spans)


def repair_typos(text):
    """Telex-leftover + one-edit repair against the routing lexicon.

    Conservative by design — see the SAFETY NOTE on `_TYPO_LEXICON`.
    """
    masked, spans = _protect(text or u'')

    def _sub(m):
        word = m.group(0)
        folded = vi_text.fold_diacritics(word).lower()
        # A word that already carries diacritics was typed deliberately —
        # "sửa" must never be "repaired" into "sứa" or anything else.
        if folded != word.lower():
            return word
        telex = _strip_telex_tone(folded)
        if telex:
            return telex
        fixed = _repair_typo(folded)
        return fixed if fixed else word

    return _restore(_TOKEN_RE.sub(_sub, masked), spans)


def normalize(text, domain=True, typos=True):
    """Full pipeline: punctuation → repeats → shorthand → typos → whitespace.

    Returns the repaired text with its original alphabet (diacritics restored
    where the shorthand table knows them). Fold it with `fold()` when a match
    key is what you need.
    """
    if not text:
        return u''
    out = _clean_punctuation(text)
    out = _collapse_repeats(out)
    out = expand_shorthand(out, domain=domain)
    if typos:
        out = repair_typos(out)
    return _WS_RE.sub(u' ', out).strip()
