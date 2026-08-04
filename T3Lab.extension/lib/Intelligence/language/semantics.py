# -*- coding: utf-8 -*-
"""
Sentence semantics — what the user is DOING with the words, not which words.

The dispatcher's keyword tables answer "which domain?". They cannot answer the
four questions that decide whether an action is safe to take at all:

    negation   "đừng xóa text note"      — a delete verb inside a prohibition
    modality   "xóa được không?"          — a QUESTION about deleting, not an order
    scope      "tô đỏ tường"              — the whole project, or just this view?
    risk       "purge unused"             — irreversible; must be confirmed

Getting these wrong is not a routing inconvenience, it is a wrong edit to
someone's model. `dispatcher.py` matched "xoa" in all four sentences above and
routed every one of them to the ACTION specialist.

Also hosts the SEQUENCER table, which is what lets `graph/planner.py` see that
"thống kê tường rồi xuất pdf sheet A" is two goals and not one.

### On diacritics

Several Vietnamese function words become dangerous once folded to ASCII:

    chớ (don't)   → "cho"   collides with cho (give/for) — in half our sentences
    đừng (don't)  → "dung"  collides with dùng (use) and dựng (build a model)
    chưa (not yet)→ "chua"  collides with chữa (to fix) — a domain verb
    hãy (do!)     → "hay"   collides with hay (or)

Those markers are therefore matched on the RAW text WITH diacritics, exactly
like the existing `dispatcher._ACTION_COLOR_RAW` table handles "tô đỏ" (whose
folded form "to do" collides with English). The documented consequence is the
same one that table already accepts: diacritic-less input ("dung xoa tuong")
stays ambiguous on purpose and falls through to the model rather than being
guessed at.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re

from Intelligence.knowledge import vi_text

_NON_WORD_RE = re.compile(r'[^a-z0-9]+')
_WS_RE = re.compile(r'\s+')
# Compiled rather than passing flags= to re.sub: IronPython 2.7's re module is
# the CPython 2.7 one, where that keyword is a late addition — the rest of this
# codebase never relies on it.
_NON_WORD_U_RE = re.compile(r'[^\w\s]', re.UNICODE)


def _fold(text):
    """Folded lowercase with punctuation flattened to single spaces.

    Punctuation MUST go: every table below is matched space-delimited, and a
    trailing "?" was enough to hide "không" from the question-particle test.
    """
    folded = vi_text.fold_diacritics(text or u'').lower()
    return _NON_WORD_RE.sub(u' ', folded).strip()


def _raw_lower(text):
    """Lowercase with diacritics INTACT, punctuation flattened."""
    lowered = _NON_WORD_U_RE.sub(u' ', (text or u'').lower())
    return _WS_RE.sub(u' ', lowered).strip()


def _padded(text):
    return u' ' + text + u' '


def _has(padded, phrases):
    return any((u' ' + p + u' ') in padded for p in phrases)


# ─── Negation ─────────────────────────────────────────────────────────────────
# Vietnamese negation is PRE-verbal and lexically explicit, which makes it
# reliable to detect: the marker always sits immediately before the verb it
# kills. English adds contractions and a couple of post-auxiliary forms.
#
# Split into two families because they mean different things to a router:
#   PROHIBITIVE  "đừng xóa"      — an order NOT to act. Never route to action.
#   DECLARATIVE  "không có tường" — a statement/question containing a negation.
#                                    Still a legitimate read request.

# Diacritics required — see the module docstring.
_PROHIBITIVE_RAW = (
    u'đừng', u'đừng có', u'chớ', u'chớ có', u'khoan đã', u'thôi đừng',
    u'không được', u'không nên', u'không cần', u'khỏi cần', u'khỏi phải',
    u'không làm', u'đừng làm', u'miễn',
)

# Unambiguous once folded: no common Vietnamese word collides with these.
_PROHIBITIVE_FOLDED = (
    'khong duoc', 'khong nen', 'khong can', 'khoi can', 'khoi phai',
    'bo qua', 'khong lam', 'do not', 'dont', 'never', 'no need',
    'skip', 'without', 'avoid', 'instead of',
)

_DECLARATIVE_RAW = (u'chưa', u'chẳng', u'chưa có', u'không có', u'không còn')

_DECLARATIVE_FOLDED = (
    'khong co', 'khong con', 'not', 'none', 'nothing', 'neither',
    'cannot', 'cant', 'isnt', 'arent', 'wasnt', 'werent', 'doesnt', 'dont',
)

# "được không" / "có ... không" / "phải không" are QUESTION particles, not
# negations — Vietnamese marks yes/no questions with a trailing "không". Read
# as negation, "xóa được không?" became "do not delete".
# A BARE "a" is deliberately absent: it is a real particle ("vậy a?") but it is
# also every sheet prefix in the building industry, and "xuat pdf sheet A" was
# being read as a yes/no question because of it.
_QUESTION_TAIL = (
    'duoc khong', 'phai khong', 'dung khong', 'chua', 'khong',
    'nhi', 'ha', 'hen', 'the a', 'vay a', 'khong a',
)


def _ends_with_question_particle(padded_folded):
    """True when the text ENDS on a Vietnamese yes/no question particle."""
    tail = padded_folded.rstrip()
    return any(tail.endswith(u' ' + p) for p in _QUESTION_TAIL)


def detect_negation(text):
    """Returns {'negated', 'prohibitive', 'markers'}.

    `prohibitive` is the one that must block an action route: it means the user
    told the assistant NOT to do the thing whose verb is in the sentence.
    """
    padded_f = _padded(_fold(text))
    padded_r = _padded(_raw_lower(text))

    markers = [p for p in _PROHIBITIVE_RAW if (u' ' + p + u' ') in padded_r]
    markers += [p for p in _PROHIBITIVE_FOLDED
                if (u' ' + p + u' ') in padded_f]
    prohibitive = bool(markers)

    declarative = []
    # A trailing "không" is a question particle, not a negation — but only when
    # nothing else in the sentence negates anything.
    if not (_ends_with_question_particle(padded_f) and not prohibitive):
        declarative = [p for p in _DECLARATIVE_RAW
                       if (u' ' + p + u' ') in padded_r]
        declarative += [p for p in _DECLARATIVE_FOLDED
                        if (u' ' + p + u' ') in padded_f]

    return {'negated': bool(prohibitive or declarative),
            'prohibitive': prohibitive,
            'markers': list(markers) + list(declarative)}


# ─── Modality ─────────────────────────────────────────────────────────────────
# Bare "sao", "đâu", "ai" are deliberately absent: folded they collide with
# "sao chép" (copy), "dấu"/"đầu", and the literal word "AI" — which this
# application says constantly.

_QUESTION_WORDS_VI = (
    'bao nhieu', 'may cai', 'may cai', 'the nao', 'nhu the nao',
    'tai sao', 'vi sao', 'sao vay', 'sao the', 'o dau', 'cho nao',
    'khi nao', 'bao gio', 'cai gi', 'la gi', 'gi vay', 'gi the',
    'the nao la', 'co phai', 'nao', 'gi',
)
_QUESTION_WORDS_EN = (
    'how many', 'how much', 'how long', 'how do', 'how can', 'how to',
    'what', 'which', 'where', 'when', 'why', 'who', 'whose', 'whom',
    'is there', 'are there', 'can i', 'can you', 'could you', 'do i',
    'does it', 'should i',
)

# "hãy" needs its diacritics — folded it is "hay" (or / good).
_IMPERATIVE_RAW = (u'hãy', u'làm ơn', u'giúp', u'giùm', u'dùm',
                   u'cho tôi', u'cho mình', u'tôi muốn', u'mình muốn')
_IMPERATIVE_FOLDED = (
    'lam on', 'giup toi', 'giup minh', 'please', 'pls', 'let me',
    'i want', 'i need', 'give me', 'show me', 'go ahead',
)

_CONFIRMATION_WORDS = (
    'ok', 'oke', 'okay', 'yes', 'yeah', 'yep', 'sure', 'dong y', 'nhat tri',
    'dung roi', 'chinh xac', 'chuan', 'tiep', 'tiep tuc', 'lam di',
    'xac nhan', 'confirm', 'proceed', 'continue', 'go', 'do it',
    'khong', 'no', 'nope', 'huy', 'thoi', 'cancel', 'stop',
)


def detect_modality(text):
    """Returns one of 'question' | 'command' | 'confirmation' | 'statement'.

    Order matters: an explicit question mark or question frame beats an
    imperative verb, because "xóa được không?" is a question ABOUT deleting.
    """
    raw = (text or u'').strip()
    folded = _fold(raw)
    padded_f = _padded(folded)
    padded_r = _padded(_raw_lower(raw))

    if raw.endswith(u'?') or raw.endswith(u'？'):
        return 'question'
    # A bare "ok" / "không" / "đồng ý" answers a previous question. Checked
    # before the question frames so a lone "không" is a reply, not a question.
    if folded in _CONFIRMATION_WORDS:
        return 'confirmation'
    if _has(padded_f, _QUESTION_WORDS_EN) or _has(padded_f, _QUESTION_WORDS_VI):
        return 'question'
    if _ends_with_question_particle(padded_f):
        return 'question'
    # A prohibition IS an imperative — "đừng xóa text note" tells the
    # assistant to do something (namely, not that). Reading it as a bare
    # statement lost the only signal that a reply was even expected.
    if detect_negation(raw).get('prohibitive'):
        return 'command'
    if _has(padded_r, _IMPERATIVE_RAW) or _has(padded_f, _IMPERATIVE_FOLDED):
        return 'command'
    return 'statement'


# ─── Scope ────────────────────────────────────────────────────────────────────
# How much of the model the request covers. The ACTION specialist changes
# behaviour on this: "tô đỏ tường" across a whole project is a very different
# operation from the same words aimed at the current view.

_SCOPE_SELECTION = (
    'dang chon', 'da chon', 'duoc chon', 'cai nay', 'nhung cai nay',
    'element nay', 'doi tuong nay', 'phan tu nay', 'may cai nay',
    'selected', 'selection', 'current selection', 'this element',
    'these elements',
)
_SCOPE_VIEW = (
    'view nay', 'view hien tai', 'trong view', 'tren view', 'khung nhin nay',
    'man hinh nay', 'mat bang nay', 'ban ve nay', 'view dang mo',
    'this view', 'current view', 'active view', 'in view', 'in this view',
    'on screen',
)
_SCOPE_PROJECT = (
    'toan bo', 'tat ca', 'ca model', 'ca project', 'toan project',
    'toan model', 'ca file', 'toan du an', 'moi view', 'tat ca view',
    'toan bo model', 'toan bo project',
    'whole project', 'entire project', 'entire model', 'whole model',
    'all views', 'project wide', 'everywhere', 'globally',
)
_SCOPE_SHEET = (
    'tren sheet', 'trong sheet', 'cac sheet', 'sheet nay', 'bo ban ve',
    'on sheet', 'in sheet', 'sheet set', 'these sheets',
)


def detect_scope_detail(text, has_level=False):
    """(scope, matched_phrase). `matched_phrase` is '' for 'level'/'unspecified'.

    The phrase is returned so callers can BLANK it before extracting entities:
    "trong view này" is a scope, and letting its "view" also register as the
    Views category made the prompt hint claim the user had named a category
    they never mentioned.
    """
    padded = _padded(_fold(text))
    for scope, table in (('selection', _SCOPE_SELECTION),
                         ('project', _SCOPE_PROJECT),
                         ('view', _SCOPE_VIEW),
                         ('sheet', _SCOPE_SHEET)):
        # Longest phrase first so the blanked span covers the whole marker.
        for phrase in sorted(table, key=lambda p: -len(p)):
            if (u' ' + phrase + u' ') in padded:
                return scope, phrase
    if has_level:
        return 'level', u''
    return 'unspecified', u''


def detect_scope(text, has_level=False):
    """Coarse scope label. Defaults to 'unspecified' — never guesses 'project'.

    Guessing 'project' would license a model-wide edit the user never asked
    for; the specialist prompt asks instead when the scope is unspecified.
    """
    return detect_scope_detail(text, has_level=has_level)[0]


# ─── Risk ─────────────────────────────────────────────────────────────────────
# Operations whose result cannot be recovered with Ctrl+Z, or that publish work
# to other people. These get an explicit confirmation gate in the prompt hint.

_RISK_DESTRUCTIVE = (
    'xoa', 'xoa het', 'xoa toan bo', 'xoa sach', 'huy bo', 'go bo',
    'delete', 'remove', 'erase', 'wipe', 'drop',
)
_RISK_IRREVERSIBLE = (
    'purge', 'purge unused', 'don dep model', 'lam sach model', 'detach',
    'tach khoi central', 'compact', 'compact file',
)
_RISK_PUBLISH = (
    'sync', 'dong bo', 'sync with central', 'dong bo voi central',
    'publish', 'save to central', 'luu len central', 'relinquish',
    'tra quyen', 'giai phong workset',
)

_RISK_LEVELS = {'none': 0, 'destructive': 1, 'irreversible': 2, 'publish': 3}


def detect_risk(text):
    """Returns {'level', 'rank', 'markers'} — the highest risk family present.

    'publish' outranks the rest: a sync with central cannot be undone AND is
    visible to the whole team, so it is the one operation the assistant must
    never perform on an inferred instruction (`settings.is_sync_with_central_
    allowed` already gates it; this makes the intent visible to the prompt).
    """
    padded = _padded(_fold(text))
    found = {}
    for level, table in (('destructive', _RISK_DESTRUCTIVE),
                         ('irreversible', _RISK_IRREVERSIBLE),
                         ('publish', _RISK_PUBLISH)):
        hits = [p for p in table if (u' ' + p + u' ') in padded]
        if hits:
            found[level] = hits
    if not found:
        return {'level': 'none', 'rank': 0, 'markers': []}
    level = max(found.keys(), key=lambda k: _RISK_LEVELS[k])
    return {'level': level, 'rank': _RISK_LEVELS[level],
            'markers': found[level]}


# ─── Sequencers (multi-goal splitting) ───────────────────────────────────────
# Words that join two INDEPENDENT requests. `graph/planner.py` splits on these.
#
# "và" / "and" are deliberately absent from the strong table: they join noun
# phrases far more often than clauses in this domain ("tường và cột" is one
# goal, not two). They live in _WEAK_SEQUENCERS, which the planner only honours
# when BOTH halves independently look like commands.

_STRONG_SEQUENCERS = (
    'roi', 'sau do', 'tiep theo', 'tiep den', 'xong thi', 'sau khi xong',
    'ke tiep', 'cuoi cung', 'dong thoi',
    'then', 'after that', 'afterwards', 'and then', 'finally',
    'once done', 'when done', 'next step',
)
_WEAK_SEQUENCERS = ('va', 'and', 'cung nhu', 'as well as', 'plus')

# Punctuation that reliably separates clauses. A comma does NOT: Vietnamese
# uses it inside lists constantly ("tường, cột, dầm").
_CLAUSE_PUNCT_RE = re.compile(r'\s*(?:;|\.\s+|\n+)\s*')
# "1. ... 2. ..." or "- ... - ..." numbered/bulleted plans
_LIST_ITEM_RE = re.compile(r'(?:^|\n)\s*(?:\d{1,2}[.)]|[-*•])\s+')


def split_sequenced(text):
    """Split on STRONG sequencers and clause punctuation only.

    Returns a list of raw substrings, original order, always at least one
    element for non-empty input. Weak sequencers are reported separately by
    `weak_split()` so the planner can apply its own "both halves are commands"
    test.
    """
    raw = (text or u'').strip()
    if not raw:
        return []

    # Bulleted / numbered plans are explicit multi-step input.
    if len(_LIST_ITEM_RE.findall(raw)) >= 2:
        parts = [p.strip() for p in _LIST_ITEM_RE.split(raw) if p.strip()]
        if len(parts) >= 2:
            return parts

    chunks = [c for c in _CLAUSE_PUNCT_RE.split(raw) if c and c.strip()]

    out = []
    for chunk in chunks:
        out.extend(_split_on_words(chunk, _STRONG_SEQUENCERS))
    return [p.strip() for p in out if p.strip()]


def weak_split(text):
    """Split on WEAK sequencers ("và", "and"). Planner-gated — see above."""
    return [p.strip() for p in _split_on_words(text, _WEAK_SEQUENCERS)
            if p.strip()]


def _split_on_words(text, words):
    """Split `text` at whole-word occurrences of any of `words`.

    Matching is done on the folded text but the SLICES come from the original
    string, so diacritics and casing survive into each sub-goal. Vietnamese
    diacritic folding is 1:1 on characters, so folded offsets index the
    original string directly.
    """
    raw = text or u''
    if not raw.strip():
        return []
    folded = vi_text.fold_diacritics(raw).lower()

    cuts = []
    for word in words:
        pattern = re.compile(r'(?:^|(?<=[^a-z0-9]))'
                             + re.escape(word)
                             + r'(?=[^a-z0-9]|$)')
        for m in pattern.finditer(folded):
            cuts.append((m.start(), m.end()))
    if not cuts:
        return [raw]

    cuts.sort()
    parts = []
    pos = 0
    for start, end in cuts:
        if start < pos:            # overlapping markers — keep the first
            continue
        parts.append(raw[pos:start])
        pos = end
    parts.append(raw[pos:])
    return parts


def analyze_sentence(text, has_level=False):
    """Everything in this module in one pass — the shape `analyzer` consumes."""
    scope, scope_phrase = detect_scope_detail(text, has_level=has_level)
    return {
        'negation':     detect_negation(text),
        'modality':     detect_modality(text),
        'scope':        scope,
        'scope_phrase': scope_phrase,
        'risk':         detect_risk(text),
        'segments':     split_sequenced(text),
    }
