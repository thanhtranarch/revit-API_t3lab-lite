# -*- coding: utf-8 -*-
"""
Vietnamese / English detection that survives real T3Lab chat input.

The assistant's previous test was one line — "does the text contain a
Vietnamese diacritic?". It got three very common cases wrong:

    "xuat pdf sheet A"          typed without diacritics → read as English,
                                so the reply came back in English.
    "export cái sheet này"      code-switched → the English half is real
                                vocabulary, not noise; the sentence is
                                still Vietnamese and must be answered so.
    "Wall Type Manager"         a bare tool name → neither language really;
                                it must not flip the session language.

Detection here is evidence-based instead: diacritics, function words on both
sides, Vietnamese syllable shape, and a technical-term list that is scored as
NEUTRAL rather than English (BIM vocabulary is English in every Vietnamese
office, so "export", "sheet", "level" say nothing about the writer).

Returns a verdict plus the evidence, so callers that need a stricter bar
(pinning the reply language for a whole session) can look at `confidence`
instead of trusting a bare label.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re

from Intelligence.knowledge import vi_text

# ─── Evidence tables ──────────────────────────────────────────────────────────
# All matched against DIACRITIC-FOLDED lowercase text, because the hard case is
# precisely the input that has no diacritics left to look at.

# Vietnamese function words that do NOT collide with common English once
# folded. "the" (thế), "la" (là), "co" (có), "da" (đã), "con" (còn), "an",
# "cam", "ban", "tin", "hay" are all deliberately absent: each is either an
# English word or an English fragment, and counting them made English
# sentences score as Vietnamese.
_VI_MARKERS = frozenset([
    'cua', 'khong', 'duoc', 'nhung', 'voi', 'trong', 'cac', 'nay', 'roi',
    'chua', 'nao', 'thi', 'khi', 'dang', 'hoac', 'tren', 'duoi', 'nhu',
    'tai', 'vao', 'lai', 'toi', 'minh', 'nhe', 'rat', 'cung', 'nen',
    'boi', 'neu', 'giup', 'lam', 'muon', 'can', 'phai', 'hon', 'nua',
    'moi', 'dum', 'gium', 'day', 'kia', 'ay', 'nhi', 'a', 'oi',
    'bao', 'nhieu', 'gio', 'dau', 'sao', 'gi', 'ma', 'mot', 'hai',
    'nhieu', 'it', 'tat', 'toan', 'bo', 'tung', 'moi',
    'xem', 'cho', 'dat', 'sua', 'xoa', 'them', 'bot', 'doi', 'chuyen',
    'kiem', 'tra', 'thong', 'ke', 'liet', 'dem', 'tim', 'loc',
    'tuong', 'san', 'cot', 'dam', 'cua', 'mai', 'phong', 'tang',
    'ban', 've', 'khung', 'nhin', 'chu', 'thich', 'kich', 'thuoc',
])

# Multi-word Vietnamese frames — far stronger evidence than any single token.
_VI_PHRASES = (
    'bao nhieu', 'tat ca', 'toan bo', 'nhu the nao', 'the nao', 'tai sao',
    'o dau', 'bao gio', 'giup toi', 'giup minh', 'cho toi', 'cho minh',
    'lam on', 'co the', 'khong the', 'da co', 'chua co', 'thi sao',
    'duoc khong', 'dc ko', 'hay khong', 'con lai', 'hien tai',
    'ban ve', 'cong trinh', 'du an', 'mo hinh',
)

_EN_MARKERS = frozenset([
    'the', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
    'what', 'which', 'how', 'why', 'when', 'where', 'who', 'whose',
    'please', 'could', 'should', 'would', 'shall', 'might', 'must',
    'and', 'or', 'but', 'of', 'to', 'in', 'on', 'at', 'for', 'with',
    'from', 'into', 'onto', 'about', 'this', 'that', 'these', 'those',
    'my', 'your', 'our', 'their', 'its', 'me', 'you', 'they', 'them',
    'do', 'does', 'did', 'done', 'have', 'has', 'had', 'having',
    'there', 'here', 'then', 'than', 'also', 'only', 'just', 'every',
    'each', 'any', 'some', 'many', 'much', 'more', 'most', 'less',
    'give', 'want', 'need', 'make', 'made', 'take', 'let', 'know',
])

_EN_PHRASES = (
    'how many', 'how much', 'how do i', 'can you', 'could you',
    'i want', 'i need', 'give me', 'show me', 'tell me', 'what is',
    'what are', 'is there', 'are there', 'in the', 'of the', 'on the',
)

# Domain vocabulary every Vietnamese BIM user types in English. Counting these
# as English evidence is exactly the bug that made "xuat pdf sheet" read as an
# English sentence — they are scored as NEUTRAL and excluded from both sides.
_NEUTRAL_TERMS = frozenset([
    'revit', 'bim', 'cad', 'dwg', 'dgn', 'ifc', 'pdf', 'png', 'jpg', 'svg',
    'nwc', 'nwd', 'rvt', 'rfa', 'rte', 'xlsx', 'csv', 'json', 'txt', 'md',
    'sheet', 'sheets', 'view', 'views', 'level', 'levels', 'grid', 'grids',
    'wall', 'walls', 'floor', 'floors', 'roof', 'column', 'columns',
    'beam', 'beams', 'door', 'doors', 'window', 'windows', 'room', 'rooms',
    'family', 'families', 'type', 'types', 'parameter', 'parameters',
    'param', 'params', 'category', 'categories', 'element', 'elements',
    'workset', 'worksets', 'warning', 'warnings', 'schedule', 'schedules',
    'tag', 'tags', 'template', 'templates', 'filter', 'filters',
    'export', 'import', 'model', 'models', 'project', 'projects',
    'link', 'links', 'batchout', 'parasync', 't3lab', 'plugin', 'addin',
    'ok', 'oke', 'okay', 'yes', 'no', 'api', 'ai', 'llm', 'mcp',
    'dimension', 'dimensions', 'annotation', 'text', 'note', 'notes',
    'pin', 'unpin', 'purge', 'audit', 'sync', 'render', 'detail',
])

# Vietnamese-only letters. Their presence is decisive on its own.
_VI_LETTERS_RE = re.compile(
    '[à-ãè-êìíò-õùúý'
    'ăđĩũơưẠ-ỹ'
    'À-ÃÈ-ÊÌÍÒ-ÕÙÚÝ'
    'ĂĐĨŨƠƯ]')

_WORD_RE = re.compile(r'[a-z0-9]+')

# A Vietnamese syllable is short and follows a tight consonant/vowel shape.
# Used as weak backup evidence for diacritic-less input whose words are not in
# the marker table ("chuyen sang model khac").
_VI_SYLLABLE_RE = re.compile(
    r'^(?:ngh|ng|nh|ph|th|tr|ch|kh|gh|gi|qu|[bcdgklmnprstvxh])?'
    r'(?:uye|uya|uyu|oai|oay|uoi|uou|ieu|yeu|uon|uan|uat|uac|'
    r'ai|ao|au|ay|eo|eu|ia|ie|iu|oa|oe|oi|oo|ua|ue|ui|uo|uu|uy|'
    r'a|e|i|o|u|y)'
    r'(?:ngh|ng|nh|ch|[cmnptk])?$')

# Below this many scoreable tokens the verdict is a guess, not a reading.
_MIN_TOKENS_FOR_CONFIDENCE = 2


def _phrase_hits(padded, phrases):
    return sum(1 for p in phrases if (u' ' + p + u' ') in padded)


def _tokenize_pair(lowered, folded):
    """Tokenise both views at once, guaranteeing equal-length token lists.

    Splitting them independently does not work: `_WORD_RE` is ASCII, so it
    tears "tô" into ["t"] and drops "đỏ" entirely. Word boundaries are taken
    from the FOLDED view (where every Vietnamese letter is ASCII) and applied
    to both, which keeps token i of one list the same word as token i of the
    other — the alignment the accent test depends on.
    """
    if len(folded) != len(lowered):
        # Vietnamese folding is 1:1; a locale-sensitive lowercase elsewhere
        # could desynchronise. Fall back to folded-only, losing accent
        # evidence rather than misaligning it.
        toks = _WORD_RE.findall(folded)
        return toks, toks

    raw_words, words = [], []
    cur_r, cur_f = [], []
    for rc, fc in zip(lowered, folded):
        if _WORD_RE.match(fc):
            cur_r.append(rc)
            cur_f.append(fc)
        elif cur_f:
            raw_words.append(u''.join(cur_r))
            words.append(u''.join(cur_f))
            cur_r, cur_f = [], []
    if cur_f:
        raw_words.append(u''.join(cur_r))
        words.append(u''.join(cur_f))
    return raw_words, words


def detect(text):
    """Weigh the evidence for Vietnamese vs English.

    Returns a dict:
        lang        'vi' | 'en' | 'unknown'
        confidence  0.0-1.0 — how firmly the label is supported
        vi_score    weighted Vietnamese evidence
        en_score    weighted English evidence
        mixed       True when BOTH sides carry real evidence (code-switching)
        diacritics  True when Vietnamese letters are present
        tokens      number of scoreable (non-neutral) tokens

    'unknown' is returned for input with no evidence either way — a bare tool
    name, a number, an emoji. Callers must keep their previous language
    instead of flipping on it.
    """
    raw = text or u''
    diacritics = bool(_VI_LETTERS_RE.search(raw))

    lowered = raw.lower()
    folded = vi_text.fold_diacritics(lowered)
    raw_words, words = _tokenize_pair(lowered, folded)
    padded = u' ' + u' '.join(words) + u' '

    # A word the user typed WITH diacritics is Vietnamese, whatever it folds
    # to. Without this, "tô đỏ tường" scored an English hit — "đỏ" folds to
    # "do", which is in _EN_MARKERS — and the sentence came back as
    # code-switched.
    accented = set()
    for rw, fw in zip(raw_words, words):
        if rw != fw:
            accented.add(fw)

    scoreable = [w for w in words if w not in _NEUTRAL_TERMS]

    vi_words = sum(1 for w in scoreable
                   if w in _VI_MARKERS or w in accented)
    en_words = sum(1 for w in scoreable
                   if w in _EN_MARKERS and w not in accented)
    vi_phr = _phrase_hits(padded, _VI_PHRASES)
    en_phr = _phrase_hits(padded, _EN_PHRASES)

    # Syllable shape only counts for words that are evidence for NEITHER side,
    # otherwise "in", "on", "to" (valid Vietnamese syllables AND English
    # function words) would be counted twice, once on each side.
    unknown_words = [w for w in scoreable
                     if w not in _VI_MARKERS and w not in _EN_MARKERS
                     and w not in accented and not w.isdigit()]
    vi_shape = sum(1 for w in unknown_words
                   if len(w) <= 7 and _VI_SYLLABLE_RE.match(w))

    # Weights: a phrase is worth three bare words; syllable shape is the
    # weakest signal and can never outvote real vocabulary on its own.
    vi_score = vi_words * 1.0 + vi_phr * 3.0 + vi_shape * 0.25
    en_score = en_words * 1.0 + en_phr * 3.0

    # Diacritics settle it. Nothing else can produce them by accident, and a
    # code-switched sentence with even one of them is a Vietnamese sentence.
    if diacritics:
        vi_score += 6.0

    total = vi_score + en_score
    if total <= 0.0:
        return {'lang': 'unknown', 'confidence': 0.0,
                'vi_score': 0.0, 'en_score': 0.0, 'mixed': False,
                'diacritics': diacritics, 'tokens': len(scoreable)}

    lang = 'vi' if vi_score >= en_score else 'en'
    margin = abs(vi_score - en_score) / total
    # Code-switching means REAL vocabulary from both languages, so the weaker
    # side needs two words, not one. At a threshold of one, a single folded
    # collision was enough to call an ordinary Vietnamese sentence mixed.
    mixed = bool(vi_words + vi_phr > 0 and en_words + en_phr > 0
                 and min(vi_score, en_score) >= 2.0)

    confidence = margin
    if diacritics:
        confidence = max(confidence, 0.9)
    elif len(scoreable) < _MIN_TOKENS_FOR_CONFIDENCE:
        confidence *= 0.5

    return {'lang': lang, 'confidence': round(min(confidence, 1.0), 3),
            'vi_score': round(vi_score, 2), 'en_score': round(en_score, 2),
            'mixed': mixed, 'diacritics': diacritics,
            'tokens': len(scoreable)}


def is_vietnamese(text, default=True):
    """Convenience boolean.

    `default` is returned for input with no evidence either way, so a bare
    "BatchOut" keeps whatever language the session was already using.
    """
    verdict = detect(text)
    if verdict['lang'] == 'unknown':
        return default
    return verdict['lang'] == 'vi'
