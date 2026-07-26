# -*- coding: utf-8 -*-
"""
Vietnamese-aware text utilities for the knowledge stack.

Pure Python — runs under IronPython 2.7 (in Revit) and CPython 3
(dev/test_knowledge.py). No external dependencies.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re

# ─── Diacritic folding ────────────────────────────────────────────────────────
# Vietnamese letters → ASCII base letter (tường → tuong, đ → d).

_FOLD_GROUPS = {
    'a': 'áàảãạăắằẳẵặâấầẩẫậ',
    'e': 'éèẻẽẹêếềểễệ',
    'i': 'íìỉĩị',
    'o': 'óòỏõọôốồổỗộơớờởỡợ',
    'u': 'úùủũụưứừửữự',
    'y': 'ýỳỷỹỵ',
    'd': 'đ',
}

_FOLD_MAP = {}
for _base, _chars in _FOLD_GROUPS.items():
    for _ch in _chars:
        _FOLD_MAP[_ch] = _base
        _FOLD_MAP[_ch.upper()] = _base.upper()


def fold_diacritics(text):
    """Strip Vietnamese diacritics: u'tường' -> u'tuong', u'đ' -> u'd'."""
    if not text:
        return u''
    return u''.join(_FOLD_MAP.get(ch, ch) for ch in text)


# ─── Tokenizer ────────────────────────────────────────────────────────────────

_SPLIT_RE = re.compile(r'[^a-z0-9]+')

# Folded lowercase stopwords. Kept small and high-frequency-only so that
# domain words (wall, tuong, dam...) always survive.
_STOPWORDS_VI = frozenset([
    'va', 'la', 'cua', 'cho', 'trong', 'mot', 'cac', 'nhung', 'duoc',
    'voi', 'den', 'tu', 'khong', 'thi', 'ma', 'nay', 'khi', 'se',
    'da', 'dang', 'hay', 'hoac', 'tren', 'duoi', 'nhu', 'tai',
    'vao', 'con', 'lai', 'toi', 'ban', 'minh', 'anh', 'em', 'nhe',
    'rat', 'cung', 'nen', 'vi', 'boi', 'neu', 'de',
])
_STOPWORDS_EN = frozenset([
    'the', 'an', 'and', 'or', 'of', 'to', 'in', 'on', 'for', 'is',
    'are', 'was', 'were', 'be', 'been', 'it', 'this', 'that', 'with',
    'as', 'at', 'by', 'from', 'will', 'would', 'can', 'could', 'do',
    'does', 'did', 'have', 'has', 'had', 'which', 'who', 'all',
    'any', 'each', 'other', 'some', 'such', 'nor', 'not', 'than',
    'too', 'very', 'just', 'about', 'into', 'over', 'under',
])
_STOPWORDS = _STOPWORDS_VI | _STOPWORDS_EN


def tokenize(text, bigrams=False):
    """lower + fold diacritics + split on non-alphanumerics + drop short/stop.

    Returns a list of folded lowercase tokens (len >= 2, no stopwords).

    bigrams=True additionally appends adjacent word bigrams ("lan_can",
    "chieu_cao") after the unigrams. Vietnamese words are multi-syllable and
    written space-separated ("tường chịu lực", "dầm sàn"); indexing bigrams
    lets BM25 reward true phrase matches, not just scattered syllables — a
    large precision gain that needs no embeddings, so it helps EVERY model.
    """
    if not text:
        return []
    folded = fold_diacritics(text).lower()
    raw = [t for t in _SPLIT_RE.split(folded) if len(t) >= 2]
    unigrams = [t for t in raw if t not in _STOPWORDS]
    if not bigrams:
        return unigrams
    # Bigrams form from the RAW sequence (stopwords included) so a phrase
    # stays connected even through a syllable that happens to be a stopword —
    # e.g. Vietnamese "lan can" (railing) keeps the bigram "lan_can" although
    # "can" collides with the English stopword and is dropped as a unigram.
    grams = [raw[i] + u'_' + raw[i + 1] for i in range(len(raw) - 1)]
    return unigrams + grams


def word_match_score(a_tokens, b_tokens):
    """Overlap ratio between two token lists, on the shorter side: 0.0–1.0.

    Used for fuzzy sheet-name matching and dispatcher heuristics.
    """
    set_a = set(a_tokens or [])
    set_b = set(b_tokens or [])
    if not set_a or not set_b:
        return 0.0
    inter = len(set_a & set_b)
    return float(inter) / float(min(len(set_a), len(set_b)))
