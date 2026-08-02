# -*- coding: utf-8 -*-
"""
Conversation-aware retrieval queries.

Retrieval used to run on the raw user message alone, which works for a first
question and fails for the follow-up that naturally follows it:

    user: "chiều cao lan can theo tiêu chuẩn là bao nhiêu?"
    bot:  "1.1 m cho ban công… [1]"
    user: "còn cầu thang thì sao?"

The second message carries "cầu thang" and nothing else — none of the words
that made the first search find the right document. BM25 scores it against the
whole corpus and returns noise, or nothing.

This module widens such a message with the subject of the turn it continues,
and ONLY such a message: a self-contained question is already the best query
for itself, and padding it with history would drag in stale terms and hurt
precision. The test for "is this a follow-up" reuses the anaphora and
continuation logic that already exists for routing rather than inventing a
second, disagreeing one.

Pure Python — no Revit, no network. CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Query Builder"

import re

from Intelligence.knowledge import vi_text

# A message longer than this stands on its own — whatever it refers back to,
# it also carries enough of its own subject to search with. Matches the spirit
# of routing.MAX_ANSWER_WORDS (8) but is deliberately looser: a follow-up
# QUESTION ("còn tiêu chuẩn cho cầu thang ngoài trời thì sao?") is wordier than
# a follow-up ANSWER ("ok").
_MAX_FOLLOWUP_WORDS = 12

# Terms lifted from earlier turns, at most. Enough to re-anchor the search
# without letting the previous question dominate the new one.
_MAX_CARRIED_TERMS = 6

# How far back to look. Two turns is one exchange: the question that set the
# subject and the answer that elaborated it.
_MAX_LOOKBACK = 4

# Phrases that mark a message as continuing the previous one even when no
# pronoun appears — "còn X thì sao" is the canonical Vietnamese follow-up.
_FOLLOWUP_MARKERS = (
    'con ', 'the con ', 'vay con ', 'thi sao', 'the nao', 'thi the nao',
    'what about', 'how about', 'and for', 'and the', 'same for',
    'tuong tu', 'con lai', 'thi sao a',
)


def _words(text):
    return [w for w in re.split(r'\s+', (text or '').strip()) if w]


def is_followup(text, history=None):
    """True when `text` needs an earlier turn to make sense as a search query.

    Two independent signals, both cheap:
      * anaphora — reuses nlu_engine's pronoun test, which already knows the
        difference between "cái này là gì" (stands in for a noun) and "cái
        này hay đấy" / "this project needs cleanup" (a determiner in front of
        one). Falling back to a local check keeps this importable if the NLU
        module is unavailable.
      * an explicit follow-up marker ("còn … thì sao", "what about …").

    Long messages are never follow-ups for retrieval purposes: they carry
    their own subject.
    """
    raw = (text or '').strip()
    if not raw:
        return False
    if len(_words(raw)) > _MAX_FOLLOWUP_WORDS:
        return False

    folded = vi_text.fold_diacritics(raw).lower()
    folded = re.sub(r'\s+', ' ', folded).strip()
    padded = ' ' + folded + ' '
    for marker in _FOLLOWUP_MARKERS:
        if marker in padded:
            return True

    try:
        from Intelligence.nlu_engine import _is_pronoun_query
        return bool(_is_pronoun_query(folded))
    except Exception:
        # Minimal stand-in: a bare pronoun with no noun after it.
        return folded in ('no', 'cai nay', 'cai do', 'it', 'that', 'this')


# Question scaffolding that survives vi_text's stopword list but carries no
# topic. Lifting these into the query is pure noise for BM25.
_INTERROGATIVES = frozenset([
    'bao', 'nhieu', 'sao', 'gi', 'nao', 'dau', 'may', 'the',
    'what', 'how', 'why', 'when', 'where', 'many', 'much',
    'theo', 'phai', 'khac', 'them', 'giup', 'minh', 'toi',
])


def salient_terms(text, limit=_MAX_CARRIED_TERMS, exclude=None):
    """Topic words of `text`, in order of appearance, de-duplicated.

    Order of appearance beats any cleverness available for free here: the
    subject of a sentence tends to be stated early, and Vietnamese words are
    multi-syllable, so ranking by token length just picks whichever syllable
    happens to be longest ("thieu" out of "tối thiểu") rather than the topic.
    vi_text.tokenize has already dropped stopwords and short tokens; this
    additionally drops question scaffolding, which is frequent in exactly the
    turns being mined and means nothing to BM25.
    """
    seen = set(exclude or [])
    out = []
    for tok in vi_text.tokenize(text or ''):
        if '_' in tok:                  # skip bigrams; the index forms its own
            continue
        if tok in seen or tok in _INTERROGATIVES:
            continue
        seen.add(tok)
        out.append(tok)
        if len(out) >= limit:
            break
    return out


def build_retrieval_query(text, history=None, max_terms=_MAX_CARRIED_TERMS):
    """The query to actually search with.

    Returns `text` unchanged unless it reads as a follow-up, in which case the
    subject of the recent turns is appended. The user's own words always come
    first and are never removed, so a widened query can only ever add recall —
    BM25 scores the original terms exactly as it did before.
    """
    raw = (text or '').strip()
    if not raw or not history:
        return raw
    if not is_followup(raw, history):
        return raw

    recent = [t for t in list(history)[-_MAX_LOOKBACK:]
              if isinstance(t, dict) and t.get('role') in ('user', 'assistant')]

    # The user's own previous question framed the topic; the assistant's reply
    # only elaborates on it (and may have drifted into other documents). Mine
    # the user turns first so the carried terms are the ones the user chose.
    ordered = ([t for t in reversed(recent) if t.get('role') == 'user']
               + [t for t in reversed(recent) if t.get('role') == 'assistant'])

    own = set(vi_text.tokenize(raw))
    carried = []
    for turn in ordered:
        content = turn.get('content') or ''
        if isinstance(content, list):
            content = ' '.join(
                b.get('text', '') for b in content
                if isinstance(b, dict) and b.get('type') == 'text')
        for term in salient_terms(content, limit=max_terms,
                                  exclude=own | set(carried)):
            carried.append(term)
            if len(carried) >= max_terms:
                break
        if len(carried) >= max_terms:
            break

    if not carried:
        return raw
    return u'{} {}'.format(raw, u' '.join(carried))
