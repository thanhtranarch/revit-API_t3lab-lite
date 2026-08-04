# -*- coding: utf-8 -*-
"""
VERIFIER primitive — checks, scores and validates what a node produced.

Without it a graph is optimistic: a node that "succeeded" by replying "Xin lỗi,
tôi không thể truy cập model" is recorded as DONE, its empty answer flows into
the reducer, and the user gets a confident summary of nothing. That failure
mode is invisible without a verifier, because no exception was raised.

Everything here is DETERMINISTIC on purpose. A second LLM call to grade the
first one doubles the cost and latency of every node and introduces a grader
that can be wrong in the same direction as the thing it grades. These checks
catch the failures that actually happen in this codebase:

    empty        the node produced nothing at all
    refusal      the model declined ("I can't", "tôi không thể")
    unanswered   a counting question came back with no number in it
    tool_error   the output is a raw tool error that leaked into prose
    truncated    the reply stops mid-sentence (budget exhausted)

A verdict is advisory: the orchestrator decides what to do with a low score
(usually one retry on the fallback route). Nothing here modifies the model.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re

from Intelligence.knowledge import vi_text

# Score below which the orchestrator should re-route and retry.
RETRY_THRESHOLD = 0.5

# A node that produced less than this is not an answer, it is an artefact.
_MIN_USEFUL_CHARS = 12

_REFUSAL_PHRASES = (
    # Vietnamese, folded
    'toi khong the', 'minh khong the', 'khong the thuc hien',
    'khong thuc hien duoc', 'khong tim thay cong cu', 'khong co quyen',
    'toi khong biet', 'khong ho tro', 'chua ho tro', 'khong truy cap duoc',
    'xin loi toi khong',
    # English
    'i cannot', 'i can not', 'i am unable', 'im unable', 'i do not have',
    'i dont have access', 'not supported', 'no tool available',
    'unable to access', 'sorry i cant', 'sorry i cannot',
)

_TOOL_ERROR_MARKERS = (
    'traceback (most recent call last)', 'exception:', 'error:',
    '"error"', "'error'", 'tool_error', 'failed to execute tool',
    'no such tool', 'invalid tool name',
)

_NUMBER_RE = re.compile(r'\d')

# A counting/listing question expects a number in the answer.
_COUNT_QUESTION = (
    'bao nhieu', 'may cai', 'so luong', 'thong ke', 'dem', 'liet ke',
    'danh sach', 'how many', 'how much', 'count', 'list', 'total',
    'statistics', 'breakdown',
)

# Sentence-ending punctuation. An answer that stops without one, mid-word, was
# almost certainly cut off by the token budget.
_ENDINGS = u'.!?:;)]}"\'`…。！？'


def _fold(text):
    folded = vi_text.fold_diacritics(text or u'').lower()
    return re.sub(r'[^a-z0-9]+', u' ', folded).strip()


class Verdict(object):
    """Result of verifying one node."""

    __slots__ = ('ok', 'score', 'reasons', 'node_id')

    def __init__(self, node_id, ok=True, score=1.0, reasons=None):
        self.node_id = node_id
        self.ok      = bool(ok)
        self.score   = float(score)
        self.reasons = list(reasons or [])

    @property
    def should_retry(self):
        return self.score < RETRY_THRESHOLD

    def to_dict(self):
        return {'node': self.node_id, 'ok': self.ok,
                'score': round(self.score, 2), 'reasons': self.reasons}

    def __repr__(self):
        return '<Verdict {} ok={} score={:.2f} {}>'.format(
            self.node_id, self.ok, self.score, ','.join(self.reasons))


class Verifier(object):
    """Deterministic quality checks over a NodeResult."""

    def check(self, node, result, utterance=None):
        """Score one node's output. Returns a Verdict.

        `utterance` is the analysed sub-goal; when present it tells the
        verifier what KIND of answer was expected (a count question wants a
        number) instead of guessing from the output alone.
        """
        reasons = []
        score = 1.0

        if result is None:
            return Verdict(node.id, ok=False, score=0.0, reasons=['no_result'])

        if result.error:
            reasons.append('error')
            score -= 0.6

        text = (result.output or u'').strip()
        if not text:
            return Verdict(node.id, ok=False, score=0.0, reasons=['empty'])
        if len(text) < _MIN_USEFUL_CHARS:
            reasons.append('too_short')
            score -= 0.4

        folded = _fold(text)
        padded = u' ' + folded + u' '

        if any((u' ' + p + u' ') in padded for p in _REFUSAL_PHRASES):
            reasons.append('refusal')
            score -= 0.5

        low = text.lower()
        if any(m in low for m in _TOOL_ERROR_MARKERS):
            reasons.append('tool_error')
            score -= 0.4

        # A counting question whose answer contains no digit did not answer it.
        goal_folded = _fold(node.goal)
        goal_padded = u' ' + goal_folded + u' '
        wants_number = any((u' ' + p + u' ') in goal_padded
                           for p in _COUNT_QUESTION)
        if utterance is not None:
            wants_number = wants_number or bool(
                utterance.quantity.get('all')) or utterance.is_question
        if wants_number and not _NUMBER_RE.search(text):
            reasons.append('no_number')
            score -= 0.3

        # Truncation: ends mid-word with no terminal punctuation. Only flagged
        # for answers long enough that the budget could plausibly be the cause.
        if len(text) > 200 and text[-1] not in _ENDINGS:
            reasons.append('truncated')
            score -= 0.2

        score = max(0.0, min(1.0, score))
        return Verdict(node.id, ok=(score >= RETRY_THRESHOLD), score=score,
                       reasons=reasons)

    def check_all(self, graph, state, utterances=None):
        """Verify every node that produced a result. Returns {node_id: Verdict}."""
        out = {}
        results = state.results()
        for node in graph.nodes():
            res = results.get(node.id)
            if res is None:
                continue
            utt = (utterances or {}).get(node.id)
            out[node.id] = self.check(node, res, utterance=utt)
        return out
