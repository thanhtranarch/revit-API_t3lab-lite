# -*- coding: utf-8 -*-
"""
Diversity re-ranking for retrieval results.

BM25 (and the semantic fusion on top of it) ranks chunks independently, so the
strongest few are routinely near-duplicates: consecutive chunks of one page,
or the same clause repeated in a document's summary and its body. The tool
agent only receives two or three excerpts, so spending all of them on one
passage means the model sees one fact three times and misses the others.

This applies the cheap half of MMR: keep the ranking, but cap how much of the
budget any single document may take, and skip a candidate whose text is
substantially already covered by one already selected. Everything is lexical
and pure Python — no embeddings, no model, nothing that could add latency to
the retrieval path.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Rerank"

from Intelligence.knowledge import vi_text

# Chunks from one document allowed into the final set. Two lets a document
# support an answer from two places; three would let one file own a top-3.
DEFAULT_PER_DOC = 2

# Token overlap above which a candidate adds nothing the selected set lacks.
# Measured on the shorter side (vi_text.word_match_score), so a short chunk
# fully contained in a longer one counts as redundant.
DEFAULT_REDUNDANCY = 0.82


def diversify(candidates, top_k, doc_of=None, text_of=None,
              per_doc=DEFAULT_PER_DOC, redundancy=DEFAULT_REDUNDANCY):
    """Pick `top_k` of `candidates` (best-first) with per-document capping.

    candidates : ranked list, best first. Items may be anything; `doc_of` and
                 `text_of` extract the document id and the text.
    doc_of     : item -> document id, or None to skip document capping.
    text_of    : item -> text, or None to skip the redundancy check.

    The caps are a preference, not a hard limit: if applying them cannot fill
    `top_k`, the best rejected candidates are added back in their original
    order. Returning fewer results than the caller asked for would be a worse
    outcome than returning a slightly repetitive one.
    """
    items = list(candidates or [])
    if top_k is None:
        return items
    if top_k <= 0:
        return []
    if len(items) <= 1:
        return items[:top_k]

    chosen = []
    chosen_tokens = []
    deferred = []
    per_doc_count = {}

    for item in items:
        if len(chosen) >= top_k:
            break

        doc = None
        if doc_of is not None:
            try:
                doc = doc_of(item)
            except Exception:
                doc = None
        if doc is not None and per_doc_count.get(doc, 0) >= per_doc:
            deferred.append(item)
            continue

        tokens = None
        if text_of is not None:
            try:
                tokens = set(vi_text.tokenize(text_of(item) or ''))
            except Exception:
                tokens = None
        if tokens:
            if any(vi_text.word_match_score(tokens, prev) >= redundancy
                   for prev in chosen_tokens):
                deferred.append(item)
                continue

        chosen.append(item)
        chosen_tokens.append(tokens or set())
        if doc is not None:
            per_doc_count[doc] = per_doc_count.get(doc, 0) + 1

    # Backfill: a strict cap that starves the answer helps nobody.
    if len(chosen) < top_k:
        for item in deferred:
            if len(chosen) >= top_k:
                break
            chosen.append(item)

    return chosen[:top_k]
