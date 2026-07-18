# -*- coding: utf-8 -*-
"""
BM25 inverted index — the always-available lexical retrieval channel.

Pure Python (IronPython 2.7 + CPython 3), JSON-serializable, no deps.
Okapi BM25 with k1=1.5, b=0.75 and the "+1" idf variant so common terms
never go negative.

Chunk keys follow "<doc_id>#p<page>#c<seq>" so per-document removal can
work by key prefix.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import math

from Intelligence.knowledge import vi_text

K1 = 1.5
B = 0.75


def make_chunk_key(doc_id, page, seq):
    return u'{}#p{}#c{}'.format(doc_id, page, seq)


def doc_id_of(chunk_key):
    return chunk_key.split(u'#', 1)[0]


class BM25Index(object):
    """Inverted index over chunks. Not internally locked — callers
    (KnowledgeStore) serialize access."""

    def __init__(self):
        self.postings = {}   # term -> list of [chunk_key, tf]
        self.doclen = {}     # chunk_key -> token count
        self._avgdl = 0.0
        self._dirty_avgdl = False

    # ── mutation ──────────────────────────────────────────────────────────

    def add_document(self, doc_id, chunks):
        """Index a document's chunks (list of {"page","seq","text"}).

        Re-adding an existing doc_id replaces it.
        """
        self.remove_document(doc_id)
        for chunk in chunks:
            key = make_chunk_key(doc_id, chunk["page"], chunk["seq"])
            tokens = vi_text.tokenize(chunk["text"])
            if not tokens:
                continue
            self.doclen[key] = len(tokens)
            tf = {}
            for tok in tokens:
                tf[tok] = tf.get(tok, 0) + 1
            for term, count in tf.items():
                self.postings.setdefault(term, []).append([key, count])
        self._dirty_avgdl = True

    def remove_document(self, doc_id):
        """Drop every chunk whose key starts with `doc_id#`."""
        prefix = doc_id + u'#'
        had = [k for k in self.doclen if k.startswith(prefix)]
        if not had:
            return
        for key in had:
            del self.doclen[key]
        empty_terms = []
        for term, plist in self.postings.items():
            kept = [p for p in plist if not p[0].startswith(prefix)]
            if kept:
                self.postings[term] = kept
            else:
                empty_terms.append(term)
        for term in empty_terms:
            del self.postings[term]
        self._dirty_avgdl = True

    # ── query ─────────────────────────────────────────────────────────────

    @property
    def size(self):
        return len(self.doclen)

    def _avg_doclen(self):
        if self._dirty_avgdl:
            n = len(self.doclen)
            total = 0
            for v in self.doclen.values():
                total += v
            self._avgdl = (float(total) / n) if n else 0.0
            self._dirty_avgdl = False
        return self._avgdl

    def search(self, query_text, top_k=8, allowed_docs=None):
        """Return [(chunk_key, score)] sorted best-first.

        allowed_docs: optional set of doc_ids — restrict scoring to chunks
        of those documents (used for attached-file retrieval).
        """
        terms = vi_text.tokenize(query_text)
        n = len(self.doclen)
        if not terms or not n:
            return []
        avgdl = self._avg_doclen() or 1.0
        allowed = set(allowed_docs) if allowed_docs else None

        scores = {}
        for term in set(terms):
            plist = self.postings.get(term)
            if not plist:
                continue
            df = len(plist)
            idf = math.log(1.0 + (n - df + 0.5) / (df + 0.5))
            for key, tf in plist:
                if allowed is not None and doc_id_of(key) not in allowed:
                    continue
                dl = self.doclen.get(key, 0)
                denom = tf + K1 * (1.0 - B + B * dl / avgdl)
                if denom <= 0:
                    continue
                scores[key] = scores.get(key, 0.0) + idf * tf * (K1 + 1.0) / denom

        ranked = sorted(scores.items(), key=lambda kv: -kv[1])
        return ranked[:top_k]

    # ── persistence ───────────────────────────────────────────────────────

    def to_dict(self):
        return {
            "version":  1,
            "postings": self.postings,
            "doclen":   self.doclen,
        }

    @classmethod
    def from_dict(cls, d):
        idx = cls()
        if not isinstance(d, dict):
            return idx
        idx.postings = d.get("postings", {}) or {}
        idx.doclen = d.get("doclen", {}) or {}
        idx._dirty_avgdl = True
        return idx
