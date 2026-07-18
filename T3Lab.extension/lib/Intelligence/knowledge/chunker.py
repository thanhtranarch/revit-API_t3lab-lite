# -*- coding: utf-8 -*-
"""
Page-aware text chunking for the knowledge stack.

Chunks never span pages, so every retrieved excerpt carries an exact
(file, page) citation. Pure Python — IronPython 2.7 + CPython 3.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals


def chunk_pages(pages, target_words=650, overlap_words=120):
    """Chunk a list of (page_no, text) into overlapping word windows.

    Args:
        pages: list of (page_no, unicode_text) — from rag_processor.extract_pdf_pages.
        target_words: window size in words.
        overlap_words: words shared between consecutive chunks (same page only).

    Returns:
        list of {"page": int, "seq": int, "text": unicode}; `seq` is unique
        across the whole document.
    """
    if overlap_words >= target_words:
        overlap_words = target_words // 4
    step = target_words - overlap_words

    chunks = []
    seq = 0
    for page_no, text in (pages or []):
        words = (text or u'').split()
        if not words:
            continue
        start = 0
        while start < len(words):
            window = words[start:start + target_words]
            # Avoid a trailing crumb: if the remainder is tiny and we already
            # emitted a chunk for this page, merge it into the previous one.
            if chunks and chunks[-1]["page"] == page_no and len(window) < overlap_words:
                chunks[-1]["text"] = chunks[-1]["text"] + u' ' + u' '.join(window)
                break
            chunks.append({
                "page": page_no,
                "seq":  seq,
                "text": u' '.join(window),
            })
            seq += 1
            if start + target_words >= len(words):
                break
            start += step
    return chunks


def chunk_text(text, target_words=650, overlap_words=120):
    """Chunk a plain text (.txt/.md) as a single pseudo-page 0."""
    return chunk_pages([(0, text)], target_words, overlap_words)
