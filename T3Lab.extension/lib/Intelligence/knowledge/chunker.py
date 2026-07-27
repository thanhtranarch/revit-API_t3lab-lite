# -*- coding: utf-8 -*-
"""
Page-aware text chunking for the knowledge stack.

Chunks never span pages, so every retrieved excerpt carries an exact
(file, page) citation. Pure Python — IronPython 2.7 + CPython 3.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import re

# ─── Heading detection (step 1 of the RAG cycle: cut by MEANING) ──────────────
# Splitting a standards document into blind 650-word windows cuts tables and
# clause lists in half, so a retrieved excerpt can show discipline codes with
# the heading that says what they are left on the previous chunk. Sections are
# detected first, and each chunk carries its heading, so an excerpt is
# self-describing when the model reads it.

_RE_MD_HEAD = re.compile(r'^\s{0,3}#{1,6}\s+\S')
# "1. Title", "1.2 Title", "2.3.1) Title", "A. Title" — the trailing dot/paren
# is OPTIONAL once the number itself is dotted, which is how most standards
# documents write sub-clauses ("1.2 File Naming").
_RE_NUM_HEAD = re.compile(
    r'^\s{0,6}(?:\d{1,2}(?:\.\d{1,2}){1,3}[\.\)]?|\d{1,2}[\.\)]|[A-Z][\.\)]'
    r'|[IVXLC]{1,5}[\.\)])\s+\S')
_MAX_HEAD_WORDS = 12
_MAX_HEAD_CHARS = 90
# Sections shorter than this are folded back into the previous section: a
# mis-detected heading (a short table cell) then costs nothing.
_MIN_SECTION_WORDS = 25


def is_heading(line):
    """True when a line looks like a section heading in extracted document text."""
    s = (line or '').strip()
    if not s or len(s) > _MAX_HEAD_CHARS:
        return False
    if _RE_MD_HEAD.match(s):
        return True
    if _RE_NUM_HEAD.match(s):
        return True
    words = s.split()
    if not words or len(words) > _MAX_HEAD_WORDS:
        return False
    # sentence punctuation means prose, not a heading (a trailing ':' is fine)
    if s.endswith(('.', ',', ';')):
        return False
    letters = [c for c in s if c.isalpha()]
    if len(letters) < 2:
        return False
    if all(c.isupper() for c in letters):
        return True                      # ALL-CAPS title
    if s.endswith(':'):
        return True                      # "Discipline Code:"
    # short Title-Case line ("Discipline", "Central Model", "Link Model")
    if len(words) <= 5 and words[0][:1].isupper():
        caps = sum(1 for w in words if w[:1].isupper())
        if caps >= max(1, len(words) - 1):
            return True
    return False


def split_sections(text):
    """Split page text into [(heading, body)] on heading-looking lines.

    Returns a single ('', text) section when no headings are found.
    """
    lines = (text or '').split('\n')
    sections = []
    head, buf = '', []
    for ln in lines:
        if is_heading(ln):
            if buf and ' '.join(buf).strip():
                sections.append((head, '\n'.join(buf).strip()))
            head = ln.strip().lstrip('#').strip()
            buf = []
        else:
            buf.append(ln)
    if buf and ' '.join(buf).strip():
        sections.append((head, '\n'.join(buf).strip()))
    if not sections:
        body = (text or '').strip()
        return [('', body)] if body else []

    # Fold tiny sections into the previous one. Standards documents are full
    # of short lines that look like headings (table cells, code columns);
    # merging them back keeps chunks meaningful instead of shattering a code
    # table into one chunk per row.
    merged = []
    for head, body in sections:
        if (merged and len(body.split()) < _MIN_SECTION_WORDS):
            p_head, p_body = merged[-1]
            merged[-1] = (p_head,
                          p_body + '\n' + ((head + ': ') if head else '') + body)
        else:
            merged.append((head, body))
    return merged


def chunk_pages(pages, target_words=650, overlap_words=120, semantic=True):
    """Chunk a list of (page_no, text) into overlapping word windows.

    Args:
        pages: list of (page_no, unicode_text) — from rag_processor.extract_pdf_pages.
        target_words: window size in words.
        overlap_words: words shared between consecutive chunks (same page only).
        semantic: split on headings first and prefix each chunk with its
            heading, so a chunk never silently loses the clause it belongs to.

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
        if not (text or '').strip():
            continue
        blocks = (split_sections(text) if semantic
                  else [('', text)])
        for head, body in blocks:
            words = (body or u'').split()
            if not words:
                continue
            prefix = (head + u': ') if head else u''
            start = 0
            while start < len(words):
                window = words[start:start + target_words]
                # Avoid a trailing crumb: if the remainder is tiny and we
                # already emitted a chunk for this page, merge it back.
                if (chunks and chunks[-1]["page"] == page_no
                        and len(window) < overlap_words):
                    chunks[-1]["text"] = (chunks[-1]["text"] + u' '
                                          + u' '.join(window))
                    break
                chunks.append({
                    "page": page_no,
                    "seq":  seq,
                    "text": prefix + u' '.join(window),
                })
                seq += 1
                if start + target_words >= len(words):
                    break
                start += step
    return chunks


def chunk_text(text, target_words=650, overlap_words=120):
    """Chunk a plain text (.txt/.md) as a single pseudo-page 0."""
    return chunk_pages([(0, text)], target_words, overlap_words)
