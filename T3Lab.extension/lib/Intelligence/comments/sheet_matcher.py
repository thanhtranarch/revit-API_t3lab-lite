# -*- coding: utf-8 -*-
"""
Sheet matcher — trace a PDF drawing back to a Revit sheet.

Candidates come from the PDF filename and title-block text (patterns
like A-101, S201, M-3001A near SHEET/DWG/BAN VE keywords); matching
prefers exact SheetNumber equality (also separator-insensitive), then
folded-token name overlap.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import os
import re

from Intelligence.knowledge import vi_text

_SHEET_NO_RE = re.compile(r'\b([A-Z]{1,3}[-_ ]?\d{2,4}[A-Z]?)\b')
_KEYWORDS = ('SHEET', 'DWG', 'DRAWING', 'BAN VE', 'SO BAN VE', 'KY HIEU',
             'SHEET NO', 'SHT')
_NAME_MATCH_THRESHOLD = 0.55


def _canon(number):
    """Canonical separator-free uppercase form: 'a-101' -> 'A101'."""
    return re.sub(r'[-_ ]', '', (number or '').upper())


def extract_sheet_candidates(page_texts, pdf_filename):
    """Ranked candidate sheet-number strings.

    Order: filename hits, then hits near title-block keywords, then the
    rest (deduped by canonical form).
    """
    ranked = []
    seen = set()

    def _add(cand):
        c = _canon(cand)
        if c and c not in seen:
            seen.add(c)
            ranked.append(cand.strip())

    # '_' is a word char — 'A-101_MatBang' would kill the \b boundary,
    # so separators are normalized to spaces before scanning.
    base = os.path.splitext(os.path.basename(pdf_filename or ''))[0]
    for m in _SHEET_NO_RE.finditer(base.upper().replace('_', ' ')):
        _add(m.group(1))

    near_kw = []
    others = []
    for _page, text in (page_texts or []):
        up = vi_text.fold_diacritics(text or '').upper().replace('_', ' ')
        for m in _SHEET_NO_RE.finditer(up):
            window = up[max(0, m.start() - 40):m.end() + 40]
            if any(kw in window for kw in _KEYWORDS):
                near_kw.append(m.group(1))
            else:
                others.append(m.group(1))
    for cand in near_kw + others:
        _add(cand)
    return ranked


def match_sheets(candidates, revit_sheets, pdf_filename=''):
    """Best Revit sheet for the candidate numbers.

    Args:
        candidates: from extract_sheet_candidates.
        revit_sheets: [{"name","number","id"}] — revit_list_sheets result.
        pdf_filename: optional, for name-overlap fallback.

    Returns {"number","name","id","score"} or None.
    """
    sheets = [s for s in (revit_sheets or []) if s.get('number')]
    if not sheets:
        return None

    by_canon = {}
    for s in sheets:
        by_canon.setdefault(_canon(s['number']), s)

    # 1) exact (separator-insensitive) number match, candidate order wins
    for cand in (candidates or []):
        hit = by_canon.get(_canon(cand))
        if hit:
            return {'number': hit['number'], 'name': hit.get('name', ''),
                    'id': hit.get('id'), 'score': 1.0}

    # 2) fuzzy: folded-token overlap between filename and sheet names
    fname_tokens = vi_text.tokenize(
        os.path.splitext(os.path.basename(pdf_filename or ''))[0]
        .replace('_', ' ').replace('-', ' '))
    if not fname_tokens:
        return None
    best = None
    best_score = 0.0
    for s in sheets:
        score = vi_text.word_match_score(
            fname_tokens, vi_text.tokenize(s.get('name', '')))
        if score > best_score:
            best_score = score
            best = s
    if best is not None and best_score >= _NAME_MATCH_THRESHOLD:
        return {'number': best['number'], 'name': best.get('name', ''),
                'id': best.get('id'), 'score': round(best_score, 2)}
    return None
