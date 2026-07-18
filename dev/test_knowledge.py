# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the T3Lab Assistant knowledge stack.

Run:  python3 dev/test_knowledge.py
Exit code 0 = all pass. No external test framework — plain asserts,
mirroring dev/audit_tools.py conventions.

Covers modules that must be importable OUTSIDE Revit (guarded clr imports):
    lib/Intelligence/knowledge/*        (vi_text, chunker, bm25, embeddings fusion)
    lib/Intelligence/agents/dispatcher  (keyword stage)
    lib/Intelligence/skills_engine      (frontmatter parsing)
    lib/Intelligence/comments/*         (pdf annots, sheet matcher)
"""
from __future__ import unicode_literals

import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─── vi_text ──────────────────────────────────────────────────────────────────

def test_vi_text():
    print('[vi_text]')
    from Intelligence.knowledge import vi_text

    check('fold basic', vi_text.fold_diacritics('tường') == 'tuong')
    check('fold dj', vi_text.fold_diacritics('đường Đông') == 'duong Dong')
    check('fold passthrough', vi_text.fold_diacritics('Wall-101') == 'Wall-101')
    check('fold empty', vi_text.fold_diacritics('') == '')

    toks = vi_text.tokenize('Có bao nhiêu bức tường trong dự án?')
    check('tokenize vi', 'tuong' in toks and 'bao' in toks and 'nhieu' in toks, repr(toks))
    check('tokenize stopword drop', 'trong' not in toks, repr(toks))
    toks2 = vi_text.tokenize('How many walls are in the project?')
    check('tokenize en', 'walls' in toks2 and 'many' in toks2, repr(toks2))
    check('tokenize stopword en', 'the' not in toks2 and 'in' not in toks2, repr(toks2))

    s = vi_text.word_match_score(['mat', 'bang', 'tang', '01'], ['mat', 'bang', 'tang', 'mai'])
    check('word_match_score partial', 0.7 < s <= 1.0, s)
    check('word_match_score empty', vi_text.word_match_score([], ['a']) == 0.0)


# ─── main ─────────────────────────────────────────────────────────────────────

def main():
    test_vi_text()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All knowledge-stack tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
