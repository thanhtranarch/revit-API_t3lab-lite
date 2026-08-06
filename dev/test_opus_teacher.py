# -*- coding: utf-8 -*-
"""
CPython 3 tests for the Opus-5 teacher distillation path.

Run:  python3 dev/test_opus_teacher.py
Exit 0 = all pass. Plain asserts, no Revit/WPF/LLM needed.

Covers:
  * local_llm.pick_tool_capable — Qwen tool-capable auto-pick (Part A).
  * study_engine teacher gating — requires='teacher' skipped without has_teacher
    while local enrichers still work (Part B1).
  * opus_teacher pure selection + run() with a fake teacher chat_fn (Part B2):
    terse-pattern harvesting, seed fill, corpus dedup, quality tag, never-raise.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_teacher_test_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence import local_llm
from Intelligence.learning.study_engine import StudyEngine, Enricher, Gate
from Intelligence.learning.enrichers import opus_teacher as ot

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─── Part A — Qwen tool-capable auto-pick ────────────────────────────────────────

def test_pick_tool_capable():
    print('[local_llm: pick_tool_capable]')
    check('sweet spot 14b wins',
          local_llm.pick_tool_capable(['qwen3:8b', 'qwen3:14b', 'llama3.2:1b'])
          == 'qwen3:14b')
    check('falls to 8b when no 14b',
          local_llm.pick_tool_capable(['qwen2.5:1.5b', 'qwen3:8b']) == 'qwen3:8b')
    check('capable non-qwen at/above floor',
          local_llm.pick_tool_capable(['llama3:8b', 'mistral:7b']) == 'llama3:8b')
    check('only tiny → pick_best fallback (never None on non-empty)',
          local_llm.pick_tool_capable(['qwen2.5:0.5b', 'llama3.2:1b']) is not None)
    check('empty → None', local_llm.pick_tool_capable([]) is None)
    check('recommended tool model is a real sweet-spot tag',
          local_llm.recommended_tool_model() == 'qwen3:14b')


# ─── Part B1 — study engine teacher gating ───────────────────────────────────────

def test_teacher_gating():
    print('[study_engine: teacher requires gating]')
    seen = []
    eng = StudyEngine([
        Enricher('local', lambda: seen.append('local') or {'status': 'ok'},
                 generative=True),
        Enricher('teach', lambda: seen.append('teach') or {'status': 'ok',
                                                            'added': 1},
                 requires='teacher'),
    ], budget_per_hour=20)

    # No local model, no teacher → nothing runnable.
    r = eng.run_cycle(Gate(enabled=True, idle=True,
                           has_local_model=False, has_teacher=False))
    check('nothing runnable when neither resource present',
          r['status'] == 'no_runnable_enricher', r)

    # Teacher available → teacher enricher runs even with no local model.
    r = eng.run_cycle(Gate(enabled=True, idle=True,
                           has_local_model=False, has_teacher=True))
    check('teacher runs when has_teacher, without local model',
          r.get('ran') == 'teach' and r.get('added') == 1, r)
    check('local enricher was skipped (no local model)', 'local' not in seen, seen)

    # Legacy generative alias still means requires="local".
    e = Enricher('legacy', lambda: {'status': 'ok'}, generative=True)
    check('generative=True aliases requires="local"', e.requires == 'local')


# ─── Part B2 — pure candidate selection ──────────────────────────────────────────

_PATTERNS = {
    'k1': {'last_raw': u'mở batchout', 'last_message': u'Đang mở BatchOut...',
           'hits': 5},
    'k2': {'last_raw': u'cách đặt tên workset', 'last_message': u'', 'hits': 2},
    'k3': {'last_raw': u'giải thích LOD',
           'last_message': (u'LOD is a detailed rich stored answer that is well '
                            u'over forty characters long already'), 'hits': 9},
}


def test_terse_and_selection():
    print('[opus_teacher: candidate selection]')
    terse = ot.terse_pattern_prompts(_PATTERNS)
    check('terse harvests only poor-answer patterns',
          terse == [u'mở batchout', u'cách đặt tên workset'], terse)

    cands = ot.select_candidates(_PATTERNS, seen_users=set(), max_n=3)
    check('patterns first, then seeds fill the budget',
          cands[:2] == [u'mở batchout', u'cách đặt tên workset']
          and len(cands) == 3, cands)

    # Already-taught prompt is skipped.
    seen = set([ot._fold(u'mở batchout')])
    cands2 = ot.select_candidates(_PATTERNS, seen_users=seen, max_n=3)
    check('folded seen user-turn is excluded',
          all(ot._fold(c) != ot._fold(u'mở batchout') for c in cands2), cands2)


# ─── Part B2 — run() with a fake teacher ─────────────────────────────────────────

def test_run_writes_teacher_rows():
    print('[opus_teacher: run() writes quality=teacher rows]')
    rows = []

    def fake_add(messages, source, quality='ok', revit_version=None, meta=None):
        rows.append((source, quality, messages[-1]['content']))
        return True, u'added'

    def fake_chat(prompt):
        return u'Gold Opus answer for: {}'.format(prompt)

    res = ot.run(chat_fn=fake_chat, add_example=fake_add,
                 load_patterns=lambda: _PATTERNS,
                 iter_examples=lambda: iter([]), max_new=2)
    check('run reports added=2', res.get('added') == 2, res)
    check('every row tagged source+quality teacher',
          rows and all(s == 'opus_teacher' and q == 'teacher'
                       for s, q, _ in rows), rows)
    check('assistant content is the teacher answer',
          rows[0][2].startswith(u'Gold Opus answer for:'), rows[0][2])


def test_run_dedupes_against_corpus():
    print('[opus_teacher: run() skips prompts already in corpus]')
    taught = []

    def fake_add(messages, source, quality='ok', revit_version=None, meta=None):
        taught.append(messages[1]['content'])
        return True, u'added'

    # Corpus already contains "mở batchout" as a user turn.
    corpus = [{'messages': [{'role': 'user', 'content': u'mở batchout'},
                            {'role': 'assistant', 'content': u'x'}]}]
    ot.run(chat_fn=lambda p: u'ans', add_example=fake_add,
           load_patterns=lambda: _PATTERNS,
           iter_examples=lambda: iter(corpus), max_new=5)
    check('previously-taught prompt not re-taught',
          u'mở batchout' not in taught, taught)


def test_run_never_raises():
    print('[opus_teacher: run() is failure-contained]')

    def boom(prompt):
        raise ValueError('teacher exploded')

    res = ot.run(chat_fn=boom, add_example=lambda *a, **k: (True, u'added'),
                 load_patterns=lambda: _PATTERNS,
                 iter_examples=lambda: iter([]), max_new=1)
    check('teacher exception contained into status', res.get('added') == 0
          and 'failed' in res.get('status', ''), res)
    check('missing chat_fn → no-op, not crash',
          ot.run(chat_fn=None).get('added') == 0)


if __name__ == '__main__':
    test_pick_tool_capable()
    test_teacher_gating()
    test_terse_and_selection()
    test_run_writes_teacher_rows()
    test_run_dedupes_against_corpus()
    test_run_never_raises()
    print('')
    if FAILURES:
        print('FAILED ({}): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        sys.exit(1)
    print('All opus-teacher tests passed.')
