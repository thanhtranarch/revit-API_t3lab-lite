# -*- coding: utf-8 -*-
"""
CPython 3 tests for Intelligence/learning/study_engine.py.

Run:  python3 dev/test_study_engine.py
Exit 0 = all pass. Plain asserts.

Locks down the idle-cycle policy the self-study loop depends on: the gate
(disabled / busy / cancelled / no-runnable), round-robin over enrichers,
generative-skip when no local model, the rolling hourly budget, and failure
containment. All driven by an injected clock + fake enrichers — no Revit/LLM.
"""
from __future__ import unicode_literals

import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence.learning.study_engine import StudyEngine, Enricher, Gate

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def _counter(calls, key, result=None):
    def fn():
        calls[key] = calls.get(key, 0) + 1
        return result or {'status': 'ok', 'added': 1}
    return fn


def test_gate_blocks():
    print('[study_engine: gate blocks disabled/busy/cancel]')
    calls = {}
    eng = StudyEngine([Enricher('a', _counter(calls, 'a'))], budget_per_hour=10)
    check('disabled → skip',
          eng.run_cycle(Gate(enabled=False, idle=True))['status'] == 'disabled')
    check('not idle → busy',
          eng.run_cycle(Gate(enabled=True, idle=False))['status'] == 'busy')
    check('cancel → cancelled',
          eng.run_cycle(Gate(enabled=True, idle=True,
                             cancel=lambda: True))['status'] == 'cancelled')
    check('nothing ran', calls.get('a', 0) == 0, calls)


def test_round_robin_and_generative_skip():
    print('[study_engine: round-robin + generative skip]')
    calls = {}
    eng = StudyEngine([
        Enricher('a', _counter(calls, 'a')),
        Enricher('b', _counter(calls, 'b')),
        Enricher('gen', _counter(calls, 'gen'), generative=True),
    ], budget_per_hour=10)

    g = Gate(enabled=True, idle=True, has_local_model=False)
    seq = [eng.run_cycle(g)['ran'] for _ in range(4)]
    check('generative skipped without local model', 'gen' not in seq, seq)
    check('cycles through a,b,a,b', seq == ['a', 'b', 'a', 'b'], seq)

    g2 = Gate(enabled=True, idle=True, has_local_model=True)
    ran = eng.run_cycle(g2)['ran']            # cursor now at 'gen'
    check('generative runs with local model', ran == 'gen', ran)


def test_budget_window():
    print('[study_engine: rolling hourly budget]')
    calls = {}
    clk = {'t': 0.0}
    eng = StudyEngine([Enricher('a', _counter(calls, 'a'))],
                      budget_per_hour=2, clock=lambda: clk['t'])
    g = Gate(enabled=True, idle=True)
    check('run 1', eng.run_cycle(g)['ran'] == 'a')
    check('run 2', eng.run_cycle(g)['ran'] == 'a')
    check('run 3 over budget',
          eng.run_cycle(g)['status'] == 'budget_exhausted')
    clk['t'] += 3601                      # slide past the hour window
    check('budget resets after an hour', eng.run_cycle(g)['ran'] == 'a')


def test_all_generative_no_model_is_noop():
    print('[study_engine: all-generative + no model → no_runnable]')
    eng = StudyEngine([Enricher('g', lambda: {'status': 'ok'}, generative=True)],
                      budget_per_hour=10)
    res = eng.run_cycle(Gate(enabled=True, idle=True, has_local_model=False))
    check('reports no runnable enricher',
          res['status'] == 'no_runnable_enricher', res)


def test_failure_contained():
    print('[study_engine: a throwing enricher is contained]')
    def boom():
        raise ValueError('kaboom')
    calls = {}
    eng = StudyEngine([Enricher('bad', boom),
                       Enricher('good', _counter(calls, 'good'))],
                      budget_per_hour=10)
    g = Gate(enabled=True, idle=True)
    r1 = eng.run_cycle(g)
    check('failure surfaced, not raised', r1['ran'] == 'bad'
          and r1['status'].startswith('error'), r1)
    r2 = eng.run_cycle(g)
    check('next enricher still runs', r2['ran'] == 'good', r2)


def main():
    test_gate_blocks()
    test_round_robin_and_generative_skip()
    test_budget_window()
    test_all_generative_no_model_is_noop()
    test_failure_contained()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All study-engine tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
