# -*- coding: utf-8 -*-
"""
CPython 3 tests for Intelligence/learning/dataset.py (self-study training corpus).

Run:  python3 dev/test_learning_dataset.py
Exit 0 = all pass. Plain asserts, mirroring dev/test_assistant_llm.py.

The dataset persists to %APPDATA%/T3LabAI/training/dataset.jsonl, so APPDATA is
sandboxed to a temp dir first and dataset.reset() clears the in-memory dedup
index between cases.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_ds_test_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence.learning import dataset as D

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def test_add_and_dedupe():
    print('[dataset: add + content dedupe]')
    D.reset()
    ok, note = D.add_qa(u'thong ke tuong', u'Co 120 buc tuong.',
                        source='model_snapshot', revit_version=2023)
    check('first add ok', ok and note == u'added', note)
    ok2, note2 = D.add_qa(u'thong ke tuong', u'Co 120 buc tuong.',
                          source='model_snapshot')
    check('identical Q&A dedupes', ok2 and note2 == u'duplicate', note2)
    # different system prompt, same user+assistant → still a duplicate
    ok3, note3 = D.add_qa(u'thong ke tuong', u'Co 120 buc tuong.',
                          source='x', system=u'You are T3Lab.')
    check('system prompt ignored for dedupe', note3 == u'duplicate', note3)
    check('count is 1', D.stats()['count'] == 1, D.stats())


def test_rejects_degenerate():
    print('[dataset: rejects malformed examples]')
    D.reset()
    check('no assistant turn rejected',
          D.add_example([{'role': 'user', 'content': 'hello there'}],
                        source='x')[0] is False)
    check('too-short content rejected',
          D.add_example([{'role': 'user', 'content': 'x'},
                         {'role': 'assistant', 'content': 'y'}],
                        source='x')[0] is False)
    check('bad role rejected',
          D.add_example([{'role': 'tool', 'content': 'aaaa'},
                         {'role': 'assistant', 'content': 'bbbb'}],
                        source='x')[0] is False)
    check('nothing was written', D.stats()['count'] == 0, D.stats())


def test_stats_and_export():
    print('[dataset: stats + export_sft]')
    D.reset()
    D.add_qa(u'q one here', u'answer one', source='api_facts', quality='high')
    D.add_qa(u'q two here', u'answer two', source='telemetry_miner')
    s = D.stats()
    check('count 2', s['count'] == 2, s)
    check('by_source split', s['by_source'].get('api_facts') == 1
          and s['by_source'].get('telemetry_miner') == 1, s['by_source'])
    check('by_quality split', s['by_quality'].get('high') == 1
          and s['by_quality'].get('ok') == 1, s['by_quality'])

    out = os.path.join(tempfile.mkdtemp(), 'sft.jsonl')
    n = D.export_sft(out)
    check('export wrote 2', n == 2, n)
    import io as _io, json as _json
    lines = [l for l in _io.open(out, encoding='utf-8').read().splitlines() if l]
    check('export lines have messages only',
          all(set(_json.loads(l).keys()) == {'messages'} for l in lines), lines)


def test_survives_malformed_line():
    print('[dataset: a junk line never breaks iteration]')
    D.reset()
    D.add_qa(u'good q here', u'good a', source='x')
    # inject a broken line directly
    import io as _io
    with _io.open(D.dataset_file(), 'a', encoding='utf-8') as f:
        f.write(u'{ this is not json\n')
    check('iter skips the junk line', sum(1 for _ in D.iter_examples()) == 1,
          list(D.iter_examples()))
    check('stats still counts the good one', D.stats()['count'] == 1)


def main():
    test_add_and_dedupe()
    test_rejects_degenerate()
    test_stats_and_export()
    test_survives_malformed_line()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All learning-dataset tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
