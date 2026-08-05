# -*- coding: utf-8 -*-
"""
CPython 3 tests for the self-study enrichers' pure helpers + injectable run().

Run:  python3 dev/test_learning_enrichers.py
Exit 0 = all pass. Plain asserts.

Covers the parts that DON'T need Revit/LLM: telemetry_miner's SFT distillation +
👎 pruning, api_facts' fact→Q&A transform, and model_snapshot's digest→grounding
text (which deliberately emits NO training rows for transient model counts).
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_enr_test_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence.learning.enrichers import telemetry_miner as TM
from Intelligence.learning.enrichers import api_facts as AF
from Intelligence.learning.enrichers import model_snapshot as MS

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─── telemetry_miner ────────────────────────────────────────────────────────────

def test_miner_sft_and_prune():
    print('[telemetry_miner: distill + prune]')
    patterns = {
        'export pdf sheet a': {'intent': 'export_direct', 'hits': 5,
                               'last_raw': u'export pdf sheet A',
                               'last_message': u'Da xuat 5 sheet A.'},
        'xoa text note': {'intent': 'delete_element', 'hits': 1,
                          'last_raw': u'xoa text note',
                          'last_message': u'Da xoa.'},
        'no target': {'intent': 'export_direct', 'hits': 9,
                      'last_raw': u'x', 'last_message': u''},
    }
    negatives = {'xoa text note': {'intents': ['delete_element']}}

    ex = TM.sft_examples_from_patterns(patterns)
    users = [u for (u, a, q) in ex]
    check('example with a real answer is emitted',
          u'export pdf sheet A' in users, users)
    check('pattern with empty last_message is skipped',
          all(u != u'x' for u in users), users)
    check('high-hit pattern tagged high',
          any(q == 'high' for (u, a, q) in ex if u == u'export pdf sheet A'), ex)

    check('👎 route is prunable',
          TM.prunable_keys(patterns, negatives) == {'xoa text note'})
    # a 👎 on a DIFFERENT intent must not prune
    check('👎 on another intent is not pruned',
          TM.prunable_keys(patterns,
                           {'xoa text note': {'intents': ['create_element']}})
          == set())


def test_miner_run_injected():
    print('[telemetry_miner: run() with injected stores]')
    patterns = {
        'good cmd here': {'intent': 'export_direct', 'hits': 4,
                          'last_raw': u'good cmd here',
                          'last_message': u'Done.'},
        'bad cmd here': {'intent': 'delete_element', 'hits': 2,
                         'last_raw': u'bad cmd here', 'last_message': u'Oops.'},
    }
    negatives = {'bad cmd here': {'intents': ['delete_element']}}
    saved = {}
    added = []

    def fake_add(messages, source, quality='ok', revit_version=None, meta=None):
        added.append(messages[0]['content'])
        return True, u'added'

    res = TM.run(add_example=fake_add,
                 load_patterns=lambda: dict(patterns),
                 save_patterns=lambda p: saved.update(p),
                 load_feedback=lambda: negatives)
    check('run reports ok', res.get('status') == 'ok', res)
    check('one route pruned', res.get('pruned') == 1, res)
    check('surviving pattern distilled', added == [u'good cmd here'], added)
    check('pruned key removed from saved patterns',
          'bad cmd here' not in saved and 'good cmd here' in saved, saved)


# ─── api_facts ──────────────────────────────────────────────────────────────────

def test_api_facts_transform():
    print('[api_facts: fact → Q&A]')
    info = {
        'revit_version': 2024,
        'export_api': {
            'pdf_export': {'signature': 'Export(folder, viewIds, PDFExportOptions)',
                           'note': 'Use PDFExportOptions.FileName'},
        },
        'dwg_export_options': {
            'exporting_areas': {'type': 'bool', 'available': True},
        },
    }
    rows = AF.api_facts_to_sft(info, 2024)
    check('emits at least the two facts', len(rows) >= 2, len(rows))
    joined = u' '.join(u'{} {}'.format(u, a) for (u, a) in rows)
    check('carries the version tag', u'Revit 2024' in joined, joined[:80])
    check('carries the signature', u'PDFExportOptions' in joined)
    check('empty info yields nothing', AF.api_facts_to_sft({}) == [])


# ─── model_snapshot ─────────────────────────────────────────────────────────────

def test_snapshot_grounding_only():
    print('[model_snapshot: grounding text, no transient training rows]')
    digest = {'element_counts': {'Walls': 120, 'Doors': 40, 'Rooms': 0},
              'total_views': 85, 'project': 'Tower A', 'ts': '2026-08-05'}
    text = MS.digest_to_grounding_text(digest)
    check('grounding mentions the project', u'Tower A' in text, text)
    check('non-zero counts included', u'Walls: 120' in text, text)
    check('zero-count category omitted', u'Rooms' not in text, text)
    check('empty digest → empty text', MS.digest_to_grounding_text({}) == u'')

    # run() must add ZERO training rows (counts are transient, not trainable)
    def fake_exec(name, args):
        if name == 'analyze_model_statistics':
            return {'element_counts': {'Walls': 5}, 'total_views': 3,
                    'project': 'P'}
        return {}
    res = MS.run(execute_tool=fake_exec, doc_key='unit-test-doc')
    check('run caches, trains nothing',
          res.get('added') == 0 and res.get('cached') is True, res)


def main():
    test_miner_sft_and_prune()
    test_miner_run_injected()
    test_api_facts_transform()
    test_snapshot_grounding_only()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All learning-enricher tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
