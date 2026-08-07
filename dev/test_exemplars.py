# -*- coding: utf-8 -*-
"""
CPython 3 tests for the portable teacher-exemplar layer (Part G).

Run:  python3 dev/test_exemplars.py
Exit 0 = all pass. No Revit/WPF/LLM needed.

Covers: trajectory/qa converters, select (quality filter + dedup + cap + kind
order), build_exemplar_block (style + char budget), promote round-trip,
load/save, never-raise, and that build_agent_system_prompt(local=True) picks up
the block while cloud does not.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_exempl_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence.learning import exemplars as X

# Redirect the exemplar store to a throwaway file so tests never touch the
# git-tracked lib/Intelligence/config/teacher_exemplars.json.
_TMP_STORE = os.path.join(tempfile.mkdtemp(prefix='t3lab_exstore_'),
                          'teacher_exemplars.json')
X._exemplars_file = lambda: _TMP_STORE

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def _traj(user, calls, final=None, quality='teacher', source='mcp_teacher'):
    msgs = [{'role': 'user', 'content': user}]
    for i, (name, args) in enumerate(calls):
        msgs.append({'role': 'assistant', 'content': '',
                     'tool_calls': [{'id': 'call_%d' % i, 'name': name,
                                     'arguments': args}]})
        msgs.append({'role': 'tool', 'content': '{}', 'tool_call_id': 'call_%d' % i,
                     'name': name})
    if final:
        msgs.append({'role': 'assistant', 'content': final})
    return {'meta': {'source': source, 'quality': quality}, 'messages': msgs}


def _qa(user, answer, quality='teacher', source='opus_teacher'):
    return {'meta': {'source': source, 'quality': quality},
            'messages': [{'role': 'system', 'content': 's'},
                         {'role': 'user', 'content': user},
                         {'role': 'assistant', 'content': answer}]}


def test_converters():
    print('[exemplars: converters]')
    ex = X.trajectory_to_exemplar(
        _traj('tag walls', [('revit_list_views', '{}'),
                            ('tag_all_walls', '{"view":"L1"}')],
              final='Đã tag 12 tường.')['messages'])
    check('trajectory -> tool kind', ex and ex['kind'] == 'tool', ex)
    check('trace has call + reply',
          'call revit_list_views()' in ex['trace']
          and 'call tag_all_walls(view=L1)' in ex['trace']
          and 'reply "' in ex['trace'], ex['trace'])

    qa = X.qa_to_exemplar(_qa('what is LOD', 'Level of Development.')['messages'])
    check('qa -> qa kind + reply', qa and qa['kind'] == 'qa'
          and qa['trace'].startswith('reply "'), qa)

    check('trajectory with no tool calls -> None',
          X.trajectory_to_exemplar([{'role': 'user', 'content': 'hi'},
                                     {'role': 'assistant', 'content': 'ok'}]) is None)
    check('no user -> None', X.qa_to_exemplar(
        [{'role': 'assistant', 'content': 'x'}]) is None)


def test_select():
    print('[exemplars: select (filter + dedup + cap + order)]')
    rows = [
        _traj('tag walls', [('tag_all_walls', '{}')], final='done'),
        _traj('tag walls', [('tag_all_walls', '{}')], final='done'),   # dup user
        _qa('what is LOD', 'Level of Development'),
        _qa('low quality', 'x', quality='ok'),                          # dropped
        _traj('count doors', [('ai_element_filter', '{"category":"Doors"}')],
              final='84'),
    ]
    sel = X.select_exemplars(rows)
    users = [e['user'] for e in sel]
    check('duplicate user dropped', users.count('tag walls') == 1, users)
    check('low-quality row dropped', 'low quality' not in users, users)
    check('tool exemplars come before qa',
          [e['kind'] for e in sel] == ['tool', 'tool', 'qa'],
          [e['kind'] for e in sel])

    capped = X.select_exemplars(rows, max_n=1)
    check('max_n caps total', len(capped) == 1, capped)


def test_block():
    print('[exemplars: build block + budget]')
    sel = [{'user': 'a', 'trace': 'call x(); reply "y"', 'kind': 'tool'},
           {'user': 'b', 'trace': 'reply "z"', 'kind': 'qa'}]
    block = X.build_exemplar_block(sel)
    check('has header', X.BLOCK_HEADER in block, block[:40])
    check('has User: lines',
          'User: "a" -> call x()' in block and 'User: "b" -> reply' in block, block)
    check('empty exemplars -> empty block', X.build_exemplar_block([]) == '')
    tight = X.build_exemplar_block(sel, max_chars=len(X.BLOCK_HEADER) + 30)
    check('char budget drops overflow but keeps >=1',
          tight and tight.count('\n') == 1, repr(tight))


def test_promote_and_store():
    print('[exemplars: promote round-trip + load/save]')
    rows = [_traj('tag walls', [('tag_all_walls', '{}')], final='done'),
            _qa('what is LOD', 'Level of Development')]
    res = X.promote_from_dataset(iter_examples=lambda: rows)
    check('promote ok, count 2', res.get('status') == 'ok'
          and res.get('count') == 2, res)
    loaded = X.load_exemplars()
    check('load reads what promote saved', len(loaded) == 2, loaded)
    check('block from stored file non-empty',
          X.build_exemplar_block(loaded) != '')
    # reset the store to empty so the test leaves no artifact behind
    X.save_exemplars([])
    check('save empty resets store', X.load_exemplars() == [])


def test_never_raises():
    print('[exemplars: failure contained]')
    def boom():
        raise IOError('disk gone')
    res = X.promote_from_dataset(iter_examples=boom)
    check('bad iterator contained', 'status' in res and res['count'] == 0, res)
    check('load on junk never raises', isinstance(X.load_exemplars(), list))


def test_prompt_injection():
    print('[exemplars: injected into local prompt only]')
    X.save_exemplars([{'user': 'tag walls', 'trace': 'call tag_all_walls(); reply "ok"',
                       'kind': 'tool'}])
    try:
        from Intelligence.agent_loop import build_agent_system_prompt
        local = build_agent_system_prompt(local=True)
        cloud = build_agent_system_prompt(local=False)
        check('local prompt contains the block', X.BLOCK_HEADER in local)
        check('cloud prompt does NOT contain the block', X.BLOCK_HEADER not in cloud)
        from Intelligence.agents.specialists import build_specialist_prompt, get_spec
        sp = build_specialist_prompt(get_spec('revit_data'), local=True)
        check('specialist local has block exactly once',
              sp.count(X.BLOCK_HEADER) == 1, sp.count(X.BLOCK_HEADER))
    finally:
        X.save_exemplars([])


if __name__ == '__main__':
    test_converters()
    test_select()
    test_block()
    test_promote_and_store()
    test_never_raises()
    test_prompt_injection()
    print('')
    if FAILURES:
        print('FAILED ({}): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        sys.exit(1)
    print('All exemplar tests passed.')
