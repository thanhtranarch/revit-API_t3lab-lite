# -*- coding: utf-8 -*-
"""
CPython 3 tests for "Opus teaches Qwen via MCP" (Part E).

Run:  python3 dev/test_mcp_teacher.py
Exit 0 = all pass. Plain asserts, no Revit/WPF needed.

Covers the pure/testable seams of the teaching-capture pipeline:
  * core.teaching — sandbox document match, sandbox write-guard decision,
    error-result detection, raw-session round-trip.
  * dataset.py — agentic schema (role 'tool' + assistant tool_calls survive,
    dedup by tool name/args, export_sft keeps tool turns) and add_trajectory.
  * mcp_session_miner — labeled -> dataset + .done, unlabeled -> .nolabel,
    empty -> .done, idempotent, never raises.
The live server + Revit + Claude Desktop path is on the in-Revit checklist.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_mcpteach_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from core import teaching as T
from Intelligence.learning import dataset as ds
from Intelligence.learning.enrichers import mcp_session_miner as M

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─── core.teaching pure helpers ──────────────────────────────────────────────────

def test_sandbox_match():
    print('[teaching: sandbox document match]')
    check('explicit path wins',
          T.is_sandbox('P', 'C:/x/sb.rvt', {'path': 'C:/x/sb.rvt'}) is True)
    check('explicit title wins',
          T.is_sandbox('Scratch', '', {'title': 'Scratch'}) is True)
    check('explicit mismatch rejects',
          T.is_sandbox('Real', 'C:/x/real.rvt', {'path': 'C:/x/sb.rvt'}) is False)
    check('no mark + name hint accepts',
          T.is_sandbox('My Sandbox', '', None) is True)
    check('no mark + real name rejects',
          T.is_sandbox('Tower A', 'C:/proj/towerA.rvt', None) is False)


def test_guard_decision():
    print('[teaching: sandbox write-guard decision]')
    check('block write outside sandbox during teaching',
          T.should_block_write(True, True, False) is True)
    check('allow read', T.should_block_write(True, False, False) is False)
    check('allow sandbox write', T.should_block_write(True, True, True) is False)
    check('teaching off never blocks',
          T.should_block_write(False, True, False) is False)


def test_error_detection():
    print('[teaching: tool-result error detection]')
    check('error key', T.is_error_result({'error': 'x'}) is True)
    check('isError flag', T.is_error_result({'isError': True}) is True)
    check('success dict', bool(T.is_error_result({'success': True})) is False)
    check('non-dict', bool(T.is_error_result('ok')) is False)


def test_session_roundtrip():
    print('[teaching: raw session round-trip]')
    d = T.session_dir()
    p = os.path.join(d, 'rt.jsonl')
    T.append_step_line(p, u'liệt kê view', {'tool': 'revit_list_views',
                                            'arguments': {}, 'result': {'count': 5},
                                            'is_error': False})
    T.append_step_line(p, u'liệt kê view', {'tool': 'revit_get_active_view',
                                            'arguments': {}, 'result': {'name': 'L1'},
                                            'is_error': False})
    goal, steps = T.read_session(p)
    check('goal recovered', goal == u'liệt kê view', goal)
    check('two steps recovered, ordered',
          len(steps) == 2 and steps[0]['tool'] == 'revit_list_views', steps)


# ─── dataset agentic schema ──────────────────────────────────────────────────────

def test_agentic_schema():
    print('[dataset: agentic schema round-trip]')
    steps = [
        {'tool': 'revit_get_selected_elements', 'arguments': {},
         'result': {'selected_count': 3}, 'is_error': False},
        {'tool': 'revit_override_color',
         'arguments': {'category': 'Walls', 'color': 'red'},
         'result': {'success': True}, 'is_error': False},
    ]
    ok, note = ds.add_trajectory(u'tô đỏ tường đang chọn', steps,
                                 final=u'Đã tô đỏ 3 tường.', source='mcp_teacher')
    check('add_trajectory stored', ok and note == u'added', note)

    row = [r for r in ds.iter_examples()
           if r['meta']['source'] == 'mcp_teacher'][0]
    roles = [m['role'] for m in row['messages']]
    check('roles interleave user/assistant/tool',
          roles == ['user', 'assistant', 'tool', 'assistant', 'tool', 'assistant'],
          roles)
    asst = [m for m in row['messages']
            if m['role'] == 'assistant' and m.get('tool_calls')]
    check('assistant tool_calls preserved with name+arguments',
          asst[0]['tool_calls'][0]['name'] == 'revit_get_selected_elements'
          and 'arguments' in asst[0]['tool_calls'][0], asst[0]['tool_calls'])
    tool_turn = [m for m in row['messages'] if m['role'] == 'tool'][0]
    check('tool turn keeps id + name + content',
          tool_turn.get('tool_call_id') and tool_turn.get('name')
          and tool_turn.get('content'), tool_turn)

    # Dedup on identical trajectory.
    ok2, note2 = ds.add_trajectory(u'tô đỏ tường đang chọn', steps,
                                   final=u'Đã tô đỏ 3 tường.', source='mcp_teacher')
    check('identical trajectory dedups', ok2 and note2 == u'duplicate', note2)

    # Different tool args -> different hash -> stored.
    steps3 = list(steps)
    steps3[1] = {'tool': 'revit_override_color',
                 'arguments': {'category': 'Floors', 'color': 'red'},
                 'result': {'success': True}, 'is_error': False}
    ok3, note3 = ds.add_trajectory(u'tô đỏ tường đang chọn', steps3,
                                   final=u'x', source='mcp_teacher')
    check('different tool args are NOT a duplicate', ok3 and note3 == u'added', note3)

    # export_sft keeps tool turns.
    out = os.path.join(tempfile.mkdtemp(), 'sft.jsonl')
    n = ds.export_sft(out)
    import json
    exported = [json.loads(l) for l in open(out, encoding='utf-8')]
    traj = [e for e in exported
            if any(m['role'] == 'tool' for m in e['messages'])]
    check('export_sft keeps trajectories with tool turns', len(traj) >= 1, n)


def test_backward_compat_text():
    print('[dataset: plain text example still valid]')
    ok, note = ds.add_qa(u'batchout là gì', u'BatchOut xuất sheets hàng loạt.',
                         source='telemetry_miner')
    check('flat Q&A still accepted', ok and note in (u'added', u'duplicate'), note)
    # A tool-result turn with NO user/assistant is rejected (needs a target).
    bad = ds._clean_messages([{'role': 'tool', 'content': '{}'}])
    check('lone tool turn rejected', bad is None)


# ─── session miner ───────────────────────────────────────────────────────────────

def test_miner():
    print('[mcp_session_miner: ingest sessions -> dataset]')
    d = T.session_dir()
    # labeled
    p1 = os.path.join(d, 'm1.jsonl')
    T.append_step_line(p1, u'tag all walls', {'tool': 'tag_all_walls',
                                              'arguments': {'view': 'L1'},
                                              'result': {'tagged': 12},
                                              'is_error': False})
    # unlabeled (no goal)
    p2 = os.path.join(d, 'm2.jsonl')
    T.append_step_line(p2, u'', {'tool': 'revit_get_project_info',
                                 'arguments': {}, 'result': {'n': 'X'},
                                 'is_error': False})
    # empty
    p3 = os.path.join(d, 'm3.jsonl')
    open(p3, 'w').close()

    res = M.run()
    check('miner added the labeled session', res.get('added') >= 1, res)
    names = set(os.listdir(d))
    check('labeled -> .done', 'm1.jsonl.done' in names, names)
    check('unlabeled -> .nolabel (deferred)', 'm2.jsonl.nolabel' in names, names)
    check('empty -> .done', 'm3.jsonl.done' in names, names)

    res2 = M.run()
    check('rerun is idempotent (nothing new)', res2.get('added') == 0, res2)


def test_miner_never_raises():
    print('[mcp_session_miner: failure contained]')
    def boom(*a, **k):
        raise IOError('disk gone')
    res = M.run(session_dir='/does/not/exist', read_session=boom,
                add_trajectory=boom, rename=boom)
    check('bad I/O contained into a status dict', 'status' in res, res)


if __name__ == '__main__':
    test_sandbox_match()
    test_guard_decision()
    test_error_detection()
    test_session_roundtrip()
    test_agentic_schema()
    test_backward_compat_text()
    test_miner()
    test_miner_never_raises()
    print('')
    if FAILURES:
        print('FAILED ({}): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        sys.exit(1)
    print('All mcp-teacher tests passed.')
