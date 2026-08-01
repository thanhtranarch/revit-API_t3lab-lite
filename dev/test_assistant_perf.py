# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the T3Lab Assistant's performance work.

Run:  python3 dev/test_assistant_perf.py
Exit code 0 = all pass. No external test framework — plain asserts, mirroring
dev/test_assistant_routing.py / dev/test_knowledge.py conventions.

Covers the machinery that makes a turn fast, none of which needs Revit:

    telemetry   turn timing, token accumulation, cache-hit accounting, and the
                JSONL sink (the only way to tell whether a latency change
                actually worked)
    batching    which tool calls may share ONE crossing to Revit's main thread
                — the guard that keeps a read from being hoisted above a write

The prompt-cache half of the same work is asserted in
dev/test_assistant_routing.py ("prompt cache" group), next to the other
prompt-shape tests.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'T3Lab.extension', 'lib'))

FAILURES = []


def check(label, ok, detail=None):
    line = '  {}  {}'.format('ok  ' if ok else 'FAIL', label)
    if not ok and detail is not None:
        line += '   {!r}'.format(detail)
    print(line)
    if not ok:
        FAILURES.append(label)


# ─── telemetry ────────────────────────────────────────────────────────────────

def test_turn_timer():
    print('[telemetry: TurnTimer]')
    from Intelligence.telemetry import TurnTimer

    t = TurnTimer(provider='claude', model='m', now=100.0)
    check('nothing streamed yet = no TTFT', t.ttft_ms is None)
    check('unfinished turn has no total', t.total_ms is None)

    t.mark_first_token(now=100.42)
    t.mark_first_token(now=102.0)          # must be ignored
    check('TTFT measured from send', t.ttft_ms == 420)
    check('first token wins', t.ttft_ms == 420)

    t.finish('done', now=103.5)
    check('total measured', t.total_ms == 3500)
    check('status recorded', t.status == 'done')


def test_usage_accumulates_across_iterations():
    print('[telemetry: token usage]')
    from Intelligence.telemetry import TurnTimer

    t = TurnTimer()
    check('no usage = unknown, not 0%', t.cache_hit_ratio is None)

    # An agent turn makes one call per iteration; the counters are sums.
    t.add_usage({'input_tokens': 200, 'output_tokens': 50,
                 'cache_read_input_tokens': 4800,
                 'cache_creation_input_tokens': 5000})
    t.add_usage({'input_tokens': 100, 'output_tokens': 30,
                 'cache_read_input_tokens': 5000})
    check('input tokens summed', t.input_tokens == 300)
    check('output tokens summed', t.output_tokens == 80)
    check('cache reads summed', t.cache_read_tokens == 9800)
    check('cache writes summed', t.cache_creation_tokens == 5000)
    check('hit ratio over the whole prefix',
          abs(t.cache_hit_ratio - (9800.0 / 10100.0)) < 1e-9)

    t.add_usage(None)
    t.add_usage('not a dict')
    t.add_usage({'input_tokens': 'garbage'})
    check('junk usage never raises or corrupts', t.input_tokens == 300)


def test_roundtrips_are_counted_apart_from_calls():
    """The point of batching: fewer crossings than calls. If these two ever
    collapse into one number the win becomes invisible."""
    print('[telemetry: round-trips]')
    from Intelligence.telemetry import TurnTimer

    t = TurnTimer()
    t.tool_roundtrip(3)     # one batched read run
    t.tool_roundtrip(1)     # one write, on its own
    check('crossings counted', t.roundtrips == 2)
    check('calls counted', t.tool_runs == 4)


def test_record_roundtrips_through_jsonl():
    print('[telemetry: JSONL sink]')
    from Intelligence.telemetry import TurnTimer, record, read_turns

    t = TurnTimer(provider='claude', model='m', specialist='revit_data',
                  skills=['qa-checklist'], now=100.0)
    t.mark_first_token(now=100.25)
    t.add_usage({'input_tokens': 10, 'cache_read_input_tokens': 90})
    t.tool_roundtrip(2)
    t.finish('done', now=101.0)

    path = os.path.join(tempfile.mkdtemp(), 'turns.jsonl')
    check('record writes', record(t, path=path) is True)
    check('second turn appends', record(t, path=path) is True)

    rows = read_turns(path)
    check('both turns read back', len(rows) == 2)
    row = rows[0]
    check('ttft persisted', row['ttft_ms'] == 250)
    check('total persisted', row['total_ms'] == 1000)
    check('specialist persisted', row['specialist'] == 'revit_data')
    check('skills persisted', row['skills'] == ['qa-checklist'])
    check('cache reads persisted', row['cache_read'] == 90)

    with open(path, 'a', encoding='utf-8') as f:
        f.write('{ this is not json\n\n')
    check('malformed lines skipped, good ones kept',
          len(read_turns(path)) == 2)
    check('missing file returns empty', read_turns(path + '.nope') == [])


def test_record_survives_vietnamese():
    """Every JSON writer here serializes to ASCII first: json.dump onto a
    byte-mode file raises mid-write under IronPython 2.7 as soon as the
    payload carries Vietnamese, truncating the file."""
    print('[telemetry: unicode safety]')
    from Intelligence.telemetry import TurnTimer, record, read_turns

    t = TurnTimer(provider='ollama', model='qwen', now=1.0)
    t.specialist = 'revit_action'
    t.skills = ['tô màu tường', 'kiểm tra tiêu chuẩn']
    t.finish('done', now=2.0)

    path = os.path.join(tempfile.mkdtemp(), 'vi.jsonl')
    check('vietnamese turn written', record(t, path=path) is True)
    rows = read_turns(path)
    check('vietnamese survives the round-trip',
          rows and rows[0]['skills'] == ['tô màu tường', 'kiểm tra tiêu chuẩn'])

    check('unwritable path fails quietly',
          record(t, path=os.path.join(path, 'nested', 'x.jsonl')) is False)


# ─── read batching ────────────────────────────────────────────────────────────

_WRITES = frozenset(['set_parameter', 'delete_element', 'revit_override_color'])


def _loop(batch_fn=None, with_seam=True):
    from Intelligence.agent_loop import AgentLoop
    if not with_seam:
        return AgentLoop(None, None, None)

    def _batch(calls):
        return [{'ok': name} for name, _ in calls]

    return AgentLoop(None, None, None,
                     execute_tools_batch=(batch_fn or _batch),
                     is_write_tool=lambda n: n in _WRITES)


def _calls(*names):
    return [{'name': n, 'args': {}} for n in names]


def test_only_leading_reads_are_batched():
    print('[batching: partition]')
    loop = _loop()

    check('a run of reads batches',
          loop.leading_read_run(_calls('list_levels', 'ai_element_filter',
                                       'revit_get_selected_elements'))
          == [0, 1, 2])
    check('a lone read is not worth a batch',
          loop.leading_read_run(_calls('list_levels')) == [])
    check('a write ends the run',
          loop.leading_read_run(_calls('list_levels', 'set_parameter',
                                       'ai_element_filter')) == [])
    check('a leading write batches nothing',
          loop.leading_read_run(_calls('set_parameter', 'list_levels',
                                       'ai_element_filter')) == [])
    check('reads after a write are never hoisted',
          loop.leading_read_run(_calls('list_levels', 'ai_element_filter',
                                       'delete_element', 'list_levels'))
          == [0, 1])


def test_special_tools_never_join_a_batch():
    """open_t3lab_tool is terminal and remember_fact runs locally; neither
    goes through the Revit server at all."""
    print('[batching: special tools]')
    loop = _loop()
    check('launcher ends the run',
          loop.leading_read_run(_calls('list_levels', 'open_t3lab_tool',
                                       'ai_element_filter')) == [])
    check('memory tool ends the run',
          loop.leading_read_run(_calls('list_levels', 'remember_fact',
                                       'ai_element_filter')) == [])
    check('an unnamed call ends the run',
          loop.leading_read_run([{'name': '', 'args': {}}] + _calls('a', 'b'))
          == [])


def test_duplicates_are_left_to_the_existing_guard():
    print('[batching: duplicate guard]')
    from Intelligence.agent_loop import _call_key
    loop = _loop()
    done = set([_call_key('ai_element_filter', {})])
    check('a duplicate ends the run',
          loop.leading_read_run(
              _calls('list_levels', 'ai_element_filter', 'list_levels'),
              done_calls=done) == [])
    check('the batch uses the guard key verbatim',
          _call_key('x', {'b': 1, 'a': 2}) == _call_key('x', {'a': 2, 'b': 1}))


def test_batching_is_opt_in_and_fails_safe():
    print('[batching: fallbacks]')
    check('no seam supplied = sequential as before',
          _loop(with_seam=False).leading_read_run(_calls('a', 'b', 'c')) == [])

    loop = _loop()
    check('prefetch returns per-index results',
          loop._prefetch_reads(_calls('a', 'b'), set())
          == {0: {'ok': 'a'}, 1: {'ok': 'b'}})

    def _boom(_calls_):
        raise RuntimeError('Revit went away')
    check('a failing batch falls back to sequential',
          _loop(batch_fn=_boom)._prefetch_reads(_calls('a', 'b'), set()) == {})

    def _short(_calls_):
        return ['only one']
    check('a short result list falls back to sequential',
          _loop(batch_fn=_short)._prefetch_reads(_calls('a', 'b'), set()) == {})

    def _wrong_type(_calls_):
        return {'not': 'a list'}
    check('a malformed result falls back to sequential',
          _loop(batch_fn=_wrong_type)._prefetch_reads(_calls('a', 'b'), set())
          == {})


# ─── classification round-trip ────────────────────────────────────────────────

class _CountingProvider(object):
    """Stands in for a real provider; records whether it was called at all."""

    def __init__(self, label='general'):
        self.calls = []
        self._label = label

    def pick_fast_model(self):
        return 'fast-model'

    def chat(self, messages, system, user, **kwargs):
        self.calls.append(kwargs)
        return '{"label": "%s", "skill": null}' % self._label


def test_a_strong_skill_match_skips_the_classify_call():
    """The classification call sits IN FRONT of the real turn, so every one
    that can be avoided is latency the user does not pay."""
    print('[classify: skip on strong skill]')
    from Intelligence.agents.dispatcher import AgentDispatcher
    from Intelligence.skills_engine import get_skills_engine

    engine = get_skills_engine()
    engine.scan()
    disp = AgentDispatcher()

    prov = _CountingProvider()
    dec = disp.classify('shared coordinates', provider=prov,
                        skills_engine=engine)
    check('decided from the skill itself', dec['source'] == 'skill', dec)
    check('the strong skill is carried', dec['skill'] == 'shared-coordinates')
    check('specialist comes from the frontmatter',
          dec['specialist'] in ('general', 'revit_data', 'revit_action',
                                'knowledge', 'comment', 'multi_doc',
                                'modeling', 'qa_check', 'export'), dec)
    check('no model round-trip was paid', prov.calls == [])


def test_a_weak_match_still_asks_the_model():
    print('[classify: weak match]')
    from Intelligence.agents.dispatcher import AgentDispatcher
    from Intelligence.skills_engine import get_skills_engine

    engine = get_skills_engine()
    engine.scan()
    disp = AgentDispatcher()

    prov = _CountingProvider()
    disp.classify('cobie handover', provider=prov, skills_engine=engine)
    check('a weak skill hit does not short-circuit', len(prov.calls) == 1)

    prov2 = _CountingProvider()
    disp.classify('thời tiết hôm nay thế nào', provider=prov2,
                  skills_engine=engine)
    check('an unmatched message still asks', len(prov2.calls) == 1)


def test_classify_call_is_time_bounded():
    """Without an explicit bound this inherited each provider's chat timeout:
    60s on the cloud APIs, 180s on Ollama / LM Studio."""
    print('[classify: timeout]')
    from Intelligence.agents.dispatcher import (AgentDispatcher,
                                                _CLASSIFY_TIMEOUT_MS)
    from Intelligence.skills_engine import get_skills_engine

    check('bound is short enough to be worth paying',
          0 < _CLASSIFY_TIMEOUT_MS <= 10000, _CLASSIFY_TIMEOUT_MS)

    engine = get_skills_engine()
    engine.scan()
    prov = _CountingProvider()
    AgentDispatcher().classify('thời tiết hôm nay thế nào', provider=prov,
                               skills_engine=engine)
    check('the call carries the bound',
          prov.calls and prov.calls[0].get('timeout_ms') == _CLASSIFY_TIMEOUT_MS,
          prov.calls)
    check('and still pins the fast model',
          prov.calls and prov.calls[0].get('model_override') == 'fast-model')


def test_providers_accept_a_timeout_override():
    """Every chat() takes **kwargs, but the timeout only helps if the provider
    actually forwards it to http_post."""
    print('[classify: provider plumbing]')
    import re as _re
    for name, default in (('claude_provider', '60000'),
                          ('openai_provider', '60000'),
                          ('deepseek_provider', '60000'),
                          ('ollama_provider', '180000'),
                          ('lmstudio_provider', '180000')):
        path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'T3Lab.extension', 'lib', 'Intelligence', name + '.py')
        with open(path, encoding='utf-8') as f:
            src = f.read()
        # the chat() body must read timeout_ms from kwargs, falling back to
        # the value that provider used before.
        pat = r'kwargs\.get\("timeout_ms"\)\s*\n?\s*or\s*' + default
        check('{} forwards timeout_ms (default {})'.format(name, default),
              bool(_re.search(pat, src)))


# ─── conversation window + rolling summary ────────────────────────────────────

def _long_history(n=30):
    turns = []
    for i in range(n):
        turns.append({'role': 'user',
                      'content': 'yêu cầu %d: đổi tên sheet A-%03d' % (i, i)})
        turns.append({'role': 'assistant',
                      'content': ('Đã đổi tên 12 sheet, ids [%d]' % i)
                      if i % 2 else 'Đang kiểm tra…'})
    return turns


def test_history_window_is_one_number():
    """script._add_to_history truncated to 16 while _sanitize_history took the
    last 24, so the larger limit never bound and the continuity it documented
    could not happen."""
    print('[conversation: window]')
    from Intelligence.conversation import HISTORY_LIMIT, trim_history
    from Intelligence.agent_loop import _sanitize_history

    hist = _long_history()
    kept, _ = trim_history(hist, '')
    check('window enforced', len(kept) == HISTORY_LIMIT)
    check('newest turns are the ones kept', kept[-1] == hist[-1])
    check('sanitize uses the same number',
          len(_sanitize_history(hist)) == HISTORY_LIMIT)

    short = hist[:10]
    check('under the limit nothing is touched',
          trim_history(short, 'keep me') == (short, 'keep me'))


def test_dropped_turns_are_folded_not_lost():
    print('[conversation: rolling summary]')
    from Intelligence.conversation import trim_history, MAX_SUMMARY_CHARS

    kept, summary = trim_history(_long_history(), '')
    check('a summary was produced', bool(summary))
    check('the start of the session survives', 'A-000' in summary, summary[:80])
    check('it is not in the live window',
          not any('A-000' in (t.get('content') or '') for t in kept))
    check('concrete results are carried', 'ids [1]' in summary)
    check('narration is dropped', 'Đang kiểm tra' not in summary)
    check('budget respected', len(summary) <= MAX_SUMMARY_CHARS)


def test_summary_keeps_the_recent_end_when_it_overflows():
    print('[conversation: summary budget]')
    from Intelligence.conversation import fold_summary, MAX_SUMMARY_CHARS

    huge = [{'role': 'user', 'content': 'điều kiện số %d phải giữ' % i}
            for i in range(400)]
    summary = fold_summary('', huge)
    check('bounded', len(summary) <= MAX_SUMMARY_CHARS)
    check('the newest folded turn survives', 'số 399' in summary)
    check('the oldest is what gets dropped', 'số 0 ' not in summary)

    grown = fold_summary(summary, [{'role': 'user', 'content': 'ràng buộc mới'}])
    check('folding again keeps the newest', 'ràng buộc mới' in grown)
    check('still bounded', len(grown) <= MAX_SUMMARY_CHARS)


def test_summary_block_is_labelled_and_bilingual():
    print('[conversation: prompt block]')
    from Intelligence.conversation import summary_block

    check('no summary = no block', summary_block('') == '')
    check('whitespace only = no block', summary_block('   \n  ') == '')

    en = summary_block('User asked: rename sheets')
    check('english heading', 'Earlier in this conversation' in en)
    check('marked as background, not a to-do', 'NOT tasks to redo' in en)
    check('content carried', 'rename sheets' in en)

    vi = summary_block('Người dùng hỏi: đổi tên sheet', viet=True)
    check('vietnamese heading', 'Trước đó trong hội thoại' in vi)
    check('vietnamese note', 'KHÔNG phải việc cần làm lại' in vi)


def test_summary_handles_every_content_shape():
    print('[conversation: content shapes]')
    from Intelligence.conversation import summarize_turns

    lines = summarize_turns([
        {'role': 'user', 'content': [{'type': 'text', 'text': 'xem sheet A-101'},
                                     {'type': 'image', 'source': {}}]},
        {'role': 'assistant', 'content': 'Đã tìm thấy 3 sheet'},
        {'role': 'system', 'content': 'ignored'},
        {'role': 'user', 'content': ''},
        'not a dict',
        {'role': 'assistant', 'content': None},
    ])
    check('vision turn text is mined',
          any('A-101' in l for l in lines), lines)
    check('concrete assistant line kept',
          any('3 sheet' in l for l in lines), lines)
    check('system turns ignored', not any('ignored' in l for l in lines))
    check('empty and malformed turns are skipped', len(lines) == 2, lines)


def main():
    print('')
    for fn in (test_turn_timer,
               test_usage_accumulates_across_iterations,
               test_roundtrips_are_counted_apart_from_calls,
               test_record_roundtrips_through_jsonl,
               test_record_survives_vietnamese,
               test_only_leading_reads_are_batched,
               test_special_tools_never_join_a_batch,
               test_duplicates_are_left_to_the_existing_guard,
               test_batching_is_opt_in_and_fails_safe,
               test_a_strong_skill_match_skips_the_classify_call,
               test_a_weak_match_still_asks_the_model,
               test_classify_call_is_time_bounded,
               test_providers_accept_a_timeout_override,
               test_history_window_is_one_number,
               test_dropped_turns_are_folded_not_lost,
               test_summary_keeps_the_recent_end_when_it_overflows,
               test_summary_block_is_labelled_and_bilingual,
               test_summary_handles_every_content_shape):
        fn()
        print('')

    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All assistant performance tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
