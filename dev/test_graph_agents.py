# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the graph agent layer (Intelligence/graph).

Run:  python3 dev/test_graph_agents.py
Exit code 0 = all pass. Plain asserts, no external framework — same conventions
as dev/test_assistant_routing.py and dev/test_knowledge.py.

What these lock down, in the order the layer runs:

    topology    layers() is the scheduler. If it returns the wrong grouping,
                nodes run before their inputs exist and every downstream
                guarantee is void.
    executor    the four properties the module promises: writer
                serialisation, fault containment, redundant paths,
                cooperative cancel. Each is tested by making it fail.
    verifier    the failures that raise no exception — a refusal, a counting
                question answered with no number — must still score as bad.
    reducer     a failed step has to stay VISIBLE in the summary. A reducer
                that quietly drops it is worse than no reducer.
    planner     the split has to be conservative: "tường và cột" is one goal,
                "thống kê tường rồi xuất pdf" is two.
    orchestrator end-to-end, including the adapt loop.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile
import threading
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

# Sandbox %APPDATA% BEFORE anything can write telemetry into the real one.
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_graph_test_')

from Intelligence.graph import topology                       # noqa: E402
from Intelligence.graph.executor import GraphExecutor          # noqa: E402
from Intelligence.graph.observability import GraphTrace, record, read_traces  # noqa: E402
from Intelligence.graph.planner import Planner                 # noqa: E402
from Intelligence.graph.primitives import (                    # noqa: E402
    AGENT, CANCELLED, DONE, FAILED, FALLBACK, REDUCER, SKIPPED,
    GraphState, Node, NodeResult,
)
from Intelligence.graph.reducer import Reducer                 # noqa: E402
from Intelligence.graph.router import GraphRouter              # noqa: E402
from Intelligence.graph.topology import GraphCycleError, TaskGraph  # noqa: E402
from Intelligence.graph.verifier import Verifier               # noqa: E402
from Intelligence.graph.orchestrator import GraphOrchestrator  # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def _n(nid, goal=u'goal', **kw):
    return Node(nid, goal, **kw)


# ─────────────────────────────────────────────────────────────────────────────
# Topology
# ─────────────────────────────────────────────────────────────────────────────

def test_linear_layers_are_one_node_each():
    g = topology.linear([_n('a'), _n('b'), _n('c')])
    check('linear: three layers', g.layers() == [['a'], ['b'], ['c']],
          g.layers())
    check('linear: pattern recorded', g.pattern == 'linear')


def test_fan_out_puts_branches_in_one_layer():
    g = topology.fan_out(_n('root'), [_n('b1'), _n('b2'), _n('b3')])
    layers = g.layers()
    check('fan_out: root alone first', layers[0] == ['root'], layers)
    check('fan_out: branches share a layer',
          sorted(layers[1]) == ['b1', 'b2', 'b3'], layers)


def test_fan_in_converges():
    g = topology.fan_in([_n('a'), _n('b')], _n('sink'))
    layers = g.layers()
    check('fan_in: sources first', sorted(layers[0]) == ['a', 'b'], layers)
    check('fan_in: sink last', layers[-1] == ['sink'], layers)


def test_diamond_is_three_layers():
    g = topology.diamond(_n('plan'), [_n('x'), _n('y')], _n('reduce'))
    layers = g.layers()
    check('diamond: 3 layers', len(layers) == 3, layers)
    check('diamond: middle runs together',
          sorted(layers[1]) == ['x', 'y'], layers)


def test_hub_spoke_feedback_is_not_scheduled():
    g = topology.hub_spoke(_n('hub'), [_n('s1'), _n('s2')])
    layers = g.layers()
    # A real return edge would be a cycle; it is a feedback edge, so the graph
    # still schedules.
    check('hub_spoke: schedules', layers[0] == ['hub'], layers)
    check('hub_spoke: spokes parallel', sorted(layers[1]) == ['s1', 's2'],
          layers)


def test_loop_declares_feedback_without_a_cycle():
    g = topology.loop([_n('a'), _n('b')])
    check('loop: still schedulable', g.layers() == [['a'], ['b']], g.layers())
    check('loop: feedback edge present', len(g.edges('feedback')) == 1)


def test_cycle_is_detected():
    g = TaskGraph()
    for nid in ('a', 'b', 'c'):
        g.add_node(_n(nid))
    g.add_edge('a', 'b')
    g.add_edge('b', 'c')
    g.add_edge('c', 'a')
    try:
        g.layers()
        check('cycle: raises', False, 'no exception')
    except GraphCycleError:
        check('cycle: raises', True)
    check('cycle: validate reports it', len(g.validate()) >= 1, g.validate())


def test_fallback_target_is_not_scheduled_normally():
    g = TaskGraph()
    g.add_node(_n('a'))
    g.add_node(_n('a_alt'))
    g.add_node(_n('sink'))
    g.add_edge('a', 'a_alt', kind=FALLBACK)
    g.add_edge('a', 'sink')
    g.add_edge('a_alt', 'sink')
    layers = g.layers()
    flat = [nid for layer in layers for nid in layer]
    check('fallback: dormant node not scheduled', 'a_alt' not in flat, layers)
    check('fallback: sink does not wait on it',
          layers == [['a'], ['sink']], layers)


def test_duplicate_node_id_is_rejected():
    g = TaskGraph()
    g.add_node(_n('a'))
    try:
        g.add_node(_n('a'))
        check('duplicate id rejected', False, 'accepted')
    except ValueError:
        check('duplicate id rejected', True)


# ─────────────────────────────────────────────────────────────────────────────
# Executor
# ─────────────────────────────────────────────────────────────────────────────

def test_executor_runs_in_dependency_order():
    order = []

    def runner(node, context, state):
        order.append(node.id)
        return u'result of {}'.format(node.id)

    g = topology.linear([_n('a'), _n('b'), _n('c')])
    state = GraphState()
    GraphExecutor(runner, max_parallel=1).run(g, state)
    check('executor: order respected', order == ['a', 'b', 'c'], order)
    check('executor: all done',
          all(state.succeeded(n) for n in ('a', 'b', 'c')), state.counts())


def test_executor_actually_runs_a_layer_in_parallel():
    """Three 150 ms nodes must finish in well under 450 ms."""
    barrier_hits = []
    lock = threading.Lock()

    def runner(node, context, state):
        with lock:
            barrier_hits.append(node.id)
        time.sleep(0.15)
        return u'ok'

    g = topology.fan_out(_n('root'), [_n('b1'), _n('b2'), _n('b3')])
    state = GraphState()
    start = time.time()
    GraphExecutor(runner, max_parallel=3).run(g, state)
    elapsed = time.time() - start
    # root (150ms) + branches in parallel (~150ms) ≈ 300ms; sequential ≈ 600ms.
    check('executor: layer ran in parallel', elapsed < 0.50,
          '{:.2f}s'.format(elapsed))
    check('executor: every branch ran', len(barrier_hits) == 4, barrier_hits)


def test_two_writers_never_overlap():
    concurrent = {'now': 0, 'max': 0}
    lock = threading.Lock()

    def runner(node, context, state):
        with lock:
            concurrent['now'] += 1
            concurrent['max'] = max(concurrent['max'], concurrent['now'])
        time.sleep(0.05)
        with lock:
            concurrent['now'] -= 1
        return u'ok'

    g = topology.fan_out(_n('root'),
                         [_n('w1', writer=True), _n('w2', writer=True),
                          _n('w3', writer=True)])
    GraphExecutor(runner, max_parallel=3).run(g, GraphState())
    check('executor: writers serialised', concurrent['max'] == 1,
          'max concurrent = {}'.format(concurrent['max']))


def test_failure_is_contained_and_dependents_are_skipped():
    def runner(node, context, state):
        if node.id == 'b':
            raise RuntimeError('boom')
        return u'ok'

    g = TaskGraph()
    for nid in ('a', 'b', 'c', 'z'):
        g.add_node(_n(nid))
    g.add_edge('a', 'b')
    g.add_edge('b', 'c')          # c depends on the failing node
    state = GraphState()
    GraphExecutor(runner, max_parallel=1).run(g, state)
    check('containment: failing node failed', state.status_of('b') == FAILED,
          state.status_of('b'))
    check('containment: dependent skipped', state.status_of('c') == SKIPPED,
          state.status_of('c'))
    check('containment: unrelated branch still ran', state.succeeded('z'),
          state.status_of('z'))


def test_retry_uses_max_attempts():
    calls = {'n': 0}

    def runner(node, context, state):
        calls['n'] += 1
        if calls['n'] < 3:
            raise RuntimeError('flaky')
        return u'third time lucky'

    g = topology.linear([_n('a', max_attempts=3)])
    state = GraphState()
    GraphExecutor(runner, max_parallel=1).run(g, state)
    check('retry: succeeded on attempt 3', state.succeeded('a'),
          state.status_of('a'))
    check('retry: attempts recorded',
          state.get_result('a').attempts == 3,
          state.get_result('a').attempts)


def test_fallback_path_recovers_the_original_node():
    def runner(node, context, state):
        if node.id == 'a':
            raise RuntimeError('primary route failed')
        return u'answer from the fallback route'

    g = TaskGraph()
    g.add_node(_n('a'))
    g.add_node(_n('a_alt'))
    g.add_node(_n('sink'))
    g.add_edge('a', 'a_alt', kind=FALLBACK)
    g.add_edge('a', 'sink')
    g.add_edge('a_alt', 'sink')
    state = GraphState()
    GraphExecutor(runner, max_parallel=1).run(g, state)
    check('fallback: original recorded as recovered',
          state.status_of('a') == DONE, state.status_of('a'))
    check('fallback: carries the alt answer',
          'fallback route' in state.get_result('a').output,
          state.get_result('a').output)
    check('fallback: recovery attributed',
          state.get_result('a').data.get('recovered_by') == 'a_alt',
          state.get_result('a').data)
    check('fallback: sink still ran', state.succeeded('sink'),
          state.status_of('sink'))


def test_cancel_stops_the_sweep():
    seen = []

    def runner(node, context, state):
        seen.append(node.id)
        state.cancel()          # user pressed Stop during the first node
        return u'ok'

    g = topology.linear([_n('a'), _n('b'), _n('c')])
    state = GraphState()
    GraphExecutor(runner, max_parallel=1).run(g, state)
    check('cancel: later nodes never ran', seen == ['a'], seen)
    check('cancel: they are marked cancelled',
          state.status_of('b') == CANCELLED, state.status_of('b'))


def test_context_carries_dependency_output_forward():
    seen = {}

    def runner(node, context, state):
        seen[node.id] = context
        return u'{} produced 42 walls'.format(node.id)

    g = topology.linear([_n('a', goal=u'count walls'), _n('b', goal=u'report')])
    state = GraphState(global_context=u'GLOBAL STATE HERE')
    GraphExecutor(runner, max_parallel=1).run(g, state)
    check('context: global reaches node a', 'GLOBAL STATE HERE' in seen['a'],
          seen.get('a'))
    check('context: b sees a\'s output', '42 walls' in seen['b'], seen.get('b'))
    check('context: a does not see b', 'produced' not in seen['a'],
          seen.get('a'))


def test_runner_returning_empty_is_a_failure():
    g = topology.linear([_n('a')])
    state = GraphState()
    GraphExecutor(lambda n, c, s: u'   ', max_parallel=1).run(g, state)
    check('empty answer is a failure', state.status_of('a') == FAILED,
          state.status_of('a'))


# ─────────────────────────────────────────────────────────────────────────────
# Verifier
# ─────────────────────────────────────────────────────────────────────────────

def _result(node_id, text):
    res = NodeResult(node_id, status=DONE, output=text)
    res.started = res.finished = time.time()
    return res


def test_verifier_catches_a_refusal():
    v = Verifier()
    node = _n('a', goal=u'đếm số tường trong model')
    good = v.check(node, _result('a', u'Model có 128 bức tường.'))
    bad = v.check(node, _result(
        'a', u'Xin lỗi, tôi không thể truy cập model lúc này.'))
    check('verifier: good answer passes', good.ok, good.to_dict())
    check('verifier: refusal is caught', 'refusal' in bad.reasons,
          bad.to_dict())
    check('verifier: refusal scores lower', bad.score < good.score,
          '{} vs {}'.format(bad.score, good.score))


def test_verifier_wants_a_number_for_a_counting_question():
    v = Verifier()
    node = _n('a', goal=u'có bao nhiêu cửa sổ')
    verdict = v.check(node, _result(
        'a', u'Tôi đã kiểm tra và có khá nhiều cửa sổ trong model của bạn.'))
    check('verifier: no_number flagged', 'no_number' in verdict.reasons,
          verdict.to_dict())


def test_verifier_flags_empty_output():
    v = Verifier()
    verdict = v.check(_n('a'), _result('a', u''))
    check('verifier: empty caught', verdict.reasons == ['empty'],
          verdict.to_dict())
    check('verifier: empty scores zero', verdict.score == 0.0, verdict.score)


def test_verifier_flags_a_leaked_tool_error():
    v = Verifier()
    verdict = v.check(_n('a', goal=u'liệt kê sheet'), _result(
        'a', u'Traceback (most recent call last):\n  File "x.py"\nException: '
             u'no such tool: revit_list_sheetz'))
    check('verifier: tool_error caught', 'tool_error' in verdict.reasons,
          verdict.to_dict())


def test_executor_treats_a_verifier_veto_as_failure():
    """A refusal must take the fallback path, not flow on as a success."""
    def runner(node, context, state):
        if node.id == 'a':
            return u'Xin lỗi, tôi không thể thực hiện việc này.'
        return u'Model có 128 bức tường.'

    g = TaskGraph()
    g.add_node(_n('a', goal=u'đếm tường'))
    g.add_node(_n('a_alt', goal=u'đếm tường'))
    g.add_edge('a', 'a_alt', kind=FALLBACK)
    state = GraphState()
    GraphExecutor(runner, max_parallel=1, verifier=Verifier()).run(g, state)
    check('verifier veto: fallback ran', state.succeeded('a_alt'),
          state.status_of('a_alt'))
    check('verifier veto: node recovered', state.status_of('a') == DONE,
          state.status_of('a'))
    check('verifier veto: recovered answer is the good one',
          '128' in state.get_result('a').output,
          state.get_result('a').output)


# ─────────────────────────────────────────────────────────────────────────────
# Reducer
# ─────────────────────────────────────────────────────────────────────────────

def _two_node_graph():
    g = TaskGraph(goal=u'do two things')
    g.add_node(_n('n1', goal=u'thống kê tường', kind=AGENT))
    g.add_node(_n('n2', goal=u'xuất pdf sheet A', kind=AGENT))
    g.add_node(_n('reduce', goal=u'', kind=REDUCER))
    g.add_edge('n1', 'reduce')
    g.add_edge('n2', 'reduce')
    return g


def test_reducer_keeps_a_failed_step_visible():
    g = _two_node_graph()
    state = GraphState()
    state.set_result(_result('n1', u'Model có 128 bức tường.'))
    failed = NodeResult('n2', status=FAILED, error=u'export timed out')
    failed.started = failed.finished = time.time()
    state.set_result(failed)

    out = Reducer().reduce(g, state, lang='vi')
    check('reducer: keeps the good result', u'128' in out, out)
    check('reducer: names the failed step', u'xuất pdf sheet A' in out, out)
    check('reducer: says it did not complete',
          u'Không hoàn thành' in out, out)


def test_reducer_is_bilingual():
    g = _two_node_graph()
    state = GraphState()
    state.set_result(_result('n1', u'128 walls.'))
    state.set_result(_result('n2', u'Exported 12 sheets.'))
    vi = Reducer().reduce(g, state, lang='vi')
    en = Reducer().reduce(g, state, lang='en')
    check('reducer: vietnamese headings', u'Bước 1' in vi, vi[:80])
    check('reducer: english headings', u'Step 1' in en, en[:80])


def test_reducer_ignores_dormant_fallback_nodes():
    g = _two_node_graph()
    g.add_node(_n('n1_alt', goal=u'thống kê tường', kind=AGENT))
    g.add_edge('n1', 'n1_alt', kind=FALLBACK)
    state = GraphState()
    state.set_result(_result('n1', u'128 walls.'))
    state.set_result(_result('n2', u'Exported 12 sheets.'))
    out = Reducer().reduce(g, state, lang='en')
    check('reducer: no phantom skipped step', 'Skipped' not in out, out)


def test_reducer_reports_a_total_failure_plainly():
    g = _two_node_graph()
    state = GraphState()
    for nid in ('n1', 'n2'):
        res = NodeResult(nid, status=FAILED, error=u'nope')
        res.started = res.finished = time.time()
        state.set_result(res)
    out = Reducer().reduce(g, state, lang='en')
    check('reducer: total failure is stated', out.startswith('No step'), out)


# ─────────────────────────────────────────────────────────────────────────────
# Planner
# ─────────────────────────────────────────────────────────────────────────────

def _planner():
    from Intelligence.agents.dispatcher import AgentDispatcher
    return Planner(dispatcher=AgentDispatcher())


def test_planner_leaves_a_single_goal_alone():
    for text in (u'có bao nhiêu bức tường?',
                 u'tô đỏ toàn bộ tường',
                 u'how many doors are in the model'):
        plan = _planner().plan(text)
        check('planner: single goal stays single ({})'.format(text[:24]),
              not plan.is_multi, [n.goal for n in plan.agent_nodes()])


def test_planner_does_not_split_a_noun_conjunction():
    plan = _planner().plan(u'thống kê tường và cột')
    check('planner: "tường và cột" is ONE goal', not plan.is_multi,
          [n.goal for n in plan.agent_nodes()])


def test_planner_splits_a_real_sequence():
    plan = _planner().plan(u'thống kê tường rồi xuất pdf sheet A-101')
    goals = [n.goal for n in plan.agent_nodes()]
    check('planner: two goals found', plan.is_multi, goals)
    check('planner: first is the count', u'thống kê' in goals[0], goals)
    check('planner: second is the export', u'xuất pdf' in goals[1], goals)
    check('planner: diamond pattern', plan.graph.pattern == 'diamond',
          plan.graph.pattern)


def test_planner_splits_two_commands_joined_by_and():
    plan = _planner().plan(u'thống kê tường và xuất pdf sheet A')
    check('planner: two commands split on "và"', plan.is_multi,
          [n.goal for n in plan.agent_nodes()])


def test_planner_merges_a_verbless_fragment():
    plan = _planner().plan(u'thống kê tường rồi cửa sổ')
    goals = [n.goal for n in plan.agent_nodes()]
    check('planner: verbless tail merged, not a second goal',
          not plan.is_multi, goals)


def test_planner_marks_writers_and_orders_them():
    plan = _planner().plan(u'thống kê tường rồi tô đỏ toàn bộ tường')
    nodes = plan.agent_nodes()
    check('planner: found two goals', len(nodes) == 2,
          [n.goal for n in nodes])
    if len(nodes) == 2:
        check('planner: second node is a writer', nodes[1].writer,
              nodes[1].specialist)
        # The writer must depend on the read that preceded it.
        check('planner: writer waits for the earlier step',
              nodes[0].id in plan.graph.dependencies(nodes[1].id),
              plan.graph.dependencies(nodes[1].id))


def test_planner_caps_the_node_count():
    text = u'; '.join([u'thống kê tường'] * 12)
    plan = _planner().plan(text)
    check('planner: node count capped',
          len(plan.agent_nodes()) <= 6, len(plan.agent_nodes()))


def test_planner_gives_read_nodes_a_fallback_but_not_writers():
    plan = _planner().plan(u'thống kê tường rồi tô đỏ toàn bộ tường')
    nodes = plan.agent_nodes()
    if len(nodes) == 2:
        read, write = nodes[0], nodes[1]
        check('planner: read node has a fallback',
              len(plan.graph.fallbacks_for(read.id)) == 1,
              plan.graph.fallbacks_for(read.id))
        check('planner: writer has NO fallback',
              plan.graph.fallbacks_for(write.id) == [],
              plan.graph.fallbacks_for(write.id))


def test_planner_demotes_a_prohibition_away_from_writing():
    plan = _planner().plan(u'đừng xóa text note, chỉ liệt kê chúng')
    check('planner: prohibition produces no writer node',
          not plan.has_writer(),
          [(n.goal, n.specialist, n.writer) for n in plan.agent_nodes()])


# ─────────────────────────────────────────────────────────────────────────────
# Router
# ─────────────────────────────────────────────────────────────────────────────

def test_router_flags_writers():
    from Intelligence.agents.dispatcher import AgentDispatcher
    r = GraphRouter(dispatcher=AgentDispatcher())
    read = r.route(u'có bao nhiêu bức tường')
    write = r.route(u'đổi tên sheet A-101 thành A-102')
    check('router: read is not a writer', not read['writer'], read)
    check('router: rename is a writer', write['writer'], write)


def test_router_demotes_a_prohibition():
    from Intelligence.agents.dispatcher import AgentDispatcher
    from Intelligence.language import analyze
    r = GraphRouter(dispatcher=AgentDispatcher())
    text = u'đừng xóa text note'
    decision = r.route(text, utterance=analyze(text))
    check('router: prohibition is not routed as a write',
          not decision['writer'], decision)
    check('router: demotion is attributed',
          decision['source'] == 'negation-demoted', decision)


def test_router_fallbacks():
    r = GraphRouter()
    check('router: read specialist has an alternative',
          r.alternatives('revit_data') == ['general'],
          r.alternatives('revit_data'))
    check('router: writer has none',
          r.alternatives('revit_action') == [],
          r.alternatives('revit_action'))


# ─────────────────────────────────────────────────────────────────────────────
# Orchestrator
# ─────────────────────────────────────────────────────────────────────────────

def _orchestrator(runner, **kw):
    from Intelligence.agents.dispatcher import AgentDispatcher
    return GraphOrchestrator(runner, dispatcher=AgentDispatcher(), **kw)


def test_orchestrator_declines_single_goal_turns():
    orch = _orchestrator(lambda n, c, s: u'never called')
    check('orchestrator: single goal is not handled',
          not orch.should_handle(u'có bao nhiêu bức tường?'))
    res = orch.handle(u'có bao nhiêu bức tường?')
    check('orchestrator: handle() reports it did nothing', not res.handled,
          res.to_dict())


def test_orchestrator_runs_a_multi_goal_turn_end_to_end():
    seen = []

    def runner(node, context, state):
        seen.append(node.goal)
        return u'Đã xong: {} (42 phần tử)'.format(node.goal)

    orch = _orchestrator(runner)
    text = u'thống kê tường rồi xuất pdf sheet A-101'
    check('orchestrator: multi goal is handled', orch.should_handle(text))
    res = orch.handle(text, record_trace=False)
    check('orchestrator: handled', res.handled, res.to_dict())
    check('orchestrator: both nodes ran', len(seen) == 2, seen)
    check('orchestrator: answer covers both steps',
          u'thống kê' in res.answer and u'xuất pdf' in res.answer,
          res.answer[:200])
    check('orchestrator: reports ok', res.ok, res.to_dict())


def test_orchestrator_step_callback_numbers_the_steps():
    steps = []
    orch = _orchestrator(lambda n, c, s: u'ok, 1 element',
                         on_step=lambda i, total, node: steps.append((i, total)))
    orch.handle(u'thống kê tường rồi xuất pdf sheet A-101',
                record_trace=False)
    check('orchestrator: step callback fired 1/2 then 2/2',
          steps == [(1, 2), (2, 2)], steps)


def test_fallback_route_is_preferred_over_adaptation():
    """The graph's own redundant path runs FIRST; adapting is the last resort."""
    attempts = []

    def runner(node, context, state):
        attempts.append((node.id, node.specialist))
        if node.id == 'n1':
            raise RuntimeError('specialist could not do it')
        return u'Model có 128 bức tường.'

    orch = _orchestrator(runner)
    res = orch.handle(u'thống kê tường rồi xuất pdf sheet A-101',
                      record_trace=False)
    check('fallback preferred: alt route ran', ('n1_alt', 'general') in attempts,
          attempts)
    check('fallback preferred: no adaptation was needed',
          res.trace.adaptations == [], res.trace.adaptations)
    check('fallback preferred: answer carries the recovered result',
          u'128' in res.answer, res.answer[:200])


def test_orchestrator_adapts_when_the_fallback_also_fails():
    """Last resort: re-run the node itself on the unrestricted `general` route.

    Every node fails its FIRST attempt here, so the graph's redundant path is
    exhausted (n1 fails, n1_alt fails) and only the adapt loop can recover n1.
    """
    calls = {}

    def runner(node, context, state):
        calls[node.id] = calls.get(node.id, 0) + 1
        if node.id in ('n1', 'n1_alt') and calls[node.id] == 1:
            raise RuntimeError('first attempt always fails')
        return u'Model có 128 bức tường.'

    orch = _orchestrator(runner)
    res = orch.handle(u'thống kê tường rồi xuất pdf sheet A-101',
                      record_trace=False)
    check('adapt: n1 was run a second time', calls.get('n1') == 2, calls)
    check('adapt: recorded in the trace', len(res.trace.adaptations) >= 1,
          res.trace.adaptations if res.trace else None)
    check('adapt: node was widened to general',
          res.plan.graph.get_node('n1').specialist == 'general',
          res.plan.graph.get_node('n1').specialist)
    check('adapt: final answer has the recovered result',
          u'128' in res.answer, res.answer[:200])


def test_reducer_node_never_costs_an_agent_turn():
    """The convergence node is structural — it must not reach the runner."""
    seen = []

    def runner(node, context, state):
        seen.append(node.id)
        return u'ok, 7 elements'

    _orchestrator(runner).handle(u'thống kê tường rồi xuất pdf sheet A-101',
                                 record_trace=False)
    check('structural node skipped by the runner', 'reduce' not in seen, seen)
    check('only the two agent nodes ran', seen == ['n1', 'n2'], seen)


def test_orchestrator_never_readapts_a_writer():
    attempts = []

    def runner(node, context, state):
        attempts.append((node.id, node.specialist, node.writer))
        if node.writer:
            raise RuntimeError('write failed')
        return u'ok, 3 elements'

    orch = _orchestrator(runner)
    orch.handle(u'thống kê tường rồi tô đỏ toàn bộ tường', record_trace=False)
    writer_runs = [a for a in attempts if a[2]]
    check('orchestrator: a failed writer is NOT retried',
          len(writer_runs) == 1, attempts)


def test_orchestrator_survives_a_runner_that_always_throws():
    def runner(node, context, state):
        raise RuntimeError('everything is broken')

    res = _orchestrator(runner).handle(
        u'thống kê tường rồi xuất pdf sheet A-101', record_trace=False)
    check('orchestrator: still returns a result', res.handled, res.to_dict())
    check('orchestrator: reports not ok', not res.ok, res.to_dict())
    check('orchestrator: answer explains the failure',
          len(res.answer) > 0, res.answer[:200])


# ─────────────────────────────────────────────────────────────────────────────
# Observability
# ─────────────────────────────────────────────────────────────────────────────

def test_trace_records_spans_and_round_trips_through_jsonl():
    def runner(node, context, state):
        return u'ok, 5 elements'

    trace = GraphTrace(goal=u'thống kê tường rồi xuất pdf', pattern='diamond',
                       lang='vi')
    g = _two_node_graph()
    state = GraphState()
    GraphExecutor(runner, max_parallel=1, on_event=trace.on_event).run(g, state)
    trace.close(state)

    check('trace: one span per executed node',
          len([s for s in trace.spans if s.status == DONE]) >= 2,
          [s.to_dict() for s in trace.spans])
    check('trace: summary renders', 'graph[diamond]' in trace.summary(),
          trace.summary())

    path = os.path.join(tempfile.mkdtemp(), 'trace.jsonl')
    check('trace: record() writes', record(trace, path=path) is True)
    back = read_traces(path)
    check('trace: round-trips', len(back) == 1 and back[0]['pattern'] ==
          'diamond', back)
    check('trace: vietnamese goal survived ascii-safe encoding',
          back[0]['goal'] == u'thống kê tường rồi xuất pdf', back[0]['goal'])


def test_trace_records_a_skipped_node():
    def runner(node, context, state):
        if node.id == 'n1':
            raise RuntimeError('nope')
        return u'ok'

    g = TaskGraph()
    g.add_node(_n('n1'))
    g.add_node(_n('n2'))
    g.add_edge('n1', 'n2')
    trace = GraphTrace(goal=u'x', pattern='linear')
    state = GraphState()
    GraphExecutor(runner, max_parallel=1, on_event=trace.on_event).run(g, state)
    trace.close(state)
    statuses = [s.status for s in trace.spans]
    check('trace: failure and skip both recorded',
          FAILED in statuses and SKIPPED in statuses, statuses)


# ─────────────────────────────────────────────────────────────────────────────

def main():
    groups = [
        ('topology', [
            test_linear_layers_are_one_node_each,
            test_fan_out_puts_branches_in_one_layer,
            test_fan_in_converges,
            test_diamond_is_three_layers,
            test_hub_spoke_feedback_is_not_scheduled,
            test_loop_declares_feedback_without_a_cycle,
            test_cycle_is_detected,
            test_fallback_target_is_not_scheduled_normally,
            test_duplicate_node_id_is_rejected,
        ]),
        ('executor', [
            test_executor_runs_in_dependency_order,
            test_executor_actually_runs_a_layer_in_parallel,
            test_two_writers_never_overlap,
            test_failure_is_contained_and_dependents_are_skipped,
            test_retry_uses_max_attempts,
            test_fallback_path_recovers_the_original_node,
            test_cancel_stops_the_sweep,
            test_context_carries_dependency_output_forward,
            test_runner_returning_empty_is_a_failure,
        ]),
        ('verifier', [
            test_verifier_catches_a_refusal,
            test_verifier_wants_a_number_for_a_counting_question,
            test_verifier_flags_empty_output,
            test_verifier_flags_a_leaked_tool_error,
            test_executor_treats_a_verifier_veto_as_failure,
        ]),
        ('reducer', [
            test_reducer_keeps_a_failed_step_visible,
            test_reducer_is_bilingual,
            test_reducer_ignores_dormant_fallback_nodes,
            test_reducer_reports_a_total_failure_plainly,
        ]),
        ('planner', [
            test_planner_leaves_a_single_goal_alone,
            test_planner_does_not_split_a_noun_conjunction,
            test_planner_splits_a_real_sequence,
            test_planner_splits_two_commands_joined_by_and,
            test_planner_merges_a_verbless_fragment,
            test_planner_marks_writers_and_orders_them,
            test_planner_caps_the_node_count,
            test_planner_gives_read_nodes_a_fallback_but_not_writers,
            test_planner_demotes_a_prohibition_away_from_writing,
        ]),
        ('router', [
            test_router_flags_writers,
            test_router_demotes_a_prohibition,
            test_router_fallbacks,
        ]),
        ('orchestrator', [
            test_orchestrator_declines_single_goal_turns,
            test_orchestrator_runs_a_multi_goal_turn_end_to_end,
            test_orchestrator_step_callback_numbers_the_steps,
            test_reducer_node_never_costs_an_agent_turn,
            test_fallback_route_is_preferred_over_adaptation,
            test_orchestrator_adapts_when_the_fallback_also_fails,
            test_orchestrator_never_readapts_a_writer,
            test_orchestrator_survives_a_runner_that_always_throws,
        ]),
        ('observability', [
            test_trace_records_spans_and_round_trips_through_jsonl,
            test_trace_records_a_skipped_node,
        ]),
    ]
    for group, tests in groups:
        print('\n{}'.format(group))
        for t in tests:
            try:
                t()
            except Exception as ex:
                FAILURES.append(t.__name__)
                print('  ERROR {}  {}'.format(t.__name__, ex))

    print('')
    if FAILURES:
        print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('all passed')
    return 0


if __name__ == '__main__':
    sys.exit(main())
