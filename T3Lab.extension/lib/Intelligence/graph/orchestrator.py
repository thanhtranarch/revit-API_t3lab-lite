# -*- coding: utf-8 -*-
"""
GRAPH ORCHESTRATOR — the loop that runs the five graph operations.

    1. PLAN      decompose the goal into a graph          (planner.py)
    2. DISPATCH  route each node to a specialist          (router.py)
    3. EXECUTE   run it, in parallel, surviving failures  (executor.py)
    4. OBSERVE   collect spans and score the results      (verifier.py,
                                                           observability.py)
    5. ADAPT     re-route what failed and run it again    (here)

Adaptation is the step that makes this a loop rather than a pipeline. When a
node fails verification and the executor's redundant path did not rescue it,
the orchestrator demotes the node to the plain `general` specialist — full tool
catalog, no narrowing — and runs that node alone, once. A specialist's tool
subset is the most common reason a node cannot answer (it was routed to
`revit_data` and needed a write tool, or to `export` and needed a schedule), so
widening is the correction that fits the failure. It happens at most once per
node, and never for a writer: re-running a model edit is how an edit gets
applied twice.

`should_handle()` is the gate. It returns True only when the planner found more
than one sub-goal, so a single-goal turn — which is almost every turn — never
enters this code at all and keeps today's exact behaviour.

Nothing here imports Revit or WPF; the caller supplies a `runner`.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

from Intelligence.graph import observability
from Intelligence.graph.executor import GraphExecutor
from Intelligence.graph.planner import Planner
from Intelligence.graph.primitives import DONE, FAILED, GraphState
from Intelligence.graph.reducer import Reducer
from Intelligence.graph.router import GraphRouter
from Intelligence.graph.verifier import Verifier

# The specialist an adapted node falls back to: no tool narrowing, no tightened
# budget, exactly the behaviour the assistant had before specialists existed.
_ADAPT_SPECIALIST = 'general'


class GraphResult(object):
    """What one orchestrated turn produced."""

    def __init__(self, answer=u'', plan=None, state=None, trace=None,
                 verdicts=None, handled=True):
        self.answer   = answer or u''
        self.plan     = plan
        self.state    = state
        self.trace    = trace
        self.verdicts = dict(verdicts or {})
        self.handled  = handled

    @property
    def ok(self):
        """True when at least one planned step produced an answer."""
        if self.state is None or self.plan is None:
            return False
        return any(self.state.succeeded(n.id)
                   for n in self.plan.agent_nodes())

    def failed_nodes(self):
        if self.state is None or self.plan is None:
            return []
        return [n.id for n in self.plan.agent_nodes()
                if self.state.status_of(n.id) != DONE]

    def to_dict(self):
        return {'handled': self.handled, 'ok': self.ok,
                'answer_chars': len(self.answer),
                'failed': self.failed_nodes(),
                'trace': self.trace.to_dict() if self.trace else None}


class GraphOrchestrator(object):
    """Plans, runs, reduces and adapts one user request.

    Args:
        runner:       callable(node, context, state) -> NodeResult | unicode.
        dispatcher:   an AgentDispatcher for routing (optional).
        provider:     LLM provider, used ONLY by the reducer's optional
                      distill step. Node execution is the runner's business.
        max_parallel: 1 when attached to a chat transcript (see executor).
        on_step:      callable(index, total, node) before each agent node —
                      script.py uses it to print "Bước 1/3".
    """

    def __init__(self, runner, dispatcher=None, skills_engine=None,
                 provider=None, max_parallel=1, on_step=None,
                 distill=False):
        self.runner       = runner
        self.router       = GraphRouter(dispatcher=dispatcher,
                                        skills_engine=skills_engine,
                                        allow_llm=False)
        self.planner      = Planner(router=self.router, dispatcher=dispatcher,
                                    skills_engine=skills_engine)
        self.reducer      = Reducer(provider=provider)
        self.verifier     = Verifier()
        self.max_parallel = max_parallel
        self.on_step      = on_step
        self.distill      = bool(distill)

    # ── gate ─────────────────────────────────────────────────────────────

    def plan_for(self, text, utterance=None):
        """Build a Plan without running anything. Never raises."""
        try:
            return self.planner.plan(text, utterance=utterance)
        except Exception:
            return None

    def should_handle(self, text, utterance=None, plan=None):
        """True only for genuinely multi-goal requests.

        Everything else keeps the existing single-specialist path, so this
        layer can only ever affect turns that the old path handled badly by
        construction (it answered one of the goals and dropped the rest).
        """
        plan = plan or self.plan_for(text, utterance=utterance)
        return bool(plan is not None and plan.is_multi)

    # ── run ──────────────────────────────────────────────────────────────

    def handle(self, text, global_context=u'', utterance=None, plan=None,
               lang=None, record_trace=True):
        """Run the whole loop. Returns a GraphResult.

        `handled=False` means the request was not multi-goal and the caller
        should fall through to its normal path — nothing was executed and
        nothing was shown to the user.
        """
        plan = plan or self.plan_for(text, utterance=utterance)
        if plan is None or not plan.is_multi:
            return GraphResult(handled=False)

        reply_lang = lang or plan.lang or 'vi'
        trace = observability.GraphTrace(
            goal=text, pattern=plan.graph.pattern, lang=reply_lang,
            analysis=(plan.utterance.to_dict()
                      if plan.utterance is not None else None))

        state = GraphState(goal=text, global_context=global_context)
        executor = GraphExecutor(
            runner=self._wrap_runner(plan),
            max_parallel=self.max_parallel,
            on_event=trace.on_event,
            verifier=self.verifier,
            utterances=plan.sub_utterances)

        try:
            executor.run(plan.graph, state)
        except Exception:
            # A structurally invalid graph is a bug in the planner, not a
            # runtime condition — fall through rather than fail the turn.
            trace.close(state)
            if record_trace:
                observability.record(trace)
            return GraphResult(handled=False, plan=plan, state=state,
                               trace=trace)

        # ── OBSERVE + ADAPT ───────────────────────────────────────────────
        verdicts = self.verifier.check_all(plan.graph, state,
                                           utterances=plan.sub_utterances)
        self._adapt(plan, state, executor, trace, verdicts)

        # ── REDUCE ────────────────────────────────────────────────────────
        node_ids = [n.id for n in plan.agent_nodes()]
        if self.distill:
            answer = self.reducer.reduce_with_llm(
                plan.graph, state, node_ids=node_ids, lang=reply_lang,
                goal=text)
        else:
            answer = self.reducer.reduce(plan.graph, state,
                                         node_ids=node_ids, lang=reply_lang)

        trace.close(state)
        if record_trace:
            observability.record(trace)
        return GraphResult(answer=answer, plan=plan, state=state, trace=trace,
                           verdicts=verdicts, handled=True)

    # ── internals ────────────────────────────────────────────────────────

    def _wrap_runner(self, plan):
        """Adds the "Bước i/N" callback around the caller's runner."""
        if self.on_step is None:
            return self.runner
        steps = [n.id for n in plan.agent_nodes()]
        total = len(steps)

        def _runner(node, context, state):
            if node.id in steps:
                try:
                    self.on_step(steps.index(node.id) + 1, total, node)
                except Exception:
                    pass
            return self.runner(node, context, state)

        return _runner

    def _adapt(self, plan, state, executor, trace, verdicts):
        """Re-run failed read nodes on the unrestricted `general` route.

        At most one adaptation per node, never for a writer, and never when the
        executor's own fallback path already recovered it.
        """
        for node in plan.agent_nodes():
            if state.is_cancelled():
                return
            if node.writer:
                continue
            result = state.get_result(node.id)
            if result is not None and result.status == DONE:
                continue
            if result is not None and result.status not in (FAILED,):
                continue        # skipped/cancelled — its dependency is the
                                # problem, re-running this node fixes nothing
            if node.specialist == _ADAPT_SPECIALIST:
                continue        # already unrestricted; nothing left to widen

            reason = u''
            verdict = verdicts.get(node.id)
            if verdict is not None:
                reason = u', '.join(verdict.reasons)
            trace.note_adaptation(node.id, u'respecialise -> {}'.format(
                _ADAPT_SPECIALIST), reason or (result.error if result else u''))

            node.specialist = _ADAPT_SPECIALIST
            node.skill = None
            try:
                # reset=True clears the recorded failure so the node is
                # eligible to run again under its widened route.
                executor.execute_node(plan.graph, state, node.id, reset=True)
            except Exception:
                pass

        # The reducer reads verdicts from the post-adaptation state.
        verdicts.update(self.verifier.check_all(
            plan.graph, state, utterances=plan.sub_utterances))
