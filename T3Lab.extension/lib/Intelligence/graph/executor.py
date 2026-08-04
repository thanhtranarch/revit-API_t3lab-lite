# -*- coding: utf-8 -*-
"""
EXECUTION LAYER — runs a TaskGraph: in parallel, in order, surviving failures.

The executor walks the graph one dependency layer at a time. Everything inside
a layer is independent by construction, so it all starts together; the layer
ends when its slowest node finishes. Four properties matter more than speed:

  WRITER SERIALISATION  at most one model-modifying node runs at a time, ever.
                        Two interleaved Revit transactions produce an undo
                        stack nobody can reason about, so writers queue behind
                        each other even when the graph says they are
                        independent. This is the same rule
                        `agents/task_manager.py` enforces with its writer slot;
                        it is implemented here rather than delegated because
                        the executor needs to WAIT for a layer to finish, and
                        that manager is a fire-and-forget submission queue.

  FAULT CONTAINMENT     a failed node does not fail the graph. Its dependents
                        are marked SKIPPED (they would have consumed an answer
                        that does not exist) and every unrelated branch runs to
                        completion. The reducer then reports what did and did
                        not happen.

  REDUNDANT PATHS       a failed node with a `fallback` edge activates its
                        alternative, and if THAT succeeds the result is
                        recorded under the ORIGINAL node id as well — the graph
                        recovered, so downstream nodes and the reducer see a
                        node that produced an answer, not a hole.

  COOPERATIVE CANCEL    `state.cancel()` stops the sweep between layers and is
                        passed into each node so a long-running runner can bail
                        mid-flight. Nothing is killed by force.

Attaching to a UI
-----------------
`max_parallel=1` runs the graph strictly sequentially in topological order.
That is what script.py uses, because two agent turns streaming into one chat
transcript interleave into nonsense. The parallel path is real, tested
(dev/test_graph_agents.py) and available to any headless caller — a background
task, a batch report — that is not writing into a live transcript.

Pure Python (threading only), IronPython 2.7 + CPython 3 compatible.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import threading
import time
import traceback

from Intelligence.graph.primitives import (
    AGENT, CANCELLED, DONE, FAILED, PENDING, RUNNING, SKIPPED,
    GraphContext, NodeResult,
)
from Intelligence.graph.topology import GraphCycleError

# Default parallelism for headless callers. Matches AgentTaskManager's own
# ceiling — the bottleneck is Revit's single main thread, not the pool.
DEFAULT_MAX_PARALLEL = 3


class GraphExecutor(object):
    """Runs a TaskGraph with an injected per-node runner.

    Args:
        runner:  callable(node, context_text, state) -> NodeResult | unicode.
                 A bare string is accepted and wrapped as a DONE result, so
                 simple callers (and tests) do not have to build one.
        max_parallel: nodes in flight at once. 1 = strictly sequential.
        on_event: optional callable(event_name, payload) for observability.
                  Called from WORKER threads — marshal to a UI yourself.
        verifier: optional Verifier. When supplied, a node whose output fails
                  verification is treated as failed, which is what lets a
                  refusal ("tôi không thể…") trigger the fallback path instead
                  of flowing into the reducer as a successful answer.
    """

    def __init__(self, runner, max_parallel=DEFAULT_MAX_PARALLEL,
                 on_event=None, verifier=None, utterances=None,
                 node_handlers=None):
        self.runner        = runner
        self.max_parallel  = max(1, int(max_parallel))
        self.on_event      = on_event
        self.verifier      = verifier
        self.utterances    = dict(utterances or {})
        # kind -> callable(node, context, state). Anything not listed and not
        # an AGENT is a STRUCTURAL node (a reducer or verifier marking a
        # convergence point in the graph); it completes without work, because
        # the orchestrator owns those operations and runs them over the whole
        # finished state. Sending a structural node to the agent runner cost a
        # whole extra LLM turn per multi-goal request.
        self.node_handlers = dict(node_handlers or {})

    # ── public ───────────────────────────────────────────────────────────

    def run(self, graph, state):
        """Execute `graph`, recording everything in `state`. Returns `state`.

        Never raises for a node failure — only for a structurally invalid
        graph, which is a programming error rather than a runtime condition.
        """
        try:
            layers = graph.layers()
        except GraphCycleError:
            self._emit('graph_invalid', {'goal': graph.goal})
            raise

        self._emit('graph_start', {'goal': graph.goal,
                                   'pattern': graph.pattern,
                                   'layers': len(layers)})
        for index, layer in enumerate(layers):
            if state.is_cancelled():
                self._cancel_remaining(graph, state, layers[index:])
                break
            runnable = [nid for nid in layer
                        if self._should_run(graph, state, nid)]
            self._emit('layer_start', {'index': index, 'nodes': list(layer),
                                       'runnable': runnable})
            if runnable:
                self._run_layer(graph, state, runnable)
        state.finished = time.time()
        self._emit('graph_end', {'counts': state.counts()})
        return state

    # ── scheduling ───────────────────────────────────────────────────────

    def _should_run(self, graph, state, node_id):
        """False when a dependency failed — the node is marked SKIPPED."""
        node = graph.get_node(node_id)
        if node is None:
            return False
        if state.status_of(node_id) in (DONE, FAILED, SKIPPED, CANCELLED):
            return False        # already handled (e.g. recovered by a fallback)

        blocked = []
        for dep in graph.dependencies(node_id):
            status = state.status_of(dep)
            if status in (FAILED, SKIPPED, CANCELLED):
                dep_node = graph.get_node(dep)
                if dep_node is not None and dep_node.optional:
                    continue    # an enrichment step; its absence degrades only
                blocked.append(dep)
        if blocked:
            result = NodeResult(node_id, status=SKIPPED,
                                error=u'dependency failed: {}'.format(
                                    u', '.join(blocked)))
            result.started = result.finished = time.time()
            state.set_result(result)
            self._emit('node_skipped', {'node': node_id, 'blocked_by': blocked})
            return False
        return True

    def _run_layer(self, graph, state, node_ids):
        if self.max_parallel == 1 or len(node_ids) == 1:
            for nid in node_ids:
                if state.is_cancelled():
                    self._mark_cancelled(state, nid)
                    continue
                self._execute(graph, state, nid)
            return

        # Writers are serialised against each other even inside one layer.
        # Everything else runs together up to max_parallel.
        pending = list(node_ids)
        threads = []
        writer_lock = threading.Lock()
        semaphore = threading.Semaphore(self.max_parallel)

        def _worker(nid):
            try:
                node = graph.get_node(nid)
                if node is not None and node.writer:
                    with writer_lock:
                        self._execute(graph, state, nid)
                else:
                    self._execute(graph, state, nid)
            finally:
                semaphore.release()

        for nid in pending:
            if state.is_cancelled():
                self._mark_cancelled(state, nid)
                continue
            semaphore.acquire()
            th = threading.Thread(target=_worker, args=(nid,))
            th.daemon = True
            threads.append(th)
            th.start()
        for th in threads:
            th.join()

    # ── one node ─────────────────────────────────────────────────────────

    def execute_node(self, graph, state, node_id, reset=False):
        """Run exactly one node, outside the layer sweep.

        Used by the orchestrator's ADAPT step to re-run a node after changing
        its route. `reset=True` clears the previous (failed) result first, so
        the node is eligible to run again.
        """
        if reset:
            state.set_result(NodeResult(node_id, status=PENDING))
        return self._execute(graph, state, node_id)

    def _execute(self, graph, state, node_id):
        node = graph.get_node(node_id)
        if node is None:
            return
        context = GraphContext(state, graph).for_node(node_id)

        result = NodeResult(node_id, status=RUNNING,
                            specialist=node.specialist)
        result.started = time.time()
        state.set_result(result)
        self._emit('node_start', {'node': node_id, 'kind': node.kind,
                                  'specialist': node.specialist,
                                  'writer': node.writer})

        handler = self.node_handlers.get(node.kind)
        if handler is None and node.kind != AGENT:
            # Structural node — nothing to run, and nothing to verify either
            # (an empty output here is correct, not a failure).
            result.status = DONE
            result.finished = time.time()
            state.set_result(result)
            self._emit('node_end', result.to_dict())
            return
        run = handler or self.runner

        attempts = 0
        last_error = None
        while attempts < node.max_attempts:
            attempts += 1
            if state.is_cancelled():
                result.status = CANCELLED
                break
            try:
                produced = run(node, context, state)
            except Exception:
                last_error = self._tail(traceback.format_exc())
                continue
            normalised = self._normalise(node_id, produced, node)
            if normalised.status == DONE:
                result.output = normalised.output
                result.data = normalised.data
                result.status = DONE
                break
            last_error = normalised.error or last_error
            result.output = normalised.output or result.output

        result.attempts = attempts
        if result.status == RUNNING:
            result.status = FAILED
            result.error = last_error or u'node produced no result'

        # A verifier veto is a failure: an answer that refuses, or answers a
        # counting question with no number, must take the fallback path rather
        # than flow into the reducer looking like a success.
        if result.status == DONE and self.verifier is not None:
            verdict = self.verifier.check(
                node, result, utterance=self.utterances.get(node_id))
            result.data['verdict'] = verdict.to_dict()
            if verdict.should_retry:
                result.status = FAILED
                result.error = u'verification failed: {}'.format(
                    u', '.join(verdict.reasons) or u'low score')

        result.finished = time.time()
        state.set_result(result)
        self._emit('node_end', result.to_dict())

        if result.status == FAILED:
            self._try_fallback(graph, state, node_id)

    def _try_fallback(self, graph, state, node_id):
        """Run the redundant path for a failed node, if the graph has one."""
        for alt_id in graph.fallbacks_for(node_id):
            if state.is_cancelled():
                return
            if state.status_of(alt_id) != PENDING:
                continue
            self._emit('fallback_start', {'node': node_id, 'alt': alt_id})
            self._execute(graph, state, alt_id)
            alt_result = state.get_result(alt_id)
            if alt_result is not None and alt_result.status == DONE:
                # The graph recovered. Record the answer under the ORIGINAL id
                # too, so dependents and the reducer see a node that produced
                # a result instead of a hole they have to reason about.
                recovered = NodeResult(
                    node_id, status=DONE, output=alt_result.output,
                    data=dict(alt_result.data, recovered_by=alt_id),
                    attempts=alt_result.attempts,
                    specialist=alt_result.specialist)
                recovered.started = alt_result.started
                recovered.finished = alt_result.finished
                state.set_result(recovered)
                self._emit('node_recovered', {'node': node_id, 'alt': alt_id})
                return

    # ── helpers ──────────────────────────────────────────────────────────

    def _normalise(self, node_id, produced, node):
        """Accept a NodeResult, a string, or None from the runner."""
        if produced is None:
            return NodeResult(node_id, status=FAILED,
                              error=u'runner returned nothing',
                              specialist=node.specialist)
        if isinstance(produced, NodeResult):
            return produced
        text = u'{}'.format(produced).strip()
        if not text:
            return NodeResult(node_id, status=FAILED,
                              error=u'runner returned an empty answer',
                              specialist=node.specialist)
        return NodeResult(node_id, status=DONE, output=text,
                          specialist=node.specialist)

    def _mark_cancelled(self, state, node_id):
        result = NodeResult(node_id, status=CANCELLED)
        result.started = result.finished = time.time()
        state.set_result(result)
        self._emit('node_cancelled', {'node': node_id})

    def _cancel_remaining(self, graph, state, layers):
        for layer in layers:
            for nid in layer:
                if state.status_of(nid) == PENDING:
                    self._mark_cancelled(state, nid)

    @staticmethod
    def _tail(text, limit=600):
        return u'{}'.format(text or u'')[-limit:]

    def _emit(self, name, payload):
        if self.on_event is None:
            return
        try:
            self.on_event(name, payload)
        except Exception:
            pass
