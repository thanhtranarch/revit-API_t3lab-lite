# -*- coding: utf-8 -*-
"""
Graph primitives — the building blocks every other module in this package
composes: NODE, EDGE, STATE, CONTEXT and the result a node produces.

Design rules that the rest of the package relies on:

- A Node holds *what to do*, never *how to do it*. Execution is supplied by
  the caller's `runner`, so the same graph runs against the real agent loop in
  Revit and against a scripted stub in the tests.
- An Edge is typed. `sequence` means "dst runs after src"; `fallback` means
  "dst runs ONLY IF src failed" (the redundant path that makes the graph fault
  tolerant); `feedback` means "src's result informs dst on a LATER turn" and is
  ignored by the scheduler — it exists so the adaptation loop can be drawn in
  the graph rather than hidden in code.
- GraphState is the only mutable thing, and every access goes through its lock.
  Nodes run on worker threads; a plain dict here would be a data race the day
  two branches finish at once.
- GraphContext gives each node LOCAL context (what its dependencies produced)
  on top of GLOBAL context (the conversation, the Revit state). That is the
  whole "local decisions, global coherence" property: a node reasons over its
  own inputs without needing to know the rest of the graph exists.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import threading
import time

# ─── Node kinds ───────────────────────────────────────────────────────────────
AGENT    = 'agent'      # runs a specialist turn (the only kind that costs LLM)
ROUTER   = 'router'     # decides where work flows
REDUCER  = 'reducer'    # aggregates and distills sibling results
VERIFIER = 'verifier'   # checks, scores, validates
MEMORY   = 'memory'     # reads or writes persistent memory

NODE_KINDS = (AGENT, ROUTER, REDUCER, VERIFIER, MEMORY)

# ─── Edge kinds ───────────────────────────────────────────────────────────────
SEQUENCE = 'sequence'   # dst depends on src — scheduling edge
FALLBACK = 'fallback'   # dst runs only when src failed — redundant path
FEEDBACK = 'feedback'   # informs a later turn — never scheduled

EDGE_KINDS = (SEQUENCE, FALLBACK, FEEDBACK)

# ─── Node lifecycle ───────────────────────────────────────────────────────────
PENDING   = 'pending'
RUNNING   = 'running'
DONE      = 'done'
FAILED    = 'failed'
SKIPPED   = 'skipped'    # a dependency failed — never attempted
CANCELLED = 'cancelled'

TERMINAL_STATES = (DONE, FAILED, SKIPPED, CANCELLED)


class Node(object):
    """One unit of work in the graph.

    Args:
        node_id:     unique within its graph.
        goal:        the request this node is responsible for, in the user's
                     own words — it becomes the node's prompt.
        kind:        one of NODE_KINDS.
        specialist:  SpecialistSpec name ('revit_data', 'export'...) or None
                     for the provider default.
        skill:       matched skill id, or None.
        writer:      True when the node may MODIFY the model. The executor
                     serialises writers against each other, and never runs two
                     at once, because two interleaved transactions confuse
                     Revit's undo stack.
        max_attempts: retries on failure, including the first try.
        optional:    when True, a failure here does not skip dependents — used
                     for enrichment steps whose absence degrades the answer
                     without invalidating it.
    """

    __slots__ = ('id', 'goal', 'kind', 'specialist', 'skill', 'writer',
                 'max_attempts', 'optional', 'meta')

    def __init__(self, node_id, goal, kind=AGENT, specialist=None, skill=None,
                 writer=False, max_attempts=1, optional=False, meta=None):
        if kind not in NODE_KINDS:
            raise ValueError('unknown node kind: {}'.format(kind))
        self.id           = node_id
        self.goal         = goal or u''
        self.kind         = kind
        self.specialist   = specialist
        self.skill        = skill
        self.writer       = bool(writer)
        self.max_attempts = max(1, int(max_attempts))
        self.optional     = bool(optional)
        self.meta         = dict(meta or {})

    def to_dict(self):
        return {'id': self.id, 'goal': self.goal, 'kind': self.kind,
                'specialist': self.specialist, 'skill': self.skill,
                'writer': self.writer, 'optional': self.optional}

    def __repr__(self):
        return '<Node {} {} spec={} writer={}>'.format(
            self.id, self.kind, self.specialist, self.writer)


class Edge(object):
    """A typed relationship between two nodes."""

    __slots__ = ('src', 'dst', 'kind', 'label')

    def __init__(self, src, dst, kind=SEQUENCE, label=u''):
        if kind not in EDGE_KINDS:
            raise ValueError('unknown edge kind: {}'.format(kind))
        self.src   = src
        self.dst   = dst
        self.kind  = kind
        self.label = label or u''

    def to_dict(self):
        return {'src': self.src, 'dst': self.dst, 'kind': self.kind,
                'label': self.label}

    def __repr__(self):
        return '<Edge {} -{}-> {}>'.format(self.src, self.kind, self.dst)


class NodeResult(object):
    """What one node produced, plus how it went."""

    __slots__ = ('node_id', 'status', 'output', 'data', 'error',
                 'attempts', 'started', 'finished', 'specialist')

    def __init__(self, node_id, status=PENDING, output=u'', data=None,
                 error=None, attempts=0, specialist=None):
        self.node_id    = node_id
        self.status     = status
        self.output     = output or u''
        self.data       = dict(data or {})
        self.error      = error
        self.attempts   = attempts
        self.specialist = specialist
        self.started    = None
        self.finished   = None

    @property
    def ok(self):
        return self.status == DONE

    @property
    def duration(self):
        if self.started is None:
            return 0.0
        return (self.finished or time.time()) - self.started

    def to_dict(self):
        return {'node': self.node_id, 'status': self.status,
                'chars': len(self.output), 'attempts': self.attempts,
                'specialist': self.specialist,
                'ms': int(round(self.duration * 1000.0)),
                'error': self.error}

    def __repr__(self):
        return '<NodeResult {} {} {}ch>'.format(
            self.node_id, self.status, len(self.output))


class GraphState(object):
    """Shared, thread-safe state for one graph run.

    `blackboard` is the free-form scratch space nodes use to leave each other
    structured values (element ids, resolved sheet sets) that do not belong in
    the prose output.
    """

    def __init__(self, goal=u'', global_context=u''):
        self.goal           = goal or u''
        self.global_context = global_context or u''
        self._lock          = threading.Lock()
        self._results       = {}      # node_id -> NodeResult
        self._blackboard    = {}
        self._cancel        = threading.Event()
        self.started        = time.time()
        self.finished       = None

    # ── results ──────────────────────────────────────────────────────────

    def set_result(self, result):
        with self._lock:
            self._results[result.node_id] = result

    def get_result(self, node_id):
        with self._lock:
            return self._results.get(node_id)

    def results(self):
        """Snapshot dict of every result recorded so far."""
        with self._lock:
            return dict(self._results)

    def status_of(self, node_id):
        res = self.get_result(node_id)
        return res.status if res is not None else PENDING

    def succeeded(self, node_id):
        return self.status_of(node_id) == DONE

    def outputs_for(self, node_ids):
        """[(node_id, output)] for the given nodes that produced anything."""
        snap = self.results()
        out = []
        for nid in node_ids:
            res = snap.get(nid)
            if res is not None and res.output:
                out.append((nid, res.output))
        return out

    # ── blackboard ───────────────────────────────────────────────────────

    def put(self, key, value):
        with self._lock:
            self._blackboard[key] = value

    def get(self, key, default=None):
        with self._lock:
            return self._blackboard.get(key, default)

    def blackboard(self):
        with self._lock:
            return dict(self._blackboard)

    # ── cancellation ─────────────────────────────────────────────────────

    def cancel(self):
        self._cancel.set()

    def is_cancelled(self):
        return self._cancel.is_set()

    # ── summary ──────────────────────────────────────────────────────────

    def counts(self):
        tally = {}
        for res in self.results().values():
            tally[res.status] = tally.get(res.status, 0) + 1
        return tally

    def to_dict(self):
        return {'goal': self.goal, 'counts': self.counts(),
                'results': [r.to_dict() for r in self.results().values()]}


class GraphContext(object):
    """Assembles the context ONE node reasons over.

    Global context is what the whole turn shares (conversation summary, live
    Revit state). Local context is what this node's dependencies produced. A
    node sees both and nothing else — it never reads a sibling's output, which
    is what keeps branches independent enough to run in parallel.
    """

    # Dependency output is quoted into the next node's prompt, so it has to be
    # bounded: a node that listed 900 elements would otherwise push the real
    # instruction out of a local model's context window.
    MAX_DEP_CHARS = 1200

    def __init__(self, state, graph):
        self.state = state
        self.graph = graph

    def local_for(self, node_id):
        """Text block of what this node's dependencies produced, or u''."""
        deps = self.graph.dependencies(node_id)
        pieces = []
        for dep_id, output in self.state.outputs_for(deps):
            node = self.graph.get_node(dep_id)
            title = (node.goal if node is not None and node.goal else dep_id)
            body = output.strip()
            if len(body) > self.MAX_DEP_CHARS:
                body = body[:self.MAX_DEP_CHARS] + u'\n[...truncated]'
            pieces.append(u'### Result of "{}"\n{}'.format(title, body))
        return u'\n\n'.join(pieces)

    def for_node(self, node_id):
        """Full context block for a node: global first, then local."""
        parts = []
        if self.state.global_context:
            parts.append(self.state.global_context)
        local = self.local_for(node_id)
        if local:
            parts.append(u'## Results from earlier steps in this request\n'
                         + local)
        return u'\n\n'.join(parts)
