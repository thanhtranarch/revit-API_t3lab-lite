# -*- coding: utf-8 -*-
"""
TaskGraph — the container, plus the standard topology patterns.

`layers()` is the load-bearing method: it returns the nodes grouped into
dependency levels, so everything inside one layer is independent and can run
together. That single computation is what turns "a graph" from a diagram into
an execution plan, and it is why the executor never needs to reason about
dependencies itself.

The pattern builders at the bottom are the shapes from the reference
architecture. They exist so the planner declares its intent ("this is a
diamond") rather than hand-wiring edges each time, and so a reader of a trace
can see which shape a turn actually used.

    linear     a → b → c            simple pipeline
    fan_out    a → {b, c, d}        parallel exploration
    fan_in     {a, b, c} → d        converge and decide
    diamond    a → {b, c} → d       explore, then verify
    hub_spoke  hub ↔ {a, b, c}      central coordination
    loop       a → b → c → a*       iterate and refine (*feedback edge)

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

from Intelligence.graph.primitives import (
    Edge, Node, SEQUENCE, FALLBACK, FEEDBACK, AGENT,
)


class GraphCycleError(ValueError):
    """Raised when scheduling edges form a cycle — the graph cannot be run."""


class TaskGraph(object):
    """A directed graph of Nodes joined by typed Edges."""

    def __init__(self, goal=u'', pattern=u'custom'):
        self.goal     = goal or u''
        self.pattern  = pattern
        self._nodes   = {}       # id -> Node
        self._order   = []       # insertion order, for stable output
        self._edges   = []

    # ── construction ─────────────────────────────────────────────────────

    def add_node(self, node):
        if node.id in self._nodes:
            raise ValueError('duplicate node id: {}'.format(node.id))
        self._nodes[node.id] = node
        self._order.append(node.id)
        return node

    def node(self, node_id, goal, **kw):
        """Convenience: build and add a Node in one call."""
        return self.add_node(Node(node_id, goal, **kw))

    def add_edge(self, src, dst, kind=SEQUENCE, label=u''):
        if src not in self._nodes:
            raise ValueError('unknown edge source: {}'.format(src))
        if dst not in self._nodes:
            raise ValueError('unknown edge target: {}'.format(dst))
        if src == dst and kind != FEEDBACK:
            raise ValueError('self-edge must be a feedback edge: {}'
                             .format(src))
        edge = Edge(src, dst, kind=kind, label=label)
        self._edges.append(edge)
        return edge

    # ── access ───────────────────────────────────────────────────────────

    def get_node(self, node_id):
        return self._nodes.get(node_id)

    def nodes(self):
        return [self._nodes[nid] for nid in self._order]

    def node_ids(self):
        return list(self._order)

    def edges(self, kind=None):
        if kind is None:
            return list(self._edges)
        return [e for e in self._edges if e.kind == kind]

    def __len__(self):
        return len(self._order)

    def __contains__(self, node_id):
        return node_id in self._nodes

    def dependencies(self, node_id):
        """Nodes that must finish before `node_id` runs (sequence edges)."""
        return [e.src for e in self._edges
                if e.kind == SEQUENCE and e.dst == node_id]

    def dependents(self, node_id):
        """Nodes waiting on `node_id` (sequence edges)."""
        return [e.dst for e in self._edges
                if e.kind == SEQUENCE and e.src == node_id]

    def fallbacks_for(self, node_id):
        """Nodes to activate when `node_id` fails (redundant paths)."""
        return [e.dst for e in self._edges
                if e.kind == FALLBACK and e.src == node_id]

    def is_fallback_target(self, node_id):
        """True when this node exists only as somebody's fallback.

        Such nodes are dormant: the executor skips them in the normal sweep
        and only runs them if the node they cover actually failed.
        """
        inbound = [e for e in self._edges if e.dst == node_id]
        return bool(inbound) and all(e.kind == FALLBACK for e in inbound)

    def writers(self):
        return [n for n in self.nodes() if n.writer]

    # ── scheduling ───────────────────────────────────────────────────────

    def layers(self):
        """Nodes grouped into dependency levels; each level is independent.

        Kahn's algorithm over SEQUENCE edges only — fallback edges are not
        dependencies (their target must not wait on a node that may never
        fail) and feedback edges are explicitly out of scheduling.

        Fallback-only targets are excluded: they have no place in the normal
        sweep, and including them would make the executor run a redundant
        path that nothing needed.

        Raises GraphCycleError when the sequence edges contain a cycle.
        """
        scheduled = [nid for nid in self._order
                     if not self.is_fallback_target(nid)]
        allowed = set(scheduled)

        indegree = dict((nid, 0) for nid in scheduled)
        outgoing = dict((nid, []) for nid in scheduled)
        for e in self._edges:
            if e.kind != SEQUENCE:
                continue
            if e.src not in allowed or e.dst not in allowed:
                continue
            indegree[e.dst] += 1
            outgoing[e.src].append(e.dst)

        # Preserve insertion order inside each layer so output is stable.
        ready = [nid for nid in scheduled if indegree[nid] == 0]
        out = []
        seen = 0
        while ready:
            out.append(list(ready))
            seen += len(ready)
            nxt = []
            for nid in ready:
                for dst in outgoing[nid]:
                    indegree[dst] -= 1
                    if indegree[dst] == 0:
                        nxt.append(dst)
            # Stable ordering: follow the graph's insertion order, not the
            # order dependencies happened to be released in.
            ready = [nid for nid in scheduled if nid in set(nxt)]
        if seen != len(scheduled):
            stuck = [nid for nid in scheduled if indegree[nid] > 0]
            raise GraphCycleError(
                'cycle among sequence edges: {}'.format(', '.join(stuck)))
        return out

    def validate(self):
        """Structural check. Returns a list of problems ([] when healthy)."""
        problems = []
        if not self._nodes:
            problems.append('graph has no nodes')
        try:
            self.layers()
        except GraphCycleError as ex:
            problems.append('{}'.format(ex))
        for nid in self._order:
            node = self._nodes[nid]
            if node.kind == AGENT and not node.goal.strip():
                problems.append('agent node {} has an empty goal'.format(nid))
        for e in self._edges:
            if e.kind == FALLBACK and e.src == e.dst:
                problems.append('fallback edge points at itself: {}'
                                .format(e.src))
        return problems

    # ── rendering ────────────────────────────────────────────────────────

    def describe(self):
        """Human-readable rendering — goes into the trace and the debug log."""
        lines = [u'TaskGraph({}) goal={!r}'.format(self.pattern, self.goal)]
        try:
            for i, layer in enumerate(self.layers()):
                lines.append(u'  layer {}:'.format(i))
                for nid in layer:
                    node = self._nodes[nid]
                    lines.append(u'    - {} [{}{}{}] {}'.format(
                        nid, node.kind,
                        u' writer' if node.writer else u'',
                        u'/' + node.specialist if node.specialist else u'',
                        node.goal[:70]))
        except GraphCycleError as ex:
            lines.append(u'  !! {}'.format(ex))
        extra = [e for e in self._edges if e.kind != SEQUENCE]
        for e in extra:
            lines.append(u'  {} -{}-> {}'.format(e.src, e.kind, e.dst))
        return u'\n'.join(lines)

    def to_dict(self):
        return {'goal': self.goal, 'pattern': self.pattern,
                'nodes': [self._nodes[n].to_dict() for n in self._order],
                'edges': [e.to_dict() for e in self._edges]}


# ─── Topology patterns ────────────────────────────────────────────────────────
# Each builder takes ready-made Nodes and returns a wired TaskGraph. They wire
# edges only — a builder never invents a node, so the planner stays the single
# place that decides what work exists.


def _seed(goal, pattern, nodes):
    g = TaskGraph(goal=goal, pattern=pattern)
    for n in nodes:
        g.add_node(n)
    return g


def linear(nodes, goal=u''):
    """a → b → c. Simple pipeline: easy to reason about, hard to adapt."""
    g = _seed(goal, 'linear', nodes)
    for prev, nxt in zip(nodes, nodes[1:]):
        g.add_edge(prev.id, nxt.id)
    return g


def fan_out(root, branches, goal=u''):
    """root → {b1, b2, ...}. Parallel exploration, more coverage."""
    g = _seed(goal, 'fan_out', [root] + list(branches))
    for b in branches:
        g.add_edge(root.id, b.id)
    return g


def fan_in(branches, sink, goal=u''):
    """{b1, b2, ...} → sink. Converge results, aggregate and decide."""
    g = _seed(goal, 'fan_in', list(branches) + [sink])
    for b in branches:
        g.add_edge(b.id, sink.id)
    return g


def diamond(root, branches, sink, goal=u''):
    """root → {b1, b2, ...} → sink. Explore, then verify — high reliability.

    The shape the planner uses for a multi-goal request: one plan node, a
    branch per sub-goal, one reducer that converges them.
    """
    g = _seed(goal, 'diamond', [root] + list(branches) + [sink])
    for b in branches:
        g.add_edge(root.id, b.id)
        g.add_edge(b.id, sink.id)
    return g


def hub_spoke(hub, spokes, goal=u''):
    """hub → each spoke → hub. Central coordination over shared context.

    Modelled as hub → spoke (scheduling) plus spoke → hub (feedback), because
    a real return edge would be a cycle and nothing could be scheduled.
    """
    g = _seed(goal, 'hub_spoke', [hub] + list(spokes))
    for s in spokes:
        g.add_edge(hub.id, s.id)
        g.add_edge(s.id, hub.id, kind=FEEDBACK, label=u'reports to hub')
    return g


def loop(nodes, goal=u'', max_cycles=2):
    """a → b → c, plus a feedback edge c → a. Iterate and refine.

    The feedback edge is not scheduled — it declares that the last node's
    findings feed the first on a subsequent pass. `max_cycles` is carried in
    the graph's pattern metadata for the orchestrator to honour.
    """
    g = linear(nodes, goal=goal)
    g.pattern = 'loop'
    if len(nodes) >= 2:
        g.add_edge(nodes[-1].id, nodes[0].id, kind=FEEDBACK,
                   label=u'refine (max {} cycles)'.format(max_cycles))
    return g
