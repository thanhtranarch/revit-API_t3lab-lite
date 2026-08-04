# -*- coding: utf-8 -*-
"""
Graph-engineering layer for the T3Lab Assistant.

The assistant used to be a LINE: classify the message → pick one specialist →
run one agent loop → print one answer. That shape cannot express a request as
ordinary as "thống kê tường rồi xuất pdf sheet A-101", which is two goals with
a dependency between them; the line did the first one and dropped the second.

This package models a turn as a GRAPH instead — nodes that hold state and
behaviour, edges that define flow and dependency — and gives it the machinery
a graph needs: a planner that decomposes the goal, an executor that runs
independent branches together, a reducer that converges them, a verifier that
scores the result, and a trace that records all of it so the next turn can
adapt.

Module map (mirrors the reference architecture it was built from):

    primitives.py     Node · Edge · GraphState · GraphContext · NodeResult
    topology.py       TaskGraph + the standard patterns (linear, fan-out,
                      fan-in, diamond, hub-and-spoke, loop)
    router.py         ROUTER   — where does this work go?
    planner.py        PLAN     — goal → graph
    reducer.py        REDUCER  — many results → one answer
    verifier.py       VERIFIER — is the answer actually good?
    executor.py       EXECUTE  — run the graph, in parallel, surviving failures
    observability.py  OBSERVE  — spans, traces, JSONL on disk
    orchestrator.py   the loop that runs all five operations and adapts

Nothing here imports Revit or WPF. The caller injects a `runner` callable that
knows how to execute one node; that keeps the whole package testable under
CPython 3 (dev/test_graph_agents.py) and keeps Revit's threading rules where
they belong — in script.py.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

from Intelligence.graph.primitives import (      # noqa: F401
    Node, Edge, NodeResult, GraphState, GraphContext,
    AGENT, ROUTER, REDUCER, VERIFIER, MEMORY,
    SEQUENCE, FALLBACK, FEEDBACK,
    PENDING, RUNNING, DONE, FAILED, SKIPPED, CANCELLED,
)
from Intelligence.graph.topology import (        # noqa: F401
    TaskGraph, GraphCycleError,
    linear, fan_out, fan_in, diamond, hub_spoke, loop,
)

__all__ = [
    'Node', 'Edge', 'NodeResult', 'GraphState', 'GraphContext',
    'AGENT', 'ROUTER', 'REDUCER', 'VERIFIER', 'MEMORY',
    'SEQUENCE', 'FALLBACK', 'FEEDBACK',
    'PENDING', 'RUNNING', 'DONE', 'FAILED', 'SKIPPED', 'CANCELLED',
    'TaskGraph', 'GraphCycleError',
    'linear', 'fan_out', 'fan_in', 'diamond', 'hub_spoke', 'loop',
]
