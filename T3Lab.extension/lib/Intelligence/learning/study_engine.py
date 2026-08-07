# -*- coding: utf-8 -*-
"""
study_engine — the idle-cycle brain of the self-study layer.

One `run_cycle(gate)` does ONE small unit of work: it checks the gate + a
rolling budget, picks the next enricher round-robin (skipping generative ones
when no local model is reachable), runs it, and records the outcome. The idle
loop (an app-level Idling hook in startup.py, or a window DispatcherTimer) calls
run_cycle repeatedly; keeping each call to one enricher means the loop never
holds the machine for long and always yields between units.

All policy — gating, budget accounting, round-robin, generative-skip — is pure
and driven by an injected clock + gate, so it is CPython-3 unit-testable without
Revit/WPF/LLM. The enrichers themselves are injected too.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Study Engine"

import time


class Enricher(object):
    """One named unit of self-study work.

    fn() -> dict (must include 'status'; 'added' optional).

    Resource requirement, decided by the scheduler before running:
      requires=None       — pure/zero-cost, always runnable.
      requires='local'    — needs a LOCAL LLM (skipped when none reachable).
                            `generative=True` is the backward-compatible alias.
      requires='teacher'  — needs a CLOUD teacher (Claude/Opus) AND the opt-in
                            switch; skipped when no teacher is available. This is
                            the only enricher class that can spend API tokens.
    """

    __slots__ = ('name', 'fn', 'generative', 'requires')

    def __init__(self, name, fn, generative=False, requires=None):
        self.name = name
        self.fn = fn
        # `generative=True` predates `requires` and meant "needs a local model".
        if requires is None and generative:
            requires = 'local'
        self.requires = requires
        # Keep the legacy attribute truthful for any external reader.
        self.generative = bool(generative or requires == 'local')


class Gate(object):
    """What the host tells the engine about the current moment.

    enabled          — the agents.self_study kill switch.
    idle             — assistant not busy AND composer empty (host's job).
    has_local_model  — a local Ollama/LM Studio model is reachable.
    cancel           — callable -> True to abort mid-cycle (user activity).
    """

    __slots__ = ('enabled', 'idle', 'has_local_model', 'has_teacher', 'cancel')

    def __init__(self, enabled=False, idle=False, has_local_model=False,
                 cancel=None, has_teacher=False):
        self.enabled = bool(enabled)
        self.idle = bool(idle)
        self.has_local_model = bool(has_local_model)
        # A cloud teacher (Claude/Opus) is reachable AND the opt-in switch is on.
        self.has_teacher = bool(has_teacher)
        self.cancel = cancel or (lambda: False)


class StudyEngine(object):
    """Round-robin scheduler over enrichers with a rolling hourly budget."""

    def __init__(self, enrichers, budget_per_hour=12, clock=None):
        self._enrichers = list(enrichers or [])
        self._budget = int(budget_per_hour)
        self._clock = clock or time.time
        self._idx = 0            # round-robin cursor
        self._runs = []          # timestamps of recent cycles (for the budget)

    # ── budget ───────────────────────────────────────────────────────────────
    def _prune_runs(self, now):
        cutoff = now - 3600.0
        self._runs = [t for t in self._runs if t >= cutoff]

    def budget_left(self):
        now = self._clock()
        self._prune_runs(now)
        return max(0, self._budget - len(self._runs))

    # ── selection ────────────────────────────────────────────────────────────
    def _runnable(self, e, has_local_model, has_teacher):
        """True when enricher `e`'s resource requirement is currently met."""
        if e.requires == 'local':
            return has_local_model
        if e.requires == 'teacher':
            return has_teacher
        return True

    def _next_enricher(self, has_local_model, has_teacher=False):
        """Advance the cursor to the next runnable enricher, or None.

        Skips enrichers whose resource requirement is unmet (a 'local' enricher
        with no local model, a 'teacher' enricher with no cloud teacher). Scans
        at most one full lap so a set with nothing runnable returns None instead
        of looping forever.
        """
        n = len(self._enrichers)
        if n == 0:
            return None
        for _ in range(n):
            e = self._enrichers[self._idx % n]
            self._idx = (self._idx + 1) % n
            if not self._runnable(e, has_local_model, has_teacher):
                continue
            return e
        return None

    # ── one cycle ────────────────────────────────────────────────────────────
    def run_cycle(self, gate):
        """Run at most one enricher. Returns a result dict; never raises."""
        if not gate.enabled:
            return {'ran': None, 'status': 'disabled'}
        if not gate.idle:
            return {'ran': None, 'status': 'busy'}
        try:
            if gate.cancel():
                return {'ran': None, 'status': 'cancelled'}
        except Exception:
            pass
        if self.budget_left() <= 0:
            return {'ran': None, 'status': 'budget_exhausted'}

        enricher = self._next_enricher(gate.has_local_model, gate.has_teacher)
        if enricher is None:
            return {'ran': None, 'status': 'no_runnable_enricher'}

        # Charge the budget for the attempt (even a failing enricher spent a
        # slice of the idle window) and run it, containing any failure.
        self._runs.append(self._clock())
        try:
            result = enricher.fn() or {}
        except Exception as ex:
            return {'ran': enricher.name, 'status': u'error: {}'.format(ex),
                    'added': 0}
        return {'ran': enricher.name,
                'status': result.get('status', 'ok'),
                'added': int(result.get('added', 0) or 0),
                'detail': result}
