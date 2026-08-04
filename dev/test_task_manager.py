# -*- coding: utf-8 -*-
"""
CPython 3 test harness for Intelligence/agents/task_manager.py.

Run:  python3 dev/test_task_manager.py
Exit 0 = all pass. Plain asserts, mirroring dev/test_assistant_llm.py.

task_manager is the parallel-execution substrate the assistant's task cards run
on. These lock down the guarantees the UI relies on: bounded parallelism, ONE
writer at a time (interleaved Revit transactions corrupt Undo), cooperative
cancel, and on_done firing only after result/error are in place.
"""
from __future__ import unicode_literals

import os
import sys
import threading
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

from Intelligence.agents import task_manager as TM

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def _drain(mgr, timeout=5.0):
    """Block until no task is active, or the timeout trips."""
    deadline = time.time() + timeout
    while mgr.active_count() > 0 and time.time() < deadline:
        time.sleep(0.01)


# ─── bounded parallelism ──────────────────────────────────────────────────────

def test_runs_in_parallel_but_bounded():
    print('[task_manager: bounded parallelism]')
    mgr = TM.AgentTaskManager(max_parallel=3)

    live = {'now': 0, 'peak': 0}
    lock = threading.Lock()

    def work(task):
        with lock:
            live['now'] += 1
            live['peak'] = max(live['peak'], live['now'])
        time.sleep(0.12)
        with lock:
            live['now'] -= 1
        return 'ok'

    t0 = time.time()
    for _ in range(3):
        mgr.submit('read', work, writer=False)
    _drain(mgr)
    elapsed = time.time() - t0

    check('all three ran at once', live['peak'] == 3, live['peak'])
    check('finished ~concurrently (well under 3x serial)', elapsed < 0.30,
          round(elapsed, 3))


def test_respects_max_parallel():
    print('[task_manager: max_parallel cap]')
    mgr = TM.AgentTaskManager(max_parallel=2)
    live = {'now': 0, 'peak': 0}
    lock = threading.Lock()

    def work(task):
        with lock:
            live['now'] += 1
            live['peak'] = max(live['peak'], live['now'])
        time.sleep(0.08)
        with lock:
            live['now'] -= 1

    for _ in range(5):
        mgr.submit('r', work)
    _drain(mgr)
    check('never more than 2 running at once', live['peak'] == 2, live['peak'])


# ─── writer serialization ─────────────────────────────────────────────────────

def test_writers_never_overlap():
    print('[task_manager: one writer at a time]')
    mgr = TM.AgentTaskManager(max_parallel=3)
    writers_live = {'now': 0, 'peak': 0}
    lock = threading.Lock()

    def writer(task):
        with lock:
            writers_live['now'] += 1
            writers_live['peak'] = max(writers_live['peak'], writers_live['now'])
        time.sleep(0.08)
        with lock:
            writers_live['now'] -= 1

    # Three writers + reads all submitted together; writers must still serialize.
    for _ in range(3):
        mgr.submit('write', writer, writer=True)
        mgr.submit('read', lambda t: time.sleep(0.02), writer=False)
    _drain(mgr)
    check('at most one writer ever ran concurrently',
          writers_live['peak'] == 1, writers_live['peak'])


# ─── cancellation ─────────────────────────────────────────────────────────────

def test_cancel_pending_fires_on_done_cancelled():
    print('[task_manager: cancel a pending task]')
    mgr = TM.AgentTaskManager(max_parallel=1)
    seen = {'done': []}

    def slow(task):
        time.sleep(0.15)

    def on_done(task):
        seen['done'].append((task.id, task.status))

    # First task occupies the single slot; second stays PENDING and is cancelled
    # before it can start.
    t1 = mgr.submit('busy', slow, on_done=on_done)
    t2 = mgr.submit('waiting', slow, on_done=on_done)
    ok = mgr.cancel(t2.id)
    check('cancel returns True', ok)
    _drain(mgr)

    statuses = dict(seen['done'])
    check('pending task ended CANCELLED',
          statuses.get(t2.id) == TM.CANCELLED, statuses)
    check('the running task still completed',
          statuses.get(t1.id) == TM.DONE, statuses)


def test_on_done_runs_after_result_is_set():
    print('[task_manager: on_done sees terminal result]')
    mgr = TM.AgentTaskManager(max_parallel=2)
    captured = {}

    def work(task):
        return 42

    def on_done(task):
        # status/result must already be published when on_done fires.
        captured['status'] = task.status
        captured['result'] = task.result

    mgr.submit('compute', work, on_done=on_done)
    _drain(mgr)
    check('status DONE inside on_done', captured.get('status') == TM.DONE,
          captured)
    check('result visible inside on_done', captured.get('result') == 42,
          captured)


def test_failure_is_contained():
    print('[task_manager: a failing task is FAILED, not fatal]')
    mgr = TM.AgentTaskManager(max_parallel=2)
    out = {}

    def boom(task):
        raise ValueError('kaboom')

    def fine(task):
        return 'ok'

    def on_done(task):
        out[task.title] = task.status

    mgr.submit('bad', boom, on_done=on_done)
    mgr.submit('good', fine, on_done=on_done)
    _drain(mgr)
    check('failing task marked FAILED', out.get('bad') == TM.FAILED, out)
    check('sibling task still DONE', out.get('good') == TM.DONE, out)


def test_clear_finished():
    print('[task_manager: clear_finished drops terminal tasks]')
    mgr = TM.AgentTaskManager(max_parallel=2)
    mgr.submit('a', lambda t: 'x')
    mgr.submit('b', lambda t: 'y')
    _drain(mgr)
    check('two tasks listed before clear', len(mgr.list_tasks()) == 2)
    mgr.clear_finished()
    check('registry empty after clear', len(mgr.list_tasks()) == 0)


# ─── parallel-eligibility policy ──────────────────────────────────────────────

def test_eligibility_policy():
    print('[task_manager: eligible_for_parallel guard]')
    E = TM.eligible_for_parallel
    check('multi read, enabled -> parallel',
          E(is_multi=True, node_count=2, has_writer=False, enabled=True))
    check('a writer present -> sequential',
          not E(is_multi=True, node_count=2, has_writer=True, enabled=True))
    check('disabled kill switch -> sequential',
          not E(is_multi=True, node_count=2, has_writer=False, enabled=False))
    check('single goal -> sequential',
          not E(is_multi=False, node_count=1, has_writer=False, enabled=True))
    check('multi flag but <2 nodes -> sequential',
          not E(is_multi=True, node_count=1, has_writer=False, enabled=True))


def main():
    test_runs_in_parallel_but_bounded()
    test_respects_max_parallel()
    test_writers_never_overlap()
    test_cancel_pending_fires_on_done_cancelled()
    test_on_done_runs_after_result_is_set()
    test_failure_is_contained()
    test_clear_finished()
    test_eligibility_policy()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All task-manager tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
