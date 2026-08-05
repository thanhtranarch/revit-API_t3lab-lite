# -*- coding: utf-8 -*-
"""
trainer — launch the out-of-Revit fine-tune job + decide when to auto-run it.

Weight training runs in CPython + GPU (tools/train/finetune_local.py), never in
Revit. This module is the bridge: it reports dataset/last-train status, decides
(purely, testably) whether an auto-train is warranted, and launches the trainer
as a DETACHED external process so it outlives — and never blocks — Revit.

Manual path: the `/train` chat command (script.py) calls stats()/launch().
Auto path: the idle loop calls maybe_auto_train() when agents.self_train_auto is
on.

Pure helper (should_auto_train) is CPython-3 testable. Everything else is
guarded and never raises.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Trainer Launch"

import io
import json
import os

# Auto-train only once the corpus has grown by at least this many NEW examples
# since the last run — retraining on a handful of new rows is pure waste.
MIN_NEW_FOR_AUTO = 200


def _repo_root():
    # …/T3Lab.extension/lib/Intelligence/learning/trainer.py → repo root
    p = os.path.abspath(__file__)
    for _ in range(5):
        p = os.path.dirname(p)
    return p


def _train_script():
    return os.path.join(_repo_root(), 'tools', 'train', 'finetune_local.py')


def _training_dir():
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI', 'training')


# ─── Status ─────────────────────────────────────────────────────────────────────

def dataset_stats():
    try:
        from Intelligence.learning import dataset as ds
        return ds.stats()
    except Exception:
        return {'count': 0, 'by_source': {}, 'by_quality': {}}


def last_train():
    """The last_train.json the trainer wrote, or {} if never trained."""
    try:
        path = os.path.join(_training_dir(), 'last_train.json')
        if os.path.exists(path):
            with io.open(path, 'r', encoding='utf-8') as f:
                return json.load(f) or {}
    except Exception:
        pass
    return {}


# ─── Pure decision ──────────────────────────────────────────────────────────────

def should_auto_train(enabled, count, last_count, running,
                      min_new=MIN_NEW_FOR_AUTO):
    """Whether the idle loop should kick off a fine-tune now.

    Conservative on purpose: the opt-in switch must be on, no train may be in
    flight, and the corpus must have grown by >= min_new since last time.
    """
    if not enabled or running:
        return False
    try:
        return (int(count) - int(last_count or 0)) >= int(min_new)
    except Exception:
        return False


# ─── Launch (detached) ──────────────────────────────────────────────────────────

def _find_python():
    try:
        from config.settings import get_settings
        cfg = get_settings().get_agent_option('self_train_python', None)
        if cfg:
            return cfg
    except Exception:
        pass
    return 'python'


_running = False


def is_running():
    return _running


def launch(force=False):
    """Spawn the trainer as a detached process. Returns (ok, note).

    Best-effort and non-blocking — training takes hours, so we never wait on it.
    A crash in the child never touches Revit.
    """
    global _running
    script = _train_script()
    if not os.path.exists(script):
        return False, u'trainer script not found: {}'.format(script)
    try:
        import subprocess
        args = [_find_python(), script]
        if force:
            args.append('--force')
        kwargs = {'cwd': _repo_root()}
        # Detach so the job survives Revit closing. DETACHED_PROCESS on Windows;
        # start_new_session elsewhere (dev boxes).
        flags = getattr(subprocess, 'DETACHED_PROCESS', 0) | \
            getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0)
        if flags:
            kwargs['creationflags'] = flags
        else:
            kwargs['start_new_session'] = True
        subprocess.Popen(args, **kwargs)
        _running = True
        return True, u'training launched'
    except Exception as ex:
        return False, u'{}'.format(ex)


def maybe_auto_train():
    """Called by the idle loop. Launches training iff warranted. {status}."""
    try:
        from config.settings import get_settings
        enabled = bool(get_settings().get_agent_option('self_train_auto', False))
    except Exception:
        enabled = False
    if not enabled:
        return {'status': 'auto-train off'}
    count = dataset_stats().get('count', 0)
    last_count = int(last_train().get('examples', 0) or 0)
    if not should_auto_train(enabled, count, last_count, is_running()):
        return {'status': 'not warranted', 'count': count,
                'last': last_count}
    ok, note = launch()
    return {'status': u'launched' if ok else u'launch failed: {}'.format(note)}
