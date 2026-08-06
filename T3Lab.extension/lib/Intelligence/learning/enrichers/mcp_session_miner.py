# -*- coding: utf-8 -*-
"""
mcp_session_miner — turn recorded MCP teaching sessions into training data.

When teaching capture is on, the in-Revit server streams every external tool
call (Opus driving Revit through the MCP bridge) into
%APPDATA%/T3LabAI/training/mcp_sessions/<id>.jsonl. On t3lab_end_teaching the
server already converts the live buffer to an SFT trajectory; this enricher is
the SAFETY NET that ingests any session file (including ones the teacher never
closed) into the dataset, then marks it processed so it is not re-scanned.

Zero-cost and zero-LLM: it only reshapes files already on disk. The dataset's
content-hash dedup makes double-ingest (server + miner) harmless. Sessions with
no declared goal are set aside (renamed .nolabel) rather than trained on a
synthetic prompt — a bad target is worse than a deferred one.

Pure helpers are injectable so this is CPython-3 unit-testable. run() never raises.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "MCP Session Miner"

import os

SOURCE = 'mcp_teacher'

MIN_GOAL_CHARS = 3
_DONE_SUFFIX    = '.done'
_NOLABEL_SUFFIX = '.nolabel'


def _session_files(directory):
    """Unprocessed *.jsonl session files in `directory` (skips .done/.nolabel)."""
    out = []
    try:
        for name in sorted(os.listdir(directory)):
            if name.endswith('.jsonl'):
                out.append(os.path.join(directory, name))
    except Exception:
        pass
    return out


def run(session_dir=None, read_session=None, add_trajectory=None,
        rename=None, revit_version=None, max_files=20, **_ignored):
    """Ingest recorded MCP teaching sessions into the SFT dataset.

    All I/O injected (defaults wired to core.teaching + dataset) so tests drive
    it with fakes. Returns {status, added, processed}. Never raises.
    """
    try:
        if session_dir is None or read_session is None:
            from core import teaching as _t
            session_dir = session_dir or _t.session_dir()
            read_session = read_session or _t.read_session
        if add_trajectory is None:
            from Intelligence.learning import dataset as _ds
            add_trajectory = _ds.add_trajectory
        if rename is None:
            rename = os.rename
    except Exception as ex:
        return {'status': u'unavailable: {}'.format(ex), 'added': 0,
                'processed': 0}

    added = 0
    processed = 0
    try:
        for path in _session_files(session_dir)[:max_files]:
            goal, steps = read_session(path)
            if not steps:
                _safe_rename(rename, path, path + _DONE_SUFFIX)
                processed += 1
                continue
            if len(u'{}'.format(goal or u'').strip()) < MIN_GOAL_CHARS:
                # No goal declared — defer rather than train on a synthetic one.
                _safe_rename(rename, path, path + _NOLABEL_SUFFIX)
                continue
            try:
                ok, note = add_trajectory(
                    goal, steps, source=SOURCE, quality='teacher',
                    revit_version=revit_version)
                if ok and note == u'added':
                    added += 1
            except Exception:
                pass
            _safe_rename(rename, path, path + _DONE_SUFFIX)
            processed += 1
    except Exception as ex:
        return {'status': u'partial: {}'.format(ex), 'added': added,
                'processed': processed}

    return {'status': 'ok', 'added': added, 'processed': processed}


def _safe_rename(rename, src, dst):
    try:
        rename(src, dst)
    except Exception:
        pass
