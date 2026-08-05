# -*- coding: utf-8 -*-
"""
model_snapshot — cache a digest of the OPEN Revit model for chat-start grounding.

Deliberately NOT a source of training examples. A model's element counts are
transient and model-specific; training the weights to answer "how many walls? →
120" would teach the model to hallucinate a fixed number for every project. The
value here is a *cached digest* the assistant injects as CONTEXT at the start of
a chat, so answers begin already knowing the model — grounding, not training.

The read tools run through the injected `execute_tool` (server._execute_tool,
which marshals to Revit's main thread and is safe when Revit is busy). Pure
helpers (digest_to_grounding_text) are CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Model Snapshot"

import io
import json
import os
import time


def _snapshot_dir():
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    d = os.path.join(base, 'T3LabAI', 'model_snapshot')
    if not os.path.isdir(d):
        try:
            os.makedirs(d)
        except Exception:
            pass
    return d


def _snapshot_file(doc_key):
    safe = u''.join(c for c in u'{}'.format(doc_key or u'default')
                    if c.isalnum() or c in (u'-', u'_'))[:80] or u'default'
    return os.path.join(_snapshot_dir(), u'{}.json'.format(safe))


# ─── Build (Revit reads, injected) ──────────────────────────────────────────────

def build_digest(execute_tool):
    """Call read-only stats tools → a compact digest dict. {} on failure."""
    digest = {}
    try:
        stats = execute_tool('analyze_model_statistics', {}) or {}
        if isinstance(stats, dict) and 'element_counts' in stats:
            digest['element_counts'] = stats.get('element_counts') or {}
            digest['total_views'] = stats.get('total_views')
            digest['project'] = stats.get('project')
    except Exception:
        pass
    try:
        health = execute_tool('get_model_health', {}) or {}
        if isinstance(health, dict) and 'error' not in health:
            digest['health'] = health
    except Exception:
        pass
    digest['ts'] = time.strftime('%Y-%m-%d %H:%M')
    return digest


# ─── Pure: digest → grounding text ──────────────────────────────────────────────

def digest_to_grounding_text(digest):
    """A short context block for the system prompt. '' when nothing usable."""
    if not isinstance(digest, dict):
        return u''
    counts = digest.get('element_counts') or {}
    parts = [u'{}: {}'.format(k, v) for k, v in sorted(counts.items()) if v]
    if not parts and not digest.get('total_views'):
        return u''
    lines = [u'## Open model snapshot (cached {})'.format(
        digest.get('ts') or u'')]
    if digest.get('project'):
        lines.append(u'Project: {}'.format(digest['project']))
    if parts:
        lines.append(u'Element counts — {}.'.format(u'; '.join(parts)))
    if digest.get('total_views'):
        lines.append(u'Total views: {}.'.format(digest['total_views']))
    lines.append(u'(A snapshot from when the assistant was idle; re-query a tool '
                 u'for live numbers if precision matters.)')
    return u'\n'.join(lines)


# ─── Persistence ────────────────────────────────────────────────────────────────

def cache_digest(digest, doc_key):
    try:
        payload = json.dumps(digest, ensure_ascii=True, indent=2)
        if isinstance(payload, bytes):
            payload = payload.decode('ascii')
        with io.open(_snapshot_file(doc_key), 'w', encoding='utf-8') as f:
            f.write(payload)
        return True
    except Exception:
        return False


def load_digest(doc_key):
    try:
        path = _snapshot_file(doc_key)
        if os.path.exists(path):
            with io.open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}


# ─── Orchestration ──────────────────────────────────────────────────────────────

def run(execute_tool=None, doc_key=None, **_ignored):
    """Refresh + cache the model digest. Adds 0 training rows by design."""
    if execute_tool is None:
        try:
            from core.server import get_t3labai_server
            execute_tool = get_t3labai_server()._execute_tool
        except Exception as ex:
            return {'status': u'server unavailable: {}'.format(ex), 'added': 0}
    digest = build_digest(execute_tool)
    if not (digest.get('element_counts') or digest.get('total_views')):
        return {'status': 'no model data', 'added': 0}
    cache_digest(digest, doc_key)
    return {'status': 'cached', 'added': 0, 'cached': True}
