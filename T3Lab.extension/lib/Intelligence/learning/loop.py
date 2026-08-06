# -*- coding: utf-8 -*-
"""
loop — the host-facing facade of the self-study layer.

Assembles the StudyEngine with the real enrichers, reads the live gate
(settings + activity + local-model availability), and exposes ONE throttled
entry, `on_idling_tick()`, that the app-level Idling hook in startup.py calls.
Each due tick runs exactly one enricher on a background daemon thread so the
Revit UI thread is never held (model reads inside enrichers marshal back through
the server's ExternalEvent, which is safe when Revit is busy).

Everything is guarded so a study failure can never disturb Revit. Off by default:
the `agents.self_study` setting must be turned on, and generative enrichers only
run when the ACTIVE provider is local (Ollama/LM Studio) — enforcing the
"local only, no API cost" contract.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Study Loop"

import threading

from Intelligence.learning import activity
from Intelligence.learning.study_engine import StudyEngine, Enricher, Gate

# The assistant must be quiet this long before a study cycle runs, and cycles
# are spaced at least this far apart even while Revit sits idle for hours.
IDLE_SECONDS     = 120
THROTTLE_SECONDS = 300

_engine = None
_engine_lock = threading.Lock()
_run_lock = threading.Lock()   # guards against overlapping background cycles
_running = False
_last_tick = 0.0


# ─── Gate inputs ────────────────────────────────────────────────────────────────

def _self_study_enabled():
    # Default ON: the local enrichers are zero-API-cost and heavily throttled
    # (IDLE_SECONDS/THROTTLE_SECONDS + hourly budget). The teacher enricher,
    # which DOES spend API tokens, is gated separately behind agents.opus_teacher.
    try:
        from config.settings import get_settings
        return bool(get_settings().get_agent_option('self_study', True))
    except Exception:
        return True


def _revit_version():
    try:
        from pyrevit import HOST_APP
        return int(HOST_APP.version)
    except Exception:
        return 2023


def _active_local_provider():
    """The active provider iff it is a LOCAL one (Ollama/LM Studio), else None.

    Generative study work is only allowed on a local model — that is the whole
    "no API cost" contract — so a cloud active provider reports 'no local model'
    and generative enrichers are skipped.
    """
    try:
        from Intelligence.llm_router import LLMRouter
        p = LLMRouter().get_active_provider()
        if p is not None and getattr(p, 'NAME', '') in ('ollama', 'lmstudio'):
            return p
    except Exception:
        pass
    return None


def _teacher_available():
    """True when the Opus-5 teacher may run: the opt-in switch is on AND the
    Claude provider has a key.

    Deliberately no HTTP here (the gate is read every cycle) — is_configured()
    is key-presence only. The teacher enricher itself resolves the Opus model
    (a possible /v1/models fetch) when it actually runs, on the study thread.
    """
    try:
        from config.settings import get_settings
        if not bool(get_settings().get_agent_option('opus_teacher', False)):
            return False
        from Intelligence.llm_router import LLMRouter
        p = LLMRouter().get_provider('claude')
        return p is not None and p.is_configured()
    except Exception:
        return False


def _teacher_chat_fn():
    """Return chat_fn(prompt)->answer using the Claude provider pinned to Opus,
    or None when unavailable. Resolving the Opus model may do one HTTP fetch —
    only ever called from the background study thread.
    """
    try:
        from Intelligence.llm_router import LLMRouter
        from Intelligence.learning.enrichers import opus_teacher
        provider = LLMRouter().get_provider('claude')
        if provider is None or not provider.is_configured():
            return None
        opus = None
        try:
            opus = provider.pick_quality_model()
        except Exception:
            opus = None

        def _chat(prompt, _p=provider, _m=opus):
            kw = {'max_tokens': opus_teacher._ANSWER_TOKENS}
            if _m:
                kw['model_override'] = _m
            return _p.chat([], opus_teacher.TEACHER_SYSTEM_PROMPT, prompt, **kw)
        return _chat
    except Exception:
        return None


# ─── Engine assembly ────────────────────────────────────────────────────────────

def _build_engine():
    from Intelligence.learning.enrichers import (telemetry_miner, api_facts,
                                                 model_snapshot, knowledge_refresh,
                                                 opus_teacher, mcp_session_miner)
    ver = _revit_version()

    def _mine():
        return telemetry_miner.run(revit_version=ver)

    def _api():
        return api_facts.run(revit_version=ver)

    def _snapshot():
        return model_snapshot.run(doc_key=_doc_key())

    def _knowledge():
        return knowledge_refresh.run()

    def _mcp_sessions():
        return mcp_session_miner.run(revit_version=ver)

    def _teacher():
        chat_fn = _teacher_chat_fn()
        if chat_fn is None:
            return {'status': 'no teacher', 'added': 0}
        return opus_teacher.run(chat_fn=chat_fn, revit_version=ver)

    return StudyEngine([
        Enricher('telemetry_miner', _mine, generative=False),
        Enricher('model_snapshot',  _snapshot, generative=False),
        Enricher('api_facts',       _api, generative=False),
        Enricher('knowledge_refresh', _knowledge, generative=True),
        # Ingest recorded MCP teaching sessions (Opus via Claude Desktop) into
        # the SFT corpus. Zero-cost file reshaping — always runnable.
        Enricher('mcp_session_miner', _mcp_sessions, generative=False),
        # DISTILLATION: opt-in, cloud teacher — runs even while the active
        # provider is local Qwen (it calls Claude directly). requires='teacher'.
        Enricher('opus_teacher',    _teacher, requires='teacher'),
    ])


def _doc_key():
    try:
        from core.server import get_t3labai_server
        srv = get_t3labai_server()
        getter = getattr(srv, '_doc_key', None) or getattr(srv, 'doc_key', None)
        if getter:
            return getter()
    except Exception:
        pass
    return 'default'


def get_engine():
    global _engine
    if _engine is None:
        with _engine_lock:
            if _engine is None:
                _engine = _build_engine()
    return _engine


def _make_gate():
    return Gate(
        enabled=_self_study_enabled(),
        idle=activity.is_idle(IDLE_SECONDS),
        has_local_model=(_active_local_provider() is not None),
        has_teacher=_teacher_available(),
        cancel=(lambda: not activity.is_idle(IDLE_SECONDS)),
    )


def run_one_cycle():
    """Run one study cycle now (background thread). Returns the result dict."""
    try:
        return get_engine().run_cycle(_make_gate())
    except Exception as ex:
        return {'ran': None, 'status': u'error: {}'.format(ex)}


# ─── Throttled tick (called from the Idling hook, UI thread) ─────────────────────

def on_idling_tick(now=None):
    """Cheap, non-blocking. Spawns a background cycle at most every THROTTLE.

    Safe to call as fast as Revit's Idling event fires — it early-returns unless
    a full throttle window has passed and no cycle is already running.
    """
    global _last_tick, _running
    import time as _t
    now = now if now is not None else _t.time()

    # Fast pre-checks that avoid spinning up a thread for nothing.
    if _running:
        return
    if (now - _last_tick) < THROTTLE_SECONDS:
        return
    if not _self_study_enabled():
        _last_tick = now
        return
    if not activity.is_idle(IDLE_SECONDS, now):
        return

    _last_tick = now

    def _work():
        global _running
        with _run_lock:
            if _running:
                return
            _running = True
        try:
            run_one_cycle()
            # Opt-in: once the corpus has grown enough, kick off a detached
            # fine-tune. Self-gated on agents.self_train_auto + a growth
            # threshold, so this is a cheap no-op on almost every cycle.
            try:
                from Intelligence.learning import trainer as _trainer
                _trainer.maybe_auto_train()
            except Exception:
                pass
        finally:
            with _run_lock:
                _running = False

    try:
        th = threading.Thread(target=_work)
        th.daemon = True
        th.start()
    except Exception:
        pass
