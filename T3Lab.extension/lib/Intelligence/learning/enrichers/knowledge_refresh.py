# -*- coding: utf-8 -*-
"""
knowledge_refresh — keep the project-standards digest current while idle.

Reuses Intelligence.knowledge.context_digest.build_context_file, which is
already INCREMENTAL (a doc whose fingerprint is unchanged is reused; the pass-2
readability fix makes a readable<->unreadable flip rebuild the overview). Running
it during idle means the "Project standards summary" the assistant grounds on is
fresh by the time the user asks — grounding value, no training rows emitted here
(the digest's own LLM summaries are the artifact).

Generative: needs a chat_fn, so the study engine only schedules it when a local
model is reachable. Never raises.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Knowledge Refresh"

SOURCE = 'knowledge_refresh'


def _active_knowledge_dirs():
    """Linked knowledge folders for the active project, or []. Guarded."""
    try:
        from config.project_store import ProjectStore
        ps = ProjectStore()
        pid = ps.get_active_project_id()
        if not pid:
            return []
        getter = getattr(ps, 'get_knowledge_dirs', None)
        if getter:
            return list(getter(pid) or [])
    except Exception:
        pass
    return []


def run(chat_fn=None, folders=None, build_fn=None, **_ignored):
    """Incrementally refresh the standards digest for each linked folder.

    Returns {status, added, refreshed}. added is 0 by design — the refreshed
    CONTEXT.md / RAG index is the artifact, consumed as grounding at chat time.
    """
    if build_fn is None:
        try:
            from Intelligence.knowledge.context_digest import build_context_file
            build_fn = build_context_file
        except Exception as ex:
            return {'status': u'digest unavailable: {}'.format(ex), 'added': 0}

    if chat_fn is None:
        try:
            from Intelligence.llm_router import LLMRouter
            provider = LLMRouter().get_active_provider()
            if provider is not None:
                chat_fn = lambda system, user, max_tokens=700: (
                    provider.chat([], system, user, max_tokens=max_tokens))
        except Exception:
            chat_fn = None

    dirs = folders if folders is not None else _active_knowledge_dirs()
    if not dirs:
        return {'status': 'no linked folders', 'added': 0, 'refreshed': 0}

    refreshed = 0
    for folder in dirs:
        try:
            build_fn(folder, chat_fn=chat_fn)
            refreshed += 1
        except Exception:
            continue
    return {'status': 'ok', 'added': 0, 'refreshed': refreshed}
