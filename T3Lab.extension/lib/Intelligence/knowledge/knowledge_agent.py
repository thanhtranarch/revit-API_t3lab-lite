# -*- coding: utf-8 -*-
"""
KnowledgeAgent — retrieval + one chat call, with citations.

Deliberately tool-free so it works on EVERY provider, including small
local models: retrieve top-k excerpts from the knowledge store(s), build
a grounded prompt, make a single chat/stream call.

UI-free and transport-free: the caller supplies `chat_fn(system, query)`
(script.py passes its _stream_llm_turn seam), so this module stays
CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

_EXCERPT_CHARS = 900     # max chars of one excerpt shown to the model
_TOP_K = 6

# The rules were written in Vietnamese with an "always answer in English" line
# bolted on, so the model was told to answer in a language the surrounding
# instructions did not use. Both variants are spelled out now and the caller
# picks one.
_SYSTEM_PROMPT_VI = (
    "Bạn là trợ lý tài liệu kỹ thuật của T3Lab bên trong Revit.\n"
    "NGUYÊN TẮC:\n"
    "- CHỈ trả lời dựa trên các trích đoạn tài liệu được cung cấp bên dưới.\n"
    "- Luôn ghi nguồn bằng chỉ số [n] ngay sau thông tin lấy từ trích đoạn [n].\n"
    "- Nếu các trích đoạn không đủ để trả lời, nói rõ là tài liệu không đề cập"
    " — TUYỆT ĐỐI không bịa.\n"
    "- Luôn trả lời bằng tiếng Việt.\n"
    "- Trả lời ngắn gọn, đúng trọng tâm; dùng gạch đầu dòng khi liệt kê."
)

_SYSTEM_PROMPT_EN = (
    "You are T3Lab's technical-document assistant inside Revit.\n"
    "RULES:\n"
    "- Answer ONLY from the document excerpts provided below.\n"
    "- Always cite with the index [n] right after information taken from"
    " excerpt [n].\n"
    "- If the excerpts are not enough to answer, say plainly that the"
    " documents do not cover it — NEVER invent anything.\n"
    "- Always answer in English.\n"
    "- Be brief and to the point; use bullets when listing."
)

# Back-compat default for callers that do not pass a language.
_SYSTEM_PROMPT = _SYSTEM_PROMPT_EN


def get_system_prompt(viet=False):
    """Knowledge-agent system prompt in the requested language."""
    return _SYSTEM_PROMPT_VI if viet else _SYSTEM_PROMPT_EN


class KnowledgeAgent(object):
    """Args:
        store: primary KnowledgeStore (default: the active project-aware store).
        fallback_store: fills remaining top-k slots (default: global store
            when `store` is project-scoped).
        embedder: optional OllamaEmbedder for the semantic channel.
    """

    def __init__(self, store=None, fallback_store=None, embedder=None):
        if store is None:
            from Intelligence.knowledge.knowledge_store import (
                get_active_store, get_global_store)
            store = get_active_store()
            glob = get_global_store()
            if fallback_store is None and store is not glob:
                fallback_store = glob
        self.store = store
        self.fallback_store = fallback_store
        self.embedder = embedder

    # ── retrieval ─────────────────────────────────────────────────────────

    def retrieve(self, question, top_k=_TOP_K):
        """Project store first; global store fills the remaining slots."""
        hits = []
        if self.store is not None:
            try:
                hits = self.store.search(question, top_k=top_k,
                                         embedder=self.embedder)
            except Exception:
                hits = []
        if self.fallback_store is not None and len(hits) < top_k:
            try:
                extra = self.fallback_store.search(
                    question, top_k=top_k - len(hits), embedder=self.embedder)
                seen = set(h['key'] for h in hits)
                for h in extra:
                    if h['key'] not in seen:
                        hits.append(h)
            except Exception:
                pass
        return hits

    # ── prompting ─────────────────────────────────────────────────────────

    def build_prompt(self, question, hits, project_instructions='',
                     skills_block='', viet=False):
        """Returns (system_prompt, user_query), in the requested language."""
        sys_parts = [get_system_prompt(viet)]
        if project_instructions:
            sys_parts.append(("HƯỚNG DẪN DỰ ÁN:\n" if viet
                              else "PROJECT INSTRUCTIONS:\n")
                             + project_instructions)
        if skills_block:
            sys_parts.append(skills_block)

        lines = ["TRÍCH ĐOẠN TÀI LIỆU:" if viet else "DOCUMENT EXCERPTS:"]
        for i, hit in enumerate(hits):
            if hit.get('page'):
                page_note = ((" — trang {}" if viet else " — page {}")
                             .format(hit['page']))
            else:
                page_note = ""
            lines.append("[{}] {}{}:\n{}".format(
                i + 1, hit.get('file', '?'), page_note,
                (hit.get('text') or '')[:_EXCERPT_CHARS]))
        lines.append(("CÂU HỎI: " if viet else "QUESTION: ") + (question or ''))
        return "\n\n".join(sys_parts), "\n\n".join(lines)

    # ── answer ────────────────────────────────────────────────────────────

    def answer(self, question, history, chat_fn,
               project_instructions='', skills_block='', top_k=_TOP_K,
               viet=False):
        """Full pipeline. `chat_fn(system_prompt, user_query)` returns the
        response text (may stream internally) or None.

        Returns:
            {"status": "no_hits"}                        nothing retrieved
            {"status": "llm_failed"}                     chat_fn empty
            {"status": "done", "text", "citations"}      success
        """
        hits = self.retrieve(question, top_k=top_k)
        if not hits:
            return {"status": "no_hits"}

        system_prompt, query = self.build_prompt(
            question, hits, project_instructions, skills_block, viet=viet)
        text = None
        try:
            text = chat_fn(system_prompt, query)
        except Exception:
            text = None
        if not text or not (text or '').strip():
            return {"status": "llm_failed"}

        citations = []
        seen = set()
        for i, hit in enumerate(hits):
            marker = (hit.get('file', ''), hit.get('page', 0))
            if marker in seen:
                continue
            seen.add(marker)
            citations.append({
                'n':    i + 1,
                'file': hit.get('file', ''),
                'page': hit.get('page', 0),
                'path': hit.get('path', ''),
            })
        return {"status": "done", "text": text, "citations": citations}


def build_reference_block(hits, excerpt_chars=800, max_items=3):
    """Format retrieved excerpts as a compact grounding block for the tool-
    calling agent (general/data/action/... specialists), NOT the knowledge
    specialist.

    Returns a system-prompt fragment, or '' when there is nothing to inject.
    The header tells the agent to FOLLOW these documented standards when the
    request touches them (and name the source), but to ignore them when
    irrelevant — so a stray lexical hit on an unrelated command is harmless.
    """
    if not hits:
        return u''
    lines = [
        u"## Project reference material (from the knowledge base)",
        u"When the request involves a documented standard, dimension, code "
        u"or rule below, FOLLOW it and name the source file in your reply. "
        u"If these excerpts are not relevant to the request, ignore them. "
        u"Never invent a standard that is not shown here.",
    ]
    for i, hit in enumerate(hits[:max_items]):
        page_note = (u" — p.{}".format(hit['page'])
                     if hit.get('page') else u"")
        lines.append(u"[{}] {}{}:\n{}".format(
            i + 1, hit.get('file', '?'), page_note,
            (hit.get('text') or u'')[:excerpt_chars]))
    return u"\n\n".join(lines)


def format_citation_line(citations):
    """Compact markdown sources line appended under the answer."""
    if not citations:
        return ""
    parts = []
    for c in citations:
        page = " tr.{}".format(c['page']) if c.get('page') else ""
        parts.append("[{}] {}{}".format(c['n'], c['file'], page))
    return "\n\n---\nNguồn: " + " · ".join(parts)
