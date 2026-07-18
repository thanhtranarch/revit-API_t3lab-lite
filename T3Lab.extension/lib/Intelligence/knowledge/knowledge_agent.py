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

_SYSTEM_PROMPT = (
    "Bạn là trợ lý tài liệu kỹ thuật của T3Lab bên trong Revit.\n"
    "NGUYÊN TẮC:\n"
    "- CHỈ trả lời dựa trên các trích đoạn tài liệu được cung cấp bên dưới.\n"
    "- Luôn ghi nguồn bằng chỉ số [n] ngay sau thông tin lấy từ trích đoạn [n].\n"
    "- Nếu các trích đoạn không đủ để trả lời, nói rõ là tài liệu không đề cập"
    " — TUYỆT ĐỐI không bịa.\n"
    "- Trả lời bằng đúng ngôn ngữ của câu hỏi (tiếng Việt hoặc tiếng Anh).\n"
    "- Trả lời ngắn gọn, đúng trọng tâm; dùng gạch đầu dòng khi liệt kê."
)


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
                     skills_block=''):
        """Returns (system_prompt, user_query)."""
        sys_parts = [_SYSTEM_PROMPT]
        if project_instructions:
            sys_parts.append("HƯỚNG DẪN DỰ ÁN:\n" + project_instructions)
        if skills_block:
            sys_parts.append(skills_block)

        lines = ["TRÍCH ĐOẠN TÀI LIỆU:"]
        for i, hit in enumerate(hits):
            page_note = (" — trang {}".format(hit['page'])
                         if hit.get('page') else "")
            lines.append("[{}] {}{}:\n{}".format(
                i + 1, hit.get('file', '?'), page_note,
                (hit.get('text') or '')[:_EXCERPT_CHARS]))
        lines.append("CÂU HỎI: " + (question or ''))
        return "\n\n".join(sys_parts), "\n\n".join(lines)

    # ── answer ────────────────────────────────────────────────────────────

    def answer(self, question, history, chat_fn,
               project_instructions='', skills_block='', top_k=_TOP_K):
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
            question, hits, project_instructions, skills_block)
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


def format_citation_line(citations):
    """Compact markdown sources line appended under the answer."""
    if not citations:
        return ""
    parts = []
    for c in citations:
        page = " tr.{}".format(c['page']) if c.get('page') else ""
        parts.append("[{}] {}{}".format(c['n'], c['file'], page))
    return "\n\n---\nNguồn: " + " · ".join(parts)
