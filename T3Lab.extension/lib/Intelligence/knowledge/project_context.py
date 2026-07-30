# -*- coding: utf-8 -*-
"""
project_context — build ONE project-wide BIM standards context, NotebookLM-style.

Implements the 4-step RAG cycle across EVERY linked folder at once, instead of
summarising one file at a time (context_digest.py):

  1. Ingestion  — documents are read and cut by MEANING (chunker.split_sections
                  keeps a clause with its heading), page-tagged.
  2. Indexing   — KnowledgeStore stores chunks in BM25 + optional embedding
                  vectors, each keyed to (file, page) for exact citation.
  3. Retrieval  — for each BIM standard TOPIC below, a hybrid search (keyword
                  BM25 fused with semantic vectors via RRF) pulls the most
                  relevant passages from ALL documents, not just one file.
  4. Generation — the passages go to the LLM under a strict "answer ONLY from
                  these excerpts" instruction, and every rule it writes carries
                  a [n] marker resolved to `file — p.N`.

The result is PROJECT_CONTEXT.md: the project's real naming codes, folder
structure, LOD and workflows in one place, each traceable to a source page.

UI-free and transport-free — `chat_fn(system, user)` is injected, so this
module is CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import io
import os
import time

OUT_NAME = 'PROJECT_CONTEXT.md'

_TOP_K = 8
_EXCERPT_CHARS = 2400
_MAX_CTX_CHARS = 40000
# A discipline-code table alone is ~700 tokens, so the old 700-token cap cut
# answers off mid-table. Same root cause as the CONTEXT.md truncation.
_MAX_TOKENS = 3000


# (section title, retrieval query). The query is deliberately keyword-rich:
# BM25 needs the vocabulary that actually appears in BIM standards, while the
# semantic channel handles paraphrases.
BIM_TOPICS = [
    (u"File & model naming",
     u"file naming convention model file name code prefix discipline code "
     u"project code originator volume level type role number example"),
    (u"Drawing & document numbering",
     u"drawing number sheet number document numbering series revision "
     u"suitability status code P01 C01 S0 S1 A1 example"),
    (u"Folder structure & CDE",
     u"folder structure directory project BIM folder central model link model "
     u"CDE common data environment WIP shared published archived"),
    (u"Model structure & splitting",
     u"model structure model split zone volume strategy discipline model "
     u"federated central file worksharing unit"),
    (u"Worksets & worksharing",
     u"workset naming worksets worksharing allocation link workset grid level"),
    (u"LOD / LOI / level of detail",
     u"LOD level of development level of detail LOI LOIN 100 200 300 350 400 "
     u"500 stage matrix element"),
    (u"Clash detection & coordination",
     u"clash detection matrix tolerance clash report coordination meeting "
     u"priority hard soft clearance"),
    (u"Submission & delivery workflow",
     u"submission workflow deliverable milestone approval RFA submit review "
     u"process step responsible"),
    (u"Export formats & deliverables",
     u"export format PDF DWG IFC NWD file format version deliverable naming "
     u"sheet size plot"),
    (u"Object properties & parameters",
     u"parameter attribute shared parameter object attribute matrix property "
     u"COBie data field"),
]

_NO_INFO = u'NO_USEFUL_INFO'

_SYSTEM = (
    u"You are a BIM Manager compiling a standards quick-reference for ONE "
    u"project.\n"
    u"MANDATORY RULES:\n"
    u"- Use ONLY the information in the EXCERPTS provided. Never infer, never "
    u"add outside knowledge, never import another company's standard.\n"
    u"- After EACH rule, cite the source as [n], matching the excerpt used.\n"
    u"- Reproduce codes, symbols, numbers and examples exactly as written "
    u"(never translate a code).\n"
    u"- Write in ENGLISH even when the excerpts are in another language, but "
    u"keep codes, identifiers and proper names verbatim.\n"
    u"- If the excerpts hold NOTHING on this topic, reply with exactly one "
    u"line: " + _NO_INFO + u"\n"
    u"- Where two documents conflict, give both with their [n] and prefer the "
    u"newer version/date.\n"
    u"- Reproduce code tables IN FULL — never abbreviate, never write '...', "
    u"never stop partway. Completeness matters more than brevity.\n"
    u"- NEVER guess what a code stands for. Give a meaning only where the "
    u"excerpts state it.\n"
    u"- Omit placeholder lines ('Not specified', 'Not mentioned', 'None') "
    u"entirely — an absent heading already means absent.\n"
    u"- Markdown bullets or tables. Length is not capped. No preamble, no "
    u"conclusion."
)


def _fmt_excerpts(hits):
    lines = []
    for i, h in enumerate(hits, start=1):
        page = h.get('page') or 0
        where = u"{0}{1}".format(h.get('file') or u'?',
                                 u" — page {0}".format(page) if page else u"")
        lines.append(u"[{0}] {1}:\n{2}".format(
            i, where, (h.get('text') or u'')[:_EXCERPT_CHARS]))
    return u"\n\n".join(lines)


def _sources_line(hits, used_only=True, text=u''):
    """`[n] file — p.N` list, optionally only the markers the answer cites."""
    out, seen = [], set()
    for i, h in enumerate(hits, start=1):
        if used_only and (u"[{}]".format(i) not in (text or u'')):
            continue
        key = (h.get('file'), h.get('page'))
        if key in seen:
            continue
        seen.add(key)
        page = h.get('page') or 0
        out.append(u"[{}] {}{}".format(
            i, h.get('file') or u'?',
            u" — p.{}".format(page) if page else u""))
    return u" · ".join(out)


def build_project_context(store, chat_fn, embedder=None, topics=None,
                          progress_cb=None, top_k=_TOP_K):
    """Run the retrieval+generation cycle for every topic.

    Returns {"markdown", "topics_found", "topics_total", "sources"}.
    `store` is a KnowledgeStore (already scanned); `chat_fn(system, user)`
    returns text or None.
    """
    topics = topics or BIM_TOPICS
    blocks = []
    found = 0
    all_sources = set()

    for title, query in topics:
        if progress_cb:
            try:
                progress_cb(title)
            except Exception:
                pass
        try:
            hits = store.search(query, top_k=top_k, embedder=embedder)
        except Exception:
            hits = []
        if not hits:
            continue
        answer = None
        if chat_fn:
            user = (u"CHỦ ĐỀ CẦN TỔNG HỢP: {}\n\nTRÍCH ĐOẠN TÀI LIỆU:\n{}"
                    .format(title, _fmt_excerpts(hits)[:_MAX_CTX_CHARS]))
            try:
                answer = chat_fn(_SYSTEM, user, _MAX_TOKENS)
            except Exception:
                answer = None
        answer = (answer or u'').strip()
        try:
            from Intelligence.knowledge.context_digest import strip_filler
            answer = strip_filler(answer)
        except Exception:
            pass
        flat = answer.upper().replace(' ', '_')
        # old Vietnamese sentinel still honoured for prompts/caches predating
        # the switch to English context files
        if not answer or _NO_INFO in flat or 'KHONG_CO_THONG_TIN' in flat:
            continue
        found += 1
        src = _sources_line(hits, used_only=True, text=answer)
        if not src:                       # answer cited nothing explicitly
            src = _sources_line(hits, used_only=False)
        for h in hits:
            if h.get('file'):
                all_sources.add(h['file'])
        blocks.append(u"## {0}\n\n{1}\n\n*Sources:* {2}\n".format(
            title, answer, src))

    header = [
        u"# Project BIM Context",
        u"",
        u"> Consolidated automatically from ALL documents linked to this "
        u"project (hybrid keyword + semantic retrieval; answers drawn only "
        u"from the retrieved excerpts, with page citations). **Do not edit by "
        u"hand** — a rescan overwrites this file. Always open the source "
        u"document to confirm before applying anything here.",
        u"",
        u"- Topics with data: **{0}/{1}**".format(found, len(topics)),
        u"- Documents cited: **{0}**".format(len(all_sources)),
        u"- Updated: {0}".format(time.strftime('%Y-%m-%d %H:%M:%S')),
        u"",
    ]
    if not blocks:
        header.append(u"_No standards content could be extracted. Check that "
                      u"the folder is linked and scanned (Rescan), and see "
                      u"the 'Documents that could NOT be read' section of "
                      u"each folder's CONTEXT.md for per-file reasons._")
    return {
        'markdown': u"\n".join(header + blocks),
        'topics_found': found,
        'topics_total': len(topics),
        'sources': sorted(all_sources),
    }


def write_project_context(out_dir, result, out_name=OUT_NAME):
    """Write the markdown to out_dir/PROJECT_CONTEXT.md. Returns path or None."""
    if not out_dir or not result:
        return None
    try:
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)
        path = os.path.join(out_dir, out_name)
        with io.open(path, 'w', encoding='utf-8', newline='\n') as f:
            f.write(result.get('markdown') or u'')
        return path
    except Exception:
        return None
