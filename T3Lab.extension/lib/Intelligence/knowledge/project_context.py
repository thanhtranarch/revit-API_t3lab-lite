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
_EXCERPT_CHARS = 1100
_MAX_CTX_CHARS = 12000


# (section title, retrieval query). The query is deliberately keyword-rich:
# BM25 needs the vocabulary that actually appears in BIM standards, while the
# semantic channel handles paraphrases.
BIM_TOPICS = [
    (u"Đặt tên file & model",
     u"file naming convention model file name code prefix discipline code "
     u"project code originator volume level type role number example"),
    (u"Đánh số bản vẽ & tài liệu",
     u"drawing number sheet number document numbering series revision "
     u"suitability status code P01 C01 S0 S1 A1 example"),
    (u"Cấu trúc thư mục & CDE",
     u"folder structure directory project BIM folder central model link model "
     u"CDE common data environment WIP shared published archived"),
    (u"Cấu trúc model & phân chia",
     u"model structure model split zone volume strategy discipline model "
     u"federated central file worksharing unit"),
    (u"Workset & worksharing",
     u"workset naming worksets worksharing allocation link workset grid level"),
    (u"LOD / LOI / mức độ chi tiết",
     u"LOD level of development level of detail LOI LOIN 100 200 300 350 400 "
     u"500 stage matrix element"),
    (u"Clash & phối hợp",
     u"clash detection matrix tolerance clash report coordination meeting "
     u"priority hard soft clearance"),
    (u"Quy trình giao nộp & submission",
     u"submission workflow deliverable milestone approval RFA submit review "
     u"process step responsible"),
    (u"Định dạng xuất & giao nộp file",
     u"export format PDF DWG IFC NWD file format version deliverable naming "
     u"sheet size plot"),
    (u"Thuộc tính & tham số đối tượng",
     u"parameter attribute shared parameter object attribute matrix property "
     u"COBie data field"),
]

_SYSTEM = (
    u"Bạn là BIM Manager, đang lập bản tra cứu tiêu chuẩn cho MỘT dự án.\n"
    u"NGUYÊN TẮC BẮT BUỘC:\n"
    u"- CHỈ dùng thông tin trong các TRÍCH ĐOẠN được cung cấp. Tuyệt đối không "
    u"suy diễn, không thêm kiến thức ngoài, không lấy chuẩn của công ty khác.\n"
    u"- Sau MỖI quy tắc, ghi chỉ số nguồn dạng [n] đúng với trích đoạn đã dùng.\n"
    u"- Giữ nguyên mã, ký hiệu, con số, ví dụ y hệt bản gốc (không dịch mã).\n"
    u"- Nếu các trích đoạn KHÔNG chứa thông tin cho chủ đề này, trả lời đúng "
    u"một dòng: KHONG_CO_THONG_TIN\n"
    u"- Nếu hai tài liệu mâu thuẫn, nêu cả hai kèm [n] và ưu tiên bản có "
    u"version/ngày mới hơn.\n"
    u"- Trình bày gạch đầu dòng hoặc bảng markdown, ngắn gọn, tối đa ~220 từ. "
    u"Không mở bài, không kết luận."
)


def _fmt_excerpts(hits):
    lines = []
    for i, h in enumerate(hits, start=1):
        page = h.get('page') or 0
        where = u"{}{}".format(h.get('file') or u'?',
                               u" — trang {}".format(page) if page else u"")
        lines.append(u"[{}] {}:\n{}".format(
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
                answer = chat_fn(_SYSTEM, user, 700)
            except Exception:
                answer = None
        answer = (answer or u'').strip()
        if not answer or 'KHONG_CO_THONG_TIN' in answer.upper().replace(' ', '_'):
            continue
        found += 1
        src = _sources_line(hits, used_only=True, text=answer)
        if not src:                       # answer cited nothing explicitly
            src = _sources_line(hits, used_only=False)
        for h in hits:
            if h.get('file'):
                all_sources.add(h['file'])
        blocks.append(u"## {}\n\n{}\n\n*Nguồn:* {}\n".format(title, answer, src))

    header = [
        u"# Project BIM Context",
        u"",
        u"> Tổng hợp tự động từ TOÀN BỘ tài liệu đã link vào project "
        u"(tìm kiếm hybrid từ khoá + ngữ nghĩa, trả lời chỉ dựa trên trích "
        u"đoạn, có dẫn nguồn trang). **Không sửa tay** — sẽ bị ghi đè khi "
        u"quét lại. Luôn mở file gốc để xác nhận trước khi áp dụng.",
        u"",
        u"- Chủ đề có dữ liệu: **{}/{}**".format(found, len(topics)),
        u"- Tài liệu được trích: **{}**".format(len(all_sources)),
        u"- Cập nhật: {}".format(time.strftime('%Y-%m-%d %H:%M:%S')),
        u"",
    ]
    if not blocks:
        header.append(u"_Chưa trích được nội dung tiêu chuẩn nào. Kiểm tra "
                      u"rằng thư mục đã được link và quét (Rescan), và tài "
                      u"liệu có lớp text (PDF scan cần OCR)._")
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
