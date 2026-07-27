# -*- coding: utf-8 -*-
"""
ContextDigest — build a human-readable CONTEXT.md for a linked folder.

When the user links an external folder as project knowledge, the RAG index
alone is invisible: you cannot open it, review it, or hand it to a colleague.
This module walks the linked folder, pulls a lead excerpt out of every
indexable document, and writes a single digest markdown INTO that folder
(default `<folder>/context/CONTEXT.md`) — so the folder carries its own
plain-text summary of what it contains, right next to the source files.

The digest is also picked up by the normal KnowledgeStore scan, so a query
that matches the summary can lead the model to the right source document.

Feedback-loop guard: the output subfolder is NEVER walked, so a rescan can
never summarise a previous digest into a new one.

Pure Python + guarded imports — importable under CPython 3 for tests.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import io
import json
import os
import re
import time

TEXT_EXTS = ('.txt', '.md')
PDF_EXT = '.pdf'
INDEXABLE_EXTS = TEXT_EXTS + (PDF_EXT,)

DEFAULT_OUT_SUBDIR = 'context'
DEFAULT_OUT_NAME = 'CONTEXT.md'
INDEX_NAME = 'context_index.json'   # machine-readable sidecar next to CONTEXT.md

_EXCERPT_CHARS = 400        # lead excerpt kept per document (no-LLM fallback)
_MAX_FILES = 300            # hard cap so a huge share can't hang the UI thread
_MAX_BYTES = 20 * 1024 * 1024   # skip files bigger than this
_LLM_INPUT_CHARS = 14000    # per-document text budget sent to the LLM


def _norm(p):
    return os.path.normcase(os.path.abspath(p))


def _clean(text, limit=_EXCERPT_CHARS):
    """Collapse whitespace and trim to `limit` chars on a word boundary."""
    words = (text or '').split()
    if not words:
        return ''
    out = ' '.join(words)
    if len(out) <= limit:
        return out
    cut = out[:limit]
    if ' ' in cut:
        cut = cut[:cut.rfind(' ')]
    return cut + '…'


def _lead_text(path, max_pages=3):
    """First readable text of a document, or '' when it yields nothing."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in TEXT_EXTS:
            with io.open(path, 'r', encoding='utf-8', errors='replace') as f:
                return f.read(8000)
        if ext == PDF_EXT:
            from Intelligence import rag_processor
            pages = rag_processor.extract_pdf_pages(path, max_pages=max_pages)
            return '\n'.join((t or '') for _p, t in (pages or []))
    except Exception:
        return ''
    return ''


def _full_text(path, limit=_LLM_INPUT_CHARS):
    """As much readable text as the LLM pass is allowed to see."""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in TEXT_EXTS:
            with io.open(path, 'r', encoding='utf-8', errors='replace') as f:
                return f.read(limit)
        if ext == PDF_EXT:
            from Intelligence import rag_processor
            pages = rag_processor.extract_pdf_pages(path, max_pages=200)
            out = []
            total = 0
            for _p, t in (pages or []):
                t = (t or '').strip()
                if not t:
                    continue
                out.append(t)
                total += len(t)
                if total >= limit:
                    break
            return '\n'.join(out)[:limit]
    except Exception:
        return ''
    return ''


# ─── LLM extraction (RAG text → the rules that actually matter) ───────────────
# A raw lead excerpt of a BEP appendix is a cover page — useless. What the
# assistant needs is the NAMING CODES, FILE FORMATS and WORKFLOWS buried in the
# middle of the document, so each file is passed through the LLM with a
# targeted extraction brief.

_EXTRACT_SYSTEM = (
    u"Bạn là chuyên gia BIM đọc tài liệu tiêu chuẩn dự án (BEP/IEP/guideline).\n"
    u"Nhiệm vụ: TRÍCH thông tin THỰC SỰ CÓ trong văn bản, phục vụ việc áp dụng "
    u"đúng chuẩn dự án.\n"
    u"Ưu tiên trích, nếu tài liệu có:\n"
    u"- Quy tắc ĐẶT TÊN (file, model, sheet, view, workset, family) + cấu trúc mã, "
    u"ý nghĩa từng trường, VÍ DỤ cụ thể.\n"
    u"- BẢNG MÃ: mã bộ môn/discipline, mã công ty, mã loại tài liệu, mã tầng/zone, "
    u"revision/suitability. Giữ nguyên dạng bảng `Mã — Ý nghĩa`.\n"
    u"- ĐỊNH DẠNG FILE & giao nộp: đuôi file, phiên bản phần mềm, khổ giấy, "
    u"cấu trúc thư mục.\n"
    u"- QUY TRÌNH (workflow): các bước, ai làm, mốc thời gian, điều kiện duyệt.\n"
    u"- NGƯỠNG/THAM SỐ: LOD, dung sai clash, tần suất phối hợp.\n"
    u"QUY TẮC:\n"
    u"- CHỈ dùng thông tin trong văn bản. TUYỆT ĐỐI không suy diễn, không bổ sung "
    u"kiến thức ngoài.\n"
    u"- Giữ nguyên mã/ký hiệu/con số y như bản gốc.\n"
    u"- Nếu văn bản không chứa thông tin thuộc nhóm nào thì BỎ QUA nhóm đó.\n"
    u"- Nếu cả tài liệu không có thông tin hữu ích, trả lời đúng một dòng: "
    u"KHONG_CO_THONG_TIN\n"
    u"- Trả lời bằng markdown gạch đầu dòng, ngắn gọn, tối đa ~250 từ. Không mở bài."
)

_SUMMARY_SYSTEM = (
    u"Bạn là BIM Manager. Dưới đây là tóm tắt đã trích từ nhiều tài liệu tiêu "
    u"chuẩn của MỘT dự án.\n"
    u"Hãy tổng hợp thành bản tra cứu nhanh cho cả dự án, gom theo đúng các mục: "
    u"**Đặt tên & mã**, **Định dạng file & giao nộp**, **Cấu trúc thư mục**, "
    u"**Quy trình/workflow**, **Ngưỡng & tham số**.\n"
    u"QUY TẮC: chỉ dùng thông tin đã cho; ghi kèm tên file nguồn trong ngoặc sau "
    u"mỗi quy tắc; mâu thuẫn giữa các tài liệu thì nêu rõ cả hai và ưu tiên bản "
    u"có version/ngày mới hơn. Không bịa. Markdown, tối đa ~400 từ."
)


def _default_chat_fn():
    """LLMRouter-backed chat seam, or None when no provider is usable."""
    try:
        from Intelligence.llm_router import LLMRouter
        router = LLMRouter()
    except Exception:
        return None

    def _chat(system_prompt, user_content, max_tokens=700):
        try:
            return router.chat([], system_prompt, user_content,
                               max_tokens=max_tokens)
        except Exception:
            return None
    return _chat


def _llm_extract(chat_fn, filename, text):
    """Ask the LLM for the rules inside one document. '' when nothing."""
    if not chat_fn or not text or not text.strip():
        return ''
    user = u"TÊN FILE: {}\n\nNỘI DUNG:\n{}".format(filename, text)
    try:
        out = chat_fn(_EXTRACT_SYSTEM, user, 700)
    except Exception:
        return ''
    out = (out or '').strip()
    if not out or 'KHONG_CO_THONG_TIN' in out.upper().replace(' ', '_'):
        return ''
    return out


def iter_documents(folder, out_dir=None):
    """Yield indexable file paths under `folder`, skipping the output dir."""
    skip = _norm(out_dir) if out_dir else None
    for dirpath, dirnames, filenames in os.walk(folder):
        if skip and _norm(dirpath).startswith(skip):
            dirnames[:] = []          # never descend into the digest folder
            continue
        # keep hidden/system folders out of the digest
        dirnames[:] = [d for d in dirnames if not d.startswith('.')]
        for fname in sorted(filenames):
            if os.path.splitext(fname)[1].lower() in INDEXABLE_EXTS:
                yield os.path.join(dirpath, fname)


def build_context_file(folder, out_subdir=DEFAULT_OUT_SUBDIR,
                       out_name=DEFAULT_OUT_NAME, max_files=_MAX_FILES,
                       progress_cb=None, use_llm=True, chat_fn=None):
    """Scan `folder` and write the digest markdown inside it.

    use_llm: run each document through the LLM to pull out the naming codes,
    file formats and workflows (RAG text in → rules out). Falls back to a
    plain lead excerpt per file when no provider is reachable, so the digest
    is always produced.

    Returns {"path", "files", "skipped", "folder", "llm"} or None when the
    folder is unusable / nothing could be written.
    """
    if not folder or not os.path.isdir(folder):
        return None
    out_dir = os.path.join(folder, out_subdir)
    try:
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)
    except Exception:
        return None

    if use_llm and chat_fn is None:
        chat_fn = _default_chat_fn()
    llm_on = bool(use_llm and chat_fn)

    rows, skipped, n, n_llm = [], 0, 0, 0
    for path in iter_documents(folder, out_dir):
        if n >= max_files:
            skipped += 1
            continue
        try:
            if os.path.getsize(path) > _MAX_BYTES:
                skipped += 1
                continue
        except Exception:
            skipped += 1
            continue
        if progress_cb:
            try:
                progress_cb(os.path.basename(path))
            except Exception:
                pass
        rel = os.path.relpath(path, folder)
        summary = ''
        if llm_on:
            summary = _llm_extract(chat_fn, rel, _full_text(path))
            if summary:
                n_llm += 1
        if not summary:
            summary = _clean(_lead_text(path))
            if not summary:
                skipped += 1
                continue
        rows.append((rel, summary))
        n += 1

    # consolidated cross-document view — the actual "project standard" answer
    overview = ''
    if llm_on and rows:
        if progress_cb:
            try:
                progress_cb('tổng hợp…')
            except Exception:
                pass
        joined = '\n\n'.join(u"### {}\n{}".format(r, s) for r, s in rows)
        try:
            overview = (chat_fn(_SUMMARY_SYSTEM, joined[:_LLM_INPUT_CHARS],
                                900) or '').strip()
        except Exception:
            overview = ''

    lines = [
        '# Context — {}'.format(os.path.basename(folder.rstrip('\\/')) or folder),
        '',
        '> Tự động tạo bởi T3Lab Assistant — trích các quy tắc (đặt tên, mã, định '
        'dạng file, quy trình) từ tài liệu trong thư mục này để trợ lý tra cứu '
        'nhanh. **Không sửa tay** (sẽ bị ghi đè khi quét lại). Luôn đối chiếu file '
        'gốc trước khi áp dụng.',
        '',
        '- Thư mục nguồn: `{}`'.format(folder),
        '- Số tài liệu đã xử lý: **{}**{}'.format(
            len(rows), ' (bỏ qua {})'.format(skipped) if skipped else ''),
        '- Cách trích: **{}**'.format(
            u'LLM đọc toàn văn ({}/{} tài liệu)'.format(n_llm, len(rows))
            if llm_on else u'trích đoạn thô (chưa bật LLM)'),
        '- Cập nhật: {}'.format(time.strftime('%Y-%m-%d %H:%M:%S')),
        '',
    ]
    if overview:
        lines.append('## ⭐ Tóm tắt chuẩn dự án (tổng hợp)')
        lines.append('')
        lines.append(overview)
        lines.append('')
    if rows:
        lines.append('## Danh mục tài liệu')
        lines.append('')
        for rel, _x in rows:
            lines.append('- `{}`'.format(rel))
        lines.append('')
        lines.append('## Chi tiết theo từng tài liệu')
        lines.append('')
        for rel, summary in rows:
            lines.append('### {}'.format(rel))
            lines.append('')
            lines.append(summary)
            lines.append('')
    else:
        lines.append('_Chưa có tài liệu .pdf/.txt/.md đọc được trong thư mục này._')
        lines.append('')

    out_path = os.path.join(out_dir, out_name)
    try:
        with io.open(out_path, 'w', encoding='utf-8', newline='\n') as f:
            f.write('\n'.join(lines))
    except Exception:
        return None

    # machine-readable sidecar so the UI can report exact counts without
    # re-walking a (possibly remote) folder on the UI thread
    try:
        meta = {
            'files': len(rows),
            'skipped': skipped,
            'llm': n_llm,
            'updated': time.strftime('%Y-%m-%d %H:%M:%S'),
            'docs': [r for r, _s in rows],
        }
        # ensure_ascii=True + decode: json.dumps returns bytes-str on py2,
        # and io.open(encoding=...) demands unicode (same as knowledge_store)
        payload = json.dumps(meta, ensure_ascii=True, indent=1)
        if isinstance(payload, bytes):
            payload = payload.decode('ascii')
        with io.open(os.path.join(out_dir, INDEX_NAME), 'w',
                     encoding='utf-8', newline='\n') as f:
            f.write(payload)
    except Exception:
        pass
    return {'path': out_path, 'files': len(rows), 'skipped': skipped,
            'folder': folder, 'llm': n_llm}


def read_context_stats(folder, out_subdir=DEFAULT_OUT_SUBDIR,
                       out_name=DEFAULT_OUT_NAME):
    """Read back an EXISTING digest without re-walking the folder.

    Cheap enough for the UI thread even when `folder` is a network share:
    one small sidecar read (or one markdown read as fallback for digests
    written before the sidecar existed).

    Returns {"files", "skipped", "llm", "updated", "path", "exists"};
    exists=False when the folder has never been scanned.
    """
    out_dir = os.path.join(folder or '', out_subdir)
    md_path = os.path.join(out_dir, out_name)
    stats = {'files': 0, 'skipped': 0, 'llm': 0, 'updated': '',
             'path': md_path, 'exists': False}

    idx = os.path.join(out_dir, INDEX_NAME)
    if os.path.isfile(idx):
        try:
            with io.open(idx, 'r', encoding='utf-8') as f:
                data = json.load(f)
            stats.update({
                'files': int(data.get('files') or 0),
                'skipped': int(data.get('skipped') or 0),
                'llm': int(data.get('llm') or 0),
                'updated': data.get('updated') or '',
                'exists': True,
            })
            return stats
        except Exception:
            pass

    # fallback: parse the numbers back out of CONTEXT.md
    if os.path.isfile(md_path):
        try:
            with io.open(md_path, 'r', encoding='utf-8',
                         errors='replace') as f:
                head = f.read(4000)
            stats['exists'] = True
            m = re.search(r'(?:Số tài liệu đã (?:xử lý|tóm tắt))'
                          r'[^*\d]*\*\*(\d+)\*\*', head)
            if m:
                stats['files'] = int(m.group(1))
            m = re.search(r'\(bỏ qua (\d+)\)', head)
            if m:
                stats['skipped'] = int(m.group(1))
            m = re.search(r'LLM đọc toàn văn \((\d+)/', head)
            if m:
                stats['llm'] = int(m.group(1))
            m = re.search(r'Cập nhật:\s*([0-9\-: ]+)', head)
            if m:
                stats['updated'] = m.group(1).strip()
        except Exception:
            pass
    return stats


def build_for_dirs(folders, progress_cb=None):
    """Build a digest for each folder. Returns the list of result dicts."""
    out = []
    for d in (folders or []):
        try:
            res = build_context_file(d, progress_cb=progress_cb)
        except Exception:
            res = None
        if res:
            out.append(res)
    return out
