# -*- coding: utf-8 -*-
"""
CommentAgent — the "tự động hoàn thiện cmt" pipeline.

analyze(pdf):
    1. extract PDF markup annotations (pdf_annots)
    2. trace the drawing to a Revit sheet in the open model(s)
       (sheet_matcher + revit_list_sheets/list_open_documents via the
       injected execute_tool)
    3. ONE batched LLM call turns the annotations into per-item
       resolution proposals (action_type/description/instruction)
    4. returns a report dict; execution happens later, per item, through
       the normal revit_action specialist (script.py button callbacks) —
       inheriting the confirm cards + TransactionGroup for free.

Transport-free and UI-free: provider/router and execute_tool are
injected, so the module is CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import json
import os
import re

from Intelligence.comments import pdf_annots, sheet_matcher

_MAX_ITEMS = 25            # annotations sent to the LLM in one batch
_PROPOSAL_MAX_TOKENS = 1600

_ACTION_TYPES = ('fix_parameter', 'move_element', 'fix_tag', 'fix_dimension',
                 'create_element', 'delete_element', 'place_note', 'info',
                 'manual')

_ANALYST_PROMPT = (
    "Bạn là chuyên gia BIM phân tích comment review trên bản vẽ.\n"
    "Với MỖI annotation dưới đây, đề xuất một phương án xử lý trong Revit.\n"
    "action_type chọn đúng một trong: fix_parameter, move_element, fix_tag, "
    "fix_dimension, create_element, delete_element, place_note, info, manual.\n"
    "- info: comment chỉ là câu hỏi/xác nhận, không cần sửa model.\n"
    "- manual: cần người xử lý, AI không đủ thông tin.\n"
    "description: 1 câu ngắn mô tả phương án, viết bằng ngôn ngữ của comment.\n"
    "instruction: câu lệnh chi tiết cho một AI agent thao tác trong Revit "
    "(nêu rõ sheet, đối tượng, hành động) — chỉ khi action_type là hành động; "
    "để chuỗi rỗng cho info/manual.\n"
    "TRẢ VỀ DUY NHẤT JSON: {\"items\": [{\"id\": \"a1\", \"action_type\": "
    "\"...\", \"description\": \"...\", \"instruction\": \"...\"}]}"
)


class CommentAgent(object):

    def analyze(self, pdf_path, execute_tool, provider, router,
                progress_cb=None, skills_block='', extra_context=''):
        """Full analysis pass. Returns the report dict (plan §3.3)."""
        report = {
            'pdf': pdf_path,
            'pdf_name': os.path.basename(pdf_path or ''),
            'items': [],
            'sheet_match': None,
            'model_open': False,
            'needs_switch': False,
            'active_doc': None,
            'open_docs': [],
            'partial_extraction': False,
            'error': None,
        }

        def _prog(msg):
            if progress_cb:
                try:
                    progress_cb(msg)
                except Exception:
                    pass

        # 1 — annotations
        _prog('Đang đọc markup trong PDF...')
        records, partial = pdf_annots.extract_annotations(pdf_path)
        report['partial_extraction'] = bool(partial)
        if not records:
            report['error'] = 'no_annotations'
            return report
        records = records[:_MAX_ITEMS]
        for rec in records:
            rec['matched_sheet'] = None
            rec['proposal'] = {'action_type': 'manual',
                               'description': '', 'instruction': ''}
            rec['status'] = 'pending'
        report['items'] = records

        # 2 — trace to a Revit sheet / open model
        _prog('Đang đối chiếu sheet với model Revit...')
        sheets = []
        try:
            res = execute_tool('revit_list_sheets', {}) or {}
            sheets = res.get('sheets', []) or []
        except Exception:
            sheets = []
        try:
            docs = execute_tool('list_open_documents', {}) or {}
            report['open_docs'] = docs.get('documents',
                                           docs.get('docs', [])) or []
        except Exception:
            pass
        try:
            proj = execute_tool('revit_get_project_info', {}) or {}
            report['active_doc'] = proj.get('name') or proj.get('title')
        except Exception:
            pass

        page_texts = pdf_annots.extract_page_texts(pdf_path)
        candidates = sheet_matcher.extract_sheet_candidates(
            page_texts, pdf_path)
        match = sheet_matcher.match_sheets(candidates, sheets,
                                           pdf_filename=pdf_path)
        report['sheet_match'] = match
        report['model_open'] = bool(match)
        # sheet not in the ACTIVE doc but other docs are open → suggest
        # switching before executing anything
        report['needs_switch'] = (not match
                                  and len(report['open_docs']) > 1)
        if match:
            for rec in records:
                rec['matched_sheet'] = match

        # 3 — one batched proposal call
        _prog('Đang đề xuất phương án xử lý...')
        proposals = self._propose(records, report, provider, router,
                                  skills_block, extra_context)
        for rec in records:
            prop = proposals.get(rec['id'])
            if prop:
                rec['proposal'] = prop
        return report

    # ── LLM proposal batch ────────────────────────────────────────────────

    def _propose(self, records, report, provider, router,
                 skills_block='', extra_context=''):
        """One chat call → {annot_id: proposal}. Empty dict on failure
        (records keep their 'manual' default)."""
        payload = []
        for rec in records:
            payload.append({
                'id': rec['id'],
                'page': rec['page'],
                'type': rec['subtype'],
                'author': rec.get('author', ''),
                'subject': rec.get('subject', ''),
                'content': (rec.get('content', '') or '')[:500],
            })
        ctx_lines = []
        match = report.get('sheet_match')
        if match:
            ctx_lines.append('Sheet khớp trong model: {} — {} (id {})'.format(
                match['number'], match['name'], match['id']))
        else:
            ctx_lines.append('CHƯA khớp được sheet nào trong model đang mở.')
        if report.get('active_doc'):
            ctx_lines.append('Model đang mở: {}'.format(report['active_doc']))
        if extra_context:
            ctx_lines.append(extra_context)

        system = _ANALYST_PROMPT
        if skills_block:
            system = system + '\n\n' + skills_block
        query = 'BỐI CẢNH:\n{}\n\nANNOTATIONS:\n{}'.format(
            '\n'.join(ctx_lines),
            json.dumps(payload, ensure_ascii=False, indent=1))

        raw = None
        try:
            if provider is not None:
                raw = provider.chat([], system, query,
                                    max_tokens=_PROPOSAL_MAX_TOKENS)
        except Exception:
            raw = None
        if not raw and router is not None:
            try:
                raw = router.chat([], system, query,
                                  max_tokens=_PROPOSAL_MAX_TOKENS)
            except Exception:
                raw = None
        if not raw:
            return {}

        try:
            m = re.search(r'\{.*\}', raw, re.DOTALL)
            data = json.loads(m.group(0)) if m else {}
        except Exception:
            return {}
        out = {}
        for item in data.get('items', []) or []:
            aid = item.get('id')
            action = item.get('action_type', 'manual')
            if action not in _ACTION_TYPES:
                action = 'manual'
            if aid:
                out[aid] = {
                    'action_type': action,
                    'description': (item.get('description') or '')[:300],
                    'instruction': (item.get('instruction') or '')[:600],
                }
        return out

    # ── execution instruction builders (used by script.py buttons) ───────

    def build_run_instruction(self, item, report):
        """Seed text for the revit_action specialist ("Run" button)."""
        match = item.get('matched_sheet') or report.get('sheet_match')
        sheet_note = ('trên sheet {} — {}'.format(
            match['number'], match['name']) if match
            else 'trong model đang mở (chưa xác định được sheet)')
        instruction = (item.get('proposal') or {}).get('instruction') or ''
        base = ('Xử lý comment review [{}] {} (trang PDF {}): "{}".'.format(
            item['id'], sheet_note, item.get('page'),
            (item.get('content') or item.get('subject') or '')[:300]))
        if instruction:
            base += ' Phương án đề xuất: ' + instruction
        base += (' Sau khi xử lý xong, báo lại kết quả ngắn gọn cho comment '
                 'này.')
        return base

    def build_note_instruction(self, item, report):
        """Seed text that records the markup as a TextNote ("Note" button)."""
        match = item.get('matched_sheet') or report.get('sheet_match')
        where = ('trên sheet {} ({})'.format(match['number'], match['name'])
                 if match else 'trên sheet phù hợp trong model')
        return ('Dùng create_text_note đặt một ghi chú {} với nội dung '
                '"[CMT-{}] {} — DANG XU LY", vị trí xếp dọc mép phải sheet, '
                'không đè lên khung tên.'.format(
                    where, item['id'].lstrip('a'),
                    (item.get('content') or item.get('subject') or '')[:200]))

    # ── markdown fallback rendering ───────────────────────────────────────

    def report_to_markdown(self, report):
        """Chat-renderable summary (used when the WPF card fails)."""
        lines = []
        match = report.get('sheet_match')
        lines.append('**Phân tích comment PDF: {}**'.format(
            report.get('pdf_name', '')))
        if match:
            lines.append('Sheet khớp: **{} — {}** (score {}).'.format(
                match['number'], match['name'], match['score']))
        else:
            lines.append('Chưa khớp được sheet nào trong model đang mở.'
                         + (' Model khác đang mở — cần switch document.'
                            if report.get('needs_switch') else ''))
        if report.get('partial_extraction'):
            lines.append('_(PDF nén một phần — có thể còn markup chưa đọc được.)_')
        lines.append('')
        lines.append('| # | Trang | Người ghi | Nội dung | Đề xuất |')
        lines.append('|---|---|---|---|---|')
        for item in report.get('items', []):
            prop = item.get('proposal') or {}
            lines.append('| {} | {} | {} | {} | {} |'.format(
                item['id'], item.get('page', ''),
                (item.get('author') or '')[:20],
                (item.get('content') or item.get('subject') or '')[:60]
                .replace('|', '/').replace('\n', ' '),
                (prop.get('description') or prop.get('action_type', ''))[:60]
                .replace('|', '/')))
        return '\n'.join(lines)
