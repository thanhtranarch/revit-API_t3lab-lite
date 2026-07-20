# -*- coding: utf-8 -*-
"""
AgentDispatcher — classify a user message onto a specialist agent.

Two stages:
  1. Keyword stage — pure rules on diacritic-folded text. High precision,
     zero cost, works offline. Anything ambiguous falls to...
  2. LLM stage (optional, M4+) — ONE tiny classification call that also
     matches skills. Local-model friendly (format json, max_tokens=60).

Anything still ambiguous routes to "general" — exactly today's behavior,
so a misfire can never be worse than the status quo.

Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import json
import re

from Intelligence.knowledge import vi_text

SPECIALIST_NAMES = ('general', 'revit_data', 'revit_action',
                    'knowledge', 'comment')

# ─── Keyword tables (folded lowercase) ───────────────────────────────────────
# Phrases are matched on the folded text (substring with word boundaries),
# single words on the token list.

_COMMENT_PHRASES = ('hoan thien cmt', 'hoan thien comment', 'xu ly comment',
                    'xu ly cmt', 'ghi chu ban ve', 'comment ban ve')
_COMMENT_WORDS = ('cmt', 'comment', 'comments', 'markup', 'markups', 'bluebeam')

_KNOWLEDGE_PHRASES = ('tieu chuan', 'quy dinh', 'quy chuan', 'tai lieu',
                      'huong dan', 'theo file', 'trong pdf', 'trong tai lieu',
                      'theo spec', 'design guide')
_KNOWLEDGE_WORDS = ('tcvn', 'qcvn', 'spec', 'specification', 'standard',
                    'standards', 'document', 'documents', 'docs', 'manual',
                    'catalogue', 'catalog')

_ACTION_PHRASES = ('doi ten', 'di chuyen', 'to mau', 'danh tag', 'gan tag',
                   'dat ten', 'chinh sua', 'cap nhat', 'thay doi')
_ACTION_WORDS = ('sua', 'xoa', 'tao', 'dat', 'gan', 'doi', 'them',
                 'rename', 'delete', 'create', 'move', 'set', 'change',
                 'update', 'modify', 'draw', 'place', 'tag', 'color',
                 'export', 'xuat', 'duplicate', 'copy', 'hide', 'isolate')

_DATA_PHRASES = ('bao nhieu', 'liet ke', 'danh sach', 'thong ke', 'canh bao',
                 'how many', 'kiem tra model', 'tinh trang model', 'thong tin')
_DATA_WORDS = ('dem', 'count', 'list', 'warnings', 'warning', 'schedule',
               'statistics', 'health', 'info')

# "bản vẽ" folds to "ban ve" — mask it so its "ve" never counts as the verb
# "vẽ" (draw), and "ban" never as "bán".
_NOUN_MASK_PHRASES = ('ban ve',)


def _fold_text(text):
    folded = vi_text.fold_diacritics(text or '').lower()
    return re.sub(r'\s+', ' ', folded).strip()


def _has_phrase(folded, phrases):
    for p in phrases:
        if re.search(r'(?:^|[^a-z0-9])' + re.escape(p) + r'(?:[^a-z0-9]|$)',
                     folded):
            return True
    return False


class AgentDispatcher(object):

    def classify(self, text, has_attachments=False, attached_pdf_annotated=False,
                 provider=None, skills_engine=None, allow_llm=True):
        """Returns {"specialist", "skill", "source", "confidence"}.

        provider/skills_engine/allow_llm feed the optional LLM stage; when
        absent the decision is keyword-stage only.
        """
        folded = _fold_text(text)
        folded_masked = folded
        for mask in _NOUN_MASK_PHRASES:
            folded_masked = folded_masked.replace(mask, ' ')
        tokens = set(folded_masked.split())

        decision = self._keyword_stage(folded, folded_masked, tokens,
                                       attached_pdf_annotated)
        if decision is not None:
            skill = self._match_skill(text, skills_engine)
            decision['skill'] = skill
            return decision

        if allow_llm and provider is not None:
            llm = self._llm_stage(text, provider, skills_engine)
            if llm is not None:
                return llm

        return {'specialist': 'general', 'skill':
                self._match_skill(text, skills_engine),
                'source': 'default', 'confidence': 0.3}

    # ── stage 1: keywords ─────────────────────────────────────────────────

    def _keyword_stage(self, folded, folded_masked, tokens,
                       attached_pdf_annotated):
        # 1. PDF markup workflow — most specific
        if attached_pdf_annotated or _has_phrase(folded, _COMMENT_PHRASES) \
                or (tokens & set(_COMMENT_WORDS)):
            return {'specialist': 'comment', 'source': 'keyword',
                    'confidence': 0.9}
        # 2. Model modification — a write verb wins even when doc/count
        #    words are present ("đổi chiều cao theo tiêu chuẩn" → action)
        if _has_phrase(folded_masked, _ACTION_PHRASES) \
                or (tokens & set(_ACTION_WORDS)):
            return {'specialist': 'revit_action', 'source': 'keyword',
                    'confidence': 0.8}
        # 3. Document questions
        if _has_phrase(folded, _KNOWLEDGE_PHRASES) \
                or (tokens & set(_KNOWLEDGE_WORDS)):
            return {'specialist': 'knowledge', 'source': 'keyword',
                    'confidence': 0.85}
        # 4. Read-only model questions
        if _has_phrase(folded, _DATA_PHRASES) or (tokens & set(_DATA_WORDS)):
            return {'specialist': 'revit_data', 'source': 'keyword',
                    'confidence': 0.8}
        return None

    # ── stage 2: one tiny LLM call ────────────────────────────────────────

    def _llm_stage(self, text, provider, skills_engine):
        """One small classification call. Returns decision dict or None."""
        catalog = ''
        try:
            if skills_engine is not None:
                items = skills_engine.get_catalog()
                if items:
                    catalog = '\nSkills:\n' + '\n'.join(
                        '- {}: {}'.format(s['id'], s.get('description', ''))
                        for s in items[:12])
        except Exception:
            catalog = ''
        system = (
            'Classify the Revit-assistant request into exactly one label:\n'
            'revit_data (read/count/list model), revit_action (modify model),'
            ' knowledge (question about documents/standards), comment '
            '(PDF markup resolution), general (anything else).'
            '{}\nReply ONLY JSON: {{"label": "...", "skill": "id-or-null"}}'
            .format(catalog))
        # Latency: classification is trivial — pin the provider's fastest
        # model when one is known (e.g. Haiku / *-mini) instead of paying the
        # active model's latency for a 60-token label.
        fast = None
        try:
            fast = provider.pick_fast_model()
        except Exception:
            fast = None
        try:
            raw = provider.chat([], system, text or '', max_tokens=60,
                                temperature=0.0, model_override=fast)
        except TypeError:
            try:
                raw = provider.chat([], system, text or '', 60)
            except Exception:
                return None
        except Exception:
            return None
        if not raw:
            return None
        try:
            m = re.search(r'\{.*\}', raw, re.DOTALL)
            data = json.loads(m.group(0)) if m else {}
        except Exception:
            return None
        label = data.get('label')
        if label not in SPECIALIST_NAMES:
            return None
        skill = data.get('skill')
        if skill in ('null', '', None):
            skill = None
        return {'specialist': label, 'skill': skill,
                'source': 'llm', 'confidence': 0.6}

    # ── skills ────────────────────────────────────────────────────────────

    def _match_skill(self, text, skills_engine):
        if skills_engine is None:
            return None
        try:
            matched = skills_engine.match(text)
            return matched[0] if matched else None
        except Exception:
            return None
