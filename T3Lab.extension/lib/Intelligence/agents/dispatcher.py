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

try:
    from Intelligence.language import analyzer as _lang_analyzer
    HAS_LANGUAGE = True
except Exception:                                   # pragma: no cover
    HAS_LANGUAGE = False
    _lang_analyzer = None

SPECIALIST_NAMES = ('general', 'revit_data', 'revit_action',
                    'knowledge', 'comment',
                    'multi_doc', 'modeling', 'qa_check', 'export')

# The classification call sits IN FRONT of the real turn — every second it
# spends is a second before the user sees anything. It asks for one label in
# 60 tokens, so it has no business taking longer than this; past the limit the
# keyword/skill decision is a perfectly good answer. Without it the call
# inherited each provider's chat timeout: 60s on the cloud APIs and 180s on
# Ollama / LM Studio, whose long ceilings exist for real generations.
_CLASSIFY_TIMEOUT_MS = 6000

# ─── Keyword tables (folded lowercase) ───────────────────────────────────────
# Phrases are matched on the folded text (substring with word boundaries),
# single words on the token list.

_COMMENT_PHRASES = ('hoan thien cmt', 'hoan thien comment', 'xu ly comment',
                    'xu ly cmt', 'ghi chu ban ve', 'comment ban ve')
_COMMENT_WORDS = ('cmt', 'comment', 'comments', 'markup', 'markups', 'bluebeam')
# "Comments" is also a built-in Revit PARAMETER, and this stage runs first —
# so "lấy giá trị parameter Comments của element 12345" was handed to the PDF
# markup-resolution agent. A parameter context disarms the bare words; an
# attached annotated PDF still wins outright, and 'cmt'/'markup'/'bluebeam'
# are unambiguous enough to keep.
_COMMENT_PARAM_CONTEXT = ('parameter', 'parameters', 'thong so', 'tham so',
                          'gia tri', 'value', 'field', 'shared parameter')
_COMMENT_STRONG_WORDS = ('cmt', 'markup', 'markups', 'bluebeam')

_KNOWLEDGE_PHRASES = ('tieu chuan', 'quy dinh', 'quy chuan', 'tai lieu',
                      'huong dan', 'theo file', 'trong pdf', 'trong tai lieu',
                      'theo spec', 'design guide')
_KNOWLEDGE_WORDS = ('tcvn', 'qcvn', 'spec', 'specification', 'standard',
                    'standards', 'document', 'documents', 'docs', 'manual',
                    'catalogue', 'catalog')

_ACTION_PHRASES = ('doi ten', 'di chuyen', 'to mau', 'danh tag', 'gan tag',
                   'dat ten', 'chinh sua', 'cap nhat', 'thay doi',
                   'danh dau',
                   # View templates and worksets read as neither a build nor
                   # a query, so they used to fall through to `general`.
                   'ap view template', 'apply view template', 'ap template',
                   'gan view template', 'sang workset', 'chuyen workset',
                   'doi workset', 'set workset', 'move to workset',
                   # Splitting. As bare words 'cat'/'chia' annex ordinary
                   # Vietnamese ("cắt tóc", "chia sẻ file"), so the verb only
                   # counts next to the thing being split.
                   'chia duong', 'chia line', 'chia curve', 'chia element',
                   'chia thanh', 'cat tuong', 'cat dam', 'cat ong',
                   'cat duong', 'cat line',
                   # "điền thông số X cho Y" is a write; 'dien' alone is not
                   # usable ("diện tích" folds to "dien tich").
                   'dien thong so', 'dien gia tri', 'dien tham so')
_ACTION_WORDS = ('sua', 'xoa', 'tao', 'dat', 'gan', 'doi', 'them',
                 'rename', 'delete', 'create', 'move', 'set', 'change',
                 'update', 'modify', 'draw', 'place', 'tag', 'color',
                 'export', 'xuat', 'duplicate', 'copy', 'hide', 'isolate',
                 'highlight', 'split', 'xoay', 'rotate')

# Color commands usually name the color directly ("tô đỏ tường", "bôi xanh
# các cột") instead of saying "tô màu". The verbs are only unambiguous WITH
# their diacritics — "tô đỏ" folds to "to do", which would false-match
# English ("what to do...") — so these phrases are matched as substrings of
# the RAW lowercase text, before folding. Diacritic-less input ("to do
# tuong") stays ambiguous on purpose and falls through to the LLM stage.
_ACTION_COLOR_RAW = ('tô đỏ', 'tô xanh', 'tô vàng', 'tô cam', 'tô tím',
                     'tô hồng', 'tô đen', 'tô trắng', 'tô xám', 'tô nâu',
                     'bôi đỏ', 'bôi xanh', 'bôi vàng', 'bôi màu', 'bôi đen')

_DATA_PHRASES = ('bao nhieu', 'liet ke', 'danh sach', 'thong ke', 'canh bao',
                 'how many', 'kiem tra model', 'tinh trang model', 'thong tin',
                 # Quantity take-off and geometry extents are reads; both
                 # used to reach no keyword at all.
                 'khoi luong', 'take off', 'material quantity',
                 'kich thuoc bao', 'bounding box',
                 'lay gia tri', 'doc gia tri', 'gia tri parameter',
                 'get parameter', 'read parameter')
_DATA_WORDS = ('dem', 'count', 'list', 'warnings', 'warning', 'schedule',
               'statistics', 'health', 'info', 'takeoff', 'quantities')

# Multi-document: several models open at once (phrases only — the bare word
# "model" is far too generic to route on).
_MULTIDOC_PHRASES = ('so sanh model', 'so sanh 2 model', 'so sanh hai model',
                     'giua cac model', 'cac model dang mo', 'model dang mo',
                     'hai model', '2 model', 'nhieu model', 'model khac',
                     'chuyen sang model', 'chuyen model', 'doi model',
                     'switch model', 'switch document', 'compare model',
                     'compare models', 'both models', 'other model',
                     'open documents', 'cac file dang mo', 'file rvt khac',
                     'lien ket model', 'lien ket cac model',
                     # Opening/closing a model is document work, not an edit.
                     # "mở file rvt này" matched nothing; "model nào mở gần
                     # đây" matched the action word 'gan' (gắn) instead.
                     'mo file', 'mo model', 'mo project', 'mo ban ve',
                     'dong model', 'dong file', 'dong project',
                     'open file', 'open model', 'open project',
                     'close model', 'close file', 'close document',
                     'model gan day', 'file gan day', 'mo gan day',
                     'recent model', 'recent file', 'recent document',
                     'recent project')

# Geometry creation (image/text-to-model) — must outrank the generic
# action verbs ('tao', 'dung' folds...) so builds get the modeling budget.
_MODELING_PHRASES = ('dung model', 'dung nha', 'dung cong trinh',
                     'dung toa nha', 'dung lai', 'build model',
                     'build from image', 'image to model', 'model tu anh',
                     'model tu ban ve', 'pdf to model', 'tao tuong',
                     'dung tuong', 've tuong', 'tao level', 'tao luoi',
                     'tao grid', 'tao san', 'tao mai', 'place wall',
                     'create wall', 'create level', 'create grid',
                     # Joining geometry is modeling work and matched no
                     # keyword table at all before.
                     'join', 'unjoin', 'join geometry', 'noi hinh hoc')

# The literal table above can only list the exact wordings somebody thought
# of, and every build request it misses lands in `general` or `revit_action`
# — neither of which HAS a geometry-creation tool. The model then answers
# with the nearest thing it can see: "dựng hệ grid 5x5" was written into the
# model as a create_text_note reading "Grid 5x5". So build verb + geometry
# noun is matched as a pair instead, with only a classifier/quantity allowed
# between them ("dựng HỆ grid", "tạo HỆ lưới trục", "add 3 levels").
#
# Adjacency is what keeps this off the action specialist's turf: only the
# noun directly governed by the verb counts, so "tạo sheet", "tạo view" and
# "thêm tham số cho tường" (them → tham, not a geometry noun) still route to
# revit_action. "ve" (vẽ) is deliberately NOT a verb here — it folds
# together with the far more common "về" ("hỏi về tường"); the literal
# 've tuong' above covers the one wording that is worth the risk.
_BUILD_VERBS = ('tao', 'dung', 'them', 'dat', 'create', 'build', 'draw',
                'add', 'place', 'make', 'generate', 'model')
# Optional filler between verb and noun: Vietnamese classifiers, articles,
# and a count ("3 tang", "5 luoi").
_BUILD_FILLERS = ('he thong', 'he', 'cac', 'mot', 'a', 'an', 'the', 'new',
                  'lai', 'them', 'ra')
# Geometry the MODELING specialist owns. Sheets, views and schedules are
# deliberately absent — those are revit_action work.
_GEOMETRY_NOUNS = ('grids', 'grid', 'luoi truc', 'luoi', 'truc', 'levels',
                   'level', 'cao do', 'tang', 'tuong', 'walls', 'wall',
                   'san', 'slab', 'floors', 'floor', 'mai', 'roof',
                   'cot', 'columns', 'column', 'dam', 'beams', 'beam',
                   'phong', 'rooms', 'room', 'nha', 'toa nha',
                   'cong trinh', 'model',
                   # Hosted families. "đặt cửa vào tường" is a build, but
                   # 'dat' + 'cua' only reads that way when they are
                   # adjacent — "xóa cửa của tường" has no build verb.
                   'cua so', 'cua', 'doors', 'door', 'windows', 'window')
_MODELING_BUILD_RE = re.compile(
    r'(?:^|[^a-z0-9])'
    r'(?:' + '|'.join(_BUILD_VERBS) + r')'
    r'(?:\s+(?:' + '|'.join(_BUILD_FILLERS) + r'))?'
    r'(?:\s+\d+)?'
    r'\s+(?:' + '|'.join(_GEOMETRY_NOUNS) + r')'
    r'(?:[^a-z0-9]|$)')

# Direct exports — outranks action ('xuat'/'export' are action words too).
_EXPORT_PHRASES = ('xuat pdf', 'xuat dwg', 'xuat ifc', 'xuat anh',
                   'xuat sheet', 'xuat ban ve', 'in ra pdf',
                   'export pdf', 'export dwg', 'export ifc',
                   'export image', 'export sheet', 'export sheets',
                   'xuat du lieu', 'xuat room', 'xuat phong',
                   'export room', 'export data')

# QA / audit — outranks data so audits get the QA prompt + highlight tools.
_QA_PHRASES = ('kiem tra loi', 'kiem tra chinh ta', 'check chinh ta',
               'soat loi', 'audit model', 'model health', 'suc khoe model',
               'kiem tra warning', 'xu ly warning', 'don dep model',
               'model cleanup', 'purge', 'qa model', 'kiem tra tieu chuan',
               'check naming', 'kiem tra dat ten')

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
                 provider=None, skills_engine=None, allow_llm=True,
                 analysis=None):
        """Returns {"specialist", "skill", "source", "confidence"}.

        provider/skills_engine/allow_llm feed the optional LLM stage; when
        absent the decision is keyword-stage only.

        `analysis` is a pre-built `language.Utterance` for this text. When None
        one is built here. Its NORMALISED form is what the keyword stage reads,
        which is how teencode and typos reach these tables at all — "ktra
        warning ko" carries no keyword until it becomes "kiểm tra warning
        không". The raw text is still what the caller sends to the model; only
        this routing decision sees the repaired version.
        """
        utt = analysis
        if utt is None and HAS_LANGUAGE:
            try:
                utt = _lang_analyzer.analyze(text)
            except Exception:
                utt = None
        match_text = text
        if utt is not None and utt.normalized:
            match_text = utt.normalized

        folded = _fold_text(match_text)
        folded_masked = folded
        for mask in _NOUN_MASK_PHRASES:
            folded_masked = folded_masked.replace(mask, ' ')
        tokens = set(folded_masked.split())
        raw_lower = re.sub(r'\s+', ' ', (match_text or '').lower()).strip()

        decision = self._keyword_stage(folded, folded_masked, tokens,
                                       raw_lower, attached_pdf_annotated)
        if decision is None:
            decision = self._entity_stage(utt)
        if decision is not None:
            decision = self._apply_negation(decision, utt)
            skill = self._match_skill(text, skills_engine)
            decision['skill'] = skill
            return decision

        # The keyword stage was not sure. Before paying for a model round-trip
        # ahead of the real turn, check whether a skill matched STRONGLY — a
        # multi-word declared trigger is evidence in its own right, and the
        # skill's frontmatter already names the specialist to run it under.
        strong = self._strong_skill(text, skills_engine)
        if strong is not None:
            sid, spec = strong
            if spec:
                return {'specialist': spec, 'skill': sid,
                        'source': 'skill', 'confidence': 0.7}

        if allow_llm and provider is not None:
            llm = self._llm_stage(match_text, provider, skills_engine)
            if llm is not None:
                return self._apply_negation(llm, utt)

        return {'specialist': 'general', 'skill':
                self._match_skill(text, skills_engine),
                'source': 'default', 'confidence': 0.3}

    # ── stage 1b: entity signals ──────────────────────────────────────────

    def _entity_stage(self, utt):
        """A decision from the language analysis when no keyword matched.

        The keyword tables need a VERB. A perfectly clear question that happens
        to use none of them ("cửa sổ tầng 3?") matched nothing and went to the
        model to be classified — a round-trip to learn what the sentence
        already said. A named Revit category plus a question is a read request;
        nothing else here is confident enough to route on.
        """
        if utt is None or not utt.categories:
            return None
        if utt.is_question:
            return {'specialist': 'revit_data', 'source': 'entity',
                    'confidence': 0.6}
        return None

    # ── negation guard ────────────────────────────────────────────────────

    @staticmethod
    def _apply_negation(decision, utt):
        """Demote a write route when the sentence FORBIDS the action it names.

        "đừng xóa text note" matches the delete verb and lands on the action
        specialist, which is the one specialist equipped to do the thing the
        user just said not to do. The read specialist can answer ("I won't
        delete anything — here they are") and cannot act.
        """
        if utt is None or utt.is_safe_to_act:
            return decision
        if decision.get('specialist') not in ('revit_action', 'modeling',
                                              'export'):
            return decision
        out = dict(decision)
        out['specialist'] = 'revit_data'
        out['source'] = 'negation-demoted'
        out['confidence'] = min(out.get('confidence', 0.8), 0.6)
        return out

    # A skill scoring at least this much matched on a multi-word DECLARED
    # trigger (10/word + 3 for being declared), not a single guessed keyword.
    # See skills_engine.match_scored for the scale.
    _STRONG_SKILL_SCORE = 25.0

    def _strong_skill(self, text, skills_engine):
        """(skill_id, specialist) when one skill matched decisively, else None.

        Only used to skip the classification round-trip; a weak match still
        goes to the model.
        """
        if skills_engine is None:
            return None
        try:
            scored = skills_engine.match_scored(text)
        except Exception:
            return None
        if not scored or scored[0][1] < self._STRONG_SKILL_SCORE:
            return None
        # An ambiguous top-2 is not decisive — let the model arbitrate.
        if len(scored) > 1 and (scored[0][1] - scored[1][1]) < 5.0:
            return None
        sid = scored[0][0]
        try:
            return sid, skills_engine.specialist_for(sid)
        except Exception:
            return None

    # ── stage 1: keywords ─────────────────────────────────────────────────

    def _keyword_stage(self, folded, folded_masked, tokens, raw_lower,
                       attached_pdf_annotated):
        # 1. PDF markup workflow — most specific
        _comment_words = _COMMENT_WORDS
        if _has_phrase(folded, _COMMENT_PARAM_CONTEXT):
            _comment_words = _COMMENT_STRONG_WORDS
        if attached_pdf_annotated or _has_phrase(folded, _COMMENT_PHRASES) \
                or (tokens & set(_comment_words)):
            return {'specialist': 'comment', 'source': 'keyword',
                    'confidence': 0.9}
        # 2. Cross-model requests — before action ("đổi model" contains the
        #    action verb 'doi' but is a document switch, not an edit)
        if _has_phrase(folded, _MULTIDOC_PHRASES):
            return {'specialist': 'multi_doc', 'source': 'keyword',
                    'confidence': 0.85}
        # 3. Direct exports — before action ('xuat'/'export' are also
        #    action words; the export spec has the tighter budget + prompt)
        if _has_phrase(folded, _EXPORT_PHRASES):
            return {'specialist': 'export', 'source': 'keyword',
                    'confidence': 0.85}
        # 4. Geometry builds — before action ('tao tuong' would otherwise
        #    land in generic action with half the iteration budget, and
        #    without a single geometry-creation tool in its subset)
        if _has_phrase(folded, _MODELING_PHRASES) \
                or _MODELING_BUILD_RE.search(folded_masked):
            return {'specialist': 'modeling', 'source': 'keyword',
                    'confidence': 0.85}
        # 5. QA / audits — before data so audits get the QA role + tools
        if _has_phrase(folded, _QA_PHRASES):
            return {'specialist': 'qa_check', 'source': 'keyword',
                    'confidence': 0.8}
        # 6. Model modification — a write verb wins even when doc/count
        #    words are present ("đổi chiều cao theo tiêu chuẩn" → action)
        if _has_phrase(folded_masked, _ACTION_PHRASES) \
                or (tokens & set(_ACTION_WORDS)) \
                or any(p in raw_lower for p in _ACTION_COLOR_RAW):
            return {'specialist': 'revit_action', 'source': 'keyword',
                    'confidence': 0.8}
        # 7. Document questions
        if _has_phrase(folded, _KNOWLEDGE_PHRASES) \
                or (tokens & set(_KNOWLEDGE_WORDS)):
            return {'specialist': 'knowledge', 'source': 'keyword',
                    'confidence': 0.85}
        # 8. Read-only model questions
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
            '(PDF markup resolution), multi_doc (work across several open '
            'models), modeling (build geometry: levels/grids/walls/floors), '
            'qa_check (audit model: warnings/health/naming/spelling), '
            'export (export sheets to PDF/DWG/images), general (anything '
            'else).'
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
            # response_format is required now that Ollama's JSON grammar is
            # opt-in rather than hardcoded on — this stage json.loads() the
            # reply below. Cloud providers accept and ignore it.
            raw = provider.chat([], system, text or '', max_tokens=60,
                                temperature=0.0, model_override=fast,
                                response_format={"type": "json_object"},
                                timeout_ms=_CLASSIFY_TIMEOUT_MS)
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
        """Best-matching skill id, or None.

        `match()` is relevance-ranked (skills_engine.match_scored), so element
        0 is the strongest hit rather than whichever id sorts first.
        """
        matched = self.match_skills(text, skills_engine, limit=1)
        return matched[0] if matched else None

    def match_skills(self, text, skills_engine, limit=2):
        """Up to `limit` skill ids, best first.

        Two is the useful default: build_skills_block injects at most two
        bodies, so asking for more only to drop them wastes nothing but is
        also pointless.
        """
        if skills_engine is None:
            return []
        try:
            return list(skills_engine.match(text))[:limit]
        except Exception:
            return []
