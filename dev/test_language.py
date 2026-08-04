# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the bilingual language layer (Intelligence/language).

Run:  python3 dev/test_language.py
Exit code 0 = all pass. Plain asserts, no external framework — same conventions
as dev/test_assistant_routing.py.

Almost every case below is a sentence that a Vietnamese BIM office types daily
and that the previous single-line `_is_viet` / flat keyword tables got wrong.
They fall into four families:

    detection    diacritic-less and code-switched Vietnamese was read as
                 English, so the assistant answered in the wrong language
    normalising  teencode ("ktra warning ko"), telex leftovers ("xoas") and
                 typos matched no keyword table at all
    collisions   folding diacritics merges distinct Vietnamese words. "từ"
                 (from) became "tủ" (cabinet), "đến" (to) became "đèn" (light),
                 "chớ" (don't) became "cho" (give). Each produced a confident
                 wrong route
    semantics    negation, question-vs-command, scope and risk — the four
                 facts that decide whether an action is safe to take at all
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_lang_test_')

from Intelligence.language import (                        # noqa: E402
    analyze, detect, is_vietnamese, normalize, resolve_ellipsis,
)
from Intelligence.language import entities, normalizer, semantics  # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─────────────────────────────────────────────────────────────────────────────
# Language detection
# ─────────────────────────────────────────────────────────────────────────────

def test_diacritic_less_vietnamese_is_not_english():
    """The original bug: no diacritics → read as English → replied in English."""
    for text in (u'xuat pdf sheet A',
                 u'thong ke tuong theo tang',
                 u'kiem tra warning trong model',
                 u'liet ke cac sheet dang mo'):
        verdict = detect(text)
        check('detect: "{}" is Vietnamese'.format(text[:28]),
              verdict['lang'] == 'vi', verdict)


def test_code_switched_vietnamese_stays_vietnamese():
    for text in (u'export cái sheet này ra pdf',
                 u'tô đỏ toàn bộ wall trong view hiện tại',
                 u'check giúp mình mấy cái warning'):
        verdict = detect(text)
        check('detect: code-switch "{}" is vi'.format(text[:26]),
              verdict['lang'] == 'vi', verdict)


def test_bim_jargon_is_not_english_evidence():
    """"sheet", "export", "level" are what every Vietnamese BIM user types."""
    verdict = detect(u'export sheet level grid workset')
    check('detect: pure jargon has no language', verdict['lang'] == 'unknown',
          verdict)


def test_real_english_is_detected():
    for text in (u'how many walls are there?',
                 u'please export all the sheets to pdf',
                 u'which level is this element on'):
        verdict = detect(text)
        check('detect: "{}" is English'.format(text[:26]),
              verdict['lang'] == 'en', verdict)


def test_accented_word_never_counts_as_english():
    """"đỏ" folds to "do", which is an English function word."""
    verdict = detect(u'tô đỏ toàn bộ tường')
    check('detect: accents block the English reading',
          verdict['en_score'] == 0.0, verdict)
    check('detect: not reported as code-switching', not verdict['mixed'],
          verdict)


def test_bare_tool_name_does_not_flip_the_language():
    check('detect: bare tool name is unknown',
          detect(u'BatchOut')['lang'] == 'unknown', detect(u'BatchOut'))
    check('is_vietnamese honours the session default',
          is_vietnamese(u'BatchOut', default=True) is True)
    check('is_vietnamese honours an English session',
          is_vietnamese(u'BatchOut', default=False) is False)


def test_reply_language_falls_back_rather_than_guessing():
    utt = analyze(u'BatchOut')
    check('reply_language: keeps the session language',
          utt.reply_language(fallback='vi') == 'vi', utt.lang)
    check('reply_language: confident input wins',
          analyze(u'có bao nhiêu tường').reply_language('en') == 'vi')


# ─────────────────────────────────────────────────────────────────────────────
# Normalisation
# ─────────────────────────────────────────────────────────────────────────────

def test_teencode_is_expanded():
    cases = [
        (u'ktra warning ko', u'kiểm tra', u'không'),
        (u'xuat pdf dc ko', u'được', u'không'),
        (u'bn cai cua so', u'bao nhiêu', None),
        (u'giup mk tke tuong', u'thống kê', None),
    ]
    for text, must, must2 in cases:
        out = normalize(text)
        ok = must in out and (must2 is None or must2 in out)
        check('normalize: "{}"'.format(text), ok, out)


def test_repeated_characters_collapse():
    check('normalize: okkkkk', u'ok' in normalize(u'okkkkk').lower(),
          normalize(u'okkkkk'))


def test_telex_leftovers_are_stripped():
    check('normalize: "xoas" -> "xoa"', u'xoa' in normalize(u'xoas het tuong'),
          normalize(u'xoas het tuong'))
    check('normalize: a real word ending in s is untouched',
          u'walls' in normalize(u'delete walls'), normalize(u'delete walls'))


def test_typo_repair_is_conservative():
    check('typo: "deleet" -> "delete"',
          u'delete' in normalize(u'deleet the walls'),
          normalize(u'deleet the walls'))
    # Guard list — these are ordinary words one edit from a lexicon entry.
    for word in (u'sweet', u'wells', u'small', u'still'):
        out = normalize(word)
        check('typo: "{}" is left alone'.format(word), out == word, out)


def test_typo_repair_never_touches_accented_words():
    text = u'sửa chiều cao tường'
    check('typo: accented text is untouched', normalize(text) == text,
          normalize(text))


def test_protected_spans_survive_normalisation():
    for text in (u'xuat sheet A-101 ra pdf',
                 u'mo file C:\\Projects\\Toa nha A.rvt',
                 u'dat chieu cao 3000mm',
                 u'doi ten thanh "Tuong ngoai 200"'):
        out = normalize(text)
        for token in (u'A-101', u'C:\\Projects\\Toa nha A.rvt', u'3000mm',
                      u'"Tuong ngoai 200"'):
            if token in text:
                check('protect: "{}" survives'.format(token), token in out, out)


def test_singular_domain_words_are_not_their_own_typo():
    """"sheet" is correct; listing only "sheets" made it repair itself."""
    for word in (u'sheet', u'level', u'view', u'floor', u'stair'):
        out = normalize(u'export {} A'.format(word))
        check('normalize: "{}" unchanged'.format(word), word in out, out)


# ─────────────────────────────────────────────────────────────────────────────
# Entities — the folding collisions
# ─────────────────────────────────────────────────────────────────────────────

def test_window_beats_door():
    check('entities: "cửa sổ" is Windows, not Doors',
          entities.extract_categories(u'có bao nhiêu cửa sổ') == ['Windows'],
          entities.extract_categories(u'có bao nhiêu cửa sổ'))
    check('entities: "cua so" (no diacritics) is still Windows',
          entities.extract_categories(u'bao nhieu cua so') == ['Windows'],
          entities.extract_categories(u'bao nhieu cua so'))


def test_folding_collisions_do_not_produce_categories():
    collisions = [
        (u'dựng model từ ảnh',            'Casework'),      # từ vs tủ
        (u'đến tầng 3 kiểm tra',          'LightingFixtures'),  # đến vs đèn
        (u'chiều cao của tường là bao nhiêu', 'Doors'),      # của vs cửa
        (u'thế nào là tường chịu lực',    'Tags'),           # thế vs thẻ
    ]
    for text, forbidden in collisions:
        cats = entities.extract_categories(text)
        check('entities: "{}" is not {}'.format(text[:24], forbidden),
              forbidden not in cats, cats)


def test_real_categories_still_resolve():
    cases = [
        (u'thống kê tường', 'Walls'),
        (u'đếm cột bê tông', 'Columns'),
        (u'liệt kê dầm', 'StructuralFraming'),
        (u'count the windows', 'Windows'),
        (u'tag all rooms', 'Rooms'),
        (u'kiểm tra lan can', 'Railings'),
        (u'xoa het text note', 'TextNotes'),
    ]
    for text, expected in cases:
        cats = entities.extract_categories(text)
        check('entities: "{}" -> {}'.format(text[:24], expected),
              expected in cats, cats)


def test_colors_need_their_diacritics():
    check('colors: "tô đỏ" is red',
          entities.extract_colors(u'tô đỏ tường') == ['red'],
          entities.extract_colors(u'tô đỏ tường'))
    check('colors: "tìm" is not purple',
          entities.extract_colors(u'tìm tường') == [],
          entities.extract_colors(u'tìm tường'))
    check('colors: "mau do" resolves without diacritics',
          entities.extract_colors(u'to mau do tuong') == ['red'],
          entities.extract_colors(u'to mau do tuong'))
    check('colors: "xanh dương" is blue, not green',
          entities.extract_colors(u'tô xanh dương') == ['blue'],
          entities.extract_colors(u'tô xanh dương'))


def test_levels_and_sheets():
    lv = entities.extract_levels(u'có bao nhiêu cửa ở tầng 3')
    check('levels: "tầng 3"', [l['label'] for l in lv] == ['Level 3'], lv)
    check('levels: basement',
          [l['kind'] for l in entities.extract_levels(u'tầng hầm 2')]
          == ['basement'], entities.extract_levels(u'tầng hầm 2'))

    sh = entities.extract_sheets(u'xuất sheet A-101 và A-102')
    check('sheets: numbers found', sh['numbers'] == ['A-101', 'A-102'], sh)
    check('sheets: prefix from "sheet A"',
          entities.extract_sheets(u'xuất các sheet A')['prefixes'] == ['A'],
          entities.extract_sheets(u'xuất các sheet A'))
    check('sheets: a Vietnamese word is not a prefix',
          entities.extract_sheets(u'export cái sheet này')['prefixes'] == [],
          entities.extract_sheets(u'export cái sheet này'))


def test_measures_normalise_to_millimetres():
    m = entities.extract_measures(u'đặt chiều cao 3m và bề dày 200mm')
    check('measures: two found', len(m) == 2, m)
    check('measures: 3m -> 3000mm', m[0]['mm'] == 3000.0, m)
    check('measures: 200mm stays', m[1]['mm'] == 200.0, m)


def test_quantity_intent():
    q = entities.extract_quantity(u'tô đỏ tất cả tường')
    check('quantity: "tất cả" is all', q['all'] is True, q)
    q2 = entities.extract_quantity(u'chọn 5 cái cột')
    check('quantity: explicit count', q2['count'] == 5, q2)
    q3 = entities.extract_quantity(u'tường cao lớn hơn 3000')
    check('quantity: comparator', q3['comparator'] == ('>', 3000.0), q3)


# ─────────────────────────────────────────────────────────────────────────────
# Semantics
# ─────────────────────────────────────────────────────────────────────────────

def test_prohibitions_are_detected():
    for text in (u'đừng xóa text note',
                 u'chớ động vào tường',
                 u'không cần xuất pdf',
                 u'do not delete the walls',
                 u'no need to export'):
        neg = semantics.detect_negation(text)
        check('negation: "{}" is prohibitive'.format(text[:26]),
              neg['prohibitive'], neg)


def test_ordinary_sentences_are_not_prohibitions():
    """Every one of these folds onto a prohibitive marker."""
    for text in (u'cho tôi xem danh sách sheet',   # cho vs chớ
                 u'dựng model từ bản vẽ',           # dựng vs đừng
                 u'dùng view template nào',         # dùng vs đừng
                 u'chữa lỗi chính tả',              # chữa vs chưa
                 u'xóa được không?'):               # question, not an order
        neg = semantics.detect_negation(text)
        check('negation: "{}" is NOT prohibitive'.format(text[:26]),
              not neg['prohibitive'], neg)


def test_question_particle_is_not_a_negation():
    neg = semantics.detect_negation(u'xóa được không?')
    check('negation: trailing "không" is a question particle',
          not neg['negated'], neg)
    check('modality: it is a question',
          semantics.detect_modality(u'xóa được không?') == 'question')


def test_modality():
    cases = [
        (u'có bao nhiêu tường?', 'question'),
        (u'how many walls', 'question'),
        (u'hãy xóa tất cả text note', 'command'),
        (u'please export the sheets', 'command'),
        (u'đừng xóa text note', 'command'),
        (u'ok', 'confirmation'),
        (u'đồng ý', 'confirmation'),
    ]
    for text, expected in cases:
        got = semantics.detect_modality(text)
        check('modality: "{}" -> {}'.format(text[:24], expected),
              got == expected, got)


def test_sheet_prefix_is_not_a_question_particle():
    """"sheet A" ends in a bare "a", which used to read as a question tail."""
    check('modality: "xuat pdf sheet A" is not a question',
          semantics.detect_modality(u'xuat pdf sheet A') != 'question',
          semantics.detect_modality(u'xuat pdf sheet A'))


def test_scope():
    cases = [
        (u'tô đỏ tường đang chọn', 'selection'),
        (u'tô đỏ tường trong view này', 'view'),
        (u'tô đỏ toàn bộ tường', 'project'),
        (u'tô đỏ tường', 'unspecified'),
    ]
    for text, expected in cases:
        got = semantics.detect_scope(text)
        check('scope: "{}" -> {}'.format(text[:26], expected),
              got == expected, got)


def test_scope_marker_is_not_also_a_category():
    utt = analyze(u'xóa text note trong view này')
    check('scope: "view này" does not add the Views category',
          utt.categories == ['TextNotes'], utt.categories)


def test_risk_levels():
    cases = [
        (u'xóa toàn bộ text note', 'destructive'),
        (u'purge unused families', 'irreversible'),
        (u'đồng bộ với central', 'publish'),
        (u'liệt kê các sheet', 'none'),
    ]
    for text, expected in cases:
        got = semantics.detect_risk(text)['level']
        check('risk: "{}" -> {}'.format(text[:26], expected),
              got == expected, got)


def test_sequencer_splitting():
    check('split: "rồi" separates two goals',
          len(semantics.split_sequenced(
              u'thống kê tường rồi xuất pdf sheet A')) == 2,
          semantics.split_sequenced(u'thống kê tường rồi xuất pdf sheet A'))
    check('split: a noun list is NOT split',
          len(semantics.split_sequenced(u'thống kê tường, cột, dầm')) == 1,
          semantics.split_sequenced(u'thống kê tường, cột, dầm'))
    check('split: numbered plan',
          len(semantics.split_sequenced(
              u'1. thống kê tường\n2. xuất pdf')) == 2,
          semantics.split_sequenced(u'1. thống kê tường\n2. xuất pdf'))
    check('split: "và" is left to the planner',
          len(semantics.split_sequenced(u'thống kê tường và cột')) == 1,
          semantics.split_sequenced(u'thống kê tường và cột'))


def test_split_preserves_diacritics():
    parts = semantics.split_sequenced(u'thống kê tường rồi xuất pdf')
    check('split: slices keep their diacritics',
          parts == [u'thống kê tường', u'xuất pdf'], parts)


# ─────────────────────────────────────────────────────────────────────────────
# Utterance
# ─────────────────────────────────────────────────────────────────────────────

def test_is_safe_to_act():
    check('safety: a prohibition is not safe to act on',
          not analyze(u'đừng xóa text note').is_safe_to_act)
    check('safety: a plain command is',
          analyze(u'xóa hết text note').is_safe_to_act)


def test_prompt_hint_states_what_the_model_gets_wrong():
    hint = analyze(u'đừng xóa text note trong view này').to_prompt_hint()
    check('hint: prohibition stated', u'NOT to do' in hint, hint)
    check('hint: scope stated', u'active view' in hint, hint)
    check('hint: category named', u'TextNotes' in hint, hint)

    hint2 = analyze(u'tô đỏ toàn bộ tường').to_prompt_hint()
    check('hint: project scope is called out explicitly',
          u'ENTIRE project' in hint2, hint2)
    check('hint: colour named', u'red' in hint2, hint2)

    hint3 = analyze(u'how many walls are there?').to_prompt_hint()
    check('hint: answers in English when asked in English',
          u'Answer in English.' in hint3, hint3)


def test_prompt_hint_does_not_warn_about_a_question():
    hint = analyze(u'xóa được không?').to_prompt_hint()
    check('hint: no confirmation gate on a question',
          u'Risk:' not in hint, hint)


def test_multi_goal_flag():
    check('utterance: multi-goal detected',
          analyze(u'thống kê tường rồi xuất pdf sheet A').is_multi_goal)
    check('utterance: single goal is not multi',
          not analyze(u'thống kê tường và cột').is_multi_goal)


def test_normalized_text_never_replaces_the_raw_text():
    utt = analyze(u'ktra warning ko')
    check('utterance: raw is preserved verbatim',
          utt.raw == u'ktra warning ko', utt.raw)
    check('utterance: normalised form is separate',
          utt.normalized != utt.raw, utt.normalized)


# ─────────────────────────────────────────────────────────────────────────────
# Ellipsis across turns
# ─────────────────────────────────────────────────────────────────────────────

def test_ellipsis_inherits_the_previous_verb():
    prev = analyze(u'thống kê tường theo tầng')
    cur = analyze(u'còn cửa sổ thì sao?')
    out = resolve_ellipsis(prev, cur)
    check('ellipsis: rebuilt request keeps the verb',
          u'thống kê' in out.normalized, out.normalized)
    check('ellipsis: new object substituted',
          'Windows' in out.categories, out.categories)
    check('ellipsis: transcript still shows what was typed',
          out.raw == u'còn cửa sổ thì sao?', out.raw)


def test_ellipsis_leaves_a_complete_request_alone():
    prev = analyze(u'thống kê tường')
    for text in (u'xóa hết text note',
                 u'hãy xuất pdf tất cả sheet A',
                 u'có bao nhiêu cột ở tầng 3 và tầng 4 trong toàn bộ dự án'):
        cur = analyze(text)
        out = resolve_ellipsis(prev, cur)
        check('ellipsis: "{}" untouched'.format(text[:26]), out is cur, out.raw)


def test_ellipsis_needs_a_previous_turn():
    cur = analyze(u'còn cửa sổ thì sao?')
    check('ellipsis: no previous turn is a no-op',
          resolve_ellipsis(None, cur) is cur)


# ─────────────────────────────────────────────────────────────────────────────
# Dispatcher integration
# ─────────────────────────────────────────────────────────────────────────────

def test_dispatcher_reaches_the_right_specialist_through_teencode():
    from Intelligence.agents.dispatcher import AgentDispatcher
    d = AgentDispatcher()
    cases = [
        (u'ktra warning ko', 'qa_check'),
        (u'xuat pdf sheet A', 'export'),
        (u'thong ke tuong', 'revit_data'),
        (u'to mau do tuong', 'revit_action'),
    ]
    for text, expected in cases:
        got = d.classify(text, allow_llm=False)
        check('dispatch: "{}" -> {}'.format(text[:24], expected),
              got.get('specialist') == expected, got)


def test_dispatcher_does_not_regress_on_plain_input():
    """The existing precedence rules must be untouched by the language layer."""
    from Intelligence.agents.dispatcher import AgentDispatcher
    d = AgentDispatcher()
    cases = [
        (u'so sánh 2 model đang mở', 'multi_doc'),
        (u'xuất pdf sheet A', 'export'),
        (u'dựng model từ ảnh', 'modeling'),
        (u'kiểm tra warning', 'qa_check'),
        (u'đổi tên sheet A-101', 'revit_action'),
        (u'theo tiêu chuẩn TCVN thì sao', 'knowledge'),
        (u'có bao nhiêu bức tường', 'revit_data'),
        (u'xử lý comment trong bản vẽ', 'comment'),
    ]
    for text, expected in cases:
        got = d.classify(text, allow_llm=False)
        check('dispatch (regression): "{}" -> {}'.format(text[:24], expected),
              got.get('specialist') == expected, got)


# ─────────────────────────────────────────────────────────────────────────────

def main():
    groups = [
        ('detection', [
            test_diacritic_less_vietnamese_is_not_english,
            test_code_switched_vietnamese_stays_vietnamese,
            test_bim_jargon_is_not_english_evidence,
            test_real_english_is_detected,
            test_accented_word_never_counts_as_english,
            test_bare_tool_name_does_not_flip_the_language,
            test_reply_language_falls_back_rather_than_guessing,
        ]),
        ('normalisation', [
            test_teencode_is_expanded,
            test_repeated_characters_collapse,
            test_telex_leftovers_are_stripped,
            test_typo_repair_is_conservative,
            test_typo_repair_never_touches_accented_words,
            test_protected_spans_survive_normalisation,
            test_singular_domain_words_are_not_their_own_typo,
        ]),
        ('entities', [
            test_window_beats_door,
            test_folding_collisions_do_not_produce_categories,
            test_real_categories_still_resolve,
            test_colors_need_their_diacritics,
            test_levels_and_sheets,
            test_measures_normalise_to_millimetres,
            test_quantity_intent,
        ]),
        ('semantics', [
            test_prohibitions_are_detected,
            test_ordinary_sentences_are_not_prohibitions,
            test_question_particle_is_not_a_negation,
            test_modality,
            test_sheet_prefix_is_not_a_question_particle,
            test_scope,
            test_scope_marker_is_not_also_a_category,
            test_risk_levels,
            test_sequencer_splitting,
            test_split_preserves_diacritics,
        ]),
        ('utterance', [
            test_is_safe_to_act,
            test_prompt_hint_states_what_the_model_gets_wrong,
            test_prompt_hint_does_not_warn_about_a_question,
            test_multi_goal_flag,
            test_normalized_text_never_replaces_the_raw_text,
        ]),
        ('ellipsis', [
            test_ellipsis_inherits_the_previous_verb,
            test_ellipsis_leaves_a_complete_request_alone,
            test_ellipsis_needs_a_previous_turn,
        ]),
        ('dispatcher', [
            test_dispatcher_reaches_the_right_specialist_through_teencode,
            test_dispatcher_does_not_regress_on_plain_input,
        ]),
    ]
    for group, tests in groups:
        print('\n{}'.format(group))
        for t in tests:
            try:
                t()
            except Exception as ex:
                FAILURES.append(t.__name__)
                print('  ERROR {}  {}'.format(t.__name__, ex))

    print('')
    if FAILURES:
        print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('all passed')
    return 0


if __name__ == '__main__':
    sys.exit(main())
