# -*- coding: utf-8 -*-
"""
CPython 3 tests for the "act, don't teach" guard in Intelligence/agent_loop.py.

Run:  python3 dev/test_assistant_act_not_teach.py
Exit 0 = all pass. Plain asserts, mirroring dev/test_exemplars.py.

Observed 2026-08-10: asked "tô đỏ tường", the local model replied with a
confident numbered walkthrough of a "Bulk Set Parameter" dialog that does not
exist in Revit, and called no tool. To the user that is a refusal with extra
steps — they asked the assistant to DO the work.

_announces_work could not catch it: that detector returns False above 220
characters because it targets the short "Đang xử lý..." promise. A tutorial is
long by nature, so it needed its own detector.
"""
from __future__ import unicode_literals

import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence import agent_loop as A

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# The reply that started this, trimmed to its shape.
REAL_TUTORIAL = """Tôi hiểu bạn muốn tô màu đỏ cho tất cả các bức tường trong dự án. Để thực hiện điều này, bạn có thể sử dụng công cụ "Bulk Set Parameter" trong Revit.

Dưới đây là hướng dẫn từng bước:

1. Chọn tất cả các bức tường trong dự án bằng cách sử dụng công cụ "Select All".
2. Nhấp chuột phải vào một bức tường đã chọn và chọn "Bulk Set Parameter".
3. Trong hộp thoại "Bulk Set Parameter", chọn tham số "Color" và nhập giá trị màu đỏ.
4. Nhấp chuột vào nút "Apply" để áp dụng thay đổi."""

ENGLISH_TUTORIAL = """Here are the steps to colour the walls red:

1. Select all the walls in the project.
2. Click on the Modify tab in the ribbon.
3. In the dialog box, choose the override colour.
4. Apply the change."""


def test_catches_the_real_reply():
    print('[act-not-teach: the observed tutorial is caught]')
    check('Vietnamese walkthrough caught',
          A._instructs_instead_of_acting(REAL_TUTORIAL, u'tô đỏ tường'))
    check('English walkthrough caught',
          A._instructs_instead_of_acting(ENGLISH_TUTORIAL, u'colour the walls red'))
    check('the old guard could NOT catch it (why this exists)',
          not A._announces_work(REAL_TUTORIAL))


def test_leaves_legitimate_replies_alone():
    print('[act-not-teach: real answers are not touched]')
    cases = [
        (u'Model có 18 cao độ, từ Parking đến Parapet 2.', u'liệt kê cao độ'),
        (u'Đã tô đỏ 185 tường trong view L1.', u'tô đỏ tường'),
        (u'The model has 1128 walls.', u'how many walls?'),
        (u'Không có tool nào xuất được định dạng đó.', u'xuất sang IFC5'),
        (u'', u'tô đỏ tường'),
        (None, u'tô đỏ tường'),
    ]
    for text, ask in cases:
        check(u'left alone: {}'.format((text or u'<empty>')[:38]),
              not A._instructs_instead_of_acting(text, ask))

    # One passing mention of a button is not a tutorial.
    passing = (u'Đã đổi xong. Nếu chưa thấy, nhấp chuột vào view để refresh.')
    check('a single incidental marker is not enough',
          not A._instructs_instead_of_acting(passing, u'tô đỏ tường'), passing)


def test_respects_an_explicit_how_to_question():
    """When the user ASKS how to do it themselves, a walkthrough is correct."""
    print('[act-not-teach: an explicit how-to question is honoured]')
    for ask in (u'làm thế nào để tô đỏ tường?', u'how do I colour walls red',
                u'chỉ tôi cách tạo sheet', u'what are the steps to add a level'):
        check(u'not nudged for: {}'.format(ask[:34]),
              not A._instructs_instead_of_acting(REAL_TUTORIAL, ask))
    check('_asks_for_instructions detects the ask',
          A._asks_for_instructions(u'làm sao để tô tường')
          and A._asks_for_instructions(u'How to export PDF'))
    check('a plain command is not a how-to question',
          not A._asks_for_instructions(u'tô đỏ tường'))


def test_fixup_and_prompt_rule():
    print('[act-not-teach: correction text and the prompt rule]')
    check('fixup tells it to call the tool now',
          'Call the tool that performs the request NOW' in A._HOWTO_FIXUP)
    check('fixup names the failure as a refusal',
          'refusal' in A._HOWTO_FIXUP.lower())
    check('fixup allows an honest "no tool can"',
          'no tool can do it' in A._HOWTO_FIXUP)

    prompt = A.build_agent_system_prompt(local=True, lang='vi')
    check('prompt carries the rule',
          'You DO the work' in prompt and 'never teach' in prompt.lower(),
          prompt[:80])
    check('prompt gives the concrete example',
          'revit_override_color' in prompt)
    check('prompt still allows explaining on request',
          'explicitly asked how to do it by hand' in prompt)

    # The guard must be wired into the loop, not just defined.
    src = open(A.__file__.replace('.pyc', '.py'), encoding='utf-8').read()
    check('guard is called in the run loop',
          '_instructs_instead_of_acting(text, user_content)' in src)
    check('guard shares the nudge budget',
          src.count('announce_nudges < _MAX_ANNOUNCE_NUDGES') == 2, src.count(
              'announce_nudges < _MAX_ANNOUNCE_NUDGES'))


def main():
    test_catches_the_real_reply()
    test_leaves_legitimate_replies_alone()
    test_respects_an_explicit_how_to_question()
    test_fixup_and_prompt_rule()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All act-not-teach tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
