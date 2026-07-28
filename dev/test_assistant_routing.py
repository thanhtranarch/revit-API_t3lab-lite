# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the T3Lab Assistant's tool catalog and routing.

Run:  python3 dev/test_assistant_routing.py
Exit code 0 = all pass. No external test framework — plain asserts, mirroring
dev/test_assistant_llm.py / dev/test_knowledge.py conventions.

Locks down the defects fixed in the 2026-07-28 Assistant logic pass. Every
test below FAILED before that pass. They fall into two families, both of
which are "the assistant confidently states something untrue":

    launcher   a pushbutton whose entry point sits in `if __name__ ==
               '__main__':` was imported, never run, and reported as opened
    catalog    intents advertised to the LLM pointed at pushbuttons that had
               been deleted or renamed, while a tool that DOES exist
               (PropertyLine) was unreachable

The catalog tests are deliberately written against the real extension tree
rather than a fixture: their whole job is to fail the moment the advertised
catalog drifts from what is actually on disk again.
"""
from __future__ import unicode_literals

import io
import os
import shutil
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXT = os.path.join(REPO, 'T3Lab.extension')
LIB = os.path.join(EXT, 'lib')
TAB = os.path.join(EXT, 'T3Lab.tab')
sys.path.insert(0, LIB)

# Sandbox %APPDATA% BEFORE any config/settings import, so settings.json lands
# in a throwaway dir instead of the real one.
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_test_')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ─────────────────────────────────────────────────────────────────────────────
# Fixtures
# ─────────────────────────────────────────────────────────────────────────────

def _write_button(root, name, body):
    """Create <root>/<name>.pushbutton/script.py containing `body`."""
    btn = os.path.join(root, name + '.pushbutton')
    os.makedirs(btn)
    path = os.path.join(btn, 'script.py')
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(body)
    return path


# ─────────────────────────────────────────────────────────────────────────────
# 1. Launcher actually runs the script
# ─────────────────────────────────────────────────────────────────────────────

def test_launcher_runs_main_guard():
    """The bug: 36 of 42 pushbuttons guard their entry point with
    `if __name__ == '__main__':`. The old launcher used imp.load_source with
    module name '_auto_<title>', so the guard never fired — yet it returned
    True. The assistant said "Opening ..." and nothing happened."""
    from Services.tool_discovery import make_generic_launcher, run_tool_script

    tmp = tempfile.mkdtemp(prefix='t3lab_launch_')
    try:
        sentinel = os.path.join(tmp, 'ran.txt')
        # The coding declaration matters: source must be compiled as bytes or
        # Python 2 raises "encoding declaration in Unicode string".
        body = (
            '# -*- coding: utf-8 -*-\n'
            'def main():\n'
            '    with open({!r}, "w") as f:\n'
            '        f.write("ran")\n'
            '\n'
            'if __name__ == "__main__":\n'
            '    main()\n'
        ).format(sentinel)
        path = _write_button(tmp, 'GuardedTool', body)

        ok, err = make_generic_launcher(path, 'Guarded Tool')()
        check('launcher reports success', ok is True, err)
        check('launcher actually ran the __main__ block',
              os.path.exists(sentinel),
              'sentinel not written — script was imported, not run')

        # Direct call surface used by the registry-driven catalog
        os.remove(sentinel)
        ok, err = run_tool_script(path, 'Guarded Tool')
        check('run_tool_script runs __main__', ok and os.path.exists(sentinel), err)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_launcher_reports_failure_honestly():
    """A missing path and a broken script must both report failure WITH a
    reason — the old code returned a bare False and the UI said
    "Could not open the tool. Check the console."."""
    from Services.tool_discovery import run_tool_script

    ok, err = run_tool_script(os.path.join(REPO, 'no', 'such', 'script.py'), 'Ghost')
    check('missing script fails', ok is False)
    check('missing script explains why', 'not found' in err.lower(), err)

    tmp = tempfile.mkdtemp(prefix='t3lab_launch_')
    try:
        path = _write_button(tmp, 'BrokenTool',
                             '# -*- coding: utf-8 -*-\nraise ValueError("boom")\n')
        ok, err = run_tool_script(path, 'Broken Tool')
        check('broken script fails', ok is False)
        check('broken script surfaces the error', 'boom' in err, err)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_launcher_treats_exitscript_as_success():
    """pyRevit's forms.alert(exitscript=True) raises SystemExit — a normal
    exit path, not a failure."""
    from Services.tool_discovery import run_tool_script

    tmp = tempfile.mkdtemp(prefix='t3lab_launch_')
    try:
        path = _write_button(tmp, 'ExitingTool',
                             '# -*- coding: utf-8 -*-\nimport sys\nsys.exit(0)\n')
        ok, err = run_tool_script(path, 'Exiting Tool')
        check('SystemExit counts as success', ok is True, err)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─────────────────────────────────────────────────────────────────────────────
# 2. Tool catalog matches what is actually installed
# ─────────────────────────────────────────────────────────────────────────────

def _fresh_registry():
    """Populate the tool registry in a temp file and return (module, tools)."""
    from Services import tool_discovery as td
    tmp = tempfile.mkdtemp(prefix='t3lab_reg_')
    td.REGISTRY_FILE = os.path.join(tmp, 'tool_registry.json')
    td.discover_new_tools()
    return td, td.get_registered_tools()


def test_every_registered_tool_exists_on_disk():
    """The catalog must never advertise an intent the assistant cannot honour.
    Eight intents (ParaSync, ProjectName, Workset, DimText, UpperAll, Reset
    Overrides, Grids, LoadFamilyCloud) pointed at deleted pushbuttons."""
    _td, tools = _fresh_registry()
    check('registry is not empty', len(tools) > 10, '{} tools'.format(len(tools)))
    missing = [t.get('intent') for t in tools
               if not (t.get('script_path') and os.path.exists(t['script_path']))]
    check('every registered intent has a script on disk', not missing, missing)


def test_propertyline_is_reachable():
    """PropertyLine.pushbutton exists but was listed in _SKIP_BUTTONS as
    "already hard-coded in TOOL_LAUNCHERS" — where it never appeared. The
    assistant could not open it at all."""
    _td, tools = _fresh_registry()
    intents = set(t.get('intent') for t in tools)
    check('PropertyLine is discoverable',
          any('propertyline' in (i or '') for i in intents),
          sorted(i for i in intents if 'prop' in (i or '')))


def test_registry_prunes_vanished_buttons():
    """Registration was append-only: a renamed or deleted pushbutton stayed in
    the registry forever and kept being offered to the model."""
    td, _tools = _fresh_registry()
    reg = td.load_registry()
    reg['tools']['Ghost.pushbutton'] = {
        'button': 'Ghost.pushbutton', 'intent': 'open_ghost',
        'script_path': os.path.join(REPO, 'no', 'such', 'script.py'),
        'title': 'Ghost',
    }
    td.save_registry(reg)
    td.discover_new_tools()
    check('vanished button is pruned from the registry',
          'Ghost.pushbutton' not in td.load_registry().get('tools', {}))


def test_skip_list_only_hides_real_buttons():
    """_SKIP_BUTTONS must not name buttons that do not exist — a stale entry
    there is how PropertyLine stayed invisible."""
    from Services.tool_discovery import _SKIP_BUTTONS, scan_all_pushbuttons
    on_disk = set(t['button'] for t in scan_all_pushbuttons())
    stale = sorted(b for b in _SKIP_BUTTONS if b not in on_disk)
    check('no stale entries in _SKIP_BUTTONS', not stale, stale)


def test_builtin_tools_are_installed():
    """nlu_engine._BUILTIN_TOOLS covers only tools with a dedicated launcher;
    each must correspond to something really shipped."""
    from Intelligence.nlu_engine import _BUILTIN_TOOLS
    intents = set(t[0] for t in _BUILTIN_TOOLS)
    check('_BUILTIN_TOOLS is only the special launchers',
          intents == set(['open_batchout', 'open_loadfamily']), sorted(intents))
    check('BatchOut ships',
          os.path.exists(os.path.join(TAB, 'Views & Sheets.panel',
                                      'BatchOut.pushbutton', 'script.py')))
    check('Family Manager ships',
          os.path.exists(os.path.join(TAB, 'Modeling & Datum.panel',
                                      'ManaFami.pushbutton', 'script.py')))


def test_renamed_tools_resolve_to_real_tools():
    """"workset" used to hit the hardcoded open_workset trigger (score 35) and
    never reached ManaWorkset, which is the tool that replaced it."""
    _fresh_registry()
    from Intelligence.nlu_engine import resolve_tool

    match, cands = resolve_tool('manaworkset')
    check('manaworkset resolves',
          bool(match) and 'workset' in (match.get('intent') or ''),
          '{} / {}'.format(match, cands))
    check('resolved workset tool is not the deleted intent',
          not match or match.get('intent') != 'open_workset',
          match)

    match, _c = resolve_tool('manaviews')
    check('manaviews resolves', bool(match) and 'view' in (match.get('intent') or ''),
          match)


def _server_tool_names():
    """Tool names registered in core/server.py, read from source.

    core.server cannot be imported headlessly (it pulls in the Revit API), so
    the names are lifted from the schema literals.
    """
    import re as _re
    path = os.path.join(LIB, 'core', 'server.py')
    with io.open(path, encoding='utf-8') as f:
        return set(_re.findall(r"['\"]name['\"]:\s*['\"]([a-z0-9_]+)['\"]", f.read()))


def test_no_dead_intent_advertised_anywhere():
    """Drift lock. Every open_* intent named in a static prompt or table must
    be either a special launcher or a live registry intent."""
    _td, tools = _fresh_registry()
    live = set(t.get('intent') for t in tools)
    live.update(['open_batchout', 'open_batchout_configured', 'open_loadfamily'])
    # The MCP Revit tools registered in core/server.py share the open_* shape
    # (open_document). They are a separate catalog, validated at runtime by
    # t3lab_agent.is_mcp_intent() against the live server registry.
    live.update(_server_tool_names())

    import re as _re
    sources = [
        os.path.join(LIB, 'Intelligence', 'nlu_engine.py'),
        os.path.join(LIB, 'Intelligence', 'local_llm.py'),
        os.path.join(LIB, 'Intelligence', 't3lab_agent.py'),
        os.path.join(LIB, 'Intelligence', 't3lab_assistant.py'),
    ]
    dead = {}
    for path in sources:
        with io.open(path, encoding='utf-8') as f:
            for lineno, line in enumerate(f, 1):
                # Skip the comments that deliberately name the removed intents.
                if line.lstrip().startswith('#'):
                    continue
                # Quoted only: bare open_verb / open_documents are local
                # identifiers, not intents.
                for intent in _re.findall(r'["\']([a-z]*open_[a-z0-9_]+)["\']', line):
                    if intent.startswith('open_') and intent not in live:
                        dead.setdefault(os.path.basename(path), []).append(
                            '{}:{}'.format(intent, lineno))
    check('no dead open_* intent is advertised', not dead, dead)


# ─────────────────────────────────────────────────────────────────────────────
# 3. Language is consistent
# ─────────────────────────────────────────────────────────────────────────────

def test_agent_prompt_language_follows_setting():
    """The agent prompt used to hardcode "Always reply in English", which
    contradicted the UI once the assistant's own strings followed the user."""
    from Intelligence.agent_loop import build_agent_system_prompt

    en = build_agent_system_prompt(u"ctx", lang='en')
    vi = build_agent_system_prompt(u"ctx", lang='vi')
    auto = build_agent_system_prompt(u"ctx", lang='auto')

    check('en pins English', 'reply in English' in en)
    check('vi pins Vietnamese', 'reply in Vietnamese' in vi)
    check('vi does not also demand English', 'reply in English' not in vi)
    check('auto mirrors the user', 'SAME language' in auto)
    check('auto pins neither language',
          'Always reply in English' not in auto
          and 'Always reply in Vietnamese' not in auto)


def test_specialist_prompt_forwards_language():
    from Intelligence.agents.specialists import build_specialist_prompt
    vi = build_specialist_prompt(None, u"ctx", lang='vi')
    check('specialist prompt honours lang', 'reply in Vietnamese' in vi)


def test_knowledge_prompt_is_single_language():
    """The knowledge prompt was written in Vietnamese with an
    "always answer in English" line bolted on."""
    from Intelligence.knowledge.knowledge_agent import get_system_prompt
    vi, en = get_system_prompt(True), get_system_prompt(False)
    check('vi knowledge prompt does not demand English',
          'in English' not in vi)
    check('en knowledge prompt is English', 'Always answer in English' in en)
    check('en knowledge prompt has no Vietnamese rules',
          'NGUYÊN TẮC' not in en)


def test_assistant_cards_are_bilingual():
    """Every visible card string had a single Vietnamese form, which is why
    the window still mixed languages after the English lock."""
    from GUI import AssistantCards as AC
    for key, pair in AC._TEXT.items():
        check('card string "{}" has both languages'.format(key),
              len(pair) == 2 and pair[0] and pair[1] and pair[0] != pair[1],
              pair)
    check('action labels cover the same actions',
          set(AC._ACTION_LABELS_VI) == set(AC._ACTION_LABELS_EN))


def test_reply_language_setting_roundtrip():
    import config.settings as CS
    CS.T3LabAISettings._instance = None
    s = CS.T3LabAISettings()
    check('default is auto', s.get_reply_language() == 'auto')
    s.set_reply_language('vi')
    check('vi persists', s.get_reply_language() == 'vi')
    s.set_reply_language('klingon')
    check('unknown value falls back to auto', s.get_reply_language() == 'auto')


# ─────────────────────────────────────────────────────────────────────────────
# 4. Routing ladder
# ─────────────────────────────────────────────────────────────────────────────

_PREV = {'specialist': 'revit_action', 'skill': 'model-cleanup'}


def test_continuation_needs_a_trailing_question():
    """The old test accepted a '?' anywhere in the last 250 characters, so a
    question mark in a mid-message aside opened the continuation path."""
    from Intelligence import routing as R

    check('question at the end counts',
          R.ends_with_question(u"Bạn muốn áp dụng cho toàn bộ project?"))
    check('trailing markdown does not hide it',
          R.ends_with_question(u"Which scope do you mean?**"))
    check('mid-message question does not count',
          not R.ends_with_question(
              u"Bạn hỏi 'cái gì?' thì đây là kết quả.\nĐã tô đỏ 128 tường."))
    check('statement is not a question',
          not R.ends_with_question(u"Đã tô đỏ 128 tường."))
    check('empty is not a question', not R.ends_with_question(u""))


def test_continuation_accepts_real_answers():
    from Intelligence import routing as R
    q = u"Bạn muốn áp dụng cho toàn bộ project hay chỉ view hiện tại?"

    for reply in (u"1", u"2", u"b", u"ok", u"vâng", u"toàn bộ project",
                  u"chỉ view hiện tại", u"phương án 2", u"cả hai"):
        check(u'continues on "{}"'.format(reply),
              R.is_continuation(q, reply, _PREV), reply)


def test_continuation_rejects_new_commands():
    """"tô vàng sàn" after a clarifying question is a NEW command. It is 13
    characters, so the old length test passed it and only the keyword escape
    hatch stopped it — one check standing alone."""
    from Intelligence import routing as R
    q = u"Bạn muốn áp dụng cho toàn bộ project hay chỉ view hiện tại?"

    check('new colour command is not a continuation',
          not R.is_continuation(q, u"tô vàng sàn", _PREV))
    check('new export command is not a continuation',
          not R.is_continuation(q, u"xuất pdf toàn bộ G sheet", _PREV))
    check('delete command is not a continuation',
          not R.is_continuation(q, u"xóa hết tường tầng 2", _PREV))
    check('shape check stands without the keyword hatch',
          not R.looks_like_answer(u"tô vàng sàn"),
          'action verb must disqualify an answer')


def test_continuation_rejects_closing_questions():
    from Intelligence import routing as R
    for closer in (u"Bạn có muốn thực hiện một hành động khác không?",
                   u"Anything else I can help with?",
                   u"Cần gì thêm không?"):
        check(u'closing question does not carry over',
              not R.is_continuation(closer, u"1", _PREV), closer)


def test_continuation_guards():
    from Intelligence import routing as R
    q = u"Which scope?"
    check('no previous decision → no carryover',
          not R.is_continuation(q, u"1", None))
    check('fresh keyword hit → no carryover',
          not R.is_continuation(q, u"1", _PREV, fresh_keyword_hit=True))
    check('empty reply → no carryover', not R.is_continuation(q, u"  ", _PREV))
    check('long reply → no carryover',
          not R.is_continuation(q, u"x" * 120, _PREV))


def test_learned_pattern_defers_to_nlu():
    """A learned pattern used to take the turn outright before the NLU ran, so
    a stale mapping outranked a confident read of the live tool catalog."""
    from Intelligence import routing as R
    learned = {'intent': 'open_batchout', 'params': {}}

    check('wins when the NLU has no opinion',
          R.learned_pattern_wins(learned, None))
    check('wins when the NLU says unknown',
          R.learned_pattern_wins(learned, {'intent': 'unknown'}))
    check('wins when both agree',
          R.learned_pattern_wins(learned, {'intent': 'open_batchout'}))
    check('loses when the NLU names another tool',
          not R.learned_pattern_wins(learned, {'intent': 'check_spelling'}))
    check('loses to an authoritative catalog answer',
          not R.learned_pattern_wins(
              learned, {'intent': 'help', '_authoritative': True}))
    check('nothing learned → never wins',
          not R.learned_pattern_wins(None, None))


def test_dispatcher_precedence_conflicts():
    """The dispatcher's keyword stage is ordering-sensitive and its comments
    document specific conflicts it had to be tuned for. None of them had a
    test, so a table edit could silently undo the tuning."""
    from Intelligence.agents.dispatcher import AgentDispatcher
    d = AgentDispatcher()

    def spec(text):
        return d.classify(text, allow_llm=False).get('specialist')

    cases = [
        # "đổi model" contains the action verb 'doi' but is a document switch
        (u"đổi model sang file kia", 'multi_doc'),
        (u"so sánh 2 model", 'multi_doc'),
        # 'xuat'/'export' are action words too; the export spec must win
        (u"xuất pdf toàn bộ sheet", 'export'),
        # 'tao tuong' would otherwise land in generic action
        (u"tạo tường tầng 2", 'modeling'),
        # audits get the QA role, not the data role
        (u"kiểm tra warning trong model", 'qa_check'),
        # a write verb wins over doc/count words
        (u"đổi chiều cao theo tiêu chuẩn", 'revit_action'),
        # colour phrases only match WITH diacritics
        (u"tô đỏ tường", 'revit_action'),
        # read-only questions
        (u"có bao nhiêu cửa", 'revit_data'),
        # document questions
        (u"theo tiêu chuẩn TCVN thì sao", 'knowledge'),
        # PDF markup workflow
        (u"xử lý comment bản vẽ", 'comment'),
    ]
    for text, want in cases:
        got = spec(text)
        check(u'"{}" → {}'.format(text, want), got == want,
              'got {}'.format(got))

    # "bản vẽ" folds to "ban ve" — its "ve" must not count as the verb "vẽ".
    check('"bản vẽ" alone is not an action',
          spec(u"bản vẽ này là gì") != 'revit_action',
          spec(u"bản vẽ này là gì"))
    # Diacritic-less "to do" must stay ambiguous rather than false-matching
    # English ("what to do...").
    check('"what to do" is not a colour command',
          spec(u"what to do next") != 'revit_action',
          spec(u"what to do next"))


def test_spellcheck_fix_detection():
    from Intelligence import routing as R
    for args in (u"fix them", u"apply the corrections", u"sửa hết",
                 u"sửa lỗi chính tả", u"cập nhật lại"):
        check(u'"{}" asks for a fix'.format(args),
              R.wants_spellcheck_fix(args), args)
    for args in (u"", u"scan the project", u"kiểm tra toàn bộ dự án"):
        check(u'"{}" is a scan'.format(args),
              not R.wants_spellcheck_fix(args), args)


# ─────────────────────────────────────────────────────────────────────────────
# Runner
# ─────────────────────────────────────────────────────────────────────────────

TESTS = [
    ('launcher', [
        test_launcher_runs_main_guard,
        test_launcher_reports_failure_honestly,
        test_launcher_treats_exitscript_as_success,
    ]),
    ('catalog', [
        test_every_registered_tool_exists_on_disk,
        test_propertyline_is_reachable,
        test_registry_prunes_vanished_buttons,
        test_skip_list_only_hides_real_buttons,
        test_builtin_tools_are_installed,
        test_renamed_tools_resolve_to_real_tools,
        test_no_dead_intent_advertised_anywhere,
    ]),
    ('language', [
        test_agent_prompt_language_follows_setting,
        test_specialist_prompt_forwards_language,
        test_knowledge_prompt_is_single_language,
        test_assistant_cards_are_bilingual,
        test_reply_language_setting_roundtrip,
    ]),
    ('routing', [
        test_continuation_needs_a_trailing_question,
        test_continuation_accepts_real_answers,
        test_continuation_rejects_new_commands,
        test_continuation_rejects_closing_questions,
        test_continuation_guards,
        test_learned_pattern_defers_to_nlu,
        test_dispatcher_precedence_conflicts,
        test_spellcheck_fix_detection,
    ]),
]


def main():
    for group, tests in TESTS:
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
