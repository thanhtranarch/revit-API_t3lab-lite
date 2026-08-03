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


def test_api_context_runner_normalises_every_launcher_shape():
    """Launchers return (ok, err), a bare bool, or None. run_in_api_context is
    the single place that folds all three into (ok, error_text) — and it must
    never raise, because the caller is a WPF click handler with no handler
    above it."""
    from Services.revit_context import run_in_api_context, ensure_api_context

    seen = []
    def _done(ok, err):
        seen.append((ok, err))

    run_in_api_context(lambda: (True, u''), _done)
    run_in_api_context(lambda: (False, u'no window'), _done)
    run_in_api_context(lambda: True, _done)
    run_in_api_context(lambda: None, _done)

    def _boom():
        raise ValueError('boom')
    run_in_api_context(_boom, _done)

    check('tuple success', seen[0] == (True, u''), seen[0])
    check('tuple failure keeps the reason', seen[1] == (False, u'no window'), seen[1])
    check('bare True is success', seen[2][0] is True, seen[2])
    check('None is success', seen[3][0] is True, seen[3])
    check('an exception is a reported failure, not a crash',
          seen[4][0] is False and 'boom' in seen[4][1], seen[4])

    ok, err = ensure_api_context()
    check('ensure_api_context degrades quietly outside Revit',
          ok is False and bool(err), err)


def test_raise_failure_only_discards_its_own_task():
    """With more than one assistant window, the task queue can hold another
    caller's still-good task when Raise() fails for THIS call. The recovery
    path must remove only the entry it just queued, not whatever is at the
    front of the queue — otherwise it silently drops someone else's task and
    its on_done callback never fires."""
    from Services import revit_context as RC

    seen = []
    RC._TASKS.put(('other-token', lambda: (True, u'other'),
                   lambda ok, err: seen.append(('other', ok, err))))

    real_event, RC._EVENT = RC._EVENT, _FakeRaisingEvent()
    try:
        def _done(ok, err):
            seen.append(('mine', ok, err))
        RC.run_in_api_context(lambda: (True, u'mine'), _done)
    finally:
        RC._EVENT = real_event

    check('the other caller\'s task is still queued, untouched',
          RC._TASKS.qsize() == 1, RC._TASKS.qsize())
    token, func, on_done = RC._TASKS.get_nowait()
    check('it is still the other task, not a stray copy',
          token == 'other-token', token)
    check('only the failing call\'s own outcome was reported',
          seen == [('mine', True, u'mine')], seen)


class _FakeRaisingEvent(object):
    def Raise(self):
        raise RuntimeError('event disposed')


def test_tools_are_launched_through_the_api_context():
    """Drift lock for the crash that made every tool unopenable from chat.

    A pushbutton script is a Revit command: it may call ExternalEvent.Create
    or open a Transaction. The assistant's callbacks run while Revit is IDLE,
    which is not a "standard API execution", so calling a launcher directly
    threw

        Attempting to create an ExternalEvent outside of a standard API execution

    right after the assistant had announced it was opening the tool (BCF
    Reader, ManaLoca, BatchOut). Every launcher call must go through
    _launch_tool → run_in_api_context."""
    import re as _re
    path = os.path.join(TAB, 'Support.panel', 'T3LabAssistant.pushbutton',
                        'script.py')
    with io.open(path, encoding='utf-8') as f:
        src = f.read()

    direct = []
    for lineno, line in enumerate(src.splitlines(), 1):
        if line.lstrip().startswith('#'):
            continue
        if _re.search(r'TOOL_LAUNCHERS\s*(\[[^\]]+\]|\.get\([^)]*\))\s*\(\s*\)', line):
            direct.append(lineno)
    check('no launcher is invoked outside the API context hop', not direct, direct)
    check('_launch_tool marshals through run_in_api_context',
          _re.search(r'def _launch_tool[\s\S]{0,1400}?run_in_api_context', src)
          is not None)


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
    Overrides, Grids, LoadFamilyCloud) pointed at deleted pushbuttons.

    A urlbutton has no script.py at all — its hyperlink is the target."""
    _td, tools = _fresh_registry()
    check('registry is not empty', len(tools) > 10, '{} tools'.format(len(tools)))
    missing = [t.get('intent') for t in tools
               if not (t.get('url')
                       or (t.get('script_path') and os.path.exists(t['script_path'])))]
    check('every registered intent has a launchable target', not missing, missing)


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


def test_every_ribbon_button_is_openable():
    """Full coverage lock: every launchable button on the ribbon must be in the
    registry, except the ones with a dedicated launcher (BatchOut, ManaFami)
    and the assistant itself.

    Before this, `.urlbutton` folders were not scanned at all — "open Autodesk
    Forma" answered that no such tool existed while the button sat on the
    ribbon two panels away."""
    td, tools = _fresh_registry()
    on_disk = set(t['button'] for t in td.scan_all_buttons())
    registered = set(t['button'] for t in tools)
    gap = sorted(on_disk - registered - set(td._SKIP_BUTTONS))
    check('every ribbon button is registered', not gap, gap)
    check('url buttons are discoverable',
          any(t.get('kind') == 'url' for t in tools),
          sorted(t['button'] for t in tools if t.get('kind') == 'url'))


def test_skipped_buttons_are_pruned_from_an_old_registry():
    """_SKIP_BUTTONS was applied only when REGISTERING. ManaFami was registered
    before it was added to the list, so the stale entry survived every rescan
    and collided with the open_loadfamily special launcher — "Family Manager",
    "Mana Fami" and "ManaFami" all resolved to nothing (two exact matches =
    ambiguous)."""
    td, _tools = _fresh_registry()
    reg = td.load_registry()
    skipped = sorted(td._SKIP_BUTTONS)[0]
    reg['tools'][skipped] = {
        'button': skipped, 'intent': 'open_stale', 'title': 'Stale',
        'script_path': os.path.join(TAB, 'Views & Sheets.panel',
                                    'BatchOut.pushbutton', 'script.py'),
    }
    td.save_registry(reg)
    td.discover_new_tools()
    check('a skipped button is pruned even when already registered',
          skipped not in td.load_registry().get('tools', {}))


def _bundle_title(button_dir):
    """The ribbon label from bundle.yaml, or None."""
    import re as _re
    path = os.path.join(button_dir, 'bundle.yaml')
    if not os.path.exists(path):
        return None
    with io.open(path, encoding='utf-8') as f:
        for line in f:
            m = _re.match(r'\s*title\s*:\s*(.+)', line)
            if m:
                return (m.group(1).strip().strip('"\'')
                        .replace('\\n', ' ').strip())
    return None


def test_every_tool_resolves_by_every_name_it_shows():
    """A tool must be openable by each name the user can see: its script
    __title__, its folder name, and the RIBBON LABEL.

    The ribbon label is not always the script title — the Wall_Adjust Base
    button reads "Auto Adj Base Offset" on the ribbon and "Auto Adjust Base
    Offset" in its script — and typing what the ribbon showed resolved to
    nothing at all."""
    td, tools = _fresh_registry()
    from Intelligence.nlu_engine import resolve_tool

    by_button = dict((t['button'], t) for t in tools)
    misses = []
    for scanned in td.scan_all_buttons():
        entry = by_button.get(scanned['button'])
        if not entry:
            continue                      # skip-listed: covered by its own test
        names = [entry['title'], td._strip_suffix(scanned['button'])]
        label = _bundle_title(scanned['path'])
        if label:
            names.append(label)
        for name in names:
            for phrase in (name, u'open ' + name, u'mở ' + name):
                match, cands = resolve_tool(phrase)
                if not match or match.get('intent') != entry['intent']:
                    misses.append(u'{!r} → {} (want {})'.format(
                        phrase, match and match.get('intent'), entry['intent']))
    check('every tool resolves by every name it shows', not misses, misses[:6])


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

    en = build_agent_system_prompt(lang='en')
    vi = build_agent_system_prompt(lang='vi')
    auto = build_agent_system_prompt(lang='auto')

    check('en pins English', 'reply in English' in en)
    check('vi pins Vietnamese', 'reply in Vietnamese' in vi)
    check('vi does not also demand English', 'reply in English' not in vi)
    check('auto mirrors the user', 'SAME language' in auto)
    check('auto pins neither language',
          'Always reply in English' not in auto
          and 'Always reply in Vietnamese' not in auto)


def test_specialist_prompt_forwards_language():
    from Intelligence.agents.specialists import build_specialist_prompt
    vi = build_specialist_prompt(None, lang='vi')
    check('specialist prompt honours lang', 'reply in Vietnamese' in vi)


def test_system_prompt_is_static_for_prompt_cache():
    """The whole point of P1: the system block must be byte-identical across
    turns so Anthropic's cache breakpoint on it actually hits.

    Live Revit state (active view, selection — re-read by a 2s timer) used to
    be interpolated into the prompt, so the system block never repeated and
    the entire prefix was re-processed every single turn.
    """
    from Intelligence.agent_loop import (build_agent_system_prompt,
                                         build_context_block,
                                         apply_context_block,
                                         strip_context_block)

    a = build_agent_system_prompt(lang='en')
    b = build_agent_system_prompt(u"View: Level 1 | Selection: 3", lang='en')
    check('system prompt ignores live context', a == b)
    check('no stale context placeholder left behind', '{context}' not in a)

    empty = build_context_block()
    check('no live state = no extra tokens', empty == u"")

    blk = build_context_block(revit_context=u"View: Level 1",
                              knowledge_ref=u"## Reference\n[1] spec")
    check('context block carries the view', 'View: Level 1' in blk)
    check('context block carries the knowledge ref', '[1] spec' in blk)

    merged = apply_context_block(u"tô đỏ tường", blk)
    check('user text survives', merged.endswith(u"tô đỏ tường"))
    check('context comes first', merged.startswith('<<<T3LAB_LIVE_CONTEXT'))
    check('round-trips back to the raw turn',
          strip_context_block(merged) == u"tô đỏ tường")

    # Vision turns send a block list, not a string — the image blocks must
    # keep their own positions with the context prepended as text.
    blocks = apply_context_block(
        [{"type": "text", "text": "look"}, {"type": "image", "source": {}}], blk)
    check('block list stays a list', isinstance(blocks, list))
    check('context prepended as its own text block',
          blocks[0]['type'] == 'text' and 'View: Level 1' in blocks[0]['text'])
    check('image block preserved', blocks[-1]['type'] == 'image')
    check('no context = content untouched',
          apply_context_block(u"hi", u"") == u"hi")


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


def _cap_answer(q, viet):
    from Intelligence import nlu_engine as N
    return N.answer_capability_question(q, viet)['message']


def _is_overview(msg):
    return (u'directly on the Revit model' in msg
            or u'trực tiếp trên model Revit' in msg)


def test_capability_scope_is_not_a_predicate():
    """A capability question is FRAME(PREDICATE [SCOPE]). The scope names WHERE
    the work happens ("with this project", "trong model này") and must never
    select a tool on its own — "what can you do with this project" used to be
    answered with Family Loader, whose doc merely reads "…vào project"."""
    for q in (u"what can you do with this project",
              u"what can you do in this model",
              u"what can you do in revit",
              u"what can you do here",
              u"what can you do",
              u"what tools do you have"):
        check(u'"{}" → overview'.format(q), _is_overview(_cap_answer(q, False)))
    for q in (u"bạn làm được gì với dự án này",
              u"t3lab hỗ trợ gì trong model này",
              u"hiện tại bạn hỗ trợ gì",
              u"có tool nào không",
              u"bạn làm được gì"):
        check(u'"{}" → overview'.format(q), _is_overview(_cap_answer(q, True)))


def test_capability_predicate_still_finds_tools():
    """Stripping the scope must not blunt real asks — a question that carries a
    predicate still resolves against the catalog, in either language."""
    for q, viet, expect in (
            (u"is there a tool to load family",     False, u'Family Loader'),
            (u"is there a tool for dwg",            False, u'DWG'),
            (u"what can you do with dwg files",     False, u'DWG'),
            (u"do you have a tool for point cloud", False, u'Point Cloud'),
            (u"any tool for tile layout",           False, u'Tile Layout'),
            (u"có tool nào để xuất pdf không",      True,  u'BatchOut'),
            (u"có tool nào check model không",      True,  u'Model Auditor'),
            (u"có tool nào quản lý workset không",  True,  u'Workset'),
            # "in" is a homograph: Vietnamese print, English preposition
            (u"có tool nào để in sheet không",      True,  u'BatchOut')):
        msg = _cap_answer(q, viet)
        check(u'"{}" → {}'.format(q, expect),
              expect in msg and not _is_overview(msg), msg[:80])


def test_vietnamese_capability_verbs_reach_the_english_catalog():
    """The catalog is named in English, so a Vietnamese ask for the same
    capability used to find nothing ("đổi tên view" → "chưa có tool")."""
    for q, expect in ((u"có tool nào để đổi tên view không", u'View Manager'),
                      (u"có tool nào quản lý dwg không",     u'DWG Manager'),
                      (u"có tool nào quản lý workset không", u'Workset Manager'),
                      (u"có tool nào quản lý thư viện family không",
                       u'Family Loader')):
        msg = _cap_answer(q, True)
        check(u'"{}" → {}'.format(q, expect), expect in msg, msg[:80])


def test_ubiquitous_word_cannot_name_a_tool():
    """"manager" is in 16 of 44 tool names — it selects nothing. A request whose
    only evidence is such a word must not answer with three arbitrary managers
    (and must not claim no tool exists either): the honest reply is the
    overview. Rarity may damp evidence, never manufacture it."""
    msg = _cap_answer(u"có tool nào quản lý project không", True)
    check(u'"quản lý project" → overview, not 3 arbitrary managers',
          _is_overview(msg), msg[:80])
    # …while the same word plus a selective one still resolves precisely
    msg = _cap_answer(u"có tool nào quản lý workset không", True)
    check(u'"quản lý workset" still resolves precisely',
          u'Workset Manager' in msg and not _is_overview(msg), msg[:80])


def test_capability_ignores_implementation_jargon():
    """Docstrings are written for developers ("v2 - pyRevit WPFWindow
    modeless"). Those words describe the implementation, never a capability, so
    they must not name a tool to the user."""
    msg = _cap_answer(u"is there a tool for modeless windows", False)
    check(u'"modeless windows" names no tool',
          msg.startswith(u'❌'), msg[:80])


def test_abbrev_expansion_respects_word_boundaries():
    """A multi-word abbreviation must not eat into the following word.
    "what are you" → "capabilities query" used to fire inside "what are YOUR
    capabilities", producing the nonsense "capabilities queryr capabilities".
    Single-token stems keep substring semantics on purpose ("images" → "imgs").
    """
    from Intelligence import nlu_engine as N
    exp = lambda q: N._expand(N._norm(q))
    check(u'"what are your capabilities" survives expansion',
          exp(u"what are your capabilities") == u"what are your capabilities",
          exp(u"what are your capabilities"))
    check(u'"what can you do" still collapses',
          exp(u"what can you do") == u"capabilities query")
    check(u'single-token stem still expands plurals',
          exp(u"export images") == u"export imgs", exp(u"export images"))
    check(u'space-padded shorthand still expands',
          exp(u"bo sheet") == u"batchout sheet", exp(u"bo sheet"))


def test_capability_frames_cover_productive_forms():
    """The frame detector must recognise the ability question in its productive
    forms, and stay out of superficially similar questions that are not."""
    from Intelligence import nlu_engine as N
    ask = lambda q: N.is_capability_question(N._expand(N._norm(q)))
    for q in (u"what else can you do", u"what are your capabilities",
              u"list your features", u"show me all your tools",
              u"tôi có thể làm gì với t3lab", u"bạn xử lý được gì"):
        check(u'"{}" is a capability question'.format(q), ask(q))
    for q in (u"what do you think about this wall",
              u"what can you tell me about walls",
              u"tôi phải làm gì bây giờ", u"list all sheets",
              u"what can I do to fix it"):
        check(u'"{}" is not a capability question'.format(q), not ask(q))


def test_pronoun_vs_determiner():
    """"nó/it/this" is anaphora only when it stands IN PLACE OF the noun. As a
    determiner ("this project", "that view") it must not bind to the last tool
    in the history, or an unrelated request silently opens that tool."""
    from Intelligence import nlu_engine as N
    exp = lambda q: N._expand(N._norm(q))

    for q in (u"mở nó", u"nó là gì", u"cái này là gì", u"open it",
              u"what is that"):
        check(u'"{}" is anaphora'.format(q), N._is_pronoun_query(exp(q)))
    for q in (u"what can you do with this project", u"that view is wrong",
              u"this project needs cleanup",
              u"nó không quan trọng bằng việc kiểm tra toàn bộ sheet"):
        check(u'"{}" is not anaphora'.format(q),
              not N._is_pronoun_query(exp(q)))


def test_history_referent_is_a_real_tool_mention():
    """The referent for "nó" is the tool the conversation acted on — the
    assistant's own "Opening X..." line, or a message that IS a tool request.
    A name merely appearing inside prose must not bind it."""
    from Intelligence import nlu_engine as N

    def intent(hist, q):
        r = N.classify(q, history=[{'role': 'assistant', 'content': h}
                                   for h in hist])
        return (r or {}).get('intent')

    check(u'"Opening BatchOut..." + "nó là gì" → BatchOut',
          intent([u"Đang mở BatchOut..."], u"nó là gì") == 'open_batchout')
    check(u'spaced title resolves ("Đang mở DWG Manager...")',
          intent([u"Đang mở DWG Manager..."], u"cái này là gì") == 'open_manadwg')
    check(u'a name inside prose does not bind the pronoun',
          intent([u"thanks for the feedback"], u"mở nó") is None)
    check(u'small talk leaves the pronoun unresolved',
          intent([u"Xin chào!"], u"mở nó") is None)

    from Intelligence import nlu_engine as N2
    r = N2.classify(u"mở nó", history=[
        {'role': 'user', 'content': u"đang mở dwg management giúp tôi với"},
    ])
    check(u'an "Opening X..." phrase from the USER does not count as the '
          u'assistant having launched that tool',
          (r or {}).get('intent') is None, r)


def test_capability_question_beats_anaphora():
    """An explicit capability frame states its own subject — resolving its
    pronoun against the history would open a tool instead of answering."""
    from Intelligence import nlu_engine as N
    r = N.classify(u"what can you do with this project",
                   history=[{'role': 'assistant', 'content': u'Đang mở BatchOut...'}])
    check(u'capability frame is answered, not routed to the last tool',
          r and r.get('intent') == 'help' and _is_overview(r.get('message', '')),
          (r or {}).get('intent'))


# ─────────────────────────────────────────────────────────────────────────────
# Runner
# ─────────────────────────────────────────────────────────────────────────────

TESTS = [
    ('launcher', [
        test_launcher_runs_main_guard,
        test_launcher_reports_failure_honestly,
        test_launcher_treats_exitscript_as_success,
        test_api_context_runner_normalises_every_launcher_shape,
        test_raise_failure_only_discards_its_own_task,
        test_tools_are_launched_through_the_api_context,
    ]),
    ('catalog', [
        test_every_registered_tool_exists_on_disk,
        test_every_ribbon_button_is_openable,
        test_every_tool_resolves_by_every_name_it_shows,
        test_propertyline_is_reachable,
        test_registry_prunes_vanished_buttons,
        test_skipped_buttons_are_pruned_from_an_old_registry,
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
    ('prompt cache', [
        test_system_prompt_is_static_for_prompt_cache,
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
    ('semantics', [
        test_abbrev_expansion_respects_word_boundaries,
        test_capability_frames_cover_productive_forms,
        test_capability_scope_is_not_a_predicate,
        test_capability_predicate_still_finds_tools,
        test_vietnamese_capability_verbs_reach_the_english_catalog,
        test_ubiquitous_word_cannot_name_a_tool,
        test_capability_ignores_implementation_jargon,
        test_pronoun_vs_determiner,
        test_history_referent_is_a_real_tool_mention,
        test_capability_question_beats_anaphora,
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
