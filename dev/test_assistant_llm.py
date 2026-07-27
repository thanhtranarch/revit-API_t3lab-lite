# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the T3Lab Assistant / LLMs Setting stack.

Run:  python3 dev/test_assistant_llm.py
Exit code 0 = all pass. No external test framework — plain asserts,
mirroring dev/test_knowledge.py and dev/test_llm_config.py conventions.

Locks down the defects fixed in the 2026-07-27 Assistant + LLMs Setting debug
pass. Every test below FAILED before that pass; most are data-loss or
"the UI confidently shows something untrue" bugs, which is why they get a
permanent regression test rather than a one-off manual check:

    settings   two Revit sessions overwriting each other's API keys
    settings   one toggle wiping every key when settings.json is truncated
    router     status cache lying about the active model after set_model
    router     a project's provider override rewriting the global default
    ollama     a custom server URL that never survived a restart
    ollama     JSON grammar forced onto prose replies (Test Connection,
               docked-pane chat, and — silently — spell-check)

WPF/Revit-bound behaviour (the chat window, the settings dialog) cannot be
exercised headlessly; those go on the in-Revit checklist instead.
"""
from __future__ import unicode_literals

import io
import json
import os
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
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


def _fresh_settings():
    """Return a brand-new T3LabAISettings, as a separate Revit session would."""
    import config.settings as CS
    CS.T3LabAISettings._instance = None
    return CS.T3LabAISettings()


def _reset_appdata():
    """Point APPDATA at an empty dir and drop the settings singleton."""
    os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_test_')
    return _fresh_settings()


# ─── settings: concurrent sessions ────────────────────────────────────────────

def test_settings_merge_on_write():
    print('[settings: merge-on-write]')
    a = _reset_appdata()
    a.set_api_key('Claude', 'sk-ant-A')

    # Session B starts, saves its own key.
    b = _fresh_settings()
    b.set_api_key('DeepSeek', 'sk-deepseek-B')

    # Session A still holds a pre-B snapshot. Each of these setters used to
    # dump that stale dict over the whole file, destroying B's key.
    a.set_provider_model('claude', 'claude-opus-5')
    a.set_active_provider('openai')
    a.set_username('Thanh')
    a.save_window_state(10, 20, 800, 600)

    c = _fresh_settings()
    check('session B api key survives session A writes',
          c.get_api_key('DeepSeek') == 'sk-deepseek-B', c.get_api_key('DeepSeek'))
    check('session A api key survives', c.get_api_key('Claude') == 'sk-ant-A')
    check('session A model persisted',
          c.get_provider_model('claude') == 'claude-opus-5')
    check('session A provider persisted', c.get_active_provider() == 'openai')
    check('session A username persisted', c.get_username() == 'Thanh')
    check('session A window state persisted', c.get_window_state()['width'] == 800)


# ─── settings: corrupt file ───────────────────────────────────────────────────

def test_settings_corrupt_quarantine():
    print('[settings: corrupt-file quarantine]')
    s = _reset_appdata()
    s.set_api_key('Claude', 'sk-ant-REAL')
    s.set_provider_model('claude', 'claude-opus-5')
    path = s._settings_file
    d = os.path.dirname(path)

    # Truncated mid-write (Revit crash, OneDrive sync conflict).
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write('{"api_keys": {"Claude": "sk-ant-RE')

    # One toggle in LLMs Setting. This used to load defaults, patch them and
    # write them back — silently destroying every key on disk.
    s2 = _fresh_settings()
    s2.set_agent_option('quality_mode', True)

    kept = [n for n in os.listdir(d) if n.startswith('settings.corrupt-')]
    check('corrupt file quarantined, not overwritten', len(kept) == 1, kept)
    if kept:
        raw = io.open(os.path.join(d, kept[0]), encoding='utf-8').read()
        check('quarantined bytes intact (key recoverable by hand)',
              'sk-ant-RE' in raw, raw[:60])
    check('app stays writable after quarantine',
          s2.get_agent_option('quality_mode') is True)
    check('settings reported healthy after quarantine', s2.is_healthy())


def test_settings_first_run_still_writes():
    print('[settings: first run]')
    s = _reset_appdata()
    # No settings.json exists yet — absence must NOT be treated as corruption.
    check('no file on first run', not os.path.exists(s._settings_file))
    check('first write succeeds', s.set_api_key('Claude', 'sk-ant-FIRST') is True)
    check('first write round-trips',
          _fresh_settings().get_api_key('Claude') == 'sk-ant-FIRST')
    check('no spurious quarantine file',
          not [n for n in os.listdir(os.path.dirname(s._settings_file))
               if n.startswith('settings.corrupt-')])


def test_settings_unreadable_refuses_write():
    print('[settings: unreadable file]')
    s = _reset_appdata()
    s.set_api_key('Claude', 'sk-ant-KEEP')
    path = s._settings_file
    good = io.open(path, encoding='utf-8').read()

    if os.geteuid() == 0 if hasattr(os, 'geteuid') else False:
        print('  skip  running as root — chmod cannot make a file unreadable')
        return
    os.chmod(path, 0)
    try:
        s2 = _fresh_settings()
        check('unreadable file marks settings unhealthy', not s2.is_healthy())
        check('save refuses while unhealthy',
              s2.set_agent_option('quality_mode', True) is False)
    finally:
        os.chmod(path, 0o600)
    check('original bytes untouched',
          io.open(path, encoding='utf-8').read() == good)


# ─── router ───────────────────────────────────────────────────────────────────

def test_router_set_model_refreshes_cache():
    print('[router: set_model refreshes status cache]')
    _reset_appdata()
    import Intelligence.llm_router as LR
    LR.LLMRouter._instance = None
    r = LR.LLMRouter()
    r._status_cache = {'claude': {'available': True, 'model': 'OLD',
                                  'display_name': 'Claude', 'active': True,
                                  'supports_vision': True, 'probed': True}}
    import time
    r._status_ts = time.time()
    r.set_model('claude', 'claude-opus-5')
    check('settings has the new model',
          _fresh_settings().get_provider_model('claude') == 'claude-opus-5')
    # The composer model chip renders from this snapshot and nothing re-probes
    # on a timer, so a stale entry here is what the user actually sees.
    check('status snapshot has the new model',
          r.get_status_instant()['claude']['model'] == 'claude-opus-5',
          r.get_status_instant()['claude']['model'])


def test_router_partial_probe_does_not_arm_ttl():
    print('[router: probe_provider TTL]')
    _reset_appdata()
    import Intelligence.llm_router as LR
    LR.LLMRouter._instance = None
    r = LR.LLMRouter()
    r._status_cache = None
    r._status_ts = 0.0
    r.probe_provider('claude')
    # Arming the 30s TTL from a one-provider probe would serve a snapshot
    # missing the other four and blank their status dots.
    check('single-provider probe leaves TTL disarmed', r._status_ts == 0.0,
          r._status_ts)
    check('probed provider still merged into the cache',
          'claude' in (r._status_cache or {}))


def test_router_scoped_switch_does_not_persist():
    print('[router: scoped provider switch]')
    _reset_appdata()
    import Intelligence.llm_router as LR
    LR.LLMRouter._instance = None
    r = LR.LLMRouter()
    _fresh_settings().set_active_provider('claude')

    # A project workspace's provider override is scoped to that project.
    r.switch_provider('ollama', persist=False)
    check('scoped switch leaves the global default alone',
          _fresh_settings().get_active_provider() == 'claude',
          _fresh_settings().get_active_provider())
    check('scoped switch still swaps the live provider',
          r.get_active_name() == 'ollama')

    # An explicit user choice in LLMs Setting still persists.
    r.switch_provider('openai')
    check('explicit switch persists',
          _fresh_settings().get_active_provider() == 'openai')


def test_router_status_instant_is_offline():
    print('[router: get_status_instant does no I/O]')
    _reset_appdata()
    import Intelligence.llm_router as LR
    LR.LLMRouter._instance = None
    r = LR.LLMRouter()

    calls = []
    for name, prov in r._providers.items():
        def _boom(_n=name):
            calls.append(_n)
            raise AssertionError('get_active_model() called from the UI path')
        prov.get_active_model = _boom

    r._status_cache = None
    snap = r.get_status_instant()
    # This runs under Dispatcher.Invoke; for Ollama/LM Studio get_active_model
    # does real HTTP, which froze Revit while opening LLMs Setting.
    check('no provider network call from get_status_instant', not calls, calls)
    check('unprobed entries are flagged as guesses',
          all(v.get('probed') is False for v in snap.values()),
          {k: v.get('probed') for k, v in snap.items()})


# ─── ollama ───────────────────────────────────────────────────────────────────

def test_ollama_host_persists():
    print('[ollama: host persistence]')
    _reset_appdata()
    import Intelligence.ollama_provider as OP
    import Intelligence.local_llm as LL

    p = OP.OllamaProvider()
    check('defaults to localhost', p._get_host() == 'http://localhost:11434',
          p._get_host())

    p.set_host('http://192.168.1.10:11434')
    check('host written to settings',
          _fresh_settings().get_api_key('Ollama_Host') == 'http://192.168.1.10:11434')
    # A brand-new provider instance is what the next Revit session builds.
    check('survives a restart',
          OP.OllamaProvider()._get_host() == 'http://192.168.1.10:11434')
    check('local_llm agrees', LL.get_host() == 'http://192.168.1.10:11434')
    check('configured host probed first',
          OP.OllamaProvider()._candidate_hosts()[0] == 'http://192.168.1.10:11434')


def test_ollama_active_model_uses_configured_host():
    print('[ollama: get_active_model honours the host]')
    _reset_appdata()
    import Intelligence.ollama_provider as OP
    p = OP.OllamaProvider()
    p.set_host('http://192.168.1.10:11434')

    seen = {}

    def _fake_probe():
        seen['host'] = p._get_host()
        return p._get_host(), ['qwen2.5:0.5b', 'qwen3:14b']
    p._probe_tags = _fake_probe

    model = p.get_active_model()
    # This used to route through local_llm's module-level OLLAMA_HOST, so a
    # remote Ollama got a green "Ready" dot but "No model selected".
    check('ranked from the configured host', seen.get('host') ==
          'http://192.168.1.10:11434', seen)
    check('returns a real installed model', model in ('qwen2.5:0.5b', 'qwen3:14b'),
          model)


def test_local_llm_pick_best_is_pure():
    print('[local_llm: pick_best]')
    import Intelligence.local_llm as LL
    check('empty list → None', LL.pick_best([]) is None)
    check('quality mode prefers reasoning + size',
          LL.pick_best(['llama3:8b', 'qwen3:14b', 'qwen2.5:0.5b'],
                       prefer_capable=True) == 'qwen3:14b')
    check('fast mode prefers the preferred list',
          LL.pick_best(['llama3:8b', 'qwen2.5:0.5b']) == 'qwen2.5:0.5b')
    check('unknown names still return something',
          LL.pick_best(['mystery:1b']) == 'mystery:1b')


def test_wants_json_contract():
    print('[providers: _wants_json]')
    from Intelligence.llm_provider import BaseLLMProvider as B
    check('none → False', B._wants_json(None) is False)
    check('empty dict → False', B._wants_json({}) is False)
    check('json_object dict → True', B._wants_json({'type': 'json_object'}) is True)
    check('"json" string → True', B._wants_json('json') is True)
    check('text type → False', B._wants_json({'type': 'text'}) is False)


def test_ollama_json_is_opt_in():
    print('[ollama: JSON grammar is opt-in]')
    _reset_appdata()
    import Intelligence.ollama_provider as OP

    captured = []

    def _fake_post(url, payload, timeout_ms=None, **kw):
        captured.append(payload)
        return json.dumps({'message': {'content': 'hello'}})

    orig = OP.http_post
    OP.http_post = _fake_post
    try:
        p = OP.OllamaProvider()
        p.set_model('qwen3:14b')

        p.chat([], 'sys', 'plain question', max_tokens=50)
        # Test Connection, the docked pane and spell-check all land here.
        # 'format' used to be hardcoded on, so they got JSON back.
        check('plain chat sends no format', 'format' not in captured[-1],
              captured[-1].get('format'))

        p.chat([], 'sys', 'give me json', max_tokens=50,
               response_format={'type': 'json_object'})
        check('json_object opts in', captured[-1].get('format') == 'json')

        p.chat([], 'sys', 'plain again', max_tokens=50,
               response_format={'type': 'text'})
        check('text format stays off', 'format' not in captured[-1])
    finally:
        OP.http_post = orig


def test_json_callers_opt_in():
    """The two internal callers that json.loads() the reply must ask for JSON."""
    print('[callers: explicit response_format]')
    import inspect
    from Intelligence.agents import dispatcher
    from Intelligence.comments import comment_agent

    src = inspect.getsource(dispatcher.AgentDispatcher._llm_stage)
    check('dispatcher._llm_stage requests json_object',
          'response_format' in src and 'json_object' in src)

    src2 = inspect.getsource(comment_agent)
    check('comment_agent._propose requests json_object',
          src2.count('json_object') >= 1)


def main():
    test_settings_merge_on_write()
    test_settings_corrupt_quarantine()
    test_settings_first_run_still_writes()
    test_settings_unreadable_refuses_write()
    test_router_set_model_refreshes_cache()
    test_router_partial_probe_does_not_arm_ttl()
    test_router_scoped_switch_does_not_persist()
    test_router_status_instant_is_offline()
    test_ollama_host_persists()
    test_ollama_active_model_uses_configured_host()
    test_local_llm_pick_best_is_pure()
    test_wants_json_contract()
    test_ollama_json_is_opt_in()
    test_json_callers_opt_in()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All assistant/LLM tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
