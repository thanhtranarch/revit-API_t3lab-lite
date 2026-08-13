# -*- coding: utf-8 -*-
"""
CPython 3 smoke test for the Claude provider's extended-thinking payload
shapes (dev/test_knowledge.py conventions — plain asserts, no framework).

Guards against a real regression: Claude Opus 4.7/4.8, Sonnet 5 and Fable 5
reject the legacy thinking shape (`{"type": "enabled", "budget_tokens": N}`)
with an HTTP 400 — only `{"type": "adaptive"}` works there. Since this
provider resolves models live from /v1/models with no hardcoded model list,
it cannot know in advance which shape the account's selected model wants;
this test locks down the payload builder's two shapes plus the fallback
heuristic that flips between them at runtime.

Run:  python3 dev/test_llm_config.py
Exit code 0 = all pass.
"""
from __future__ import unicode_literals

import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

import tempfile
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_test_')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def test_claude_thinking_payload_shapes():
    print('claude_provider: extended-thinking payload shapes')
    from Intelligence.claude_provider import ClaudeProvider

    p = ClaudeProvider()

    adaptive = p._agent_payload(
        'claude-opus-4-8', 'sys', [], [], 1500, stream=False,
        thinking_mode='adaptive')
    check('adaptive: thinking.type == "adaptive"',
          adaptive.get('thinking') == {'type': 'adaptive'}, adaptive.get('thinking'))
    check('adaptive: no budget_tokens field',
          'budget_tokens' not in adaptive.get('thinking', {}))
    check('adaptive: output_config.effort set',
          adaptive.get('output_config', {}).get('effort') is not None,
          adaptive.get('output_config'))

    legacy = p._agent_payload(
        'claude-opus-4-5', 'sys', [], [], 1500, stream=False,
        thinking_mode='legacy')
    check('legacy: thinking.type == "enabled"',
          legacy.get('thinking', {}).get('type') == 'enabled', legacy.get('thinking'))
    check('legacy: budget_tokens present and < max_tokens',
          0 < legacy.get('thinking', {}).get('budget_tokens', 0) < legacy['max_tokens'],
          legacy.get('thinking'))

    off = p._agent_payload(
        'claude-opus-4-8', 'sys', [], [], 1500, stream=False,
        thinking_mode=None)
    check('off: no thinking key at all', 'thinking' not in off, off)
    check('off: no output_config key', 'output_config' not in off, off)


def test_thinking_rejected_heuristic():
    print('claude_provider: thinking-rejection detection')
    from Intelligence.claude_provider import ClaudeProvider

    p = ClaudeProvider()
    rejected = RuntimeError(
        u'API error: {"type":"error","error":{"type":"invalid_request_error",'
        u'"message":"thinking.budget_tokens: Extra inputs are not permitted"}}')
    unrelated = RuntimeError('API error: {"type":"error","error":{"type":'
                             '"authentication_error","message":"invalid x-api-key"}}')
    check('flags a thinking-shape 400', p._thinking_rejected(rejected))
    check('does not flag an unrelated error', not p._thinking_rejected(unrelated))


def test_legacy_beta_header_only_on_legacy_mode():
    print('claude_provider: legacy-only beta header')
    # Re-derive the same closure logic chat_agent() builds, since _headers()
    # is a local closure and not a standalone method.
    from Intelligence.claude_provider import _THINK_BETA

    def _headers(mode):
        h = {'x-api-key': 'k', 'anthropic-version': '2023-06-01'}
        if mode == 'legacy':
            h['anthropic-beta'] = _THINK_BETA
        return h

    check('adaptive mode: no beta header', 'anthropic-beta' not in _headers('adaptive'))
    check('legacy mode: beta header present', _headers('legacy').get('anthropic-beta') == _THINK_BETA)
    check('no thinking: no beta header', 'anthropic-beta' not in _headers(None))


def test_no_dated_model_ids_hardcoded():
    """Regression: no request payload should pin a dated model snapshot.

    Every provider in Intelligence/ resolves models live from /v1/models —
    intentional, so an account's actual available models (and Anthropic's own
    snapshot rotation behind an alias like "claude-haiku-4-5") are never
    second-guessed by a stale hardcoded ID. batch_out_assistant.py used to be
    the one exception (a standalone raw-HTTP caller, outside the LLMRouter
    provider stack) pinned to "claude-haiku-4-5-20251001" — silently stale the
    day Anthropic rotates that alias to a newer snapshot. Locks down that it
    stays on the un-dated alias.
    """
    print('batch_out_assistant: no dated Claude model ID')
    import re
    import io as _io

    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(repo, 'T3Lab.extension', 'lib', 'Services',
                        'batch_out_assistant.py')
    src = _io.open(path, encoding='utf-8').read()
    dated = re.findall(r'claude-[a-z0-9.-]*-20\d{6}', src)
    check('no dated Claude model snapshot in batch_out_assistant.py',
          not dated, dated)
    check('still targets the Haiku alias', '"claude-haiku-4-5"' in src)


if __name__ == '__main__':
    test_claude_thinking_payload_shapes()
    test_thinking_rejected_heuristic()
    test_legacy_beta_header_only_on_legacy_mode()
    test_no_dated_model_ids_hardcoded()

    print()
    if FAILURES:
        print('FAILED: {}'.format(', '.join(FAILURES)))
        sys.exit(1)
    print('All checks passed.')
    sys.exit(0)
