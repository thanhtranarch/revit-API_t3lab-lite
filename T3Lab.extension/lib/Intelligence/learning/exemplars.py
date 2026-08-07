# -*- coding: utf-8 -*-
"""
exemplars — a PORTABLE few-shot layer distilled from teacher data.

Fine-tuning bakes behaviour into model weights, which live in one machine's
Ollama and do not travel. This module distils the high-value teacher examples in
the SFT corpus (dataset.jsonl) into a small set of compact few-shot traces, saved
to a GIT-TRACKED file (lib/Intelligence/config/teacher_exemplars.json — same place
and precedent as learned_patterns.json). Committed, those exemplars are injected
into the local model's system prompt on ANY machine — so a plain qwen3:14b
responds in the taught style without a re-train.

The block is appended to the STATIC (cache-friendly) part of the prompt for LOCAL
models only, mirroring agent_loop._LOCAL_GENERAL_FEWSHOT and SpecialistSpec.few_shot.

Pure helpers (converters, selection, block rendering) take plain dicts so they are
CPython-3 unit-testable; promote_from_dataset() wires them to the real dataset.
Nothing here ever raises.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Teacher Exemplars"

import io
import json
import os

# Sources/qualities worth promoting into the portable prompt layer.
_GOOD_QUALITY = ('teacher', 'high')

# Keep the injected block small — it rides in the cached system prompt every turn.
DEFAULT_MAX_EXEMPLARS = 10
DEFAULT_MAX_PER_KIND  = 7
DEFAULT_MAX_CHARS     = 2000
_MAX_REPLY_CHARS      = 180
_MAX_ARG_VAL_CHARS    = 24
_MAX_ARGS_PER_CALL    = 4
_MAX_CALLS_PER_TRACE  = 4

BLOCK_HEADER = u"## Learned examples (from teaching — follow this style)"


# ─── Storage (git-tracked, like learned_patterns.json) ───────────────────────────

def _exemplars_file():
    # …/Intelligence/learning/exemplars.py -> …/Intelligence/config/teacher_exemplars.json
    here = os.path.dirname(os.path.abspath(__file__))          # learning/
    config_dir = os.path.join(os.path.dirname(here), 'config')  # Intelligence/config
    if not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir)
        except Exception:
            pass
    return os.path.join(config_dir, 'teacher_exemplars.json')


def load_exemplars():
    """Return the list of stored exemplar dicts (or [] on any problem)."""
    try:
        path = _exemplars_file()
        if os.path.exists(path):
            with io.open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                return [e for e in data if isinstance(e, dict) and e.get('user')]
    except Exception:
        pass
    return []


def save_exemplars(exemplars):
    """Persist the exemplar list. ASCII-serialize-then-write (IronPython-safe)."""
    try:
        payload = json.dumps(list(exemplars or []), ensure_ascii=True, indent=1)
        if isinstance(payload, bytes):
            payload = payload.decode('ascii')
        with io.open(_exemplars_file(), 'w', encoding='utf-8') as f:
            f.write(payload)
        return True
    except Exception:
        return False


# ─── Pure conversion helpers ─────────────────────────────────────────────────────

def _fold(text):
    try:
        return u" ".join(u"{}".format(text or u"").split()).lower()
    except Exception:
        return u""


def _clip(text, n):
    text = u"{}".format(text or u"").strip()
    return text if len(text) <= n else text[:n].rstrip() + u"…"


def _compact_args(arguments):
    """Render a tool call's arguments as a short 'k=v, k=v' string."""
    args = arguments
    if not isinstance(args, dict):
        try:
            args = json.loads(args) if args else {}
        except Exception:
            args = {}
    if not isinstance(args, dict) or not args:
        return u""
    parts = []
    for k in list(args.keys())[:_MAX_ARGS_PER_CALL]:
        v = args[k]
        if isinstance(v, (list, dict)):
            v = u"[…]" if isinstance(v, list) else u"{…}"
        parts.append(u"{}={}".format(k, _clip(v, _MAX_ARG_VAL_CHARS)))
    return u", ".join(parts)


def _first_user(messages):
    for m in messages or []:
        if isinstance(m, dict) and m.get('role') == 'user':
            return u"{}".format(m.get('content') or u"").strip()
    return u""


def trajectory_to_exemplar(messages):
    """A tool-use row -> {user, trace, kind:'tool'}, or None if not usable.

    trace mirrors _LOCAL_GENERAL_FEWSHOT: "call a(args); call b(args); reply …".
    """
    user = _first_user(messages)
    if not user:
        return None
    calls = []
    final = u""
    for m in messages or []:
        if not isinstance(m, dict):
            continue
        if m.get('role') == 'assistant':
            for tc in m.get('tool_calls') or []:
                name = u"{}".format(tc.get('name') or u"").strip()
                if not name:
                    continue
                ca = _compact_args(tc.get('arguments'))
                calls.append(u"{}({})".format(name, ca))
            txt = u"{}".format(m.get('content') or u"").strip()
            if txt and not (m.get('tool_calls')):
                final = txt
    if not calls:
        return None
    calls = calls[:_MAX_CALLS_PER_TRACE]
    trace = u"; ".join(u"call " + c for c in calls)
    if final:
        trace += u'; reply "{}"'.format(_clip(final, _MAX_REPLY_CHARS))
    return {'user': user, 'trace': trace, 'kind': 'tool'}


def qa_to_exemplar(messages):
    """A plain Q&A row -> {user, trace, kind:'qa'}, or None if not usable."""
    user = _first_user(messages)
    if not user:
        return None
    answer = u""
    for m in messages or []:
        if isinstance(m, dict) and m.get('role') == 'assistant':
            answer = u"{}".format(m.get('content') or u"").strip()
    if not answer:
        return None
    return {'user': user, 'trace': u'reply "{}"'.format(
        _clip(answer, _MAX_REPLY_CHARS)), 'kind': 'qa'}


def _row_to_exemplar(row):
    """Pick the right converter by row shape. Trajectory (has tool turns) wins."""
    messages = row.get('messages') if isinstance(row, dict) else None
    if not messages:
        return None
    has_tool = any(isinstance(m, dict) and m.get('role') == 'tool'
                   for m in messages)
    return trajectory_to_exemplar(messages) if has_tool \
        else qa_to_exemplar(messages)


def select_exemplars(rows, max_n=DEFAULT_MAX_EXEMPLARS,
                     max_per_kind=DEFAULT_MAX_PER_KIND):
    """Distil dataset rows into a de-duplicated exemplar list.

    Keeps only teacher/high-quality rows. Tool-use exemplars come first (they
    teach the thing small models get wrong — calling the right tool); Q&A fills
    the rest. Dedup is by the folded user prompt.
    """
    tool_ex, qa_ex = [], []
    seen = set()
    for row in rows or []:
        meta = row.get('meta') if isinstance(row, dict) else None
        quality = (meta or {}).get('quality') if isinstance(meta, dict) else None
        if quality not in _GOOD_QUALITY:
            continue
        ex = _row_to_exemplar(row)
        if not ex:
            continue
        key = _fold(ex['user'])
        if not key or key in seen:
            continue
        seen.add(key)
        (tool_ex if ex['kind'] == 'tool' else qa_ex).append(ex)

    out = tool_ex[:max_per_kind] + qa_ex[:max_per_kind]
    return out[:max_n]


def build_exemplar_block(exemplars, max_chars=DEFAULT_MAX_CHARS):
    """Render exemplars as a compact few-shot block, or u'' when empty.

    Same shape as _LOCAL_GENERAL_FEWSHOT so the local model reads one consistent
    style. Bounded by max_chars — drops trailing exemplars that would overflow.
    """
    if not exemplars:
        return u""
    lines = [BLOCK_HEADER]
    for ex in exemplars:
        if not isinstance(ex, dict) or not ex.get('user') or not ex.get('trace'):
            continue
        line = u'User: "{}" -> {}'.format(ex['user'].strip(), ex['trace'].strip())
        candidate = u"\n".join(lines + [line])
        if len(candidate) > max_chars and len(lines) > 1:
            break
        lines.append(line)
    if len(lines) <= 1:
        return u""
    return u"\n".join(lines)


# ─── Promotion (dataset -> committed exemplar file) ──────────────────────────────

def promote_from_dataset(iter_examples=None, save=None,
                         max_n=DEFAULT_MAX_EXEMPLARS,
                         max_per_kind=DEFAULT_MAX_PER_KIND):
    """Distil the current dataset's teacher rows into teacher_exemplars.json.

    I/O injected (defaults wired to the real dataset + file) so tests drive it
    with plain rows. Returns {status, count}. Never raises.
    """
    try:
        if iter_examples is None:
            from Intelligence.learning import dataset as _ds
            iter_examples = _ds.iter_examples
        if save is None:
            save = save_exemplars
    except Exception as ex:
        return {'status': u'unavailable: {}'.format(ex), 'count': 0}
    try:
        rows = list(iter_examples())
    except Exception as ex:
        return {'status': u'read failed: {}'.format(ex), 'count': 0}
    exemplars = select_exemplars(rows, max_n=max_n, max_per_kind=max_per_kind)
    try:
        ok = save(exemplars)
    except Exception as ex:
        return {'status': u'save failed: {}'.format(ex), 'count': 0}
    return {'status': 'ok' if ok else 'save failed', 'count': len(exemplars)}
