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

from core import jsonsafe

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
    """Persist the exemplar list. ASCII-serialize-then-write via jsonsafe
    (IronPython-safe — plain json.dumps(ensure_ascii=True) raises there)."""
    try:
        payload = jsonsafe.dumps(list(exemplars or []), indent=1)
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


def _tool_result_failed(text):
    """True when a tool turn's content reads as a failure.

    Deliberately shallow — the result is JSON produced by core/server.py, whose
    failure shape is a top-level "error" key, and the MCP transport wraps
    exceptions as "Error executing tool: ...".
    """
    low = u"{}".format(text or u"").lstrip()[:400].lower()
    return low.startswith('error') or '"error"' in low or "'error'" in low


def trajectory_to_exemplar(messages):
    """A tool-use row -> {user, trace, kind:'tool'}, or None if not usable.

    trace mirrors _LOCAL_GENERAL_FEWSHOT: "call a(args); call b(args); reply …".

    Calls whose tool turn came back an ERROR are dropped, so the exemplar shows
    only the path that worked. The captured trajectory keeps the failure — it is
    real signal for a fine-tune, where the model also sees the error text — but
    a few-shot line is read as an instruction, and one distilled verbatim told
    the model to `call revit_override_color(color=...)` (the call that gets
    rejected for missing ids) BEFORE fetching the ids. That is precisely the
    "did work nobody asked for" behaviour the guard elsewhere has to clean up.
    """
    user = _first_user(messages)
    if not user:
        return None

    # Which call ids produced an error, so their assistant turn can be skipped.
    failed_ids = set()
    for m in messages or []:
        if isinstance(m, dict) and m.get('role') == 'tool' \
                and _tool_result_failed(m.get('content')):
            failed_ids.add(m.get('tool_call_id'))

    calls = []
    final = u""
    for m in messages or []:
        if not isinstance(m, dict):
            continue
        if m.get('role') == 'assistant':
            for tc in m.get('tool_calls') or []:
                name = u"{}".format(tc.get('name') or u"").strip()
                if not name or tc.get('id') in failed_ids:
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


def _has_diacritics(text):
    """True when the prompt carries Vietnamese tone marks.

    Deliberately crude — this is a spread heuristic for picking exemplars, not
    language detection (Intelligence.language does that, and is not importable
    from this leaf module). Unaccented Vietnamese counts as "plain" here, which
    is fine: what must not happen is a block with no accented example at all.
    """
    return any(ord(c) > 127 for c in text or u'')


def _tool_of(ex):
    """First tool name in an exemplar's trace ('' for a plain Q&A)."""
    return (ex.get('trace') or u'').split(u'(')[0].strip()


def _is_multi_step(ex):
    """True when the trace chains more than one tool call."""
    return (ex.get('trace') or u'').count(u'call ') > 1


def _by_teaching_value(items):
    """Reorder one language's exemplars by how much they teach:

      1. multi-step traces — chaining calls, and recovering when the first one
         is rejected, is precisely what a small model gets wrong; a single
         `call list_levels()` it can already copy from the static few-shot;
      2. then the first appearance of each tool, so the block spans the API;
      3. then repeats.

    Without step 1 the whole block came out single-call: the corrective
    sequences (override_color rejected for missing ids -> fetch ids -> retry,
    create_view ignoring view_name -> rename_element) sat at the back of the
    queue and never fitted in the 7 tool slots.
    """
    multi, first, rest, seen = [], [], [], set()
    for ex in items:
        name = _tool_of(ex)
        if _is_multi_step(ex):
            multi.append(ex)
        elif name and name in seen:
            rest.append(ex)
        else:
            first.append(ex)
        if name:
            seen.add(name)
    return multi + first + rest


def _spread(items, limit):
    """Take `limit` items, alternating between languages, distinct tools first.

    Plain truncation used to decide this, and the result was a block that
    silently excluded a whole language: the earliest-taught tasks happened to
    be English, they filled all 7 tool slots, and the six Vietnamese
    demonstrations taught right after them never reached the model — in an
    office that types Vietnamese.

    Language is the OUTER dimension on purpose. Bucketing by (language, tool)
    together looks more thorough but behaves identically to truncation when
    every task uses a different tool: each bucket holds one item, so the
    round-robin just walks them in insertion order and English still wins every
    slot. Alternate languages first, and let tool variety order each language's
    own queue.
    """
    groups = {}
    order = []
    for ex in items:
        key = _has_diacritics(ex.get('user'))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(ex)

    queues = [_by_teaching_value(groups[k]) for k in order]
    out = []
    depth = 0
    while len(out) < limit and any(len(q) > depth for q in queues):
        for q in queues:
            if len(q) > depth:
                out.append(q[depth])
                if len(out) >= limit:
                    break
        depth += 1
    return out


def select_exemplars(rows, max_n=DEFAULT_MAX_EXEMPLARS,
                     max_per_kind=DEFAULT_MAX_PER_KIND):
    """Distil dataset rows into a de-duplicated exemplar list.

    Keeps only teacher/high-quality rows. Tool-use exemplars come first (they
    teach the thing small models get wrong — calling the right tool); Q&A fills
    the rest. Dedup is by the folded user prompt, then each kind is thinned by
    _spread so the block covers distinct tools and BOTH languages rather than
    whichever tasks happened to be demonstrated first.
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

    out = _spread(tool_ex, max_per_kind) + _spread(qa_ex, max_per_kind)
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
