# -*- coding: utf-8 -*-
"""
Training-dataset store for the assistant's self-study layer.

One JSONL file at %APPDATA%/T3LabAI/training/dataset.jsonl, one example per line:

    {"messages": [{"role": "system", "content": "..."},
                  {"role": "user",   "content": "..."},
                  {"role": "assistant", "content": "..."}],
     "meta": {"source": "telemetry_miner", "quality": "ok",
              "revit_version": 2023, "ts": "2026-08-05"}}

This is the chat-SFT shape Unsloth / Axolotl / llama.cpp all consume, so
export_sft() is a near-passthrough. The enrichers append here; the external
fine-tune job (tools/train/) reads it.

Write discipline mirrors assistant_memory: ASCII-serialize then write, so an
IronPython 2.7 UnicodeEncodeError mid-write can never truncate the file. The
serializer is `core.jsonsafe.dumps`, NOT `json.dumps(ensure_ascii=True)` —
IronPython's ASCII escaper raises on the very non-ASCII input it exists to
handle, which silently dropped every Vietnamese example. Append is
line-oriented (one dumps per line) so a crash costs at most the last line,
never the whole corpus. Dedup is by a content hash of the user+assistant turns,
cached in memory and rebuilt from the file on first use.

Pure Python + guarded imports — CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Learning Dataset"

import hashlib
import io
import json
import os
import threading
import time

from core import jsonsafe

# Corpus ceiling. Beyond this the oldest lines are dropped on the next add —
# a training set, not an archive; stale examples from a renamed tool would only
# teach the model a route that no longer exists.
MAX_EXAMPLES   = 5000
MIN_CHARS      = 3        # reject degenerate user/assistant turns
# 'tool' joins the set so an agentic trajectory (Opus driving Revit via MCP:
# user goal -> assistant tool_calls -> tool results -> ... -> final answer) can
# be stored and trained on, not just flat Q&A. Assistant turns may now carry
# `tool_calls`; tool turns carry the tool result. Plain text examples are
# unchanged — they simply never use the new roles/fields.
_VALID_ROLES   = ('system', 'user', 'assistant', 'tool')

_LOCK = threading.Lock()
# Dedup index: content-hash -> True. None until first load; rebuilt from disk.
_seen = None


# ─── Storage ────────────────────────────────────────────────────────────────────

def _appdata_dir():
    """%APPDATA%/T3LabAI (or ~/T3LabAI off Windows). Mirrors telemetry/settings."""
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def _dataset_dir():
    d = os.path.join(_appdata_dir(), 'training')
    if not os.path.isdir(d):
        try:
            os.makedirs(d)
        except Exception:
            pass
    return d


def dataset_file():
    return os.path.join(_dataset_dir(), 'dataset.jsonl')


# ─── Validation + hashing ───────────────────────────────────────────────────────

def _clean_tool_calls(raw):
    """Canonicalize an assistant turn's tool_calls to [{id,name,arguments}].

    Accepts the two shapes in play: the OpenAI wire form
    ({id,type,function:{name,arguments}}) and the recorder's flat
    {id,name,arguments} (arguments a dict OR a json string). `arguments` is
    always stored as a JSON STRING so the on-disk row is a stable, trainer-
    ready shape. Returns [] when nothing usable is present.
    """
    out = []
    for c in (raw or []):
        if not isinstance(c, dict):
            continue
        fn = c.get('function') if isinstance(c.get('function'), dict) else c
        name = u'{}'.format(fn.get('name') or u'').strip()
        if not name:
            continue
        args = fn.get('arguments')
        if isinstance(args, (dict, list)):
            try:
                args = json.dumps(args, ensure_ascii=False)
            except Exception:
                args = u'{}'
        else:
            args = u'{}'.format(args if args is not None else u'{}')
        out.append({'id': u'{}'.format(c.get('id') or u''),
                    'name': name, 'arguments': args})
    return out


def _clean_messages(messages):
    """Return a sanitized message list, or None if unusable.

    Supports both a flat Q&A ({role in system/user/assistant, content:str}) and
    an agentic trajectory (assistant turns carrying `tool_calls`, plus role
    'tool' result turns). A usable example still needs a `user` prompt and an
    `assistant` target — but the target may be a tool-calling turn, not prose.
    """
    if not isinstance(messages, (list, tuple)):
        return None
    out = []
    roles = set()
    has_assistant_target = False
    for m in messages:
        if not isinstance(m, dict):
            return None
        role = m.get('role')
        if role not in _VALID_ROLES:
            return None
        content = u'{}'.format(m.get('content') or u'').strip()

        if role == 'assistant':
            tool_calls = _clean_tool_calls(m.get('tool_calls'))
            # An assistant turn is a valid target if it says something OR calls
            # a tool. A tool-calling turn legitimately has empty content.
            if len(content) < MIN_CHARS and not tool_calls:
                return None
            turn = {'role': role, 'content': content}
            if tool_calls:
                turn['tool_calls'] = tool_calls
            out.append(turn)
            has_assistant_target = True
        elif role == 'tool':
            # Tool result. Keep the id/name link so the trainer can pair it to
            # the call. Content may be short (e.g. "{}") — do not reject on size.
            if not content:
                content = u'{}'
            out.append({'role': role, 'content': content,
                        'tool_call_id': u'{}'.format(m.get('tool_call_id') or u''),
                        'name': u'{}'.format(m.get('name') or u'')})
        else:  # system / user
            if role == 'user' and len(content) < MIN_CHARS:
                return None
            out.append({'role': role, 'content': content})
        roles.add(role)

    if 'user' not in roles or not has_assistant_target:
        return None
    return out


def _hash(messages):
    """Content hash over the user/assistant/tool turns (system prompt ignored,
    so the same trajectory saved under two system prompts still dedups).

    Includes assistant tool_calls (name+arguments) and tool results, so two
    trajectories that differ only in which tools were called hash differently.
    """
    parts = []
    for m in messages:
        role = m['role']
        if role == 'system':
            continue
        parts.append(role)
        parts.append((m.get('content') or u'').lower().strip())
        for tc in m.get('tool_calls') or []:
            parts.append(tc.get('name') or u'')
            parts.append((tc.get('arguments') or u'').lower().strip())
    blob = u' '.join(parts).encode('utf-8')
    return hashlib.sha1(blob).hexdigest()


def _load_seen():
    """Rebuild the dedup index from disk. Caller holds _LOCK."""
    global _seen
    _seen = {}
    path = dataset_file()
    if not os.path.exists(path):
        return
    try:
        with io.open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                    msgs = _clean_messages(obj.get('messages'))
                    if msgs:
                        _seen[_hash(msgs)] = True
                except Exception:
                    continue
    except Exception:
        pass


# ─── Public API ─────────────────────────────────────────────────────────────────

def add_example(messages, source, quality='ok', revit_version=None, meta=None):
    """Append ONE training example. Returns (ok, note).

    Deduped by content hash, so an enricher that re-derives the same Q&A on a
    later cycle is a cheap no-op rather than a corpus full of duplicates.
    """
    global _seen
    clean = _clean_messages(messages)
    if clean is None:
        return False, u'invalid messages'
    h = _hash(clean)
    row = {
        'messages': clean,
        'meta': {
            'source':        u'{}'.format(source or u'unknown'),
            'quality':       u'{}'.format(quality or u'ok'),
            'revit_version': revit_version,
            'ts':            time.strftime('%Y-%m-%d'),
        },
    }
    if isinstance(meta, dict):
        row['meta'].update(meta)

    with _LOCK:
        if _seen is None:
            _load_seen()
        if h in _seen:
            return True, u'duplicate'
        try:
            payload = jsonsafe.dumps(row)
            with io.open(dataset_file(), 'a', encoding='utf-8') as f:
                f.write(payload + u'\n')
            _seen[h] = True
        except Exception as ex:
            return False, u'{}'.format(ex)

    # Trim occasionally rather than on every add (a full rewrite is costly).
    if len(_seen) > MAX_EXAMPLES:
        _trim(MAX_EXAMPLES)
    return True, u'added'


def add_qa(user, assistant, source, system=None, **kw):
    """Convenience wrapper for the common single-turn Q&A example."""
    msgs = []
    if system:
        msgs.append({'role': 'system', 'content': system})
    msgs.append({'role': 'user', 'content': user})
    msgs.append({'role': 'assistant', 'content': assistant})
    return add_example(msgs, source, **kw)


def add_trajectory(goal, steps, final=None, source='mcp_teacher',
                   quality='teacher', revit_version=None, system=None, **kw):
    """Store one agentic tool-use trajectory as an SFT example. (ok, note).

    Reconstructs the teacher's session as a chat-SFT trajectory:
        [system?] user(goal)
        -> assistant(tool_call #1) -> tool(result #1)
        -> assistant(tool_call #2) -> tool(result #2) ...
        -> assistant(final answer)?

    `steps` is the recorder's list of dicts, each:
        {'tool': str, 'arguments': dict|str, 'result': dict|str, 'is_error': bool}
    The external MCP bridge issues one tools/call at a time, so one step == one
    assistant tool-call turn + its tool-result turn. Pure/deterministic; the
    caller supplies goal/final (a bare trajectory with no goal is rejected —
    the session miner fills that in before calling).
    """
    goal = u' '.join(u'{}'.format(goal or u'').split())
    if len(goal) < MIN_CHARS:
        return False, u'no goal'
    if not steps and not final:
        return False, u'empty trajectory'

    msgs = []
    if system:
        msgs.append({'role': 'system', 'content': system})
    msgs.append({'role': 'user', 'content': goal})

    for i, step in enumerate(steps or []):
        if not isinstance(step, dict):
            continue
        name = u'{}'.format(step.get('tool') or u'').strip()
        if not name:
            continue
        call_id = u'call_{}'.format(i)
        result = step.get('result')
        if not isinstance(result, (str, bytes)) and not isinstance(result, type(u'')):
            try:
                result = json.dumps(result, ensure_ascii=False)
            except Exception:
                result = u'{}'.format(result)
        msgs.append({'role': 'assistant', 'content': u'',
                     'tool_calls': [{'id': call_id, 'name': name,
                                     'arguments': step.get('arguments')}]})
        msgs.append({'role': 'tool', 'content': u'{}'.format(result),
                     'tool_call_id': call_id, 'name': name})

    # Append a final answer only when it is substantive — a trivial one-word
    # "final" must not invalidate an otherwise-good tool-call trajectory (the
    # tool-call turns are valid targets on their own).
    if final and len(u'{}'.format(final).strip()) >= MIN_CHARS:
        msgs.append({'role': 'assistant', 'content': u'{}'.format(final)})

    return add_example(msgs, source, quality=quality,
                       revit_version=revit_version, **kw)


def iter_examples():
    """Yield every valid example dict on disk (malformed lines skipped)."""
    path = dataset_file()
    if not os.path.exists(path):
        return
    try:
        with io.open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                except Exception:
                    continue
                if _clean_messages(obj.get('messages')):
                    yield obj
    except Exception:
        return


def stats():
    """{'count', 'by_source', 'by_quality'} — cheap corpus summary."""
    count = 0
    by_source = {}
    by_quality = {}
    for obj in iter_examples():
        count += 1
        meta = obj.get('meta') or {}
        s = meta.get('source') or u'unknown'
        q = meta.get('quality') or u'ok'
        by_source[s] = by_source.get(s, 0) + 1
        by_quality[q] = by_quality.get(q, 0) + 1
    return {'count': count, 'by_source': by_source, 'by_quality': by_quality}


def export_sft(out_path):
    """Write a clean, deduped {"messages": [...]} JSONL for the trainer.

    Returns the number of examples written. Near-passthrough — the on-disk
    format is already chat-SFT — but it re-validates and de-dupes so the
    training job never sees a malformed or repeated line.
    """
    seen = {}
    n = 0
    try:
        with io.open(out_path, 'w', encoding='utf-8') as f:
            for obj in iter_examples():
                msgs = _clean_messages(obj.get('messages'))
                if not msgs:
                    continue
                h = _hash(msgs)
                if h in seen:
                    continue
                seen[h] = True
                f.write(jsonsafe.dumps({'messages': msgs}) + u'\n')
                n += 1
    except Exception:
        return n
    return n


def _trim(keep):
    """Keep only the newest `keep` lines. Caller must NOT hold _LOCK (re-enters
    via the rewrite); called opportunistically from add_example."""
    global _seen
    with _LOCK:
        path = dataset_file()
        try:
            with io.open(path, 'r', encoding='utf-8') as f:
                lines = [ln for ln in f if ln.strip()]
        except Exception:
            return
        if len(lines) <= keep:
            return
        lines = lines[-keep:]
        try:
            body = u''.join(lines if lines[-1].endswith(u'\n')
                            else [l if l.endswith(u'\n') else l + u'\n'
                                  for l in lines])
            with io.open(path, 'w', encoding='utf-8') as f:
                f.write(body)
        except Exception:
            return
        _seen = None   # force a rebuild next add


def reset():
    """Wipe the corpus. Tests + an explicit user reset only."""
    global _seen
    with _LOCK:
        try:
            if os.path.exists(dataset_file()):
                os.remove(dataset_file())
        except Exception:
            pass
        _seen = None
    return True
