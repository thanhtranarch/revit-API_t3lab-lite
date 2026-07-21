# -*- coding: utf-8 -*-
"""
Assistant Memory

Persistent cross-session memory for the T3Lab Assistant (Claude-style).
Stores short durable facts — user preferences and project conventions —
and serves them back as a system-prompt section so every new chat starts
with what previous chats already established.

Two scopes:
    global  — user preferences that apply everywhere ("answer in mm",
              "always show element ids").
    project — conventions of one project ("sheet prefix is WH-",
              "floor-to-floor is 3.3 m"), keyed by ProjectStore project id.

Facts get in three ways:
    1. Deterministic chat triggers ("remember that ...", "nhớ là ...")
       handled in script.py without an LLM round-trip.
    2. The `remember_fact` agent tool (tool_schema.make_memory_tool) —
       the model saves facts it judges durable during a conversation.
    3. The /memory command family (view / forget N / clear).

Storage: lib/Intelligence/config/assistant_memory.json — same location and
same ASCII-serialize-then-write pattern as learned_patterns.json (guards
against the IronPython 2.7 UnicodeEncodeError mid-write that truncates the
file to 0 bytes).

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Assistant Memory"

import io
import json
import os
import threading
from datetime import datetime

GLOBAL_SCOPE  = "global"
PROJECT_SCOPE = "project"

MAX_FACTS_PER_SCOPE = 30      # oldest dropped beyond this
MAX_FACT_CHARS      = 300     # longer facts are clipped
MIN_FACT_CHARS      = 4       # reject empty / degenerate saves

_LOCK = threading.Lock()


# ─── Storage ──────────────────────────────────────────────────────────────────

def _memory_file():
    lib_dir = os.path.dirname(os.path.abspath(__file__))
    config_dir = os.path.join(lib_dir, 'config')
    if not os.path.exists(config_dir):
        try:
            os.makedirs(config_dir)
        except Exception:
            pass
    return os.path.join(config_dir, 'assistant_memory.json')


def _load():
    try:
        path = _memory_file()
        if os.path.exists(path):
            with io.open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                if not isinstance(data.get('global'), list):
                    data['global'] = []
                if not isinstance(data.get('projects'), dict):
                    data['projects'] = {}
                return data
    except Exception:
        pass
    return {'version': 1, 'global': [], 'projects': {}}


def _save(data):
    try:
        payload = json.dumps(data, ensure_ascii=True, indent=2)
        if isinstance(payload, bytes):
            payload = payload.decode('ascii')
        with io.open(_memory_file(), 'w', encoding='utf-8') as f:
            f.write(payload)
        return True
    except Exception:
        return False


def _pid_key(project_id):
    return u'{}'.format(project_id) if project_id else None


# ─── Core API ─────────────────────────────────────────────────────────────────

def add_fact(text, scope=PROJECT_SCOPE, project_id=None):
    """Store one durable fact. Returns (ok, note) — note is user-showable."""
    text = u' '.join(u'{}'.format(text or u'').split())
    if len(text) < MIN_FACT_CHARS:
        return False, u'Nothing concrete to remember in that.'
    if len(text) > MAX_FACT_CHARS:
        text = text[:MAX_FACT_CHARS].rstrip()
    if scope not in (GLOBAL_SCOPE, PROJECT_SCOPE):
        scope = PROJECT_SCOPE
    if scope == PROJECT_SCOPE and not project_id:
        scope = GLOBAL_SCOPE   # no active project — keep the fact anyway
    with _LOCK:
        data = _load()
        if scope == PROJECT_SCOPE:
            bucket = data['projects'].setdefault(_pid_key(project_id), [])
        else:
            bucket = data['global']
        low = text.lower()
        for f in bucket:
            if (f.get('text') or u'').lower() == low:
                return True, u'Already in memory.'
        bucket.append({'text': text,
                       'created': datetime.now().strftime('%Y-%m-%d')})
        while len(bucket) > MAX_FACTS_PER_SCOPE:
            bucket.pop(0)
        ok = _save(data)
    if ok:
        where = (u'this project' if scope == PROJECT_SCOPE
                 else u'all projects')
        return True, u'Remembered ({}): {}'.format(where, text)
    return False, u'Could not write the memory file.'


def list_facts(project_id=None):
    """Facts visible to `project_id`: [(scope, fact_dict)], global first.

    Order is stable, so the 1-based numbering shown by /memory maps back
    to remove_fact() as long as memory is unchanged in between.
    """
    data = _load()
    out = [(GLOBAL_SCOPE, f) for f in data['global']]
    key = _pid_key(project_id)
    if key:
        for f in data['projects'].get(key, []):
            out.append((PROJECT_SCOPE, f))
    return out


def remove_fact(number, project_id=None):
    """Remove fact #number (1-based, list_facts order). (ok, removed_text)."""
    with _LOCK:
        data = _load()
        key = _pid_key(project_id)
        combined = [(GLOBAL_SCOPE, f) for f in data['global']]
        if key:
            for f in data['projects'].get(key, []):
                combined.append((PROJECT_SCOPE, f))
        try:
            idx = int(number) - 1
        except (TypeError, ValueError):
            return False, None
        if idx < 0 or idx >= len(combined):
            return False, None
        scope, fact = combined[idx]
        if scope == GLOBAL_SCOPE:
            data['global'].remove(fact)
        else:
            data['projects'][key].remove(fact)
        ok = _save(data)
    return ok, (fact.get('text') or u'')


def clear_facts(project_id=None, everything=False):
    """Clear this project's facts (+ global too when everything=True).

    Returns how many facts were removed.
    """
    with _LOCK:
        data = _load()
        n = 0
        key = _pid_key(project_id)
        if key and key in data['projects']:
            n += len(data['projects'][key])
            data['projects'][key] = []
        if everything:
            n += len(data['global'])
            data['global'] = []
        if n:
            _save(data)
    return n


# ─── Prompt / report rendering ────────────────────────────────────────────────

def build_memory_block(project_id=None, include_tool_hint=True):
    """System-prompt section carrying every remembered fact ('' when none).

    include_tool_hint=False for prompt paths that have no remember_fact
    tool (the legacy JSON-intent path) — mentioning an uncallable tool
    there would leak into the "message" field.
    """
    facts = list_facts(project_id)
    if not facts:
        return u''
    lines = [u'## Persistent memory '
             u'(facts saved in previous chats — apply them without re-asking)']
    for i, (scope, f) in enumerate(facts):
        tag = u'all projects' if scope == GLOBAL_SCOPE else u'this project'
        lines.append(u'{}. [{}] {}'.format(i + 1, tag, f.get('text') or u''))
    if include_tool_hint:
        lines.append(
            u'When the user states a NEW durable preference or convention '
            u'("from now on...", "remember...", "our standard is..."), call '
            u'the `remember_fact` tool ONCE with a short English sentence — '
            u'scope "global" for user preferences, "project" for facts about '
            u'this model. Never store one-off requests, secrets, or element '
            u'ids. If the user contradicts a remembered fact, follow the '
            u'user and save the correction.')
    return u'\n'.join(lines)


def format_memory_report(project_id=None, viet=False):
    """Markdown chat report for the /memory command."""
    facts = list_facts(project_id)
    if not facts:
        if viet:
            return (u'Tôi chưa lưu ghi nhớ nào. Nói **"nhớ là ..."** hoặc '
                    u'**"remember that ..."** để tôi ghi nhớ sở thích của bạn '
                    u'hoặc quy ước của dự án cho các lần chat sau.')
        return (u'Nothing in memory yet. Say **"remember that ..."** and I will '
                u'keep that preference or project convention for future chats.')
    lines = []
    if viet:
        lines.append(u'**Ghi nhớ hiện có** ({} mục):'.format(len(facts)))
    else:
        lines.append(u'**Current memory** ({} facts):'.format(len(facts)))
    lines.append(u'')
    lines.append(u'| # | Scope | Fact | Saved |')
    lines.append(u'|---|-------|------|-------|')
    for i, (scope, f) in enumerate(facts):
        tag = u'All projects' if scope == GLOBAL_SCOPE else u'This project'
        txt = (f.get('text') or u'').replace(u'|', u'\\|')
        lines.append(u'| {} | {} | {} | {} |'.format(
            i + 1, tag, txt, f.get('created') or u''))
    lines.append(u'')
    if viet:
        lines.append(u'Xóa 1 mục: `/memory forget <số>` · Xóa hết: '
                     u'`/memory clear`')
    else:
        lines.append(u'Remove one: `/memory forget <number>` · Remove all: '
                     u'`/memory clear`')
    return u'\n'.join(lines)
