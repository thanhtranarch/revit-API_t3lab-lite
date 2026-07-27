# -*- coding: utf-8 -*-
"""
SkillsEngine — Claude-style reusable instruction packages ("skills").

A skill is a markdown file with a small frontmatter header:

    ---
    name: sheet-naming-standard
    description: Quy tac dat ten sheet chuan T3Lab
    triggers: ten sheet, sheet number, doi ten sheet
    agents: revit_action, general
    tools: rename_element, create_sheet
    ---
    <markdown instruction body>

Progressive disclosure keeps local-model context budgets safe: routing
only ever sees name+description (get_catalog); the full body is injected
into the specialist's system prompt ONLY when a trigger matches
(build_skills_block, capped).

Skill sources (scanned in order; later sources may add, ids are unique
by first-wins):
    1. built-in   <extension>/lib/Intelligence/skills/
    2. user       %APPDATA%/T3LabAI/skills/   (<name>.md or <name>/SKILL.md)
    3. project    %APPDATA%/T3LabAI/projects/<pid>/skills/  (active project)

No PyYAML — a minimal key:value parser. Pure Python, CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import io
import os
import re
import threading

from Intelligence.knowledge import vi_text

_MAX_BODY_CHARS = 6000
_MAX_FILE_BYTES = 200 * 1024
_LIST_KEYS = ('triggers', 'agents', 'tools')


def parse_frontmatter(text):
    """Parse a '---' fenced frontmatter block.

    Returns (meta_dict, body). Files without a fence yield ({}, text).
    Lists are comma-separated (optionally in [brackets]).
    """
    lines = (text or '').splitlines()
    if not lines or lines[0].strip() != '---':
        return {}, text or ''
    meta = {}
    end = None
    for i in range(1, len(lines)):
        if lines[i].strip() == '---':
            end = i
            break
        line = lines[i]
        if ':' not in line or line.lstrip().startswith('#'):
            continue
        key, _, value = line.partition(':')
        key = key.strip().lower()
        value = value.strip()
        if key in _LIST_KEYS:
            value = value.strip('[]')
            meta[key] = [v.strip() for v in value.split(',') if v.strip()]
        else:
            meta[key] = value
    if end is None:
        return {}, text or ''
    return meta, '\n'.join(lines[end + 1:]).strip()


def _builtin_skills_dir():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'skills')


def _user_skills_dir():
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    d = os.path.join(base, 'T3LabAI', 'skills')
    if not os.path.isdir(d):
        try:
            os.makedirs(d)
        except Exception:
            pass
    return d


def _project_skills_dir():
    try:
        from config.settings import get_settings
        pid = get_settings().get_active_project()
        if pid:
            from config.project_store import ProjectStore
            return os.path.join(ProjectStore().project_dir(pid), 'skills')
    except Exception:
        pass
    return None


class SkillsEngine(object):
    """Singleton skill registry."""

    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(SkillsEngine, cls).__new__(cls)
                cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._skills = {}        # id -> meta dict
        self._bodies = {}        # id -> body cache
        self._scanned = False
        self._initialized = True

    # ── scanning ──────────────────────────────────────────────────────────

    def scan(self, extra_dirs=None):
        """(Re)scan all skill sources. Returns the number of skills found."""
        skills = {}
        sources = [('builtin', _builtin_skills_dir()),
                   ('user', _user_skills_dir())]
        proj = _project_skills_dir()
        if proj:
            sources.append(('project', proj))
        for label, d in (list(sources) +
                         [('extra', x) for x in (extra_dirs or [])]):
            if not d or not os.path.isdir(d):
                continue
            try:
                entries = sorted(os.listdir(d))
            except Exception:
                continue
            for entry in entries:
                path = os.path.join(d, entry)
                skill_path = None
                skill_id = None
                if os.path.isfile(path) and entry.lower().endswith('.md'):
                    skill_id = entry[:-3]
                    skill_path = path
                elif os.path.isdir(path):
                    cand = os.path.join(path, 'SKILL.md')
                    if os.path.isfile(cand):
                        skill_id = entry
                        skill_path = cand
                if not skill_path or skill_id in skills:
                    continue
                meta = self._read_meta(skill_id, skill_path, label)
                if meta:
                    skills[skill_id] = meta
        self._skills = skills
        self._bodies = {}
        self._scanned = True
        return len(skills)

    def _read_meta(self, skill_id, path, source):
        try:
            if os.path.getsize(path) > _MAX_FILE_BYTES:
                return None
            with io.open(path, 'r', encoding='utf-8', errors='replace') as f:
                text = f.read()
            meta, body = parse_frontmatter(text)
            name = meta.get('name') or skill_id
            description = meta.get('description') or \
                (body.strip().splitlines()[0][:120] if body.strip() else '')
            return {
                'id': skill_id,
                'name': name,
                'description': description,
                'triggers': [vi_text.fold_diacritics(t).lower()
                             for t in meta.get('triggers', [])],
                'agents': meta.get('agents', []),
                'tools': meta.get('tools', []),
                'path': path,
                'source': source,
            }
        except Exception:
            return None

    def _ensure_scanned(self):
        if not self._scanned:
            self.scan()

    # ── queries ───────────────────────────────────────────────────────────

    def _disabled(self):
        try:
            from config.settings import get_settings
            return set(get_settings().get_disabled_skills())
        except Exception:
            return set()

    def all_skills(self):
        """Every skill meta (including disabled), for the sidebar list."""
        self._ensure_scanned()
        disabled = self._disabled()
        out = []
        for sid in sorted(self._skills):
            meta = dict(self._skills[sid])
            meta['enabled'] = sid not in disabled
            out.append(meta)
        return out

    def get_catalog(self, enabled_only=True):
        """Cheap disclosure list: [{'id','name','description'}]."""
        self._ensure_scanned()
        disabled = self._disabled() if enabled_only else set()
        return [{'id': s['id'], 'name': s['name'],
                 'description': s['description']}
                for sid, s in sorted(self._skills.items())
                if sid not in disabled]

    def match(self, text):
        """Skill ids whose folded triggers appear in the folded text."""
        self._ensure_scanned()
        folded = vi_text.fold_diacritics(text or '').lower()
        folded = re.sub(r'\s+', ' ', folded)
        disabled = self._disabled()
        hits = []
        for sid in sorted(self._skills):
            if sid in disabled:
                continue
            for trig in self._skills[sid].get('triggers', []):
                if not trig:
                    continue
                if re.search(r'(?:^|[^a-z0-9])' + re.escape(trig)
                             + r'(?:[^a-z0-9]|$)', folded):
                    hits.append(sid)
                    break
        return hits

    def get_body(self, skill_id):
        """Full instruction body (lazy, cached, capped)."""
        self._ensure_scanned()
        if skill_id in self._bodies:
            return self._bodies[skill_id]
        meta = self._skills.get(skill_id)
        if not meta:
            return ''
        body = ''
        try:
            with io.open(meta['path'], 'r', encoding='utf-8',
                         errors='replace') as f:
                _, body = parse_frontmatter(f.read())
        except Exception:
            body = ''
        body = (body or '')[:_MAX_BODY_CHARS]
        self._bodies[skill_id] = body
        return body

    def filter_for_specialist(self, skill_ids, spec_name):
        """Keep only skills that apply to `spec_name` (no `agents` field
        in the frontmatter = applies everywhere)."""
        self._ensure_scanned()
        out = []
        for sid in (skill_ids or []):
            meta = self._skills.get(sid)
            if not meta:
                continue
            agents = meta.get('agents') or []
            if not agents or spec_name in agents:
                out.append(sid)
        return out

    def tools_for(self, skill_id):
        """Tool names a skill declares (validated later by tool_schema)."""
        self._ensure_scanned()
        meta = self._skills.get(skill_id)
        return list((meta or {}).get('tools', []))

    def agents_for(self, skill_id):
        """Specialist names a skill declares (empty = applies everywhere)."""
        self._ensure_scanned()
        meta = self._skills.get(skill_id)
        return list((meta or {}).get('agents', []))

    def specialist_for(self, skill_id, preferred=None):
        """Specialist to run a /slash-forced skill under.

        The router classifies the *message*, which for a bare "/skill-id"
        invocation is synthetic boilerplate — its wording (e.g. the word
        "MODIFY") used to decide the specialist, so a reference playbook
        could land on the action specialist and get write tools plus the
        "tô đỏ tường" few-shot. The frontmatter is the authority instead:
        keep `preferred` when the skill declares it, otherwise fall back to
        the first declared agent.

        Returns None when the skill is unknown or declares no agents (=
        applies everywhere → leave the router's choice alone).
        """
        agents = self.agents_for(skill_id)
        if not agents:
            return None
        if preferred and preferred in agents:
            return preferred
        return agents[0]

    def is_reference_skill(self, skill_id):
        """True for playbooks that declare no tools — standards/naming
        references that are answered from their body, not by running Revit
        tools."""
        self._ensure_scanned()
        if skill_id not in self._skills:
            return False
        return not self.tools_for(skill_id)

    def set_enabled(self, skill_id, enabled):
        try:
            from config.settings import get_settings
            return get_settings().set_skill_disabled(skill_id, not enabled)
        except Exception:
            return False


def get_skills_engine():
    return SkillsEngine()


# Shared execution contract injected once above every active skill body.
# Individual playbooks stay focused on their domain; the cross-cutting
# lessons (learned from real sessions) live here so ALL skills inherit
# them without 25 copies drifting apart.
_SKILLS_CONTRACT = """\
## Skill execution contract (applies to every active skill below)
1. FOLLOW THE PLAYBOOK: the active skill is the standard for this request \
— its naming rules, thresholds, orders of steps and report formats \
override your defaults.
2. ACT IN THIS TURN: checking/audit playbooks default to the ENTIRE \
project and start scanning immediately — no scope questions. Playbooks \
that MODIFY the model present the plan (targets + counts) and wait for \
one confirmation, then execute without re-asking.
3. NUMBERS COME FROM TOOL-COMPUTED FIELDS ONLY: `total_count` \
(ai_element_filter), `row_count`/`column_totals` (get_schedule_data), \
`element_counts` (analyze_model_statistics). Never count or sum raw \
`rows`/`elements` lists — they may be truncated before you see them. \
When `total_count` > returned count, raise `limit` and rescan, or report \
"checked X of Y" explicitly.
4. REPORT LIKE THE PLAYBOOK: use its table format, include element ids, \
and state clearly what was NOT covered and why."""


def build_skills_block(skill_ids, max_skills=2):
    """'## Active skill' prompt block for the matched skills (capped),
    preceded by the shared execution contract."""
    if not skill_ids:
        return ''
    engine = get_skills_engine()
    parts = []
    for sid in list(skill_ids)[:max_skills]:
        body = engine.get_body(sid)
        if not body:
            continue
        meta = engine._skills.get(sid) or {}
        parts.append('## Active skill: {}\n{}'.format(
            meta.get('name', sid), body))
    if not parts:
        return ''
    return '\n\n'.join([_SKILLS_CONTRACT] + parts)
