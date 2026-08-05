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

Claude compatibility
--------------------
A Claude skill ships the same `<name>/SKILL.md` layout but declares only
`name` + `description` (plus optional `allowed-tools`, `license`, `version`),
often as multi-line YAML (`description: >-`) or block lists. The frontmatter
parser below understands those forms, and `derive_triggers()` gives an
imported skill something for `match()` to fire on — otherwise a skill whose
activation is written in prose would only ever run via /slash.
See Intelligence/skill_installer.py for installing them from a repo link.

No PyYAML — a small hand-rolled subset parser. Pure Python, CPython-3
testable.

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
_LIST_KEYS = ('triggers', 'agents', 'tools', 'allowed-tools')

# Claude / community spellings folded onto the keys this engine reads.
# `allowed-tools` is NOT folded onto `tools`: it names Claude Code tools
# (Read, Bash, …), while `tools` means T3Lab Revit tools validated by
# tool_schema — merging them would make a reference playbook look executable.
_KEY_ALIASES = {
    'allowed_tools': 'allowed-tools',
    'allowedtools': 'allowed-tools',
    'trigger': 'triggers',
    'keywords': 'triggers',
    'agent': 'agents',
    'specialists': 'agents',
    'title': 'name',
    'summary': 'description',
}

_KEY_RE = re.compile(r'^([A-Za-z][A-Za-z0-9_. -]*):[ \t]*(.*)$')
_BLOCK_RE = re.compile(r'^[|>][+-]?$')


def _unquote(value):
    value = (value or '').strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
        return value[1:-1].strip()
    return value


def _split_list(value):
    """Comma / bracket list → list of strings (also strips per-item quotes)."""
    if isinstance(value, (list, tuple)):
        items = list(value)
    else:
        items = (value or '').strip().strip('[]').split(',')
    return [_unquote(v) for v in items if _unquote(v)]


def _parse_yaml_subset(lines):
    """`key: value` YAML subset: block scalars, block lists, continuations.

    Enough for real Claude frontmatter without a PyYAML dependency (there is
    none in Revit's IronPython). Anything it cannot read is skipped rather
    than raising — a malformed header must not take the whole skill down.
    """
    meta = {}
    key = None
    mode = None          # None | 'fold' | 'literal'
    buf = []
    items = None

    def flush():
        if key is None:
            return
        target = _KEY_ALIASES.get(key, key)
        if items is not None:
            meta[target] = (list(items) if target in _LIST_KEYS
                            else ', '.join(items))
            return
        if mode == 'literal':
            value = '\n'.join(buf).strip()
        else:
            value = ' '.join(buf).strip()
        meta[target] = (_split_list(value) if target in _LIST_KEYS
                        else _unquote(value))

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()
        indented = bool(line[:1] in (' ', '\t'))

        if not stripped:
            if key is not None and mode == 'literal':
                buf.append('')
            continue
        if stripped.startswith('#') and not indented:
            continue

        if indented and key is not None:
            if stripped.startswith('- ') or stripped == '-':
                if items is None:
                    items = []
                items.append(_unquote(stripped[1:].strip()))
            else:
                buf.append(stripped)
            continue

        m = _KEY_RE.match(line)
        if not m:
            if key is not None and not indented and mode in ('fold', None):
                buf.append(stripped)      # unindented continuation line
            continue

        flush()
        key = m.group(1).strip().lower()
        value = m.group(2).strip()
        buf, items, mode = [], None, None
        if _BLOCK_RE.match(value):
            mode = 'literal' if value[0] == '|' else 'fold'
        elif value:
            buf = [value]

    flush()
    return meta


def parse_frontmatter(text):
    """Parse a '---' fenced frontmatter block.

    Returns (meta_dict, body). Files without a fence yield ({}, text).
    Lists may be comma-separated (optionally in [brackets]) or YAML block
    lists; values may be quoted, folded (`>-`) or literal (`|`) scalars.
    """
    lines = (text or '').splitlines()
    if not lines or lines[0].strip() != '---':
        return {}, text or ''
    end = None
    for i in range(1, len(lines)):
        if lines[i].strip() in ('---', '...'):
            end = i
            break
    if end is None:
        return {}, text or ''
    return _parse_yaml_subset(lines[1:end]), '\n'.join(lines[end + 1:]).strip()


# ── Trigger derivation (Claude skills declare none) ─────────────────────────
# Claude leaves activation to the model reading `description`; T3Lab's
# match() needs literal phrases. Derivation stays deliberately conservative —
# an over-eager trigger silently hijacks unrelated requests, which is worse
# than a skill that only answers to /slash.

# Matches anywhere, not just at a line start: a folded YAML description
# collapses to one line, so "… merges documents. Triggers: "extract pdf", …"
# is the normal shape rather than the exception.
_TRIGGER_LINE_RE = re.compile(
    r'(?:\*\*)?\b'
    r'(?:triggers?|trigger words|kich hoat|kích hoạt|tu khoa|từ khóa)'
    r'(?:\*\*)?[ \t]*[:：][ \t]*(?P<body>.+)', re.I)
_QUOTED_RE = re.compile(r'["“”\'`]([^"“”\'`\n]{3,60})["“”\'`]')

_MIN_TRIGGER_CHARS = 4
_MAX_TRIGGER_CHARS = 60
_MAX_DERIVED_TRIGGERS = 12
# Phrases too generic to route on.
_TRIGGER_STOPWORDS = frozenset([
    'use', 'used', 'use when', 'when', 'this', 'that', 'skill', 'skills',
    'tool', 'tools', 'task', 'file', 'files', 'user', 'help', 'and', 'or',
    'etc', 'e.g', 'i.e', 'the', 'any', 'all', 'more', 'other', 'others',
])


def derive_triggers(skill_id, name, description, body=''):
    """Literal activation phrases for a skill that declares none.

    Sources, in order: an explicit "Triggers: …" line in the description or
    body (quoted phrases win over comma splitting), then the skill's own
    name/id as a phrase. Returns raw (unfolded) strings.
    """
    out = []
    seen = set()

    def _add(phrase):
        phrase = re.sub(r'\s+', ' ', (phrase or '')).strip(' .,;:!?-–—*_"\'`')
        low = phrase.lower()
        if (not phrase or low in seen or low in _TRIGGER_STOPWORDS
                or len(phrase) < _MIN_TRIGGER_CHARS
                or len(phrase) > _MAX_TRIGGER_CHARS):
            return
        seen.add(low)
        out.append(phrase)

    for text in (description or '', body or ''):
        for m in _TRIGGER_LINE_RE.finditer(text):
            chunk = m.group('body')
            quoted = _QUOTED_RE.findall(chunk)
            for item in (quoted if quoted
                         else re.split(r'[,;·|]|\s+/\s+', chunk)):
                _add(item)
            if len(out) >= _MAX_DERIVED_TRIGGERS:
                break
        if len(out) >= _MAX_DERIVED_TRIGGERS:
            break

    for candidate in ((skill_id or '').replace('-', ' ').replace('_', ' '),
                      name or ''):
        if candidate and len(candidate.split()) <= 5:
            _add(candidate)

    return out[:_MAX_DERIVED_TRIGGERS]


# Tie-break weight per skill source. A project ships a playbook precisely to
# override the generic one, so it must not lose a tie to a built-in.
_SOURCE_WEIGHT = {'project': 2.0, 'extra': 1.5, 'user': 1.0, 'builtin': 0.0}


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


def _active_pid():
    """Active project id, or None.

    Goes through ProjectStore rather than settings directly: it nulls out a
    STALE id (project folder deleted outside the app). Reading settings raw
    meant skills kept resolving against a dead project dir while every other
    subsystem had already agreed there was no active project.
    """
    try:
        from config.project_store import ProjectStore
        return ProjectStore().get_active_project_id()
    except Exception:
        return None


def _project_skills_dir():
    try:
        pid = _active_pid()
        if pid:
            from config.project_store import ProjectStore
            return os.path.join(ProjectStore().project_dir(pid), 'skills')
    except Exception:
        pass
    return None


def _project_disabled_skills():
    """Skill ids switched off for the active project (empty when none)."""
    try:
        pid = _active_pid()
        if not pid:
            return []
        from config.project_store import ProjectStore
        meta = ProjectStore().get_project(pid) or {}
        return list(meta.get('skills_disabled') or [])
    except Exception:
        return []


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
            # A Claude skill declares no triggers — without this fallback it
            # would sit in the registry and never activate on its own.
            triggers = meta.get('triggers') or []
            derived = False
            if not triggers:
                triggers = derive_triggers(skill_id, name, description, body)
                derived = bool(triggers)
            return {
                'id': skill_id,
                'name': name,
                'description': description,
                'triggers': [vi_text.fold_diacritics(t).lower()
                             for t in triggers],
                'triggers_derived': derived,
                'agents': meta.get('agents', []),
                'tools': meta.get('tools', []),
                # Claude Code tools the skill declares — informational only.
                'allowed_tools': meta.get('allowed-tools', []),
                # 'standard: project' marks a playbook whose concrete values
                # (codes, numbering, names, thresholds) are company-specific
                # and MUST come from the project's BEP/standard documents —
                # see _PROJECT_STANDARD_CONTRACT.
                'standard': (meta.get('standard') or '').strip().lower(),
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
        """Skill ids switched off — global preference UNION the active
        project's own list.

        `skills_disabled` has been written into every project.json since
        create_project existed but had no reader anywhere, so a project could
        never actually turn a skill off. Union (not override) is the safe
        reading: a project may narrow the skill set, never re-enable something
        the user disabled globally.
        """
        out = set()
        try:
            from config.settings import get_settings
            out |= set(get_settings().get_disabled_skills())
        except Exception:
            pass
        try:
            out |= set(_project_disabled_skills())
        except Exception:
            pass
        return out

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

    def match_scored(self, text):
        """[(skill_id, score)] for every skill whose triggers hit, best first.

        Only two skills ever reach the model (build_skills_block caps at
        max_skills), so the ORDER decides which playbook the model actually
        gets. This used to be `for sid in sorted(self._skills)` with the
        caller taking element 0 — i.e. decided by the alphabet, so
        `annotation-standard` beat `warning-triage` on a message about
        warnings for no reason but its first letter.

        Scoring, highest wins:
          * a trigger's word count — "hoan thien cmt" is far more specific
            evidence than "cmt", so a longer phrase outranks a shorter one;
          * how much of the message the best trigger covers;
          * how many DISTINCT triggers of that skill fired;
          * declared triggers over derived ones — derive_triggers() guesses
            from a Claude skill's prose and is deliberately loose;
          * source precedence project > user > builtin, so a project's own
            playbook wins a tie against the shipped one it overrides.

        Ties break on skill id, so the result stays deterministic.
        """
        self._ensure_scanned()
        folded = vi_text.fold_diacritics(text or '').lower()
        folded = re.sub(r'\s+', ' ', folded).strip()
        if not folded:
            return []
        total_words = len(folded.split()) or 1
        disabled = self._disabled()

        scored = []
        for sid in sorted(self._skills):
            if sid in disabled:
                continue
            meta = self._skills[sid]
            best_words = 0
            best_chars = 0
            n_hits = 0
            for trig in meta.get('triggers', []):
                if not trig:
                    continue
                if re.search(r'(?:^|[^a-z0-9])' + re.escape(trig)
                             + r'(?:[^a-z0-9]|$)', folded):
                    n_hits += 1
                    words = len(trig.split())
                    if words > best_words:
                        best_words = words
                        best_chars = len(trig)
                    elif words == best_words:
                        best_chars = max(best_chars, len(trig))
            if not n_hits:
                continue

            score = 10.0 * best_words                      # phrase specificity
            score += 2.0 * min(n_hits, 4)                   # corroboration
            score += 4.0 * (float(best_words) / total_words)  # query coverage
            score += best_chars / 100.0                     # longer of equal rank
            if not meta.get('triggers_derived'):
                score += 3.0
            score += _SOURCE_WEIGHT.get(meta.get('source'), 0.0)
            # 👍/👎 nudge: the votes are collected per skill (feedback.py) but
            # nothing read them back until here. Smooth and bounded to ±3 — it
            # re-orders ties and near-ties (a repeatedly down-voted playbook
            # sinks below an equally-triggered neutral one; an up-voted one
            # edges ahead) but can never override a strong phrase match at
            # 10 pts/word. With no votes net is 0 → +0.0, so the order is
            # byte-identical to before feedback existed.
            try:
                from Intelligence import feedback
                up, down = feedback.skill_score(sid)
                net = up - down
                if net:
                    score += 3.0 * (float(net) / (abs(net) + 3.0))
            except Exception:
                pass
            scored.append((sid, score))

        scored.sort(key=lambda pair: (-pair[1], pair[0]))
        return scored

    def match(self, text):
        """Skill ids whose folded triggers appear in the folded text.

        Same contract as before — a list of ids — but ordered by relevance
        instead of alphabetically. See match_scored().
        """
        return [sid for sid, _ in self.match_scored(text)]

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
1b. A SKILL IS NOT A TOOL: it is instructions for YOU to carry out. There \
is no `apply_playbook`, `apply_skill`, `run_playbook` or `playbook_name` \
— those are not tool or intent names, and writing a tool call as JSON in \
your reply text does nothing. Execute the playbook's steps using ONLY the \
real tools available to you, and plain prose for everything else.
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


# Injected ONLY when an active skill declares `standard: project`.
# Naming codes, sheet numbering, workset names, LOD matrices, suitability and
# revision codes, export naming and folder structures are COMPANY-SPECIFIC —
# they come from the client's/employer's BEP, EIR or in-house standard, not
# from a generic playbook. Without this block the model reads the illustrative
# tables in those skill bodies as "the" standard and confidently invents rules
# for a company that never agreed to them.
_PROJECT_STANDARD_CONTRACT = """\
## Project-standard precedence (this request touches a company-specific standard)
1. THE PROJECT DOCUMENTS ARE THE ONLY AUTHORITY: the project's BEP, EIR, \
in-house standard/guideline, or any document in the knowledge base defines \
the real codes, prefixes, numbering, names, LOD matrix and thresholds. \
Use the "Project reference material" excerpts and the project instructions \
first, follow them EXACTLY, and name the source file you took each rule from.
2. THE PLAYBOOK TABLES ARE A GENERIC TEMPLATE, NOT THE PROJECT'S RULE: every \
concrete code, prefix, example name and threshold written in the skill body \
below is an illustrative industry-typical placeholder. NEVER present it as \
this project's standard and never apply it silently.
3. WHEN NO PROJECT STANDARD IS FOUND: say so plainly in one line — no BEP / \
standard document was found in the knowledge base — then ask the user to add \
it, naming a concrete route: Settings → Projects → Link folder (link the \
project's BEP/standards folder, indexed in place), Settings → Projects → Add \
files (copy the BEP/standard in), or Settings → Knowledge → Add folder (a \
company-wide standards folder for every project). Then re-run the request. \
Offer the generic template ONLY as an explicitly-labelled temporary \
suggestion the user must confirm.
4. NEVER INVENT OR ASSUME: do not fabricate a project code, discipline \
prefix, revision/suitability code or LOD requirement. If the existing model \
already shows a convention (e.g. via revit_list_sheets, list_worksets, \
list_levels, revit_list_views), you may FOLLOW the observed existing pattern \
— state that you inferred it from the model, not from a standard document."""


def _needs_project_standard(engine, skill_ids):
    """True when any of the given skills declares `standard: project`."""
    for sid in skill_ids:
        meta = engine._skills.get(sid) or {}
        if (meta.get('standard') or '') == 'project':
            return True
    return False


def build_skills_block(skill_ids, max_skills=2):
    """'## Active skill' prompt block for the matched skills (capped),
    preceded by the shared execution contract — plus the project-standard
    precedence contract when a company-specific standard skill is active."""
    if not skill_ids:
        return ''
    engine = get_skills_engine()
    parts = []
    used = []
    for sid in list(skill_ids)[:max_skills]:
        body = engine.get_body(sid)
        if not body:
            continue
        meta = engine._skills.get(sid) or {}
        used.append(sid)
        parts.append('## Active skill: {}\n{}'.format(
            meta.get('name', sid), body))
    if not parts:
        return ''
    head = [_SKILLS_CONTRACT]
    if _needs_project_standard(engine, used):
        head.append(_PROJECT_STANDARD_CONTRACT)
    return '\n\n'.join(head + parts)
