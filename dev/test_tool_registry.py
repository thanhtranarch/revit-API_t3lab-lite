# -*- coding: utf-8 -*-
"""
CPython 3 drift locks for the MCP tool registry in lib/core/server.py.

Run:  python3 dev/test_tool_registry.py
Exit code 0 = all pass. Plain asserts, same conventions as
dev/test_assistant_routing.py. No Revit required.

WHY THIS FILE EXISTS
--------------------
A user asked the assistant to "pin all walls in view". No pin tool existed, and
`operate_element` — the tool that should have carried it — had been left out of
the specialist tool subsets, so the model picked the nearest thing it could see
(`revit_override_color`) and painted 54 walls red while reporting success.

Silently doing a different thing is the worst failure this assistant has. Every
test here mechanises one step of the "registering a tool" checklist that, when
skipped, produces exactly that class of bug:

    registry ↔ dispatch     a tool advertised but not implemented returns None,
                            and the model reports success on it
    enum coverage           a free-text operation field means an unrecognised
                            value silently falls through to some default branch
    enum ↔ dispatch         an advertised operation with no branch is a promise
                            the server cannot keep
    _WRITE_TOOLS            a transactional tool missing here throws
                            "outside of API context" on every single call
    tool subsets            a name missing from a specialist subset is invisible
                            to that specialist — the original pin bug
    prompt rules            a multi-op tool the prompt never names is a tool the
                            model reaches for by guesswork

core.server can't be imported headlessly (it pulls the Revit API), but the
registry is a pure dict literal, so ast.literal_eval reads the whole thing.
"""
from __future__ import unicode_literals

import ast
import io
import os
import re
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXT = os.path.join(REPO, 'T3Lab.extension')
LIB = os.path.join(EXT, 'lib')
SERVER = os.path.join(LIB, 'core', 'server.py')
SPECIALISTS = os.path.join(LIB, 'Intelligence', 'agents', 'specialists.py')
TOOL_SCHEMA = os.path.join(LIB, 'Intelligence', 'tool_schema.py')
NLU = os.path.join(LIB, 'Intelligence', 'nlu_engine.py')
AGENT_LOOP = os.path.join(LIB, 'Intelligence', 'agent_loop.py')
ASSISTANT = os.path.join(EXT, 'T3Lab.tab', 'Support.panel',
                         'T3LabAssistant.pushbutton', 'script.py')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def _read(path):
    with io.open(path, encoding='utf-8') as f:
        return f.read()


# ─────────────────────────────────────────────────────────────────────────────
# Source extraction
# ─────────────────────────────────────────────────────────────────────────────

_SERVER_SRC = _read(SERVER)


def _registry():
    """The `self._tools = {...}` literal inside _register_default_tools()."""
    tree = ast.parse(_SERVER_SRC)
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef):
            continue
        if node.name != '_register_default_tools':
            continue
        for stmt in node.body:
            if not isinstance(stmt, ast.Assign):
                continue
            if not isinstance(stmt.value, ast.Dict):
                continue
            return ast.literal_eval(stmt.value)
    raise AssertionError('could not locate the _tools registry literal')


REGISTRY = _registry()


def _frozenset_literal(src, name):
    """Read a module- or class-level `NAME = frozenset([...])` / `= {...}`."""
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if not (isinstance(target, ast.Name) and target.id == name):
                continue
            value = node.value
            # frozenset([...]) / set([...])
            if (isinstance(value, ast.Call)
                    and isinstance(value.func, ast.Name)
                    and value.func.id in ('frozenset', 'set')):
                if not value.args:
                    return set()
                return set(ast.literal_eval(value.args[0]))
            if isinstance(value, (ast.Set, ast.List, ast.Tuple)):
                return set(ast.literal_eval(value))
            # READ_TOOLS | frozenset([...]) — union expressions
            if isinstance(value, ast.BinOp) and isinstance(value.op, ast.BitOr):
                return _union_names(src, value)
    raise AssertionError('could not find {}'.format(name))


def _union_names(src, node):
    """Flatten `A | frozenset([...]) | {...}` into a set of strings."""
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.BitOr):
        return _union_names(src, node.left) | _union_names(src, node.right)
    if isinstance(node, ast.Name):
        return _frozenset_literal(src, node.id)
    if (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
            and node.func.id in ('frozenset', 'set')):
        return set(ast.literal_eval(node.args[0])) if node.args else set()
    if isinstance(node, (ast.Set, ast.List, ast.Tuple)):
        return set(ast.literal_eval(node))
    raise AssertionError('unsupported union expression')


def _class_attr_set(src, class_name, attr):
    """Read a class-level `attr = frozenset([...])` (e.g. _WRITE_TOOLS)."""
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if not (isinstance(node, ast.ClassDef) and node.name == class_name):
            continue
        for stmt in node.body:
            if not isinstance(stmt, ast.Assign):
                continue
            for target in stmt.targets:
                if isinstance(target, ast.Name) and target.id == attr:
                    value = stmt.value
                    if (isinstance(value, ast.Call)
                            and isinstance(value.func, ast.Name)
                            and value.func.id in ('frozenset', 'set')):
                        return set(ast.literal_eval(value.args[0]))
                    return set(ast.literal_eval(value))
    raise AssertionError('could not find {}.{}'.format(class_name, attr))


def _dispatch_branches():
    """{tool_name: source_text} for each `elif tool_name == 'x':` branch."""
    lines = _SERVER_SRC.splitlines()
    starts = []
    pat = re.compile(r"^\s*(?:el)?if\s+tool_name\s*==\s*'([a-z0-9_]+)'\s*:")
    for i, line in enumerate(lines):
        m = pat.match(line)
        if m:
            starts.append((i, m.group(1)))
    out = {}
    for idx, (line_no, name) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        out[name] = '\n'.join(lines[line_no:end])
    return out


BRANCHES = _dispatch_branches()


def _literal_assign(src, name):
    """Value of a module- or class-level `NAME = <literal>`, or None."""
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == name:
                try:
                    return ast.literal_eval(node.value)
                except Exception:
                    return None
    return None


def _dict_literal(src, name):
    value = _literal_assign(src, name)
    return value if isinstance(value, dict) else {}


def _list_literal(src, name):
    value = _literal_assign(src, name)
    return list(value) if isinstance(value, (list, tuple)) else []


def _module_constants(src):
    """{name: source_segment} for every `NAME = <literal>` assignment.

    Module- and class-level alike (ast.walk, same as _frozenset_literal).
    Used to follow a branch that switches through a lookup table declared
    outside itself.
    """
    out = {}
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name):
                try:
                    out[target.id] = ast.get_source_segment(src, node) or ''
                except Exception:
                    pass
    return out


_SERVER_CONSTANTS = _module_constants(_SERVER_SRC)

_IMPORT_RE = re.compile(r'^\s*from\s+([\w.]+)\s+import\b', re.MULTILINE)
_NAME_RE = re.compile(r'\b[A-Za-z_][A-Za-z0-9_]*\b')


def _effective_branch_source(tool):
    """Branch source PLUS whatever it delegates the decision to.

    A dispatch branch does not have to spell every value it accepts. Two
    legitimate shapes defeat a plain substring scan of the branch body:

      * delegation to a helper module — `create_view` imports
        create_single_view / SUPPORTED_VIEW_TYPES from
        core.advanced_view_manager and lets it own the nine view types;
      * a lookup table declared outside the branch — `revit_list_views`
        normalises through the module-level _VIEW_TYPE_FILTER_MAP.

    Both were reported as advertising values the server "cannot reach", which
    was never true. The invariant worth keeping is *is this advertised value
    reachable at all*, so resolve one hop: the branch body, the source of any
    in-extension module it imports, and any server.py constant it names.
    """
    branch = BRANCHES.get(tool, '')
    if not branch:
        return ''
    parts = [branch]

    for mod in _IMPORT_RE.findall(branch):
        path = os.path.join(LIB, *mod.split('.')) + '.py'
        if os.path.isfile(path):
            parts.append(_read(path))

    for name in set(_NAME_RE.findall(branch)):
        if name in _SERVER_CONSTANTS:
            parts.append(_SERVER_CONSTANTS[name])

    return '\n'.join(parts)

# Names that are deliberately not public registry entries. Read from the
# server itself rather than restated here — a second copy in the test goes
# stale the moment a pseudo-tool is added, and then reports a live internal
# tool as unreachable dead code.
INTERNAL_TOOLS = _frozenset_literal(_SERVER_SRC, '_INTERNAL_TOOLS')

# Fields whose value the dispatch switches on. A free-text value here is how
# "asked for a section, silently got a 3D view" happens.
SWITCH_FIELDS = {'operation', 'mode', 'action', 'view_type', 'element_type',
                 'data_type', 'format', 'scope', 'match', 'detail_level',
                 'discipline', 'link_kind'}


# ─────────────────────────────────────────────────────────────────────────────
# Registry ↔ dispatch
# ─────────────────────────────────────────────────────────────────────────────

def test_every_registry_tool_has_a_dispatch_branch():
    """Advertised but unimplemented: the branch chain falls through, the tool
    returns 'Tool not implemented' or None, and the model narrates a success
    that never happened."""
    missing = sorted(n for n in REGISTRY if n not in BRANCHES)
    check('every registered tool has a dispatch branch', not missing, missing)


def test_every_dispatch_branch_is_registered():
    """The converse — dead code no provider can reach."""
    extra = sorted(n for n in BRANCHES
                   if n not in REGISTRY and n not in INTERNAL_TOOLS)
    check('every dispatch branch is registered', not extra, extra)


def test_no_duplicate_tool_names():
    """A tool whose 'name' disagrees with its registry key is unroutable: the
    schema goes out under one name and the dispatch keys off the other."""
    mismatched = sorted(k for k, v in REGISTRY.items()
                        if v.get('name') != k)
    check('registry key matches the schema name', not mismatched, mismatched)


def test_every_schema_description_is_non_empty():
    """tool_schema.to_anthropic_tools falls back to the bare tool NAME when the
    description is empty — the model then has to guess from the name alone."""
    blank = sorted(k for k, v in REGISTRY.items()
                   if not (v.get('description') or '').strip())
    check('every tool has a description', not blank, blank)


# ─────────────────────────────────────────────────────────────────────────────
# Enum coverage
# ─────────────────────────────────────────────────────────────────────────────

def _switch_props():
    """Yield (tool, field, schema_prop) for every dispatch-switch field."""
    for tool, spec in sorted(REGISTRY.items()):
        props = (spec.get('inputSchema') or {}).get('properties') or {}
        for field, prop in sorted(props.items()):
            if field in SWITCH_FIELDS and prop.get('type') == 'string':
                yield tool, field, prop


def test_every_switch_field_has_an_enum():
    """Without an enum the model can send anything, and the dispatch's else
    branch quietly does something else. This is the whole Gap-C bug class."""
    bare = ['{}.{}'.format(t, f) for t, f, p in _switch_props()
            if not p.get('enum')]
    check('every dispatch-switch field declares an enum', not bare, bare)


def test_every_enum_value_appears_in_its_dispatch():
    """An advertised enum value with no branch is a promise the server cannot
    keep — the model calls it and gets an unknown-operation error at best."""
    orphans = []
    for tool, field, prop in _switch_props():
        branch = _effective_branch_source(tool)
        for value in prop.get('enum') or []:
            if "'{}'".format(value) not in branch and '"{}"'.format(value) not in branch:
                orphans.append('{}.{}={}'.format(tool, field, value))
    check('every advertised enum value exists in the dispatch',
          not orphans, orphans)


def test_every_extension_file_parses():
    """A file that does not parse cannot be imported by anything, ever.

    lib/Snippets/_revisions.py carried `revision_type=RevisionNumberType.None`
    in a signature for a long time — `None` is a keyword, so `X.None` is a
    SyntaxError — and nothing noticed, because audit_tools.py only inspects
    pushbuttons and no test covered lib/Snippets at all.

    Parse only, never import: most of this tree needs a live Revit, so
    importing under CPython 3 proves nothing. Syntax is checkable everywhere.
    """
    broken = []
    for root, dirs, files in os.walk(EXT):
        dirs[:] = [d for d in dirs
                   if d not in ('_archive', 'scratch', '__pycache__')]
        for fname in files:
            if not fname.endswith('.py'):
                continue
            path = os.path.join(root, fname)
            try:
                ast.parse(_read(path))
            except SyntaxError as ex:
                broken.append('{}:{}'.format(
                    os.path.relpath(path, REPO), ex.lineno))
    check('every .py in the extension parses', not broken, broken)


def test_list_views_enum_matches_the_filter_map():
    """revit_list_views advertises an enum but normalises through
    _VIEW_TYPE_FILTER_MAP; the two are declared separately (the registry has to
    stay ast.literal_eval-able, so the enum cannot be sorted(map)). If they
    drift, the model is told about a filter value the dispatch cannot honour —
    which is the original bug in a new disguise."""
    keys = set(_dict_literal(_SERVER_SRC, '_VIEW_TYPE_FILTER_MAP'))
    advertised = set(
        REGISTRY['revit_list_views']['inputSchema']['properties']
        ['view_type'].get('enum') or [])
    check('every filter-map key is advertised', not (keys - advertised),
          sorted(keys - advertised))
    check('every advertised filter value is mapped', not (advertised - keys),
          sorted(advertised - keys))


def test_create_view_still_advertises_every_supported_type():
    """_effective_branch_source resolves create_view's enum through
    advanced_view_manager. That is correct but looser, so pin the count here:
    the schema must keep advertising every type SUPPORTED_VIEW_TYPES can build,
    and nothing else."""
    supported = set(_list_literal(
        _read(os.path.join(LIB, 'core', 'advanced_view_manager.py')),
        'SUPPORTED_VIEW_TYPES'))
    advertised = set(
        REGISTRY['create_view']['inputSchema']['properties']
        ['view_type'].get('enum') or [])
    check('create_view advertises what it can build',
          supported == advertised, sorted(supported ^ advertised))
    check('all nine view types are still there', len(advertised) == 9,
          sorted(advertised))


# ─────────────────────────────────────────────────────────────────────────────
# Threading / transactions
# ─────────────────────────────────────────────────────────────────────────────

def test_transactional_tools_are_write_tools():
    """A tool that opens a Transaction but isn't in _WRITE_TOOLS never reaches
    Revit's main thread: it throws "Starting a transaction ... outside of API
    context is not allowed" on every call."""
    write_tools = _class_attr_set(_SERVER_SRC, 'T3LabAIServer', '_WRITE_TOOLS')
    missing = []
    for name, src in sorted(BRANCHES.items()):
        if name in INTERNAL_TOOLS or name not in REGISTRY:
            continue
        if 'Transaction(doc' in src and name not in write_tools:
            missing.append(name)
    check('every transactional tool is in _WRITE_TOOLS', not missing, missing)


def test_write_tools_are_real_tools():
    write_tools = _class_attr_set(_SERVER_SRC, 'T3LabAIServer', '_WRITE_TOOLS')
    unknown = sorted(n for n in write_tools
                     if n not in REGISTRY and n not in INTERNAL_TOOLS)
    check('_WRITE_TOOLS names only real tools', not unknown, unknown)


def test_docless_tools_are_real_tools():
    docless = _class_attr_set(_SERVER_SRC, 'T3LabAIServer', '_DOCLESS_TOOLS')
    unknown = sorted(n for n in docless if n not in REGISTRY)
    check('_DOCLESS_TOOLS names only real tools', not unknown, unknown)


# ─────────────────────────────────────────────────────────────────────────────
# Version safety
# ─────────────────────────────────────────────────────────────────────────────

def test_bic_map_is_built_defensively():
    """_BIC_NAMES holds MEMBER NAME STRINGS resolved through getattr, never
    `B.OST_x` attribute access. BuiltInCategory members come and go between
    Revit releases (OST_Toposolid is 2024+); one missing member in a literal
    raises while BUILDING the map, which takes down all nine category-accepting
    tools at once on Revit 2023."""
    tree = ast.parse(_SERVER_SRC)
    bad = []
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Assign)
                and any(isinstance(t, ast.Name) and t.id == '_BIC_NAMES'
                        for t in node.targets)):
            continue
        for sub in ast.walk(node.value):
            if isinstance(sub, ast.Attribute) and sub.attr.startswith('OST_'):
                bad.append(sub.attr)
    check('_BIC_NAMES uses guarded member-name strings', not bad, bad)


def test_bic_map_resolves_through_getattr():
    src = _SERVER_SRC
    start = src.find('def _bic_map(self)')
    body = src[start:start + 1400]
    check('_bic_map resolves members with getattr',
          'getattr(B, member' in body, 'guard missing from _bic_map')


# ─────────────────────────────────────────────────────────────────────────────
# Tool subsets — the original pin bug
# ─────────────────────────────────────────────────────────────────────────────

def test_tool_subsets_reference_real_tools():
    """A typo here is INVISIBLE at runtime: tool_schema._registry_tools falls
    back to the full list when a subset filters to empty, and a specialist just
    silently loses the tool — which is how "pin the walls" became "paint them
    red"."""
    spec_src = _read(SPECIALISTS)
    schema_src = _read(TOOL_SCHEMA)
    bad = {}
    subsets = [(SPECIALISTS, spec_src, n) for n in
               ('READ_TOOLS', 'MODIFY_TOOLS', 'EXPORT_TOOLS', 'ACTION_TOOLS',
                'MULTIDOC_TOOLS', 'MODELING_TOOLS', 'QA_TOOLS')]
    subsets.append((TOOL_SCHEMA, schema_src, 'ESSENTIAL_TOOL_NAMES'))
    for path, src, name in subsets:
        try:
            names = _frozenset_literal(src, name)
        except AssertionError as ex:
            bad[name] = str(ex)
            continue
        unknown = sorted(n for n in names if n not in REGISTRY)
        if unknown:
            bad[name] = unknown
    check('every tool subset names only real tools', not bad, bad)


def test_multiop_tools_are_visible_to_the_action_specialist():
    """Any tool with an `operation` enum is a verb the user will ask for
    directly. If the ACTION specialist can't see it, it substitutes."""
    spec_src = _read(SPECIALISTS)
    action = _frozenset_literal(spec_src, 'ACTION_TOOLS')
    multiop = set()
    for tool, field, prop in _switch_props():
        if field == 'operation':
            multiop.add(tool)
    missing = sorted(multiop - action)
    check('action specialist sees every multi-operation tool',
          not missing, missing)


def test_cap_map_references_real_tools():
    """_CAP_MAP drives the user-facing "what can you do" answer."""
    nlu_src = _read(NLU)
    start = nlu_src.find('_CAP_MAP')
    end = nlu_src.find('\n_CAP_HIDE', start)
    if end == -1:
        end = start + 20000
    block = nlu_src[start:end]
    named = set(re.findall(r'u"([a-z0-9_]+)"', block))
    # Only judge names that look like tool names we actually know about.
    unknown = sorted(n for n in named
                     if n.count('_') >= 1 and n not in REGISTRY
                     and not n.startswith('open_'))
    check('_CAP_MAP names only real tools', not unknown, unknown)


# ─────────────────────────────────────────────────────────────────────────────
# Prompt rules
# ─────────────────────────────────────────────────────────────────────────────

def test_prompt_names_every_multiop_tool():
    """The step whose omission caused the pin incident: a tool exists, is
    registered, is in the subsets — and the prompt never tells the model which
    verbs map to it, so it reaches for whatever it recognises."""
    prompt = _read(AGENT_LOOP)
    multiop = sorted({t for t, f, p in _switch_props() if f == 'operation'})
    missing = [t for t in multiop if '`{}`'.format(t) not in prompt]
    check('agent prompt names every multi-operation tool', not missing, missing)


def test_destructive_names_are_real_tools():
    """The confirmation gate and the TransactionGroup exemption are both keyed
    on hardcoded tool names in the assistant script."""
    src = _read(ASSISTANT)
    start = src.find('_group_exempt = frozenset((')
    block = src[start:src.find('))', start)] if start != -1 else ''
    named = set(re.findall(r"'([a-z0-9_]+)'", block))
    unknown = sorted(n for n in named
                     if n not in REGISTRY and n not in INTERNAL_TOOLS)
    check('_group_exempt names only real tools', not unknown, unknown)


TESTS = [
    ('syntax', [
        test_every_extension_file_parses,
    ]),
    ('registry ↔ dispatch', [
        test_every_registry_tool_has_a_dispatch_branch,
        test_every_dispatch_branch_is_registered,
        test_no_duplicate_tool_names,
        test_every_schema_description_is_non_empty,
    ]),
    ('enum coverage', [
        test_every_switch_field_has_an_enum,
        test_every_enum_value_appears_in_its_dispatch,
        test_list_views_enum_matches_the_filter_map,
        test_create_view_still_advertises_every_supported_type,
    ]),
    ('transactions', [
        test_transactional_tools_are_write_tools,
        test_write_tools_are_real_tools,
        test_docless_tools_are_real_tools,
    ]),
    ('version safety', [
        test_bic_map_is_built_defensively,
        test_bic_map_resolves_through_getattr,
    ]),
    ('tool subsets', [
        test_tool_subsets_reference_real_tools,
        test_multiop_tools_are_visible_to_the_action_specialist,
        test_cap_map_references_real_tools,
    ]),
    ('prompt rules', [
        test_prompt_names_every_multiop_tool,
        test_destructive_names_are_real_tools,
    ]),
]


def main():
    print('registry: {} tools, {} dispatch branches'.format(
        len(REGISTRY), len(BRANCHES)))
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
