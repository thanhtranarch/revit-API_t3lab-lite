# -*- coding: utf-8 -*-
"""
CPython 3 drift locks for the element finder in lib/core/server.py.

Run:  python3 dev/test_element_finder.py
Exit code 0 = all pass. No Revit required.

WHY THIS FILE EXISTS
--------------------
The assistant was asked to move "every element inside the groups named _F1_"
onto a workset. It answered that no such filter existed, then fell back to
listing the whole model — because `ai_element_filter` really could only filter
by category + parameter, and because every id it found had to be ferried back
through the model's context to reach `set_element_workset`.

Two things fix that, and both are easy to break silently:

    the filter arguments     an argument the schema advertises but the dispatch
                             branch never reads is dropped on the floor, and the
                             tool still reports success (_reject_unknown_arguments
                             only catches the opposite direction)
    the "@name" handles      element_ids must stay schema-legal for a string, or
                             a model following the schema will never pass one —
                             and _expand_id_handles must refuse an unknown handle
                             rather than expand it to [], which for the tools
                             that fall back to the active selection or the whole
                             category means doing something else entirely

core.server can't be imported headlessly (it pulls the Revit API), so the
registry is read with ast.literal_eval and the handle methods are lifted out of
the class by AST and exec'd on their own — they touch nothing but dicts.
"""
from __future__ import unicode_literals

import ast
import io
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVER = os.path.join(REPO, 'T3Lab.extension', 'lib', 'core', 'server.py')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


_SRC = io.open(SERVER, encoding='utf-8').read()
_TREE = ast.parse(_SRC)
_CLASS = next(n for n in ast.walk(_TREE)
              if isinstance(n, ast.ClassDef) and n.name == 'T3LabAIServer')


def _method(name):
    for node in _CLASS.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError('T3LabAIServer has no method {}'.format(name))


def _class_constant(name):
    """Value of a plain class-level assignment (tuple/int literal)."""
    for node in _CLASS.body:
        if isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Name) and tgt.id == name:
                    return ast.literal_eval(node.value)
    raise AssertionError('T3LabAIServer has no constant {}'.format(name))


def _registry():
    """The _tools dict literal, read without importing the Revit API."""
    fn = _method('_register_default_tools')
    for node in fn.body:
        if isinstance(node, ast.Assign) and isinstance(node.value, ast.Dict):
            return ast.literal_eval(node.value)
    raise AssertionError('_register_default_tools has no dict literal')


REGISTRY = _registry()


# ─────────────────────────────────────────────────────────────────────────────
# 1 · The filter advertises every narrowing it can actually do
# ─────────────────────────────────────────────────────────────────────────────
def test_filter_schema():
    print('\nai_element_filter schema')
    schema = REGISTRY['ai_element_filter']['inputSchema']
    props = schema['properties']
    for arg in ('category', 'categories', 'group_name', 'workset',
                'level_name', 'type_name', 'name_contains',
                'parameter_name', 'parameter_value',
                'group_by', 'fields', 'store_as', 'limit', 'offset'):
        check('declares {}'.format(arg), arg in props, sorted(props))

    # No required argument: the filter is reachable by group/workset/level
    # alone, which is the whole point — "elements in the groups named _F1_"
    # names no category at all.
    check('nothing is required', schema.get('required') == [],
          schema.get('required'))

    # A narrowing the schema advertises but the code never reads is silently
    # dropped, and the tool reports success anyway.
    src = ast.get_source_segment(_SRC, _method('_find_elements')) or ''
    branch = _SRC.split("elif tool_name == 'ai_element_filter':", 1)[1][:9000]
    for arg in props:
        seen = ("'{}'".format(arg) in src) or ("'{}'".format(arg) in branch)
        check('{} is actually read'.format(arg), seen)


def test_enums_match_constants():
    print('\nenums track the constants')
    props = REGISTRY['ai_element_filter']['inputSchema']['properties']
    fields_enum = set(props['fields']['items']['enum'])
    group_enum = set(props['group_by']['items']['enum'])
    check('fields enum == _FILTER_FIELDS',
          fields_enum == set(_class_constant('_FILTER_FIELDS')),
          sorted(fields_enum ^ set(_class_constant('_FILTER_FIELDS'))))
    check('group_by enum == _GROUP_BY_KEYS',
          group_enum == set(_class_constant('_GROUP_BY_KEYS')),
          sorted(group_enum ^ set(_class_constant('_GROUP_BY_KEYS'))))
    defaults = set(_class_constant('_FILTER_DEFAULT_FIELDS'))
    check('default fields are all real fields',
          defaults <= set(_class_constant('_FILTER_FIELDS')), sorted(defaults))


# ─────────────────────────────────────────────────────────────────────────────
# 2 · "@name" handles stay schema-legal everywhere they are expanded
# ─────────────────────────────────────────────────────────────────────────────
def test_id_arguments_accept_a_handle():
    print('\nid arguments accept "@name"')
    handled = set(_class_constant('_HANDLE_ARG_KEYS')) - {'element_id'}
    offenders = []
    for tool, spec in sorted(REGISTRY.items()):
        for arg, prop in (spec.get('inputSchema', {})
                          .get('properties', {}).items()):
            if arg not in handled or not isinstance(prop, dict):
                continue
            item_type = (prop.get('items') or {}).get('type')
            if item_type == 'integer':
                offenders.append('{}.{}'.format(tool, arg))
    check('every expanded id argument allows a string item',
          not offenders, offenders)


def test_expansion_semantics():
    print('\n_expand_id_handles')
    ns = {}
    body = ['class Srv(object):']
    for const in ('_SELSET_MAX', '_HANDLE_ARG_KEYS'):
        body.append('    {} = {!r}'.format(const, _class_constant(const)))
    for name in ('_selset_key', '_selset_bucket', '_selset_store',
                 '_selset_get', '_expand_id_handles'):
        # get_source_segment drops the def's own indentation but keeps the
        # body's, so only the first line needs putting back inside the class.
        body.append('    ' + ast.get_source_segment(_SRC, _method(name)))
    exec(compile('\n'.join(body), '<finder>', 'exec'), ns)

    class Doc(object):
        Title = 'demo.rvt'

    srv, doc = ns['Srv'](), Doc()
    srv._selset_store(doc, 'sel', [11, 22, 33])

    args = {'element_ids': ['@sel']}
    check('unknown-free expansion returns None',
          srv._expand_id_handles(doc, 'set_element_workset', args) is None)
    check('handle expands to the stored ids', args['element_ids'] == [11, 22, 33],
          args)

    args = {'element_ids': [7, '@sel']}
    srv._expand_id_handles(doc, 'select_elements', args)
    check('literal ids and a handle mix', args['element_ids'] == [7, 11, 22, 33],
          args)

    # The failure that matters: a typo must NOT become an empty list. Several
    # tools read an empty element_ids as "use the active selection" or "use the
    # whole category", so silently emptying it means doing something else.
    args = {'element_ids': ['@typo']}
    err = srv._expand_id_handles(doc, 'delete_element', args)
    check('unknown handle is an error, not []', isinstance(err, dict)
          and 'error' in err, err)
    check('the error names the sets that DO exist',
          err.get('saved_sets') == ['sel'], err)

    # Plain calls are untouched — expansion runs for every tool, so it must be
    # a no-op when nobody passed a handle.
    args = {'element_ids': [1, 2], 'category': 'Walls'}
    check('plain ids pass through',
          srv._expand_id_handles(doc, 'select_elements', args) is None
          and args['element_ids'] == [1, 2], args)

    # A singular element_id can take a handle only when it resolved to one id.
    srv._selset_store(doc, 'one', [99])
    args = {'element_id': '@one'}
    srv._expand_id_handles(doc, 'revit_get_element_info', args)
    check('singular element_id takes a one-element set',
          args['element_id'] == 99, args)
    args = {'element_id': '@sel'}
    err = srv._expand_id_handles(doc, 'revit_get_element_info', args)
    check('singular element_id refuses a multi-element set',
          isinstance(err, dict) and 'error' in err, err)

    # Sets are per document: another model must not inherit them.
    class Other(object):
        Title = 'other.rvt'
    check('sets do not leak across documents',
          srv._selset_get(Other(), 'sel') is None)

    # Oldest sets are evicted rather than growing without bound.
    for i in range(int(_class_constant('_SELSET_MAX')) + 5):
        srv._selset_store(doc, 'set{}'.format(i), [i])
    check('bucket is capped at _SELSET_MAX',
          len(srv._selset_bucket(doc)) == int(_class_constant('_SELSET_MAX')),
          len(srv._selset_bucket(doc)))


# ─────────────────────────────────────────────────────────────────────────────
# 3 · Expansion is central, and happens before any branch reads the ids
# ─────────────────────────────────────────────────────────────────────────────
def test_expansion_is_wired_once():
    print('\nexpansion wiring')
    disp = ast.get_source_segment(_SRC, _method('_execute_tool_in_context'))
    check('dispatch calls _expand_id_handles',
          '_expand_id_handles(' in disp)
    head = disp.split("elif tool_name == 'ai_element_filter':", 1)[0]
    check('it runs before the tool branches', '_expand_id_handles(' in head)
    check('and before the teaching write-guard',
          disp.index('_expand_id_handles(') < disp.index('_teaching_enabled'))


def main():
    test_filter_schema()
    test_enums_match_constants()
    test_id_arguments_accept_a_handle()
    test_expansion_semantics()
    test_expansion_is_wired_once()
    print('')
    if FAILURES:
        print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('all passed')
    return 0


if __name__ == '__main__':
    sys.exit(main())
