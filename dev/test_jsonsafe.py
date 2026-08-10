# -*- coding: utf-8 -*-
"""
CPython 3 tests for core/jsonsafe.py (IronPython-safe ASCII JSON).

Run:  python3 dev/test_jsonsafe.py
Exit 0 = all pass. Plain asserts, mirroring dev/test_learning_dataset.py.

These tests can only prove the OUTPUT contract — that the payload is pure ASCII
and round-trips — because the bug they guard against (IronPython 2.7's
py_encode_basestring_ascii raising on non-ASCII) does not reproduce on CPython.
That asymmetry is the whole point: `json.dumps(ensure_ascii=True)` passed here
for months while silently dropping every Vietnamese string inside Revit.
"""
from __future__ import unicode_literals

import json
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from core import jsonsafe as J

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def test_output_is_ascii_and_round_trips():
    print('[jsonsafe: ASCII out, lossless in]')
    cases = [
        {'g': 'liệt kê tất cả level trong model'},
        {'note': 'Đã đổi màu 42 cửa — xong'},          # em dash + Vietnamese
        {'emoji': 'ok 😀', 'math': '≥ 8 GB'},           # astral + symbol
        ['tường', {'level': 'Tầng 1'}, 12.5, None, True],
        {'ascii only': 'nothing to escape'},
    ]
    for i, obj in enumerate(cases):
        out = J.dumps(obj)
        check('case {} is pure ASCII'.format(i),
              all(ord(c) < 128 for c in out), repr(out[:80]))
        check('case {} round-trips'.format(i), json.loads(out) == obj, out)


def test_matches_cpython_ensure_ascii():
    print('[jsonsafe: byte-identical to CPython ensure_ascii=True]')
    for obj in ({'v': 'Cửa đi 😀 — ≥'}, ['a', 'ệ'], {'k': {'n': 'Tầng'}}):
        check('same as json.dumps(ensure_ascii=True)',
              J.dumps(obj) == json.dumps(obj, ensure_ascii=True), obj)
    check('kwargs pass through (indent)',
          J.dumps({'a': 'ệ'}, indent=1)
          == json.dumps({'a': 'ệ'}, ensure_ascii=True, indent=1))
    check('caller-supplied ensure_ascii=False is ignored',
          all(ord(c) < 128 for c in J.dumps({'a': 'ệ'}, ensure_ascii=False)))


def test_surrogate_pairs():
    print('[jsonsafe: astral chars become surrogate pairs]')
    out = J.dumps('😀')
    check('two \\u escapes for one astral char', out.count('\\u') == 2, out)
    check('astral round-trips', json.loads(out) == '😀', out)


def test_escape_helper():
    print('[jsonsafe: escape() on bare text]')
    check('escapes non-ASCII', J.escape('ệ') == '\\u1ec7', J.escape('ệ'))
    check('leaves ASCII alone', J.escape('plain text') == 'plain text')
    check('accepts utf-8 bytes', J.escape('ệ'.encode('utf-8')) == '\\u1ec7')
    check('returns unicode', isinstance(J.escape('x'), type('')))


def test_tool_manifest_shape():
    """The regression that hid the teaching tools: a tool description with an
    em dash killed the whole tools/list response inside Revit."""
    print('[jsonsafe: a tool manifest with an em dash serializes]')
    manifest = {'tools': [
        {'name': 't3lab_mark_sandbox',
         'description': 'Designate a document as the teaching sandbox — '
                        'the ONLY document model-modifying tools may touch.',
         'inputSchema': {'type': 'object',
                         'properties': {'document': {'type': 'string'}}}},
    ]}
    out = J.dumps(manifest)
    check('manifest is pure ASCII', all(ord(c) < 128 for c in out))
    check('manifest round-trips', json.loads(out) == manifest)


def test_code_page_bytes():
    """Real `bytes` (a .NET byte array) still need decoding before the encoder
    sees them. Note a Revit *name* is NOT this case — on IronPython 2.7
    `str is unicode`, so "Café Kitchen - Section" arrives already decoded and
    coerce() must leave it alone; decoding it is what raises."""
    print('[jsonsafe: real byte strings survive, unicode is left alone]')
    check('unicode text is not decoded', J.coerce('Café') == 'Café')
    check('latin-1 bytes decoded, not fatal',
          J.coerce(b'Caf\xe9') == 'Café', repr(J.coerce(b'Caf\xe9')))
    check('utf-8 bytes preferred',
          J.coerce('Tầng'.encode('utf-8')) == 'Tầng')

    payload = {'views': [{'name': b'Caf\xe9', 'id': 12}], b'k': b'v'}
    out = J.dumps(payload)
    check('mixed byte payload serializes', all(ord(c) < 128 for c in out), out)
    check('view name preserved',
          json.loads(out)['views'][0]['name'] == 'Café', out)
    check('byte keys decoded', json.loads(out)['k'] == 'v', out)

    check('nested containers coerced',
          J.coerce([{'a': (b'x', b'y')}]) == [{'a': ['x', 'y']}])
    check('non-string values untouched',
          J.coerce({'n': 1, 'f': 2.5, 'b': True, 'z': None})
          == {'n': 1, 'f': 2.5, 'b': True, 'z': None})


def main():
    test_output_is_ascii_and_round_trips()
    test_code_page_bytes()
    test_matches_cpython_ensure_ascii()
    test_surrogate_pairs()
    test_escape_helper()
    test_tool_manifest_shape()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All jsonsafe tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
