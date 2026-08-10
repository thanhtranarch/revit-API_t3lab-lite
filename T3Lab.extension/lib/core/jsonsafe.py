# -*- coding: utf-8 -*-
"""
ASCII-safe JSON serialization that also works on IronPython 2.7.

`json.dumps(obj, ensure_ascii=True)` is the pattern used across this codebase to
produce a pure-ASCII payload that can then be written through
`io.open(..., encoding='utf-8')` without a mid-write UnicodeEncodeError. On
CPython that works. On **IronPython 2.7 it raises** on any non-ASCII input:

    >>> json.dumps({'g': u'liệt kê'}, ensure_ascii=True)
    UnicodeEncodeError: 'ascii' codec can't encode character u'\\u1ec7'
        ... in json/encoder.py, py_encode_basestring_ascii

IronPython's `py_encode_basestring_ascii` does not escape the non-ASCII
characters before calling `str()` on the result, so the escaper itself is what
blows up — exactly the failure it was supposed to prevent. Every Vietnamese
string therefore fails to persist, silently, under Revit only (the CPython-3
unit tests pass).

`dumps()` below sidesteps the broken escaper: serialize with
`ensure_ascii=False` (that path is sound on both runtimes) and do the \\uXXXX
escaping here. Output is byte-identical to CPython's `ensure_ascii=True`,
including surrogate pairs for astral characters, and always `unicode`.

Measured IronPython 2.7 string semantics (probed inside Revit, 2026-08-10) —
they are NOT py2's, so do not reason from CPython 2 habits:

    bytes is str    -> False        # separate types, unlike CPython 2
    str is unicode  -> True         # every str is already unicode
    u'Café'.decode('utf-8')         -> UnicodeDecodeError (encodes first)
    json.dumps(x, ensure_ascii=True)-> UnicodeDecodeError on the same input

So a Revit name like "Café Kitchen - Section" is unicode already and must NOT
be decoded — `coerce()` only touches real `bytes` (a .NET byte array), where
UTF-8-then-latin-1 is the right ladder. Calling `.decode()` on anything else is
what produces the confusing "'unknown' codec can't decode byte 0xe9".

Pure Python, no imports beyond json/re — CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "JSON Safe"

import json
import re

_NON_ASCII = re.compile('[^\x00-\x7f]')


def _escape(match):
    """Replace one non-ASCII character with its \\uXXXX escape(s)."""
    n = ord(match.group(0))
    if n > 0xFFFF:  # astral plane -> UTF-16 surrogate pair, as CPython does
        n -= 0x10000
        return '\\u%04x\\u%04x' % (0xD800 + (n >> 10), 0xDC00 + (n & 0x3FF))
    return '\\u%04x' % n


def escape(text):
    """Return `text` with every non-ASCII character replaced by \\uXXXX."""
    if isinstance(text, bytes):
        text = text.decode('utf-8')
    return _NON_ASCII.sub(_escape, text)


def coerce(obj):
    """Recursively turn byte strings into unicode. Never raises.

    UTF-8 first, then latin-1 with replacement — a Windows code-page byte such
    as 0xE9 is not valid UTF-8 on its own, and losing the accent beats losing
    the whole payload.
    """
    if isinstance(obj, bytes):
        try:
            return obj.decode('utf-8')
        except Exception:
            return obj.decode('latin-1', 'replace')
    if isinstance(obj, dict):
        return dict((coerce(k), coerce(v)) for k, v in obj.items())
    if isinstance(obj, (list, tuple, set)):
        return [coerce(v) for v in obj]
    return obj


def dumps(obj, **kwargs):
    """`json.dumps(obj, ensure_ascii=True)` that survives IronPython 2.7.

    Accepts the same keyword arguments as `json.dumps` (`indent`, `sort_keys`,
    …); any `ensure_ascii` passed in is ignored — ASCII output is the point.
    Always returns `unicode`, ready for `io.open(..., encoding='utf-8')`.
    """
    kwargs['ensure_ascii'] = False
    return escape(json.dumps(coerce(obj), **kwargs))
