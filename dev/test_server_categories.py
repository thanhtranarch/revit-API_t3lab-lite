# -*- coding: utf-8 -*-
"""
CPython 3 test for the shared category resolver that get_material_quantities
now routes through (core/server.py _resolve_bic).

Run:  python3 dev/test_server_categories.py
Exit 0 = all pass. Plain asserts, mirroring dev/test_assistant_llm.py.

get_material_quantities used to keep its own case-sensitive 4-entry map, so
"walls", "Tường" and "tường" returned a false "Unsupported category". It now
calls _resolve_bic — the same case-insensitive + Vietnamese-alias resolver every
other category tool uses. The live _bic_map() needs the Revit API, so we drive
_resolve_bic with a stub self whose _bic_map returns the four canonical names;
this exercises the resolution contract the fix depends on. The quantity numbers
themselves are on the in-Revit checklist.
"""
from __future__ import unicode_literals

import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
sys.path.insert(0, LIB)

import core.server as S

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


class _StubServer(object):
    """Minimal host for the unbound _resolve_bic — no Revit API needed."""
    _BIC_ALIASES = S.T3LabAIServer._BIC_ALIASES

    def _bic_map(self):
        # Canonical name -> a stand-in for the BuiltInCategory value. The
        # resolver only ever compares/returns these opaquely, so strings are fine.
        return {'Walls': 'OST_Walls', 'Floors': 'OST_Floors',
                'Roofs': 'OST_Roofs', 'Ceilings': 'OST_Ceilings'}


def _resolve(cat):
    return S.T3LabAIServer._resolve_bic(_StubServer(), cat)


def test_case_insensitive():
    print('[category: case-insensitive]')
    for arg in ('Walls', 'walls', 'WALLS', ' Walls '):
        bic, name, err = _resolve(arg)
        check(u'"{}" -> Walls'.format(arg),
              err is None and bic == 'OST_Walls', (bic, err))


def test_vietnamese_aliases():
    print('[category: Vietnamese aliases]')
    cases = {'tường': 'OST_Walls', 'tuong': 'OST_Walls',
             'sàn': 'OST_Floors', 'san': 'OST_Floors',
             'mái': 'OST_Roofs', 'trần': 'OST_Ceilings'}
    for arg, want in cases.items():
        bic, name, err = _resolve(arg)
        check(u'"{}" -> {}'.format(arg, want),
              err is None and bic == want, (bic, err))


def test_unknown_returns_contract():
    print('[category: unknown -> did_you_mean contract]')
    bic, name, err = _resolve('not-a-category')
    check('unknown yields an error dict', bic is None and isinstance(err, dict),
          (bic, err))
    check('error carries did_you_mean',
          isinstance(err, dict) and 'did_you_mean' in err, err)


def main():
    test_case_insensitive()
    test_vietnamese_aliases()
    test_unknown_returns_contract()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All server-category resolver tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
