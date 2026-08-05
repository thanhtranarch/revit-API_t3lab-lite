# -*- coding: utf-8 -*-
"""
CPython 3 tests for Intelligence/learning/trainer.py (Phase 3 triggers).

Run:  python3 dev/test_trainer_trigger.py
Exit 0 = all pass. Plain asserts.

Covers the pure auto-train decision + the status readers. The actual detached
launch and GPU training are on the on-machine checklist.
"""
from __future__ import unicode_literals

import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_trainer_test_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))

from Intelligence.learning import trainer as T

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def test_should_auto_train():
    print('[trainer: should_auto_train predicate]')
    mn = T.MIN_NEW_FOR_AUTO
    check('off → never', T.should_auto_train(False, 10000, 0, False) is False)
    check('running → never',
          T.should_auto_train(True, 10000, 0, True) is False)
    check('not enough new → no',
          T.should_auto_train(True, 100 + mn - 1, 100, False) is False)
    check('enough new → yes',
          T.should_auto_train(True, 100 + mn, 100, False) is True)
    check('no prior train, enough total → yes',
          T.should_auto_train(True, mn, 0, False) is True)
    check('bad inputs → no',
          T.should_auto_train(True, None, None, False) is False)


def test_last_train_reader():
    print('[trainer: last_train reader]')
    check('missing file → empty dict', T.last_train() == {})
    d = T._training_dir()
    if not os.path.isdir(d):
        os.makedirs(d)
    import io as _io
    import json as _json
    with _io.open(os.path.join(d, 'last_train.json'), 'w',
                  encoding='utf-8') as f:
        f.write(_json.dumps({'trained_at': '2026-08-05', 'examples': 321}))
    lt = T.last_train()
    check('reads examples', lt.get('examples') == 321, lt)


def test_maybe_auto_train_off_by_default():
    print('[trainer: maybe_auto_train is off unless opted in]')
    res = T.maybe_auto_train()
    check('reports auto-train off (default)',
          res.get('status') == 'auto-train off', res)


def main():
    test_should_auto_train()
    test_last_train_reader()
    test_maybe_auto_train_off_by_default()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All trainer-trigger tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
