# -*- coding: utf-8 -*-
"""
validate_dataset — gate the self-study corpus before a fine-tune run.

The in-Revit self-study loop appends chat-SFT examples to
%APPDATA%/T3LabAI/training/dataset.jsonl (Intelligence/learning/dataset.py).
Before the GPU job (finetune_local.py) spends hours on it, this checks the
corpus is worth training on: enough valid, deduped examples and no malformed
lines. Pure Python (CPython 3), no ML deps — so it is unit-testable and runs in
a second as a pre-flight.

CLI:
    python3 tools/train/validate_dataset.py [path/to/dataset.jsonl]
Exit 0 = OK to train, 1 = not ready (too few / malformed), 2 = file missing.
"""
from __future__ import unicode_literals

import io
import json
import os
import sys

# Reuse the exact validation + hashing the writer uses, so "valid here" means
# "valid to the same rules that wrote it".
_LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), 'T3Lab.extension', 'lib')
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

try:
    from Intelligence.learning.dataset import _clean_messages, _hash, dataset_file
except Exception:                                    # pragma: no cover
    _clean_messages = _hash = dataset_file = None

# Below this a LoRA run just overfits noise — refuse rather than waste a GPU.
MIN_EXAMPLES = 30
# Very long examples blow the training context; warn (don't fail) past this.
WARN_CHARS = 8000


def validate_file(path):
    """Return a report dict. Never raises.

    {
      'ok': bool,                # enough clean, deduped examples to train
      'total_lines', 'valid', 'invalid', 'duplicates', 'unique',
      'long_examples', 'reasons': [str, ...]
    }
    """
    report = {'ok': False, 'total_lines': 0, 'valid': 0, 'invalid': 0,
              'duplicates': 0, 'unique': 0, 'long_examples': 0, 'reasons': []}
    if _clean_messages is None:
        report['reasons'].append('cannot import dataset validator from lib')
        return report
    if not path or not os.path.exists(path):
        report['reasons'].append('dataset file not found: {}'.format(path))
        return report

    seen = {}
    try:
        with io.open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                report['total_lines'] += 1
                try:
                    obj = json.loads(line)
                except Exception:
                    report['invalid'] += 1
                    continue
                msgs = _clean_messages(obj.get('messages'))
                if not msgs:
                    report['invalid'] += 1
                    continue
                report['valid'] += 1
                if sum(len(m['content']) for m in msgs) > WARN_CHARS:
                    report['long_examples'] += 1
                h = _hash(msgs)
                if h in seen:
                    report['duplicates'] += 1
                else:
                    seen[h] = True
    except Exception as ex:
        report['reasons'].append('read error: {}'.format(ex))
        return report

    report['unique'] = len(seen)
    if report['invalid']:
        report['reasons'].append(
            '{} malformed line(s) will be skipped'.format(report['invalid']))
    if report['long_examples']:
        report['reasons'].append(
            '{} example(s) exceed {} chars (may truncate)'.format(
                report['long_examples'], WARN_CHARS))
    if report['unique'] < MIN_EXAMPLES:
        report['reasons'].append(
            'only {} unique examples; need >= {} to train'.format(
                report['unique'], MIN_EXAMPLES))
        report['ok'] = False
    else:
        report['ok'] = True
    return report


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    path = argv[0] if argv else (dataset_file() if dataset_file else None)
    report = validate_file(path)
    print('Dataset: {}'.format(path))
    for k in ('total_lines', 'valid', 'invalid', 'duplicates', 'unique',
              'long_examples'):
        print('  {:14s} {}'.format(k, report[k]))
    for r in report['reasons']:
        print('  - {}'.format(r))
    if report['ok']:
        print('READY to train.')
        return 0
    if not os.path.exists(path or ''):
        return 2
    print('NOT ready to train.')
    return 1


if __name__ == '__main__':
    sys.exit(main())
