# -*- coding: utf-8 -*-
"""
CPython 3 tests for the out-of-Revit training pre-flight (tools/train/).

Run:  python3 dev/test_validate_dataset.py
Exit 0 = all pass. Plain asserts.

Covers validate_dataset.validate_file (the GPU-job gate) and finetune_local's
non-training logic (render_modelfile + the --dry-run flow), so the pipeline's
CPU-only parts are locked down even though a real LoRA run needs a GPU.
"""
from __future__ import unicode_literals

import io
import json
import os
import sys
import tempfile

os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_train_test_')

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'T3Lab.extension', 'lib'))
sys.path.insert(0, os.path.join(REPO, 'tools', 'train'))

import validate_dataset as V
import finetune_local as F

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def _write_dataset(rows):
    path = os.path.join(tempfile.mkdtemp(), 'dataset.jsonl')
    with io.open(path, 'w', encoding='utf-8') as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=True) + u'\n')
    return path


def _ex(u, a):
    return {'messages': [{'role': 'user', 'content': u},
                         {'role': 'assistant', 'content': a}]}


def test_validate_too_few():
    print('[validate: below the minimum is not ready]')
    path = _write_dataset([_ex(u'question {}'.format(i),
                               u'answer number {}'.format(i))
                           for i in range(5)])
    rep = V.validate_file(path)
    check('not ok with 5 examples', rep['ok'] is False, rep)
    check('unique counted', rep['unique'] == 5, rep)


def test_validate_ready_and_dedupe():
    print('[validate: enough unique examples is ready; dupes counted]')
    rows = [_ex(u'question {}'.format(i), u'answer {}'.format(i))
            for i in range(V.MIN_EXAMPLES)]
    rows.append(_ex(u'question 0', u'answer 0'))     # a duplicate
    rows.append({'messages': [{'role': 'user', 'content': 'x'}]})  # malformed
    path = _write_dataset(rows)
    rep = V.validate_file(path)
    check('ready with >= MIN unique', rep['ok'] is True, rep)
    check('duplicate detected', rep['duplicates'] == 1, rep)
    check('malformed counted invalid', rep['invalid'] == 1, rep)
    check('unique == MIN_EXAMPLES', rep['unique'] == V.MIN_EXAMPLES, rep)


def test_validate_missing_file():
    print('[validate: missing file]')
    rep = V.validate_file('/no/such/dataset.jsonl')
    check('not ok', rep['ok'] is False)
    check('reason mentions not found',
          any('not found' in r for r in rep['reasons']), rep['reasons'])


def test_render_modelfile():
    print('[finetune: Modelfile rendering]')
    mf = F.render_modelfile('llama3.1:8b', system='You are T3Lab.')
    check('FROM line present', mf.startswith('FROM llama3.1:8b'), mf.split(chr(10))[0])
    check('system baked in', 'You are T3Lab.' in mf, mf)
    check('no leftover placeholders', '{base}' not in mf and '{system}' not in mf, mf)
    mf2 = F.render_modelfile('base.gguf', adapter_ref='adapter.gguf')
    check('adapter line present when given', 'ADAPTER adapter.gguf' in mf2, mf2)


def test_dry_run_flow():
    print('[finetune: --dry-run validates + writes Modelfile, no training]')
    rows = [_ex(u'question {}'.format(i), u'answer text {}'.format(i))
            for i in range(V.MIN_EXAMPLES)]
    path = _write_dataset(rows)
    out = tempfile.mkdtemp()
    rc = F.main(['--dry-run', '--dataset', path, '--out', out,
                 '--base', 'llama3.1:8b'])
    check('dry run exits 0', rc == 0, rc)
    check('Modelfile written', os.path.exists(os.path.join(out, 'Modelfile')))
    check('clean sft exported', os.path.exists(os.path.join(out, 'sft.jsonl')))
    n = sum(1 for _ in io.open(os.path.join(out, 'sft.jsonl'), encoding='utf-8'))
    check('sft has all examples', n == V.MIN_EXAMPLES, n)


def main():
    test_validate_too_few()
    test_validate_ready_and_dedupe()
    test_validate_missing_file()
    test_render_modelfile()
    test_dry_run_flow()

    print('')
    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All train-pipeline tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
