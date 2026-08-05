# -*- coding: utf-8 -*-
"""
finetune_local — LoRA fine-tune a LOCAL model on the self-study dataset.

This is the out-of-Revit half of the self-improving assistant. It runs in normal
CPython 3 (NOT IronPython) on a machine with a GPU, reads the chat-SFT corpus the
in-Revit self-study loop curated, LoRA-fine-tunes a base model, and produces a
GGUF + an Ollama Modelfile so the result is served as `t3lab-assistant:<date>` —
which the assistant's OllamaProvider can then point at.

Heavy ML deps (unsloth/transformers/torch) are imported LAZILY and only for a
real run, so `--dry-run` works anywhere: it validates the dataset, exports a
clean SFT file, and writes the Modelfile WITHOUT training. That keeps this file
importable and its non-training logic testable in CI.

Usage:
    python3 tools/train/finetune_local.py --dry-run
    python3 tools/train/finetune_local.py --base unsloth/llama-3.1-8b-instruct-bnb-4bit \
        --out ./t3lab-out --tag t3lab-assistant

See tools/train/README.md for the full recipe (base-model choice, hardware,
`ollama create`, and pointing the assistant at the new model).
"""
from __future__ import unicode_literals

import argparse
import os
import sys
import time

_LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))), 'T3Lab.extension', 'lib')
if _LIB not in sys.path:
    sys.path.insert(0, _LIB)

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TAG = 't3lab-assistant'


def _default_dataset():
    try:
        from Intelligence.learning.dataset import dataset_file
        return dataset_file()
    except Exception:
        base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
        return os.path.join(base, 'T3LabAI', 'training', 'dataset.jsonl')


def export_clean(dataset_path, out_path):
    """Re-validate + dedupe the corpus into a training-ready SFT file.

    Returns the number of examples written. Reuses the writer's own export so
    the training set is exactly what dataset.export_sft() guarantees.
    """
    from Intelligence.learning import dataset as ds
    # Point the exporter at the requested file by overriding its path resolver
    # only if a non-default dataset was passed.
    if dataset_path and dataset_path != ds.dataset_file():
        _orig = ds.dataset_file
        ds.dataset_file = lambda: dataset_path
        try:
            return ds.export_sft(out_path)
        finally:
            ds.dataset_file = _orig
    return ds.export_sft(out_path)


def render_modelfile(base_ref, adapter_ref=None, system=None):
    """Produce Modelfile text from the template. Pure — unit-testable."""
    tmpl_path = os.path.join(_HERE, 'Modelfile.tmpl')
    try:
        with open(tmpl_path, 'r') as f:
            tmpl = f.read()
    except Exception:
        tmpl = ('FROM {base}\n{adapter}'
                'SYSTEM """{system}"""\n'
                'PARAMETER temperature 0.3\n')
    system = system or ('You are T3Lab Assistant, fine-tuned on this office\'s '
                        'Revit conventions, tools, and standards.')
    adapter = 'ADAPTER {}\n'.format(adapter_ref) if adapter_ref else ''
    return (tmpl.replace('{base}', base_ref)
                .replace('{adapter}', adapter)
                .replace('{system}', system))


def write_modelfile(out_dir, base_ref, adapter_ref=None, system=None):
    path = os.path.join(out_dir, 'Modelfile')
    with open(path, 'w') as f:
        f.write(render_modelfile(base_ref, adapter_ref, system))
    return path


def _validate_or_die(dataset_path):
    from validate_dataset import validate_file           # sibling module
    report = validate_file(dataset_path)
    for r in report['reasons']:
        print('  - {}'.format(r))
    print('  unique examples: {}'.format(report['unique']))
    return report


def run(args):
    dataset_path = args.dataset or _default_dataset()
    out_dir = args.out or os.path.join(_HERE, 'out')
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)

    print('== T3Lab local fine-tune ==')
    print('dataset:', dataset_path)
    report = _validate_or_die(dataset_path)
    if not report['ok'] and not args.force:
        print('Dataset not ready (use --force to override). Aborting.')
        return 1

    sft_path = os.path.join(out_dir, 'sft.jsonl')
    n = export_clean(dataset_path, sft_path)
    print('exported {} clean examples -> {}'.format(n, sft_path))

    if args.dry_run:
        mf = write_modelfile(out_dir, args.base, system=args.system)
        print('DRY RUN — wrote Modelfile ->', mf)
        print('No training performed. Re-run without --dry-run on a GPU box.')
        return 0

    # ── Real training path — heavy deps imported only now ────────────────────
    try:
        from unsloth import FastLanguageModel     # noqa: F401
        import torch                               # noqa: F401
    except Exception as ex:
        print('Training deps missing ({}). Install unsloth+torch, or use '
              '--dry-run. See tools/train/README.md.'.format(ex))
        return 3

    print('Loading base model:', args.base)
    from unsloth import FastLanguageModel
    from datasets import load_dataset
    from trl import SFTTrainer
    from transformers import TrainingArguments

    model, tokenizer = FastLanguageModel.from_pretrained(
        model_name=args.base, max_seq_length=args.max_seq_len,
        load_in_4bit=True)
    model = FastLanguageModel.get_peft_model(
        model, r=args.lora_r, lora_alpha=args.lora_r * 2,
        target_modules=['q_proj', 'k_proj', 'v_proj', 'o_proj',
                        'gate_proj', 'up_proj', 'down_proj'])

    ds = load_dataset('json', data_files=sft_path, split='train')

    def _fmt(ex):
        return {'text': tokenizer.apply_chat_template(
            ex['messages'], tokenize=False, add_generation_prompt=False)}
    ds = ds.map(_fmt)

    trainer = SFTTrainer(
        model=model, tokenizer=tokenizer, train_dataset=ds,
        dataset_text_field='text', max_seq_length=args.max_seq_len,
        args=TrainingArguments(
            per_device_train_batch_size=args.batch,
            gradient_accumulation_steps=4,
            num_train_epochs=args.epochs, learning_rate=2e-4,
            warmup_steps=5, logging_steps=10, output_dir=out_dir,
            optim='adamw_8bit'))
    trainer.train()

    gguf_dir = os.path.join(out_dir, 'gguf')
    print('exporting merged GGUF ->', gguf_dir)
    model.save_pretrained_gguf(gguf_dir, tokenizer,
                               quantization_method='q4_k_m')
    write_modelfile(out_dir, os.path.join(gguf_dir, 'unsloth.Q4_K_M.gguf'),
                    system=args.system)
    _write_result(out_dir, dataset_path, n, args.tag)
    print('Done. Next: ollama create {} -f {}/Modelfile'.format(
        args.tag, out_dir))
    return 0


def _write_result(out_dir, dataset_path, n, tag):
    import json
    result = {'trained_at': time.strftime('%Y-%m-%d %H:%M'),
              'examples': n, 'dataset': dataset_path, 'tag': tag}
    try:
        base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
        d = os.path.join(base, 'T3LabAI', 'training')
        if not os.path.isdir(d):
            os.makedirs(d)
        with open(os.path.join(d, 'last_train.json'), 'w') as f:
            json.dump(result, f, indent=2)
    except Exception:
        pass


def build_parser():
    p = argparse.ArgumentParser(description='LoRA fine-tune a local model on '
                                            'the T3Lab self-study dataset.')
    p.add_argument('--dataset', help='dataset.jsonl (default: %%APPDATA%%/...)')
    p.add_argument('--out', help='output dir (default: tools/train/out)')
    p.add_argument('--base', default='unsloth/llama-3.1-8b-instruct-bnb-4bit',
                   help='HF base model matching your Ollama base')
    p.add_argument('--tag', default=DEFAULT_TAG, help='Ollama model tag')
    p.add_argument('--system', help='SYSTEM prompt baked into the Modelfile')
    p.add_argument('--epochs', type=int, default=1)
    p.add_argument('--batch', type=int, default=2)
    p.add_argument('--lora-r', type=int, default=16, dest='lora_r')
    p.add_argument('--max-seq-len', type=int, default=2048, dest='max_seq_len')
    p.add_argument('--dry-run', action='store_true',
                   help='validate + write Modelfile, do NOT train')
    p.add_argument('--force', action='store_true',
                   help='train even if the dataset is below the minimum')
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    return run(args)


if __name__ == '__main__':
    sys.exit(main())
