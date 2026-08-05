# -*- coding: utf-8 -*-
"""
api_facts — turn the cached Revit-API compatibility facts into training data.

Reuses Intelligence.api_learner.RevitAPILearner, which already caches
version-specific API facts (export signatures, option availability) under
~/.t3lab/api_cache/. Unlike model counts, these facts are STABLE for a given
Revit version, so they are legitimate SFT targets — the model learns the correct
API shape for the office's Revit release.

Pure helper (api_facts_to_sft) is CPython-3 testable with a plain dict.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "API Facts"

SOURCE = 'api_facts'


def api_facts_to_sft(api_info, revit_version=None):
    """[(user, assistant)] Q&A from a RevitAPILearner api_info dict.

    Walks the known compatibility sections and emits one grounded fact each.
    Only entries carrying a human-readable signature/note become examples, so a
    sparse cache simply yields fewer rows rather than junk.
    """
    out = []
    if not isinstance(api_info, dict):
        return out
    ver = revit_version or api_info.get('revit_version') or u''
    vtag = u' (Revit {})'.format(ver) if ver else u''

    export = api_info.get('export_api') or {}
    for kind, spec in sorted(export.items()):
        if not isinstance(spec, dict):
            continue
        sig = spec.get('signature')
        note = spec.get('note')
        if sig:
            out.append((
                u'What is the Revit API signature for {}{}?'.format(kind, vtag),
                u'{}{}'.format(sig, u' — {}'.format(note) if note else u'')))
        elif note:
            out.append((u'How does {} work in the Revit API{}?'.format(kind, vtag),
                        note))

    opts = api_info.get('dwg_export_options') or {}
    for prop, spec in sorted(opts.items()):
        if isinstance(spec, dict) and spec.get('type'):
            avail = spec.get('available')
            detail = u'Type {}.'.format(spec['type'])
            if avail is not None:
                detail += u' Available: {}.'.format(bool(avail))
            vals = spec.get('supported_values')
            if vals:
                detail += u' Values: {}.'.format(u', '.join(map(u'{}'.format,
                                                                vals)))
            out.append((u'DWGExportOptions.{}{} — what should I know?'.format(
                prop, vtag), detail))
    return out


def run(add_example=None, revit_version=None, max_new=60):
    """Refresh the API cache and emit stable API-fact SFT rows. {status, added}."""
    if add_example is None:
        try:
            from Intelligence.learning import dataset as _ds
            add_example = _ds.add_example
        except Exception as ex:
            return {'status': u'dataset unavailable: {}'.format(ex), 'added': 0}
    try:
        from Intelligence.api_learner import RevitAPILearner
    except Exception as ex:
        return {'status': u'api_learner unavailable: {}'.format(ex), 'added': 0}

    ver = revit_version or 2023
    try:
        learner = RevitAPILearner(ver)
        api_info = getattr(learner, 'api_info', None) or {}
    except Exception as ex:
        return {'status': u'learner init failed: {}'.format(ex), 'added': 0}

    added = 0
    try:
        for user, answer in api_facts_to_sft(api_info, ver):
            if added >= max_new:
                break
            ok, note = add_example(
                [{'role': 'user', 'content': user},
                 {'role': 'assistant', 'content': answer}],
                source=SOURCE, quality='high', revit_version=ver)
            if ok and note == u'added':
                added += 1
    except Exception as ex:
        return {'status': u'partial: {}'.format(ex), 'added': added}
    return {'status': 'ok', 'added': added}
