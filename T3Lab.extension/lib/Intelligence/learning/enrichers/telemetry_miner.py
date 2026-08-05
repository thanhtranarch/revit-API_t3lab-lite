# -*- coding: utf-8 -*-
"""
telemetry_miner — turn the assistant's own usage record into training data.

Raw telemetry (Intelligence/telemetry.py) stores timings/tokens but NO message
content, so the mineable content lives in the two stores that DO persist real
phrasings:

  * learned_patterns.json — {normkey: {intent, params, hits, last_raw,
    last_message}}: the office's actual commands that routed successfully. Each
    becomes an SFT example (user = last_raw, assistant = last_message).
  * assistant_feedback.json negatives — phrasings the user 👎'd. Feedback already
    suppresses these at route time; here we also PRUNE the matching learned
    pattern so a rejected mapping stops being replayed AND never leaks into the
    training set.

This enricher is zero-LLM and zero-cost — it only reshapes data already on disk.

Pure helpers (sft_examples_from_patterns, prunable_keys) take plain dicts so they
are CPython-3 unit-testable; run() wires them to the real stores + dataset.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Telemetry Miner"

SOURCE = 'telemetry_miner'

# A pattern seen this many times is a settled convention worth labelling 'high'.
_HIGH_HITS = 3


# ─── Pure helpers ───────────────────────────────────────────────────────────────

def sft_examples_from_patterns(patterns, min_hits=1):
    """[(user, assistant, quality)] from a learned-patterns dict.

    Only patterns that carry a real assistant `last_message` yield an example —
    without a target there is nothing to learn. Terse routing acks are still
    useful supervision (they teach the office's command → response mapping).
    """
    out = []
    for _key, data in sorted((patterns or {}).items()):
        if not isinstance(data, dict):
            continue
        if int(data.get('hits', 0)) < min_hits:
            continue
        user = u'{}'.format(data.get('last_raw') or u'').strip()
        answer = u'{}'.format(data.get('last_message') or u'').strip()
        if not user or not answer:
            continue
        quality = 'high' if int(data.get('hits', 0)) >= _HIGH_HITS else 'ok'
        out.append((user, answer, quality))
    return out


def prunable_keys(patterns, negatives):
    """Keys in `patterns` whose (phrasing, intent) the user rejected with 👎.

    `negatives` is the feedback store's negative map: {normkey: {'intents':
    [...]}}. Patterns and feedback share the same normalized-key space
    (both via t3lab_assistant._normalize_key), so a direct key match is exact.
    """
    prune = set()
    negatives = negatives or {}
    for key, data in (patterns or {}).items():
        if not isinstance(data, dict):
            continue
        entry = negatives.get(key)
        if not entry:
            continue
        rejected = entry.get('intents') or []
        if data.get('intent') in rejected:
            prune.add(key)
    return prune


# ─── Orchestration (injectable for tests) ───────────────────────────────────────

def run(add_example=None, load_patterns=None, save_patterns=None,
        load_feedback=None, revit_version=None, max_new=200):
    """Mine learned patterns → dataset; prune 👎'd routes. Returns {status, added}.

    All I/O is injected (defaults wired to the real modules) so tests drive it
    with plain dicts. Never raises.
    """
    try:
        if add_example is None:
            from Intelligence.learning import dataset as _ds
            add_example = _ds.add_example
        if load_patterns is None or save_patterns is None:
            from Intelligence import t3lab_assistant as _ta
            load_patterns = load_patterns or _ta.load_learned_patterns
            save_patterns = save_patterns or _ta.save_learned_patterns
        if load_feedback is None:
            from Intelligence import feedback as _fb
            load_feedback = lambda: (_fb.load() or {}).get('negative') or {}
    except Exception as ex:
        return {'status': u'unavailable: {}'.format(ex), 'added': 0}

    try:
        patterns = load_patterns() or {}
        negatives = load_feedback() or {}
    except Exception as ex:
        return {'status': u'load failed: {}'.format(ex), 'added': 0}

    # 1) Prune rejected routes so they neither replay nor train.
    pruned = 0
    try:
        prune = prunable_keys(patterns, negatives)
        if prune:
            patterns = dict((k, v) for k, v in patterns.items()
                            if k not in prune)
            save_patterns(patterns)
            pruned = len(prune)
    except Exception:
        pass

    # 2) Distill the surviving patterns into SFT examples.
    added = 0
    try:
        for user, answer, quality in sft_examples_from_patterns(patterns):
            if added >= max_new:
                break
            ok, note = add_example(
                [{'role': 'user', 'content': user},
                 {'role': 'assistant', 'content': answer}],
                source=SOURCE, quality=quality, revit_version=revit_version)
            if ok and note == u'added':
                added += 1
    except Exception as ex:
        return {'status': u'partial: {}'.format(ex),
                'added': added, 'pruned': pruned}

    return {'status': u'ok', 'added': added, 'pruned': pruned}
