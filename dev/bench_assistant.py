#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aggregate T3Lab Assistant turn telemetry into a before/after report.

The Assistant writes one JSONL line per turn (see
T3Lab.extension/lib/Intelligence/telemetry.py). This reads those files and
prints the numbers that decide whether a latency change actually worked:

    TTFT            time to first streamed token — what the user feels
    total           whole turn, including tool work
    roundtrips      crossings to Revit's main thread per turn (P2 batching)
    cache hit       share of prompt input tokens served from cache (P1)

Run:
    python3 dev/bench_assistant.py                      # today, default dir
    python3 dev/bench_assistant.py --days 7
    python3 dev/bench_assistant.py --dir /path/to/telemetry
    python3 dev/bench_assistant.py --since 2026-08-01 --label after

Compare two runs by capturing the output before and after a change; the
`--label` is echoed in the header so the two reports are easy to tell apart.
"""
from __future__ import print_function

import argparse
import os
import sys
import time

sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'T3Lab.extension', 'lib'))

from Intelligence import telemetry   # noqa: E402


def _percentile(values, pct):
    """Nearest-rank percentile of an already-sorted list."""
    if not values:
        return None
    k = int(round((pct / 100.0) * (len(values) - 1)))
    return values[max(0, min(k, len(values) - 1))]


def _fmt_ms(v):
    if v is None:
        return '   n/a'
    if v >= 10000:
        return '{:>5.1f}s'.format(v / 1000.0)
    return '{:>5d}ms'.format(int(v))


def _collect(directory, days, since):
    """Return every turn dict from the JSONL files in range."""
    if not os.path.isdir(directory):
        return [], []
    cutoff = None
    if since:
        try:
            cutoff = time.mktime(time.strptime(since, '%Y-%m-%d'))
        except ValueError:
            print('bad --since (want YYYY-MM-DD): {}'.format(since))
            return [], []
    elif days:
        cutoff = time.time() - days * 86400.0

    turns, files = [], []
    for name in sorted(os.listdir(directory)):
        if not name.endswith('.jsonl'):
            continue
        path = os.path.join(directory, name)
        rows = telemetry.read_turns(path)
        if cutoff is not None:
            rows = [r for r in rows if (r.get('ts') or 0) >= cutoff]
        if rows:
            files.append(name)
            turns.extend(rows)
    return turns, files


def _stat_block(title, turns):
    print('')
    print('  {}  (n={})'.format(title, len(turns)))
    print('  ' + '-' * 58)
    if not turns:
        print('    no turns')
        return

    for label, key in (('TTFT ', 'ttft_ms'), ('total', 'total_ms')):
        vals = sorted(v for v in (t.get(key) for t in turns)
                      if isinstance(v, (int, float)))
        if not vals:
            print('    {}   no measurements'.format(label))
            continue
        print('    {}   p50 {}   p95 {}   max {}'.format(
            label,
            _fmt_ms(_percentile(vals, 50)),
            _fmt_ms(_percentile(vals, 95)),
            _fmt_ms(vals[-1])))

    tool_turns = [t for t in turns if (t.get('tool_runs') or 0) > 0]
    if tool_turns:
        runs  = sum(t.get('tool_runs') or 0 for t in tool_turns)
        trips = sum(t.get('roundtrips') or 0 for t in tool_turns)
        saved = runs - trips
        print('    tools   {} calls over {} Revit round-trips '
              '({} saved by batching)'.format(runs, trips, saved))
        print('            {:.2f} calls/round-trip, {:.2f} calls/turn'.format(
            (float(runs) / trips) if trips else 0.0,
            float(runs) / len(tool_turns)))

    read  = sum(t.get('cache_read') or 0 for t in turns)
    fresh = sum(t.get('in_tok') or 0 for t in turns)
    write = sum(t.get('cache_write') or 0 for t in turns)
    out   = sum(t.get('out_tok') or 0 for t in turns)
    if read or fresh:
        print('    cache   {:.1%} of {} prompt tokens served from cache'.format(
            float(read) / float(read + fresh), read + fresh))
        print('            {} written to cache, {} output tokens'.format(
            write, out))
    else:
        print('    cache   no token usage reported '
              '(provider does not expose it)')


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--dir', default=None,
                    help='telemetry directory (default: %%APPDATA%%/T3LabAI/telemetry)')
    ap.add_argument('--days', type=int, default=1,
                    help='look back this many days (default 1)')
    ap.add_argument('--since', default=None,
                    help='only turns on/after this date (YYYY-MM-DD); overrides --days')
    ap.add_argument('--label', default='', help='tag for the report header')
    ap.add_argument('--by-provider', action='store_true',
                    help='break the report down per provider')
    args = ap.parse_args()

    directory = args.dir or telemetry.telemetry_dir()
    turns, files = _collect(directory, args.days, args.since)

    header = 'T3Lab Assistant — turn telemetry'
    if args.label:
        header += '  [{}]'.format(args.label)
    print('')
    print(header)
    print('=' * 60)
    print('  dir    {}'.format(directory))
    print('  files  {}'.format(', '.join(files) if files else '(none)'))

    if not turns:
        print('')
        print('  No turns recorded in range. Use the Assistant in Revit first,')
        print('  or widen the window with --days / --since.')
        return 0

    _stat_block('ALL', turns)

    if args.by_provider:
        seen = []
        for t in turns:
            p = t.get('provider') or '(unknown)'
            if p not in seen:
                seen.append(p)
        for p in seen:
            _stat_block('provider: {}'.format(p),
                        [t for t in turns if (t.get('provider') or '(unknown)') == p])

    print('')
    return 0


if __name__ == '__main__':
    sys.exit(main())
