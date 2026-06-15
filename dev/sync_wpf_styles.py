#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sync the T3Lab Terra shared style block into every tool XAML.

The master block lives in T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml
between the AUTO-SYNCED markers. Every XAML in T3Lab.extension/lib/GUI/Tools/
must carry a byte-identical copy of that block between the same markers.

Usage:
    python3 dev/sync_wpf_styles.py            # rewrite the block in all tool XAMLs
    python3 dev/sync_wpf_styles.py --check    # verify only; exit 1 on any mismatch
"""

import sys
import io
import os
import glob

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MASTER = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "GUI", "Resources", "WPF_styles.xaml")
TOOLS_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "GUI", "Tools")

BEGIN_MARK = "T3LAB SHARED STYLES v2 — AUTO-SYNCED"
END_MARK = "END T3LAB SHARED STYLES"


def read_lines(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return f.read().splitlines(True)


def find_block(lines, path):
    begin = end = None
    for i, line in enumerate(lines):
        if BEGIN_MARK in line:
            if begin is not None:
                raise SystemExit("%s: duplicate BEGIN marker" % path)
            begin = i
        elif END_MARK in line:
            if end is not None:
                raise SystemExit("%s: duplicate END marker" % path)
            end = i
    if begin is None or end is None or end <= begin:
        return None
    return begin, end


def main():
    check_only = "--check" in sys.argv[1:]

    master_lines = read_lines(MASTER)
    span = find_block(master_lines, MASTER)
    if span is None:
        raise SystemExit("Master file has no marker block: %s" % MASTER)
    # Block includes both marker lines so the markers themselves stay canonical.
    master_block = master_lines[span[0]:span[1] + 1]

    xaml_files = [p for p in sorted(glob.glob(os.path.join(TOOLS_DIR, "*.xaml")))
                  if "UIStandardShowcase.xaml" not in p]
    if not xaml_files:
        raise SystemExit("No XAML files found in %s" % TOOLS_DIR)

    missing, stale, synced = [], [], []
    for path in xaml_files:
        lines = read_lines(path)
        span = find_block(lines, path)
        rel = os.path.relpath(path, REPO_ROOT)
        if span is None:
            missing.append(rel)
            continue
        current = lines[span[0]:span[1] + 1]
        if current == master_block:
            continue
        if check_only:
            stale.append(rel)
        else:
            new_lines = lines[:span[0]] + master_block + lines[span[1] + 1:]
            with io.open(path, "w", encoding="utf-8") as f:
                f.write("".join(new_lines))
            synced.append(rel)

    for rel in missing:
        print("MISSING MARKERS: %s" % rel)
    if check_only:
        for rel in stale:
            print("OUT OF SYNC:     %s" % rel)
        ok = not missing and not stale
        print("%d files checked, %d out of sync, %d missing markers"
              % (len(xaml_files), len(stale), len(missing)))
        sys.exit(0 if ok else 1)
    else:
        for rel in synced:
            print("SYNCED:          %s" % rel)
        print("%d files checked, %d updated, %d missing markers"
              % (len(xaml_files), len(synced), len(missing)))
        sys.exit(1 if missing else 0)


if __name__ == "__main__":
    main()
