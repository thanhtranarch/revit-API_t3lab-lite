#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Migrate inline colors in all tool XAML files from Terra (v2) to Lumina.
"""

import os
import glob
import re

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "GUI", "Tools")

# Mapping from Terra (v2) -> Lumina color hexes (case-insensitive keys)
COLOR_MAP = {
    "#1C2B33": "#0F172A", # ink text -> slate
    "#0F766E": "#0F172A", # primary teal -> slate
    "#115E59": "#1E293B", # primary hover -> slate hover
    "#0B4F4A": "#0F172A", # primary pressed -> slate
    "#F6F8F8": "#F8FAFC", # surface bg -> white/soft gray
    "#E6EDEC": "#F1F5F9", # surface hover -> gray hover
    "#D9EBE8": "#E2E8F0", # selected tint -> surface press
    "#DDE5E7": "#E2E8F0", # border -> border gray
    "#C7D2D4": "#CBD5E1", # input border
    "#15803D": "#10B981", # success green -> emerald
    "#166534": "#059669", # success hover -> emerald hover
    "#DC2626": "#EF4444", # danger red -> rose red
    "#B91C1C": "#DC2626", # danger hover -> rose hover
    "#AEBFBC": "#E2E8F0", # disabled bg
}

def migrate_file(path):
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    original = content
    # Replace each color (case-insensitive match)
    for old, new in COLOR_MAP.items():
        # Compile pattern for case-insensitive match
        pattern = re.compile(re.escape(old), re.IGNORECASE)
        content = pattern.sub(new, content)

    if content != original:
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        print("Migrated: %s" % os.path.basename(path))
        return True
    return False

def main():
    xaml_files = glob.glob(os.path.join(TOOLS_DIR, "*.xaml"))
    print("Found %d XAML files." % len(xaml_files))
    
    migrated_count = 0
    for path in xaml_files:
        if migrate_file(path):
            migrated_count += 1
            
    print("Migration complete. %d files updated." % migrated_count)

if __name__ == "__main__":
    main()
