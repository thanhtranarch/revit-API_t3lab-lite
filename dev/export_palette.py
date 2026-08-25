#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dump the RevitTheme palettes as JSON for dev/preview_xaml.ps1.

The preview harness runs under PowerShell and cannot import the Python module,
so the two palettes are handed over as data. Keeping this a script rather than
a checked-in JSON means the preview can never drift from the module.

Usage:
    python3 dev/export_palette.py [out.json]
"""
import importlib.util
import io
import json
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULE = os.path.join(REPO, "T3Lab.extension", "lib", "GUI", "RevitTheme.py")


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(REPO, "palette.json")
    spec = importlib.util.spec_from_file_location("revit_theme", MODULE)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert set(mod._LIGHT) == set(mod._DARK), "light/dark token sets diverged"
    with io.open(out, "w", encoding="utf-8") as fh:
        fh.write(json.dumps({"light": mod._LIGHT, "dark": mod._DARK}, indent=1))
    print("%d tokens -> %s" % (len(mod._LIGHT), out))


if __name__ == "__main__":
    main()
