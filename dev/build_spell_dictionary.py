# -*- coding: utf-8 -*-
"""
Build the bundled English frequency dictionary used by the deterministic
spell-checker (lib/Services/spell_dictionary.py).

Source of the general English word list:
    pyspellchecker  (MIT licence)  https://github.com/barrust/pyspellchecker
    resources/en.json.gz  — {word: corpus_frequency}, ~160k entries.

We merge that with our own construction / architecture domain terms
(lib/Services/data/construction_terms.txt) so trade vocabulary that a general
corpus lacks (e.g. "cementitious") is both accepted AND usable as a correction
target. Domain terms get a frequency floor so they reliably win the
edit-distance ranking over obscure general words.

Output (checked into the repo, loaded read-only at runtime by IronPython):
    lib/Services/data/en_freq.json.gz   — gzipped compact JSON {word: freq}

Run from the repo root:
    python3 dev/build_spell_dictionary.py

This is a DEV/BUILD script (CPython 3). It is never imported by the Revit
runtime — only its output asset is.
"""

import io
import os
import gzip
import json
import sys
import subprocess
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "T3Lab.extension", "lib", "Services", "data")
TERMS_FILE = os.path.join(DATA_DIR, "construction_terms.txt")
OUT_FILE = os.path.join(DATA_DIR, "en_freq.json.gz")

# Domain terms are floored to this frequency so they outrank rare general
# words when suggesting a correction (tuned against acoustic=2415,
# submitted=5799 in the base corpus).
DOMAIN_FLOOR = 8000


def _load_pyspellchecker_en():
    """Return the {word: freq} dict from pyspellchecker, installing it into a
    temp dir if it is not already importable. Dev-machine only."""
    try:
        import spellchecker  # noqa: F401
    except ImportError:
        target = os.path.join(tempfile.gettempdir(), "t3lab_spellbuild")
        print("pyspellchecker not found — installing to {} ...".format(target))
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--quiet",
            "--disable-pip-version-check", "--target", target,
            "pyspellchecker",
        ])
        sys.path.insert(0, target)
        import spellchecker  # noqa: F401

    res = os.path.join(os.path.dirname(spellchecker.__file__), "resources",
                       "en.json.gz")
    with gzip.open(res, "rt") as fh:          # text mode; utf-8 default
        data = json.load(fh)
    # keys are already lowercase in pyspellchecker's resource
    return {w: int(c) for w, c in data.items()}


def _load_domain_terms():
    terms = set()
    with io.open(TERMS_FILE, "r", encoding="utf-8") as fh:
        for line in fh:
            w = line.strip().lower()
            if not w or w.startswith("#"):
                continue
            # keep hyphenated compounds AND their parts as candidates
            terms.add(w)
    return terms


def main():
    freq = _load_pyspellchecker_en()
    base_n = len(freq)
    print("base English entries: {}".format(base_n))

    terms = _load_domain_terms()
    added, boosted = 0, 0
    for w in terms:
        if len(w) < 2:
            continue
        if w in freq:
            if freq[w] < DOMAIN_FLOOR:
                freq[w] = DOMAIN_FLOOR
                boosted += 1
        else:
            freq[w] = DOMAIN_FLOOR
            added += 1
    print("domain terms: {} (new: {}, boosted: {})".format(
        len(terms), added, boosted))
    print("total entries: {}".format(len(freq)))

    if not os.path.isdir(DATA_DIR):
        os.makedirs(DATA_DIR)
    # compact, sorted for stable diffs; gzip to keep the repo lean
    payload = json.dumps(freq, ensure_ascii=False, sort_keys=True,
                         separators=(",", ":"))
    with gzip.open(OUT_FILE, "wb") as fh:
        fh.write(payload.encode("utf-8"))
    print("wrote {} ({:.0f} KB)".format(OUT_FILE, os.path.getsize(OUT_FILE) / 1024.0))

    # sanity check against the known screenshot errors / correct forms
    for w in ("cementitious", "acrylic", "waterproofing", "acoustic",
              "submitted", "setting", "advise", "advice"):
        assert w in freq, "missing correct word: {}".format(w)
    for w in ("cemtitious", "acylic", "waterproffing", "acostic",
              "sybmitted", "seting"):
        assert w not in freq, "typo leaked into dict: {}".format(w)
    print("sanity check OK")


if __name__ == "__main__":
    main()
