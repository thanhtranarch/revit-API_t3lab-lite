# -*- coding: utf-8 -*-
"""
Tests for the deterministic spell-check engine (lib/Services/spell_dictionary).

Runs under CPython 3 on the dev machine (the engine itself is written to also
run under IronPython 2.7). Ground truth is the real T3Lab model scan that
Claude-via-MCP flagged but the local LLM missed:

    CEMTITIOUS   -> CEMENTITIOUS
    ACYLIC       -> ACRYLIC
    SETING-OUT   -> SETTING-OUT
    WATERPROFFING-> WATERPROOFING
    ACOSTIC      -> ACOUSTIC
    SYBMITTED    -> SUBMITTED
    to advice    -> to advise

Run:  python3 dev/test_spell_checker.py
"""

import io
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, "T3Lab.extension", "lib", "Services"))

import spell_dictionary as SD  # noqa: E402


def _corrections(text):
    """{wrong_lower: right_lower} for every finding in `text`."""
    return dict((f["wrong"].lower(), f["right"].lower())
                for f in SD.check_text(text))


# ── True positives: the 7 errors from the real scan (in situ) ───────────────
POSITIVE = [
    # (note text as it appears on the drawing, wrong token, expected correction)
    (u"75MM THK CEMTITIOUS SCREED TO FALL",            u"cemtitious",    u"cementitious"),
    (u"ACYLIC SOLID SURFACE COUNTER TOP",              u"acylic",        u"acrylic"),
    (u"REFER TO SETING-OUT PLAN",                      u"seting",        u"setting"),
    (u"WATERPROFFING MEMBRANE UP-TURN 300MM",          u"waterproffing", u"waterproofing"),
    (u"ACOSTIC RATING STC 45",                         u"acostic",       u"acoustic"),
    (u"TO BE SYBMITTED FOR APPROVAL",                  u"sybmitted",     u"submitted"),
    (u"CONTRACTOR TO ADVICE ON SITE CONDITION",        u"to advice",     u"to advise"),
]

# ── True negatives: must NEVER be flagged ───────────────────────────────────
NEGATIVE = [
    u"FFL +100.000",
    u"SSL RC SLAB",
    u"100 THK CEMENTITIOUS SCREED",
    u"WATERPROOFING MEMBRANE",
    u"ACOUSTIC RATING STC 45",
    u"SETTING-OUT PLAN",
    u"TO BE SUBMITTED FOR APPROVAL",
    u"GRC CLADDING SYSTEM",
    u"ACMV DUCT ABOVE",
    u"M20 CONCRETE GRADE",
    u"DRG NO. A-201 REV C",
    u"TYP. DETAIL U.N.O.",
    u"NTS",
    u"GALV. M.S. BRACKET",
    u"PHASE II HANDOVER",
    u"ALUMINIUM MULLION AND TRANSOM",
    u"PLASTERBOARD CEILING BULKHEAD",
    u"SANITARYWARE BY OWNER",
    u"75x50 SKIRTING",
    u"Ø32 DOWNPIPE",
    # ── false positives reported from a real scan (2026-07-29) ─────────────
    # closed compounds / prefixed words absent from the general corpus; the
    # engine used to "correct" them into nonsense
    u"SOLID SURFACE COUNTERTOP TO WET AREA",      # was -> COUNTERTYPE
    u"CERAMIC TILE BACKSPLASH ABOVE COUNTER",     # was -> BACKSLASH
    u"PREFINISHED ALUMINIUM CAPPING",             # was -> REFINISHED
    u"STORMWATER DOWNPIPE TO DISCHARGE",
    u"DOWNLIGHTS AND UPLIGHTS TO CEILING",
    u"NON-SLIP HOMOGENEOUS TILE",
    u"WATERTIGHT SUBFRAME WITH TOPCOAT",
    # brand / software names
    u"MODEL PREPARED IN REVIT 2023",              # was -> REFIT
    u"EXPORTED FROM AUTOCAD AND NAVISWORKS",
    # non-ASCII text must not be shredded into fake ASCII tokens
    u"ALUMINIUM FAÇADES PANEL SYSTEM",            # was -> "ades" -> "ages"
    u"Tường gạch nhẹ 100mm hoàn thiện sơn nước",
]

# ── Corrections that must survive the well-formedness relaxation ────────────
# (a typo whose tail happens to be a real word must still be reported)
POSITIVE.append((u"CONCRETE PRESURE TEST", u"presure", u"pressure"))


def main():
    if not SD.is_available():
        print("FAIL: dictionary asset not loaded: {}".format(SD.load_error()))
        return 1

    failures = []

    for text, wrong, right in POSITIVE:
        corr = _corrections(text)
        if wrong not in corr:
            failures.append(u"MISS  {!r}: expected to flag {!r}, got {}".format(
                text, wrong, sorted(corr.keys())))
        elif corr[wrong] != right:
            failures.append(u"WRONG {!r}: {!r} -> {!r} (expected {!r})".format(
                text, wrong, corr[wrong], right))

    for text in NEGATIVE:
        corr = _corrections(text)
        if corr:
            failures.append(u"FALSE+ {!r}: unexpectedly flagged {}".format(
                text, corr))

    total = len(POSITIVE) + len(NEGATIVE)
    if failures:
        print(u"\n".join(failures))
        print("\n{}/{} checks passed, {} FAILED".format(
            total - len(failures), total, len(failures)))
        return 1
    print("all {} checks passed ({} typos caught, {} clean strings left "
          "alone)".format(total, len(POSITIVE), len(NEGATIVE)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
