# -*- coding: utf-8 -*-
"""
Bilingual (Vietnamese / English) language-understanding layer.

Everything the assistant knows about the *words* a user typed lives here:
language detection, normalisation (teencode, abbreviations, typos), Revit
entity extraction, and sentence semantics (negation, modality, scope, risk).

The layer is deliberately separate from routing: `agents/dispatcher.py`,
`graph/planner.py` and the NLU all consume the same `Utterance` object, so a
word learned in one place improves every consumer at once.

Pure Python — IronPython 2.7 (Revit) and CPython 3 (dev/test_language.py).

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

from Intelligence.language.analyzer import (          # noqa: F401
    Utterance, analyze, resolve_ellipsis,
)
from Intelligence.language.lang_detect import (       # noqa: F401
    detect, is_vietnamese,
)
from Intelligence.language.normalizer import (        # noqa: F401
    normalize, expand_shorthand, fold,
)

__all__ = ['Utterance', 'analyze', 'resolve_ellipsis',
           'detect', 'is_vietnamese',
           'normalize', 'expand_shorthand', 'fold']
