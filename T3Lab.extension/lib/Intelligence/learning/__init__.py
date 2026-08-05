# -*- coding: utf-8 -*-
"""
Intelligence.learning — the assistant's self-study layer.

While Revit is open and the assistant is idle, a background loop (study_engine)
runs enrichers that (a) refresh the on-disk knowledge stores and (b) curate a
local training dataset (dataset.py). A separate CPython+GPU job (tools/train/)
later fine-tunes a local Ollama model from that dataset — weight training cannot
run inside IronPython, so this package only ever COLLECTS and CURATES.

Every module here is pure Python with guarded imports so it stays importable
under both IronPython 2.7 (in Revit) and CPython 3 (dev tests).
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Intelligence.learning"
