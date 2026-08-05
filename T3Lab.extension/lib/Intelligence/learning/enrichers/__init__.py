# -*- coding: utf-8 -*-
"""
Intelligence.learning.enrichers — one self-study unit of work each.

Every enricher exposes `run(**injected)` returning {'status', 'added'} and never
raises into the study loop. Pure decision helpers are factored out so they can be
unit-tested without Revit, WPF, or an LLM.
"""
from __future__ import unicode_literals
