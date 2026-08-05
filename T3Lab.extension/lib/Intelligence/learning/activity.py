# -*- coding: utf-8 -*-
"""
activity — a process-wide "when did the assistant last do something" signal.

The self-study loop must only run when the assistant is genuinely idle. The
window knows its own `_busy` flag, but the app-level Idling hook (startup.py)
has no handle to the window, so the two communicate through this tiny module:
the window calls `note_active()` whenever it starts a turn, and the idle loop
asks `is_idle(min_seconds)`.

Kept dead simple (one timestamp + a lock) and pure, so it is CPython-3 testable
and cannot itself throw into either the UI or the Idling handler.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Learning Activity"

import threading
import time

_lock = threading.Lock()
# Last moment the user/assistant did something. Seeded to "now" so a freshly
# opened Revit waits out one idle window before the first study cycle.
_last_active = time.time()


def note_active(now=None):
    """Record that the assistant is doing (or just did) real work."""
    global _last_active
    with _lock:
        _last_active = now if now is not None else time.time()


def seconds_since_active(now=None):
    with _lock:
        last = _last_active
    return max(0.0, (now if now is not None else time.time()) - last)


def is_idle(min_seconds, now=None):
    """True when nothing has happened for at least `min_seconds`."""
    return seconds_since_active(now) >= float(min_seconds)
