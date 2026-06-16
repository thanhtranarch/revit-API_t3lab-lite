# -*- coding: utf-8 -*-
__title__ = "View\nHub"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "View Hub — Manage Views, View Templates, and Room Plans in one place."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.ViewHubDialog import show_view_hub

if __name__ == '__main__':
    show_view_hub(os.path.dirname(__file__), __revit__)
