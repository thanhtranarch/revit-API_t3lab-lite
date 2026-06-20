# -*- coding: utf-8 -*-
__title__ = "Model\nAuditor"
__author__ = "T3Lab"
__doc__ = "Model Auditor — Model check, warnings, in-place models, and material lists."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.ModelAuditorDialog import show_model_auditor

if __name__ == '__main__':
    show_model_auditor(os.path.dirname(__file__), __revit__)
