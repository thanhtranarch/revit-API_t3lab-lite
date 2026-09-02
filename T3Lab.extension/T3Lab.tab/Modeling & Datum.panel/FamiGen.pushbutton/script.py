# -*- coding: utf-8 -*-
__title__  = "FamiGen"
__author__ = "Tran Tien Thanh"
__doc__    = "FamiGen — Create Revit families from CAD blocks, JSON schema, or presets."

import os, sys

_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_lib = os.path.join(_ext_dir, 'lib')
if _lib not in sys.path:
    sys.path.insert(0, _lib)

# Force reload GUI modules to prevent caching issues in pyRevit
import GUI.FamiGenDialog
reload(GUI.FamiGenDialog)

from pyrevit import revit
from GUI.FamiGenDialog import show_family_creator

if __name__ == '__main__':
    show_family_creator(revit.doc, revit.doc.Application)
