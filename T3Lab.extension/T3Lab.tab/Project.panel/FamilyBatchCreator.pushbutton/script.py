# -*- coding: utf-8 -*-
__title__ = "Family Batch\nCreator"
__author__ = "Tran Tien Thanh"
__doc__ = "Family Batch Creator — Batch create families from CAD or JSON data."

import os, sys

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

from GUI.FamilyBatchCreatorDialog import show_family_batch_creator

if __name__ == '__main__':
    show_family_batch_creator(os.path.dirname(__file__), __revit__)
