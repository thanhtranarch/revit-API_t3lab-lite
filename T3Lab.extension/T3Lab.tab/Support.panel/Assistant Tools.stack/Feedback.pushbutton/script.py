# -*- coding: utf-8 -*-
"""
Send Feedback

Popup window that lets the user send feedback or suggestions about T3Lab
tools, delivered by email to the T3Lab team.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__title__   = "Send Feedback"
__author__  = "Tran Tien Thanh"
__version__ = "1.0.0"

import os
import sys

# Path setup — script.py lives 4 levels below T3Lab.extension/ (inside a .stack bundle)
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit, forms
from GUI.FeedbackDialog import show_feedback_dialog


def main():
    if not revit.doc:
        forms.alert("No open Revit document found.", exitscript=True)
    show_feedback_dialog()


if __name__ == '__main__':
    main()
