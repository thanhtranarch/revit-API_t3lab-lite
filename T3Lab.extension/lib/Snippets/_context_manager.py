# -*- coding: utf-8 -*-
"""
Context Manager Snippets

Code snippets for Revit transaction context managers.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Context Manager Snippets"

from Autodesk.Revit.DB import Transaction
import contextlib
import traceback

import sys, os
# ╔═╗╔═╗╔╗╔╔╦╗╔═╗═╗ ╦╔╦╗  ╔╦╗╔═╗╔╗╔╔═╗╔═╗╔═╗╦═╗╔═╗
# ║  ║ ║║║║ ║ ║╣ ╔╩╦╝ ║   ║║║╠═╣║║║╠═╣║ ╦║╣ ╠╦╝╚═╗
# ╚═╝╚═╝╝╚╝ ╩ ╚═╝╩ ╚═ ╩   ╩ ╩╩ ╩╝╚╝╩ ╩╚═╝╚═╝╩╚═╚═╝ CONTEXT MANAGERS
#====================================================================================================

@contextlib.contextmanager
def try_except(debug=False):
    """ContextManager for Try/Except statement with debug option for except.
    :param debug: if True - Exception error will be displayed with traceback.format_exc()"""
    try:
        yield
    except Exception as e:
        if debug:
            print("*"*20)
            print("Exception occured: " + traceback.format_exc())
            print("*"*20)


@contextlib.contextmanager
def ef_Transaction(doc, title, debug = True, exitscript = False ):
    t = Transaction(doc, title)
    t.Start()

    try:
        yield
        t.Commit()

    except Exception as e:
        if debug:
            print("*"*20)
            print("Exception occured - Transaction is being Rollbacked!")
            print(traceback.format_exc())
            print("*"*20)
        t.RollBack()

        if exitscript:
            print('*Script Excution stopped!*')
            sys.exit()






