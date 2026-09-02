# -*- coding: utf-8 -*-
import os
import sys

# Ensure lib directory is in sys.path
pushbutton_dir = os.path.dirname(__file__)
panel_dir = os.path.dirname(pushbutton_dir)
tab_dir = os.path.dirname(panel_dir)
extension_dir = os.path.dirname(tab_dir)
lib_dir = os.path.join(extension_dir, 'lib')

try:
    reload
except NameError:
    try:
        from importlib import reload
    except Exception:
        reload = None

if reload:
    if 'GUI.UIShowcaseDialog' in sys.modules:
        reload(sys.modules['GUI.UIShowcaseDialog'])
    elif 'UIShowcaseDialog' in sys.modules:
        reload(sys.modules['UIShowcaseDialog'])

from GUI.UIShowcaseDialog import show_ui_standard_showcase

if __name__ == '__main__':
    show_ui_standard_showcase()
