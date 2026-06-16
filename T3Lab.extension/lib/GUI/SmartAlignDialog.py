# -*- coding: utf-8 -*-
"""Smart Align Dialog — event handling for the Smart Align launcher window."""

import os
from pyrevit import forms
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from smartalign.core import Alignment

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'SmartAlign.xaml')


class SmartAlignHandler(IExternalEventHandler):
    def __init__(self):
        self.align_type = None
        self.is_distribute = False

    def Execute(self, app):
        try:
            if self.is_distribute:
                from smartalign.distribute import main as dist_main
                dist_main(self.align_type)
            else:
                from smartalign.align import main as align_main
                align_main(self.align_type)
        except Exception as ex:
            import traceback
            print("SmartAlign Execution Error: {}".format(ex))

    def GetName(self):
        return "Smart Align Handler"


class SmartAlignWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, _XAML)
        self.handler = SmartAlignHandler()
        self.ext_event = ExternalEvent.Create(self.handler)

        # Register event handlers
        self.btn_align_left.Click += self._on_align_left
        self.btn_align_center_h.Click += self._on_align_center_h
        self.btn_align_right.Click += self._on_align_right
        self.btn_align_top.Click += self._on_align_top
        self.btn_align_center_v.Click += self._on_align_center_v
        self.btn_align_bottom.Click += self._on_align_bottom
        self.btn_distribute_h.Click += self._on_distribute_h
        self.btn_distribute_v.Click += self._on_distribute_v

        self.btn_minimize.Click += self._minimize
        self.btn_close_chrome.Click += self._close_chrome

    def _trigger(self, align_type, is_distribute=False):
        self.handler.align_type = align_type
        self.handler.is_distribute = is_distribute
        self.ext_event.Raise()

    def _on_align_left(self, sender, e):
        self._trigger(Alignment.HLEFT)

    def _on_align_center_h(self, sender, e):
        self._trigger(Alignment.HCENTER)

    def _on_align_right(self, sender, e):
        self._trigger(Alignment.HRIGHT)

    def _on_align_top(self, sender, e):
        self._trigger(Alignment.VTOP)

    def _on_align_center_v(self, sender, e):
        self._trigger(Alignment.VCENTER)

    def _on_align_bottom(self, sender, e):
        self._trigger(Alignment.VBOTTOM)

    def _on_distribute_h(self, sender, e):
        self._trigger(Alignment.HDIST, is_distribute=True)

    def _on_distribute_v(self, sender, e):
        self._trigger(Alignment.VDIST, is_distribute=True)

    def _minimize(self, sender, e):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def _close_chrome(self, sender, e):
        self.Close()


def show_smart_align():
    window = SmartAlignWindow()
    window.show(modal=False)
