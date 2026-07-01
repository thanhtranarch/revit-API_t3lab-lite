# -*- coding: utf-8 -*-
"""
WPF Base

Base class for WPF-based GUI dialogs in pyRevit.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "WPF Base"

import clr
clr.AddReference("System")
clr.AddReference("PresentationFramework")

from System.Diagnostics.Process import Start
from System.Windows import Window, WindowState


class my_WPF(Window):

    def button_close(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def Hyperlink_RequestNavigate(self, sender, e):
        Start(e.Uri.AbsoluteUri)
