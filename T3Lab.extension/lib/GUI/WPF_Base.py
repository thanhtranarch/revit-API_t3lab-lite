# -*- coding: utf-8 -*-
"""
WPF Base

Base class for WPF-based GUI dialogs in pyRevit.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "WPF Base"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> .NET IMPORTS
import os, sys
from pyrevit import revit, forms

import os, clr
clr.AddReference("System")

from System.Diagnostics.Process import Start
from System.Windows.Window import DragMove
from System.Windows.Input import MouseButtonState
import wpf
from System.Windows import Application, Window, ResourceDictionary
from System import Uri

# ╦ ╦╔═╗╔═╗  ╔╦╗╔═╗╔╦╗╔═╗╦  ╔═╗╔╦╗╔═╗
# ║║║╠═╝╠╣    ║ ║╣ ║║║╠═╝║  ╠═╣ ║ ║╣
# ╚╩╝╩  ╚     ╩ ╚═╝╩ ╩╩  ╩═╝╩ ╩ ╩ ╚═╝ WPF TEMPLATE
# ==================================================
class my_WPF(Window):

    # ╔╦╗╔═╗╔╦╗╦ ╦╔═╗╔╦╗╔═╗
    # ║║║║╣  ║ ╠═╣║ ║ ║║╚═╗
    # ╩ ╩╚═╝ ╩ ╩ ╩╚═╝═╩╝╚═╝ METHODS
    #==================================================
    def add_wpf_resource(self):
        """Function to add WPF resources."""
        dir_path        = os.path.dirname(__file__)
        path_styles     = os.path.join(dir_path, 'Resources/WPF_styles.xaml')
        r               = ResourceDictionary()
        r.Source        = Uri(path_styles)
        self.Resources  = r

    # ╔═╗╦ ╦╦  ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║ ╦║ ║║  ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝╚═╝╩  ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ GUI EVENTS
    #==================================================
    def button_close(self, sender, e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.Close()

    def header_drag(self, sender, e):
        """Drag window by holding LeftButton on the header."""
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def Hyperlink_RequestNavigate(self, sender, e):
        """Forwarding for a Hyperlinks."""
        Start(e.Uri.AbsoluteUri)


