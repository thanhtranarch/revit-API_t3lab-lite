# -*- coding: utf-8 -*-
"""
T3Lab Assistant — Dockable Pane Provider

Registers the T3Lab AI Assistant as a native Revit DockablePane so it can dock
alongside the Properties panel and Project Browser. The pane hosts the FULL
assistant window (SetupDockablePane loads T3LabAssistantWindow and re-parents
its content), which is why there is no separate pane chat implementation here.

Removed 2026-07-27: AssistantPaneController (~350 lines) and its
Tools/AssistantPane.xaml (~1370 lines). SetupDockablePane never constructed the
controller, so get_pane_controller() always returned None and the entire class —
along with two live bugs inside it — was unreachable. Recoverable from git
history if a lightweight pane chat is ever wanted.
"""

from __future__ import unicode_literals

import os
import sys

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
clr.AddReference('RevitAPIUI')

from System import Guid
from Autodesk.Revit.UI import IDockablePaneProvider, DockablePaneProviderData, DockablePaneState

# ─── Path bootstrap ────────────────────────────────────────────────────────────
_GUI_DIR  = os.path.dirname(__file__)                         # lib/GUI
_LIB_DIR  = os.path.dirname(_GUI_DIR)                        # lib
_EXT_DIR  = os.path.dirname(_LIB_DIR)                        # T3Lab.extension
for _p in (_LIB_DIR, _EXT_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)


# ─── Shared pane GUID (must match startup.py) ──────────────────────────────────
ASSISTANT_PANE_GUID = Guid('7F3A9B2E-C4D1-4E8F-A6B5-1234567890AB')

# ─── IDockablePaneProvider ─────────────────────────────────────────────────────

class AssistantPaneProvider(IDockablePaneProvider):
    """
    Revit calls SetupDockablePane() the first time the pane is shown.
    We load the UserControl XAML here and attach the controller.
    """

    def SetupDockablePane(self, data):
        try:
            import imp
            
            # Load the pushbutton script.py as a module to get T3LabAssistantWindow
            tab_dir = os.path.join(_EXT_DIR, 'T3Lab.tab')
            script_path = os.path.join(
                tab_dir, 'Support.panel', 'T3LabAssistant.pushbutton', 'script.py'
            )
            if os.path.isfile(script_path):
                if _LIB_DIR not in sys.path:
                    sys.path.insert(0, _LIB_DIR)
                if _EXT_DIR not in sys.path:
                    sys.path.insert(0, _EXT_DIR)
                    
                mod = imp.load_source('t3lab_assistant_full', script_path)
                if hasattr(mod, 'T3LabAssistantWindow'):
                    # Instantiate on UI thread as docked
                    win = mod.T3LabAssistantWindow(is_docked=True)
                    
                    # Detach visual content
                    content = win.Content
                    win.Content = None
                    
                    # Keep the window class instance alive
                    self._win_ref = win
                    
                    data.FrameworkElement = content
                    
                    from Autodesk.Revit.UI import EditorInteraction, EditorInteractionType
                    data.EditorInteraction = EditorInteraction(EditorInteractionType.KeepAlive)
                    return
            
            raise Exception("pushbutton script.py not found")

        except Exception as ex:
            import logging
            logging.basicConfig()
            logger = logging.getLogger("T3LabAssistant")
            logger.error("Error setting up DockablePane: %s", ex, exc_info=True)

            from System.Windows.Controls import Border, TextBlock
            from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness, TextWrapping
            from System.Windows.Media import Brushes

            border = Border()
            border.Background = Brushes.Crimson
            border.Padding = Thickness(20)

            lbl = TextBlock()
            lbl.Text = u'T3Lab Assistant pane could not load:\n{}'.format(ex)
            lbl.Foreground = Brushes.White
            lbl.TextWrapping = TextWrapping.Wrap
            lbl.HorizontalAlignment = HorizontalAlignment.Center
            lbl.VerticalAlignment   = VerticalAlignment.Center

            border.Child = lbl
            data.FrameworkElement   = border
