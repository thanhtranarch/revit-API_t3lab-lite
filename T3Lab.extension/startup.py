# -*- coding: utf-8 -*-
"""
T3Lab Extension Startup Script
================================
Runs during pyRevit's OnStartup phase.

Responsibilities:
  1. Register the T3Lab Assistant as a native Revit DockablePane.
  2. Optionally auto-start the MCP server if the user has enabled that preference.

In the startup script context, `__revit__` is the UIControlledApplication
(available during Revit's OnStartup), which is required for DockablePane registration.
"""

from __future__ import unicode_literals

import os
import sys

# ─── Path bootstrap ────────────────────────────────────────────────────────────
_STARTUP_DIR = os.path.dirname(__file__)   # T3Lab.extension/
_LIB_DIR     = os.path.join(_STARTUP_DIR, 'lib')
for _p in (_STARTUP_DIR, _LIB_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# ─── Attempt DockablePane registration ─────────────────────────────────────────

try:
    import clr
    clr.AddReference('RevitAPIUI')
    from Autodesk.Revit.UI import DockablePaneId
    from System import Guid

    from GUI.AssistantPaneControl import ASSISTANT_PANE_GUID, AssistantPaneProvider

    pane_id = DockablePaneId(ASSISTANT_PANE_GUID)

    # __revit__ is UIControlledApplication during startup.py execution in pyRevit.
    # Use it directly if available; otherwise fall back to pyrevit.HOST_APP.
    _uictrld = None
    try:
        _uictrld = __revit__  # noqa: F821 — injected by pyRevit runtime
    except NameError:
        pass

    if _uictrld is None:
        try:
            from pyrevit import HOST_APP
            # HOST_APP.uiapp is UIApplication; HOST_APP.uicontrolledapp may exist
            _uictrld = getattr(HOST_APP, 'uicontrolledapp', None)
        except Exception:
            pass

    if _uictrld is not None and hasattr(_uictrld, 'RegisterDockablePane'):
        provider = AssistantPaneProvider()
        _uictrld.RegisterDockablePane(pane_id, 'T3Lab Assistant', provider)
    else:
        # Revit 2022+ fallback: try UIApplication.RegisterDockablePane
        try:
            from pyrevit import HOST_APP
            uiapp = HOST_APP.uiapp
            if hasattr(uiapp, 'RegisterDockablePane'):
                provider = AssistantPaneProvider()
                uiapp.RegisterDockablePane(pane_id, 'T3Lab Assistant', provider)
        except Exception:
            pass

except Exception:
    # Never crash Revit startup — silently skip pane registration.
    pass

# ─── Start file-based task watcher ─────────────────────────────────────────────
# Watches ~/T3Lab_AI_Data/task.json (and task.py) for AI-written tasks.
# Executes them in Revit context via ExternalEvent; result → result.json / result.txt.
# Zero-network alternative to the MCP HTTP server — data never leaves the machine.
try:
    from core.file_watcher import get_task_watcher
    get_task_watcher().start()
except Exception:
    pass
