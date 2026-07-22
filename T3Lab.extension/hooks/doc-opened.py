# -*- coding: utf-8 -*-
"""pyRevit hook — DocumentOpened.

Counterpart of hooks/doc-closed.py: that hook releases the MCP port when
the last document closes; this one brings the server back the moment a
document opens, so "file open ⇄ port held" stays symmetric. Honors the
same "auto_start_mcp": false opt-out as startup.py (absent = on).
Idempotent: start_server() returns immediately when already running.
"""

import os
import sys

_HOOKS_DIR = os.path.dirname(__file__)      # T3Lab.extension/hooks
_EXT_DIR   = os.path.dirname(_HOOKS_DIR)    # T3Lab.extension
_LIB_DIR   = os.path.join(_EXT_DIR, 'lib')
for _p in (_EXT_DIR, _LIB_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

try:
    from core import paths as _mcp_paths
    _auto = _mcp_paths.load_settings().get('auto_start_mcp')
    if _auto is None or _auto:
        from Services.mcp_service import MCPService
        # DocumentOpened runs on Revit's UI thread — a valid API context —
        # so the ExternalEvent for model-editing tools can be (re)created.
        MCPService.ensure_external_event()
        MCPService.start_server()
except Exception:
    # Hooks must never throw into Revit's event loop.
    pass
