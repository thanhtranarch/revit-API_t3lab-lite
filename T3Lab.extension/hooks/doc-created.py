# -*- coding: utf-8 -*-
"""pyRevit hook — DocumentCreated.

Mirror of hooks/doc-opened.py for brand-new (unsaved) documents: a fresh
project is "a file open" too, so the MCP server must come back for it.
Same "auto_start_mcp": false opt-out, same idempotent start.
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
        # DocumentCreated runs on Revit's UI thread — a valid API context —
        # so the ExternalEvent for model-editing tools can be (re)created.
        MCPService.ensure_external_event()
        MCPService.start_server()
except Exception:
    # Hooks must never throw into Revit's event loop.
    pass
