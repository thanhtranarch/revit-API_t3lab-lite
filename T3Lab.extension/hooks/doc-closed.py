# -*- coding: utf-8 -*-
"""pyRevit hook — DocumentClosed.

Releases the MCP port the moment the LAST document closes: an idle Revit
sitting at the start page must not squat a port in the shared 48884-48894
range (Revit 2023 and 2026 instances all draw from it). While other
documents remain open the server keeps running — one Revit process = one
port, shared by all its documents. hooks/doc-opened.py (and doc-created.py)
bring the server back on the next open.

Requires the AppDomain-anchored singleton in core/server.py: this hook runs
in its own IronPython engine, and only the AppDomain anchor lets it reach
the live server instance instead of a blank per-engine copy whose
stop_server() would no-op while the orphan keeps the port.
"""

import os
import sys

_HOOKS_DIR = os.path.dirname(__file__)      # T3Lab.extension/hooks
_EXT_DIR   = os.path.dirname(_HOOKS_DIR)    # T3Lab.extension
_LIB_DIR   = os.path.join(_EXT_DIR, 'lib')
for _p in (_EXT_DIR, _LIB_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _open_primary_doc_count():
    """Count non-linked documents still open, or None when undeterminable."""
    app = None
    try:
        app = __eventsender__      # noqa: F821 — DB.Application (event sender)
    except NameError:
        app = None
    if app is None or not hasattr(app, 'Documents'):
        try:
            app = __revit__.Application   # noqa: F821 — injected by pyRevit
        except Exception:
            return None
    try:
        count = 0
        for _doc in app.Documents:
            try:
                if _doc.IsLinked:
                    continue
            except Exception:
                pass
            count += 1
        return count
    except Exception:
        return None


try:
    # DocumentClosed fires AFTER the document leaves app.Documents, so 0 here
    # means the closed one was the last. None (couldn't tell) must NOT stop
    # the server — wrongly keeping the port is recoverable, wrongly cutting a
    # live session mid-work is not.
    if _open_primary_doc_count() == 0:
        from Services.mcp_service import MCPService
        MCPService.stop_server()
except Exception:
    # Hooks must never throw into Revit's event loop.
    pass
