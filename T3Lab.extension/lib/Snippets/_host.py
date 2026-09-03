# -*- coding: utf-8 -*-
"""
Host application access that survives a null document.

`revit.doc` is a getattr chain that silently yields None whenever the calling
engine has no active document — which is exactly the situation when a tool is
launched from the T3Lab Assistant's dockable pane rather than from its ribbon
button, and also when Revit is open with no project loaded.

Eight pushbuttons computed their Revit version as

    REVIT_VERSION = int(revit.doc.Application.VersionNumber)

at MODULE level, so importing the script at all raised

    'NoneType' object has no attribute 'Application'

and the tool never opened. The version does not come from the document in the
first place — the Application is reachable directly — so this module reads it
from the UIApplication and only falls back to the document.

Author: Tran Tien Thanh
"""

__author__ = "Tran Tien Thanh"
__title__ = "Host"

# Oldest release this codebase supports; used only when every probe fails, so
# that a version read can never be the thing that stops a tool from opening.
FALLBACK_VERSION = 2023


def injected_uiapp():
    """The UIApplication pyRevit injects as `__revit__`, or None.

    Under IronPython `__revit__` is a real builtin. Under the CPython engine it
    is injected into the *executing script's* globals instead, so a library
    module like this one never sees it in builtins — which silently made every
    caller fall through to "no Revit application" and turned a working model
    into a "no document is open" message. Probe both, then the call stack.
    """
    try:
        import builtins as _b
        uiapp = getattr(_b, '__revit__', None)
        if uiapp is not None:
            return uiapp
    except Exception:
        pass

    # CPython engine: __revit__ lives in the running script's globals.
    try:
        import sys
        uiapp = getattr(sys.modules.get('__main__'), '__revit__', None)
        if uiapp is not None:
            return uiapp
    except Exception:
        pass

    try:
        import inspect
        frame = inspect.currentframe()
        while frame is not None:
            uiapp = frame.f_globals.get('__revit__')
            if uiapp is not None:
                return uiapp
            frame = frame.f_back
    except Exception:
        pass

    return None


def host_uiapp():
    """Best available UIApplication: pyRevit HOST_APP first, then `__revit__`."""
    try:
        from pyrevit import HOST_APP
        if HOST_APP.uiapp is not None:
            return HOST_APP.uiapp
    except Exception:
        pass

    uiapp = injected_uiapp()
    if uiapp is not None:
        return uiapp

    try:
        from pyrevit import EXEC_PARAMS
        if EXEC_PARAMS.uiapp is not None:
            return EXEC_PARAMS.uiapp
    except Exception:
        pass

    return None


def get_revit_version(default=FALLBACK_VERSION):
    """Revit release year as an int (e.g. 2026). Never raises.

    Order matters: the Application is document-independent, so it answers
    even with no project open — which is the whole point of this helper.
    """
    uiapp = host_uiapp()
    try:
        return int(uiapp.Application.VersionNumber)
    except Exception:
        pass
    try:
        from pyrevit import HOST_APP
        return int(HOST_APP.version)
    except Exception:
        pass
    try:
        from pyrevit import revit
        return int(revit.doc.Application.VersionNumber)
    except Exception:
        return default


def resolve_doc():
    """(doc, error_text) — the document a tool should act on.

    Fallback chain, most specific first:
      1. pyrevit `revit.doc`        — correct inside a normal pushbutton engine
      2. uiapp.ActiveUIDocument     — direct API read, works when HOST_APP is stale
      3. the single open non-linked document — unambiguous, so safe to assume

    Anything else is a genuine "no target" situation and returns an actionable
    message rather than a null-document crash deeper in the call stack.
    """
    try:
        from pyrevit import revit
        doc = revit.doc
        if doc is not None:
            return doc, None
    except Exception:
        pass

    uiapp = host_uiapp()
    if uiapp is None:
        return None, (u"Cannot reach the Revit application from this engine. "
                      u"Open this tool from the T3Lab ribbon button.")

    try:
        uidoc = uiapp.ActiveUIDocument
        if uidoc is not None and uidoc.Document is not None:
            return uidoc.Document, None
    except Exception:
        pass

    try:
        docs = [d for d in uiapp.Application.Documents if not d.IsLinked]
        if len(docs) == 1:
            return docs[0], None
        if len(docs) > 1:
            return None, (u"Several models are open and none is active. "
                          u"Click the model you want, then run this again.")
    except Exception:
        pass

    return None, u"No Revit model is open. Open a project and try again."
