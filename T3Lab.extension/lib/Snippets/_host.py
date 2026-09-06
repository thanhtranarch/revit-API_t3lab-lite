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

import sys

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


class DocResult(tuple):
    """2-tuple (doc, error_text) that safely proxies attribute access to `doc`.

    Supports unpacking:
        doc, err = resolve_doc()
    as well as direct single-variable assignment:
        doc = resolve_doc()
        if doc: ...
        doc.ActiveView
        hasattr(doc, 'Application')
    """
    def __new__(cls, doc, err):
        return super(DocResult, cls).__new__(cls, (doc, err))

    @property
    def doc(self):
        return self[0]

    @property
    def error(self):
        return self[1]

    def __getattr__(self, name):
        doc = self[0]
        if doc is not None:
            return getattr(doc, name)
        raise AttributeError("'DocResult' (no active document) has no attribute '%s'" % name)

    def __bool__(self):
        return self[0] is not None

    __nonzero__ = __bool__


def get_revit_version(doc_or_default=None, default=FALLBACK_VERSION):
    """Revit release year as an int (e.g. 2026). Never raises.

    Supports:
        get_revit_version()
        get_revit_version(default=2024)
        get_revit_version(doc)
        get_revit_version(doc, default=2024)
    """
    candidate_doc = None
    target_default = default
    if doc_or_default is not None:
        if isinstance(doc_or_default, int):
            target_default = doc_or_default
        else:
            candidate_doc = doc_or_default

    if candidate_doc is not None:
        try:
            return int(candidate_doc.Application.VersionNumber)
        except Exception:
            pass

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
        return target_default


def resolve_doc(candidate=None):
    """(doc, error_text) — the document a tool should act on.

    Returns a DocResult(doc, error_text) 2-tuple that also delegates attribute
    lookups directly to `doc` if assigned to a single variable.

    Fallback chain, most specific first:
      1. Valid candidate Document passed in
      2. pyrevit `revit.doc`        — correct inside a normal pushbutton engine
      3. uiapp.ActiveUIDocument     — direct API read, works when HOST_APP is stale
      4. the single open non-linked document — unambiguous, so safe to assume

    Anything else is a genuine "no target" situation and returns an actionable
    message rather than a null-document crash deeper in the call stack.
    """
    if candidate is not None:
        try:
            if isinstance(candidate, tuple) and len(candidate) > 0:
                candidate = candidate[0]
            if candidate is not None and hasattr(candidate, 'Application') and candidate.Application is not None:
                return DocResult(candidate, None)
        except Exception:
            pass

    try:
        from pyrevit import revit
        doc = revit.doc
        if doc is not None:
            return DocResult(doc, None)
    except Exception:
        pass

    uiapp = host_uiapp()
    if uiapp is None:
        return DocResult(None, (u"Cannot reach the Revit application from this engine. "
                                u"Open this tool from the T3Lab ribbon button."))

    try:
        uidoc = uiapp.ActiveUIDocument
        if uidoc is not None and uidoc.Document is not None:
            return DocResult(uidoc.Document, None)
    except Exception:
        pass

    try:
        docs = [d for d in uiapp.Application.Documents if not d.IsLinked]
        if len(docs) == 1:
            return DocResult(docs[0], None)
        if len(docs) > 1:
            return DocResult(None, (u"Several models are open and none is active. "
                                    u"Click the model you want, then run this again."))
    except Exception:
        pass

    return DocResult(None, u"No Revit model is open. Open a project and try again.")


def resolve_uidoc(candidate=None):
    """Best available UIDocument; None when no view context exists.

    Accepts an optional candidate UIDocument to support `uidoc = resolve_uidoc(uidoc)`.
    Fallback chain, most specific first:
      1. Valid candidate UIDocument passed in
      2. pyrevit `revit.uidoc`      — correct inside a normal pushbutton engine
      3. host_uiapp().ActiveUIDocument — direct API read when pyrevit uidoc is unset
    """
    if candidate is not None:
        try:
            if hasattr(candidate, 'Document') and candidate.Document is not None:
                return candidate
        except Exception:
            pass

    try:
        from pyrevit import revit
        uidoc = revit.uidoc
        if uidoc is not None:
            return uidoc
    except Exception:
        pass

    uiapp = host_uiapp()
    if uiapp is not None:
        try:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is not None:
                return uidoc
        except Exception:
            pass

    return None
