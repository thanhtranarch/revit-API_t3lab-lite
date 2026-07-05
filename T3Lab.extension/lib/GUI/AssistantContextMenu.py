# -*- coding: utf-8 -*-
"""
T3Lab Assistant — Revit right-click context-menu entry.

Adds a "T3Lab Assistant" item to Revit's native right-click context menu (Canvas
and Project Browser) so the tool can be launched from anywhere with a right-click.
Registered once from startup.py.

Mechanism (Revit Context Menu API — Revit 2025+ only):
  * IContextMenuCreator.BuildContextMenu is invoked by Revit every time the user
    right-clicks. We attach a CommandMenuItem bound to the T3Lab Assistant
    push-button's generated external command.
  * CommandMenuItem(name, className, assemblyName) needs the IExternalCommand
    class name + assembly path. pyRevit generates exactly these for every button
    and exposes them via the ribbon wrapper (btn.class_name / btn.assembly_name),
    so we read them LIVE from the ribbon at right-click time rather than hard-code
    pyRevit's fragile "CustomCtrl_%CustomCtrl_%..." control-id string.

The Context Menu API does not exist before Revit 2025 — register() is a no-op
there (returns False) and never raises, so older hosts are unaffected.

Diagnostics: every step appends to ~/T3Lab_AI_Data/context_menu_debug.log so a
failing registration/resolution can be diagnosed from a single Revit test run.
"""

from __future__ import unicode_literals

import os
import datetime
import traceback

import clr
clr.AddReference('RevitAPIUI')

# Label shown in the context menu (single line — the ribbon title has a newline).
_MENU_LABEL = u"T3Lab Assistant"
# Identifier passed to RegisterContextMenu (docs: "FullClassName of the application").
_APP_ID = u"T3Lab.ContextMenu.AssistantCreator"
# Normalized (whitespace-stripped, lower-cased) name/title of the push-button.
_TARGETS = (u"t3labassistant",)

# ─── Debug log ─────────────────────────────────────────────────────────────────
_LOG_PATH = os.path.join(os.path.expanduser("~"), "T3Lab_AI_Data",
                         "context_menu_debug.log")


def _log(msg):
    """Append a timestamped line to the debug log. Never raises."""
    try:
        d = os.path.dirname(_LOG_PATH)
        if not os.path.isdir(d):
            os.makedirs(d)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(_LOG_PATH, "a") as f:
            f.write(u"[{}] {}\n".format(stamp, msg))
    except Exception:
        pass


# ─── Context Menu API is Revit 2025+ only ──────────────────────────────────────
try:
    from Autodesk.Revit.UI import IContextMenuCreator, CommandMenuItem
    _API_AVAILABLE = True
except Exception as _imp_ex:
    _API_AVAILABLE = False
    _log(u"IMPORT: Context Menu API unavailable (pre-2025?): {}".format(_imp_ex))

_log(u"MODULE LOADED — API_AVAILABLE={}".format(_API_AVAILABLE))

# Resolved (className, assemblyName) — cached after first successful lookup.
_CACHED_CMD = None


def _norm(s):
    """Collapse all whitespace and lower-case, so 'T3Lab\\nAssistant' == 't3labassistant'."""
    try:
        return u"".join((s or u"").split()).lower()
    except Exception:
        return u""


def _cmd_info(btn):
    """Return (className, assemblyName) for a ribbon push-button, or (None, None)."""
    try:
        cls = btn.class_name
        asm = btn.assembly_name
        if cls and asm:
            return cls, asm
    except Exception:
        pass
    return None, None


def _find_assistant(container):
    """Recursively search a ribbon container for the T3Lab Assistant button.

    Returns (className, assemblyName) or None.
    """
    try:
        children = iter(container)
    except Exception:
        return None
    for sub in children:
        if (_norm(getattr(sub, 'ui_title', u'')) in _TARGETS
                or _norm(getattr(sub, 'name', u'')) in _TARGETS):
            info = _cmd_info(sub)
            if info[0]:
                return info
        found = _find_assistant(sub)
        if found:
            return found
    return None


def _resolve_command():
    """Resolve (className, assemblyName) of the T3Lab Assistant push-button.

    Read live from the current ribbon (fully built by right-click time) and
    cached, so we never depend on pyRevit's internal control-id format.
    """
    global _CACHED_CMD
    if _CACHED_CMD:
        return _CACHED_CMD
    try:
        from pyrevit.coreutils.ribbon import get_current_ui
        ui = get_current_ui()

        # Scope to the T3Lab tab first for speed; fall back to a full walk.
        found = None
        try:
            tab = ui.find_child(u"T3Lab")
            _log(u"RESOLVE: T3Lab tab found={}".format(tab is not None))
            if tab is not None:
                found = _find_assistant(tab)
        except Exception as ex:
            _log(u"RESOLVE: tab lookup error: {}".format(ex))
            found = None
        if not found:
            found = _find_assistant(ui)

        if found:
            _CACHED_CMD = found
            _log(u"RESOLVE: OK class={} asm={}".format(found[0], found[1]))
        else:
            _log(u"RESOLVE: button NOT found in ribbon")
        return found
    except Exception as ex:
        _log(u"RESOLVE: error {}\n{}".format(ex, traceback.format_exc()))
        return None


if _API_AVAILABLE:

    class AssistantContextMenuCreator(IContextMenuCreator):
        """Adds the 'T3Lab Assistant' entry to Revit's right-click context menu."""

        def BuildContextMenu(self, menu):
            _log(u"BuildContextMenu CALLED by Revit")
            try:
                info = _resolve_command()
                if not info:
                    _log(u"BuildContextMenu: no command info — item NOT added")
                    return
                class_name, assembly_name = info
                menu.AddItem(CommandMenuItem(_MENU_LABEL, class_name, assembly_name))
                _log(u"BuildContextMenu: item ADDED")
            except Exception as ex:
                # Never break Revit's native context menu.
                _log(u"BuildContextMenu: error {}\n{}".format(ex, traceback.format_exc()))


def register(uictrl_app):
    """Register the context-menu creator. Call once during startup.

    Args:
        uictrl_app: UIControlledApplication (``__revit__`` in startup.py) or
            UIApplication — both expose RegisterContextMenu in Revit 2025+.
    Returns:
        bool: True if registered; False if the API is unavailable (pre-2025) or
        registration failed. Never raises.
    """
    # Log the host version to disambiguate "pre-2025" from "registration failed".
    ver = u"?"
    try:
        ctrl = getattr(uictrl_app, 'ControlledApplication', None) \
            or getattr(uictrl_app, 'Application', None)
        if ctrl is not None:
            ver = u"{} ({})".format(getattr(ctrl, 'VersionNumber', '?'),
                                    getattr(ctrl, 'VersionName', '?'))
    except Exception:
        pass
    _log(u"register() called — host={} API_AVAILABLE={} uictrl_type={}".format(
        ver, _API_AVAILABLE, type(uictrl_app).__name__))

    if not _API_AVAILABLE:
        _log(u"register(): API unavailable — skipping (needs Revit 2025+)")
        return False
    try:
        uictrl_app.RegisterContextMenu(_APP_ID, AssistantContextMenuCreator())
        _log(u"register(): RegisterContextMenu OK")
        return True
    except Exception as ex:
        _log(u"register(): RegisterContextMenu FAILED: {}\n{}".format(
            ex, traceback.format_exc()))
        return False
