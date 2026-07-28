# -*- coding: utf-8 -*-
"""
Tool Discovery

Auto-registers and discovers T3Lab pushbutton tools.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Tool Discovery"

import os
import re
import io
import json

# ── Paths ─────────────────────────────────────────────────────────────────────
_SERVICES_DIR = os.path.dirname(os.path.abspath(__file__))
_LIB_DIR      = os.path.dirname(_SERVICES_DIR)
_EXT_DIR      = os.path.dirname(_LIB_DIR)
_TAB_DIR      = os.path.join(_EXT_DIR, 'T3Lab.tab')
REGISTRY_FILE = os.path.join(_LIB_DIR, 'config', 'tool_registry.json')

# Bump when the entry schema changes — a mismatched on-disk registry is
# rebuilt from scratch so every entry carries the new fields (doc, xaml).
REGISTRY_VERSION = 2

# ── Buttons that must not be auto-registered ─────────────────────────────────
# Keep this list to what genuinely cannot be auto-derived. It used to also
# exclude ParaSync / ProjectName / Workset / DimText / UpperAll / Reset
# Overrides / Grids / LoadFamily* as "already hard-coded" — but those
# pushbuttons no longer exist (or never did), so the exclusions were inert.
# PropertyLine.pushbutton was the costly one: it EXISTS, was skipped here as
# hard-coded, and was never actually in TOOL_LAUNCHERS — leaving a real tool
# the assistant could not reach at all.
_SKIP_BUTTONS = {
    # Infrastructure — not user-facing tools. (Settings.pushbutton and
    # StartMCP.pushbutton were listed here too; neither exists — the real
    # buttons are LLMsSetting and MCPControl, and those SHOULD be openable.)
    'T3LabAssistant.pushbutton',
    # Have a dedicated launcher in script.py (_SPECIAL_LAUNCHERS): they take
    # arguments or open through a Dialog module rather than a script entry.
    'BatchOut.pushbutton',
    'ManaFami.pushbutton',
}


# ─────────────────────────────────────────────────────────────────────────────
# Scanning helpers
# ─────────────────────────────────────────────────────────────────────────────

def scan_all_pushbuttons():
    """
    Walk T3Lab_Lite.tab and return a list of dicts for every *.pushbutton
    that has a script.py, at any nesting depth (panel/pushbutton,
    panel/stack/pushbutton, panel/pulldown/stack/pushbutton, etc.).

    Each dict: {button, panel, script_path, title}
    """
    results = []
    if not os.path.isdir(_TAB_DIR):
        return results
    for panel in sorted(os.listdir(_TAB_DIR)):
        if not panel.endswith('.panel'):
            continue
        panel_dir = os.path.join(_TAB_DIR, panel)
        if not os.path.isdir(panel_dir):
            continue
        for root, dirs, files in os.walk(panel_dir):
            btn = os.path.basename(root)
            if not btn.endswith('.pushbutton'):
                continue
            script = os.path.join(root, 'script.py')
            if not os.path.exists(script):
                continue
            title, doc, xamls = _read_meta(script)
            title = title or btn.replace('.pushbutton', '')
            results.append({
                'button':      btn,
                'panel':       panel,
                'script_path': script,
                'title':       title,
                'doc':         doc,
                'xamls':       xamls,
            })
    return results


def _read_meta(script_path):
    """Extract (__title__, first doc line, referenced .xaml basenames) from a
    script source file. All three feed the assistant's tool catalog so a tool
    is recognisable by every name it carries (title / folder / XAML)."""
    title, doc, xamls = None, '', []
    try:
        with io.open(script_path, 'r', encoding='utf-8', errors='ignore') as f:
            src = f.read()
    except Exception:
        return title, doc, xamls

    m = re.search(r'__title__\s*=\s*["\'](.+?)["\']', src)
    if m:
        title = m.group(1).replace('\\n', ' ').strip()

    # Tooltip: explicit __doc__ assignment wins, else module docstring
    m = re.search(r'__doc__\s*=\s*u?["\'](.+?)["\']', src)
    if not m:
        m = re.search(r'^\s*u?"""(.*?)"""', src, re.S | re.M)
    if not m:
        m = re.search(r"^\s*u?'''(.*?)'''", src, re.S | re.M)
    if m:
        for ln in m.group(1).strip().splitlines():
            ln = ln.strip()
            if not ln or ln.startswith(('#', '-', '=', '~')):
                continue
            # First meaningful line that is not just the title repeated
            if title and ln.lower() == title.lower():
                continue
            doc = ln[:140]
            break

    # XAML files the script loads — their basenames are tool aliases
    for x in re.findall(r'([\w\-. ]+?\.xaml)', src):
        base = os.path.basename(x).rsplit('.xaml', 1)[0].strip()
        if base and base.lower() != 'wpf_styles' and base not in xamls:
            xamls.append(base)
    return title, doc, xamls


# ─────────────────────────────────────────────────────────────────────────────
# Registry helpers
# ─────────────────────────────────────────────────────────────────────────────

def load_registry():
    """Return the on-disk registry dict, or a blank one if absent/corrupt."""
    try:
        if os.path.exists(REGISTRY_FILE):
            with io.open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {'version': REGISTRY_VERSION, 'tools': {}}


def save_registry(reg):
    """Persist the registry dict to disk.

    Serialize to an ASCII string FIRST, then write it in one shot.
    The old `json.dump(reg, f, ensure_ascii=False)` on a bytes-mode file
    blew up with UnicodeEncodeError mid-write under IronPython 2.7 the
    moment a tool title/doc carried Vietnamese text — leaving a 0-byte
    registry, silently (exception swallowed). Every session then saw an
    empty registry: no auto-discovered tool could ever be opened by name.
    """
    try:
        d = os.path.dirname(REGISTRY_FILE)
        if not os.path.exists(d):
            os.makedirs(d)
        data = json.dumps(reg, ensure_ascii=True, indent=2, sort_keys=True)
        if isinstance(data, bytes):                     # py2 str
            data = data.decode('ascii')
        with io.open(REGISTRY_FILE, 'w', encoding='utf-8') as f:
            f.write(data)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# Name / keyword generation
# ─────────────────────────────────────────────────────────────────────────────

def _button_to_intent(btn_name):
    """
    'LoadFamily(Cloud).pushbutton'  →  'open_loadfamily_cloud'
    'Reset Overrides.pushbutton'    →  'open_reset_overrides'
    """
    name = btn_name.replace('.pushbutton', '')
    name = re.sub(r'[\s\(\)\-\[\]]', '_', name)
    name = re.sub(r'_+', '_', name).strip('_')
    return 'open_' + name.lower()


def _camel_split(text):
    """'DWGManagement' → 'DWG Management' (word boundaries for keyword gen)."""
    return re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', text)


def _gen_keywords(title, btn_name, xamls=None):
    """
    Generate lowercase keyword hints from the button name, title, and any
    XAML basenames the script references. Returns a deduplicated list
    sorted by length desc.
    """
    base = btn_name.replace('.pushbutton', '')
    parts = [base, _camel_split(base), title]
    for x in (xamls or []):
        parts.append(x)
        parts.append(_camel_split(x))
    combined = ' '.join(parts).lower()
    words = re.findall(r'[a-z][a-z0-9]*', combined)
    # Add the raw names as extra hints
    extras = [base.lower(), title.lower()] + [x.lower() for x in (xamls or [])]
    all_kw = list(set(words + extras))
    all_kw = [k for k in all_kw if len(k) > 1]
    return sorted(all_kw, key=len, reverse=True)


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def discover_new_tools():
    """
    Scan the extension for pushbuttons not yet in the registry.
    Register each new one and return the list of new entries.

    Returns:
        list of tool dicts (empty if nothing new was found)
    """
    reg = load_registry()
    # Schema migration: rebuild from scratch so every entry gains the
    # newer fields (doc, xaml) used by the assistant's tool catalog.
    if reg.get('version') != REGISTRY_VERSION:
        reg = {'version': REGISTRY_VERSION, 'tools': {}}

    # Prune entries whose script.py is gone. Registration was append-only, so a
    # renamed or deleted pushbutton stayed in the registry forever and kept
    # being offered to the model as an openable tool.
    stale = [btn for btn, e in (reg.get('tools') or {}).items()
             if not (e.get('script_path') and os.path.exists(e['script_path']))]
    for btn in stale:
        reg['tools'].pop(btn, None)

    known = set(reg.get('tools', {}).keys())

    new_tools = []
    for tool in scan_all_pushbuttons():
        btn = tool['button']
        if btn in _SKIP_BUTTONS or btn in known:
            continue
        intent = _button_to_intent(btn)
        entry = {
            'button':      btn,
            'panel':       tool['panel'],
            'script_path': tool['script_path'],
            'title':       tool['title'],
            'doc':         tool.get('doc', ''),
            'xaml':        tool.get('xamls', []),
            'intent':      intent,
            'keywords':    _gen_keywords(tool['title'], btn, tool.get('xamls')),
        }
        reg.setdefault('tools', {})[btn] = entry
        new_tools.append(entry)

    if new_tools or stale:
        save_registry(reg)

    return new_tools


def get_registered_tools():
    """Return all tools currently in the registry (list of dicts)."""
    reg = load_registry()
    return list(reg.get('tools', {}).values())


def _err_text(exc):
    """Exception message as unicode. Never raises.

    `str(exc)` is unsafe under IronPython 2.7: an error carrying non-ASCII
    text (a Vietnamese family name, a localised Windows string) throws
    UnicodeEncodeError from inside the handler.
    """
    try:
        return u"{}".format(exc)
    except Exception:
        pass
    try:
        r = repr(exc)
        return r.decode('utf-8', 'replace') if isinstance(r, bytes) else u"{}".format(r)
    except Exception:
        return u"<unprintable error>"


def run_tool_script(script_path, title=None):
    """Execute a pushbutton's script.py the way pyRevit runs it.

    Returns (ok, error_text) — error_text is u'' on success.

    pyRevit executes a button script as the __main__ module, and 36 of the
    42 T3Lab pushbuttons put their entry point inside `if __name__ ==
    '__main__':`. The previous implementation tried `imp.load_source` FIRST
    (module name `_auto_<title>`), so that guard never fired: the module
    imported, nothing ran, and the launcher still reported success. The
    caller then told the user the tool had opened while nothing happened.
    Running the source as __main__ is therefore the primary strategy, not
    the fallback.

    The source is read as BYTES on purpose: every script starts with a
    `# -*- coding: utf-8 -*-` line, and CPython/IronPython 2.7 raise
    "SyntaxError: encoding declaration in Unicode string" when such source
    is passed to compile() as unicode. Bytes honour the declaration on both
    Python 2 and 3.
    """
    if not script_path or not os.path.exists(script_path):
        return False, u"Script not found on disk: {}".format(
            script_path or u"<empty path>")

    # ── Strategy 1: run as __main__ (correct pyRevit semantics) ───────────
    first_err = u""
    try:
        with open(script_path, 'rb') as f:
            src = f.read()
        glb = {'__file__': script_path, '__name__': '__main__', '__doc__': None}
        exec(compile(src, script_path, 'exec'), glb)  # noqa
        return True, u""
    except SystemExit:
        # pyRevit's forms.alert(exitscript=True) is a normal exit path.
        return True, u""
    except Exception as ex:
        first_err = _err_text(ex)

    # ── Strategy 2: import the module and show a window it defines ────────
    # Only counts as success when a *Window / *Dialog is ACTUALLY shown —
    # a bare successful import means nothing ran.
    try:
        import imp
        safe = re.sub(r'[^a-z0-9]', '_', (title or 'tool').lower())
        mod  = imp.load_source('_auto_' + safe, script_path)
        for attr in dir(mod):
            if not (attr.endswith('Window') or attr.endswith('Dialog')):
                continue
            cls = getattr(mod, attr, None)
            if not isinstance(cls, type):
                continue
            try:
                cls().ShowDialog()
                return True, u""
            except Exception:
                continue
    except SystemExit:
        return True, u""
    except Exception:
        pass

    return False, first_err or u"Could not run {}".format(
        os.path.basename(script_path))


def make_generic_launcher(script_path, title):
    """Build a zero-argument launcher for a tool. Returns callable () → (ok, err)."""
    def _launcher():
        return run_tool_script(script_path, title)

    return _launcher


def build_system_prompt_section(tools):
    """
    Return an extra system-prompt snippet listing auto-discovered tools.
    Pass this to get_system_prompt() / parse_command().
    """
    if not tools:
        return ''
    lines = ['  ── Auto-discovered tools ────────────────────────────────────────────────────']
    for t in tools:
        lines.append('  {}   params: {{}}   (title: "{}")'.format(t['intent'], t['title']))
    return '\n'.join(lines)
