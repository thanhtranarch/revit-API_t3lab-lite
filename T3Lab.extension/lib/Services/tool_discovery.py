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

# ── Buttons that are infrastructure / already hardcoded in TOOL_LAUNCHERS ─────
_SKIP_BUTTONS = {
    # Infrastructure — not user-facing tools
    'T3LabAssistant.pushbutton',
    'Settings.pushbutton',
    'StartMCP.pushbutton',
    # Already hard-coded in script.py TOOL_LAUNCHERS
    'BatchOut.pushbutton',
    'ParaSync.pushbutton',
    'LoadFamily.pushbutton',
    'LoadFamily(Cloud).pushbutton',
    'ProjectName.pushbutton',
    'Workset.pushbutton',
    'DimText.pushbutton',
    'UpperAll.pushbutton',
    'Reset Overrides.pushbutton',
    'Grids.pushbutton',
    'PropertyLine.pushbutton',
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
            with open(REGISTRY_FILE, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return {'version': REGISTRY_VERSION, 'tools': {}}


def save_registry(reg):
    """Persist the registry dict to disk."""
    try:
        d = os.path.dirname(REGISTRY_FILE)
        if not os.path.exists(d):
            os.makedirs(d)
        with open(REGISTRY_FILE, 'w') as f:
            json.dump(reg, f, ensure_ascii=False, indent=2)
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

    if new_tools:
        save_registry(reg)

    return new_tools


def get_registered_tools():
    """Return all tools currently in the registry (list of dicts)."""
    reg = load_registry()
    return list(reg.get('tools', {}).values())


def make_generic_launcher(script_path, title):
    """
    Build a zero-argument launcher function for an auto-discovered tool.

    Strategy (tries each in order):
      1. Load the module with imp.load_source — module-level code runs the tool.
      2. If that surfaces a *Window / *Dialog class, instantiate + ShowDialog.
      3. Fall back to exec() of the raw source.

    Returns:
        callable () → bool
    """
    def _launcher():
        # ── Strategy 1: imp.load_source ───────────────────────────────────
        try:
            import imp
            safe = re.sub(r'[^a-z0-9]', '_', title.lower())
            mod  = imp.load_source('_auto_' + safe, script_path)
            # Try to find a Window/Dialog class and show it
            for attr in dir(mod):
                if attr.endswith('Window') or attr.endswith('Dialog'):
                    cls = getattr(mod, attr, None)
                    if cls and callable(cls) and isinstance(cls, type):
                        try:
                            win = cls()
                            win.ShowDialog()
                            return True
                        except Exception:
                            pass
            # Module loaded successfully (script ran at module level)
            return True
        except Exception:
            pass

        # ── Strategy 2: exec the source ───────────────────────────────────
        try:
            with open(script_path, 'r') as f:
                src = f.read()
            g = {'__file__': script_path, '__name__': '__main__'}
            exec(compile(src, script_path, 'exec'), g)  # noqa
            return True
        except SystemExit:
            return True   # clean exit is normal
        except Exception:
            pass

        return False

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
