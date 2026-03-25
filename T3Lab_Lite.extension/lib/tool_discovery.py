# -*- coding: utf-8 -*-
"""
T3Lab Tool Discovery — auto-register new pushbutton tools with the Assistant.

How it works
────────────
Every time the T3Lab Assistant starts it calls discover_new_tools().
That function scans the T3Lab_Lite.tab directory for *.pushbutton folders
that are NOT yet tracked in tool_registry.json.  For each new button it:

  1. Reads __title__ from the script.
  2. Derives a unique intent name  (e.g. "open_propertyline").
  3. Generates keyword hints for the NLU / keyword-fallback parser.
  4. Writes the entry to tool_registry.json so it is not re-processed.

Callers can then use get_registered_tools() to iterate all auto-registered
tools and inject them into TOOL_LAUNCHERS + the chat-UI quick-button strip.
"""
from __future__ import unicode_literals

import os
import re
import json

# ── Paths ─────────────────────────────────────────────────────────────────────
_LIB_DIR      = os.path.dirname(os.path.abspath(__file__))
_EXT_DIR      = os.path.dirname(_LIB_DIR)
_TAB_DIR      = os.path.join(_EXT_DIR, 'T3Lab_Lite.tab')
REGISTRY_FILE = os.path.join(_LIB_DIR, 'config', 'tool_registry.json')

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
    'UpperDimText.pushbutton',
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
    that has a script.py.

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
        for btn in sorted(os.listdir(panel_dir)):
            if not btn.endswith('.pushbutton'):
                continue
            script = os.path.join(panel_dir, btn, 'script.py')
            if not os.path.exists(script):
                continue
            title = _read_title(script) or btn.replace('.pushbutton', '')
            results.append({
                'button':      btn,
                'panel':       panel,
                'script_path': script,
                'title':       title,
            })
    return results


def _read_title(script_path):
    """Extract __title__ from the first 40 lines of a script file."""
    try:
        with open(script_path, 'r') as f:
            for i, line in enumerate(f):
                if i > 40:
                    break
                m = re.match(r'__title__\s*=\s*["\'](.+?)["\']', line)
                if m:
                    return m.group(1).replace('\\n', ' ').strip()
    except Exception:
        pass
    return None


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
    return {'version': 1, 'tools': {}}


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


def _gen_keywords(title, btn_name):
    """
    Generate lowercase keyword hints from the button name and title.
    Returns a deduplicated list sorted by length desc.
    """
    combined = (btn_name.replace('.pushbutton', '') + ' ' + title).lower()
    words = re.findall(r'[a-z][a-z0-9]*', combined)
    # Add the raw name and title as extra hints
    extras = [btn_name.replace('.pushbutton', '').lower(),
               title.lower()]
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
    reg   = load_registry()
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
            'intent':      intent,
            'keywords':    _gen_keywords(tool['title'], btn),
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
