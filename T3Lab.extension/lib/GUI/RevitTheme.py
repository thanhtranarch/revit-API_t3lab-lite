# -*- coding: utf-8 -*-
"""
Revit theme bridge for the T3Lab Assistant surface.

Two jobs:

1. **Follow Revit.** Revit 2024+ exposes ``UIThemeManager.CurrentTheme``
   (Light / Dark). The assistant lives inside a Revit DockablePane, so a
   hardcoded light surface reads as a foreign window glued into the host —
   a white rectangle glaring out of a dark Revit. ``current_theme()`` reads
   the host theme (Light on older hosts, which have no dark mode) and every
   palette lookup follows it.

2. **One palette, one place.** The assistant is a *chat* surface, not a tool
   dialog: it uses the warm "paper" split (paper background · warm-grey user
   bubble · bubble-less assistant text) instead of the Lumina tool palette in
   ``.claude/rules/ui-design-standard.md``. That deviation is deliberate and
   scoped to this one surface — the copyright amber ``#F59E0B`` and the shared
   button styles are untouched, so the Lumina audit still passes. Everything
   else routes through the tokens below rather than being re-typed as a hex
   literal in 70 different places.

Usage from XAML (re-themes live, no reload):

    Background="{DynamicResource T3ThemeChatBg}"

Usage from code (bubbles, chips, cards built in script.py):

    from GUI.RevitTheme import brush, rgb
    border.Background = brush('CardBg')

``apply(window)`` writes every token into the element's ``Resources`` as a
``SolidColorBrush`` named ``T3Theme<Token>``, so a theme flip is a single
call — WPF re-evaluates every ``DynamicResource`` that points at it.

Pure Python + guarded .NET imports, so it stays importable under CPython 3 for
the headless tests.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals


# ─── Palettes ─────────────────────────────────────────────────────────────────
# Light is the warm paper split from the reference chat UI. Dark is the same
# split rebuilt on Revit's dark-theme greys, so the pane reads as one surface
# with the host rather than as a light card dropped into a dark application.

_LIGHT = {
    # Surfaces
    'AppBg':          '#FAF9F5',   # window/pane backdrop — warm paper
    'ChatBg':         '#FAF9F5',   # chat scroll surface (same paper: seamless)
    'ComposerBg':     '#FAF9F5',   # bar hosting the composer card
    'CardBg':         '#FFFFFF',   # composer card, popups, overlays
    'CardBorder':     '#E5E3DA',   # hairline around cards
    'Divider':        '#EAE8E0',   # separators, footer rule
    'PaneEdge':       '#DCDAD1',   # outer edge when floating

    # Chat split
    'UserBubbleBg':   '#F0EEE6',   # user message — warm grey, NOT a blue pill
    'UserBubbleText': '#282722',
    'BotText':        '#3D3C38',   # assistant reply: text on paper, no bubble

    # Type
    'Ink':            '#1F1E1D',   # headings, strong text
    'Muted':          '#6F6E68',   # secondary text
    'Faint':          '#A3A199',   # placeholders, disclaimers, hints

    # Accent + status
    'Accent':         '#D97757',   # brand coral — avatar, active states
    'AccentSoft':     '#F6E7DF',   # coral tint fill
    'Blue':           '#3B82F6',   # links / info
    'Success':        '#0B8A5A',
    'Danger':         '#D23B3B',
    'Amber':          '#F59E0B',   # copyright (Lumina standard, do not retint)

    # Controls
    'IconFg':         '#8B8981',   # muted line icons (Claude-style action row)
    'IconFgHover':    '#3D3C38',
    'IconHoverBg':    '#EFEDE5',
    'InputText':      '#1F1E1D',
    'InputCaret':     '#1F1E1D',
    'CodeBg':         '#F3F1E9',
    'CodeFg':         '#1F1E1D',
    'ScrollThumb':    '#CFCCC2',
    'SelectedBg':     '#EFEDE5',
}

_DARK = {
    # Surfaces — anchored on Revit's dark chrome so the pane blends into it
    'AppBg':          '#262624',
    'ChatBg':         '#262624',
    'ComposerBg':     '#262624',
    'CardBg':         '#30302E',
    'CardBorder':     '#3E3D39',
    'Divider':        '#3A3935',
    'PaneEdge':       '#3E3D39',

    # Chat split
    'UserBubbleBg':   '#37352E',
    'UserBubbleText': '#F0EEE6',
    'BotText':        '#E3E1D9',

    # Type
    'Ink':            '#F5F4EE',
    'Muted':          '#A8A69F',
    'Faint':          '#7C7A74',

    # Accent + status
    'Accent':         '#D97757',
    'AccentSoft':     '#3A2E28',
    'Blue':           '#60A5FA',
    'Success':        '#34D399',
    'Danger':         '#F87171',
    'Amber':          '#F59E0B',

    # Controls
    'IconFg':         '#8F8D86',
    'IconFgHover':    '#E3E1D9',
    'IconHoverBg':    '#3A3935',
    'InputText':      '#F5F4EE',
    'InputCaret':     '#F5F4EE',
    'CodeBg':         '#1F1E1D',
    'CodeFg':         '#E3E1D9',
    'ScrollThumb':    '#4A4944',
    'SelectedBg':     '#3A3935',
}

_PALETTES = {'light': _LIGHT, 'dark': _DARK}

#: Resource-key prefix used for the brushes written by :func:`apply`.
RESOURCE_PREFIX = 'T3Theme'

_brush_cache = {}       # (theme, token) -> SolidColorBrush
_forced_theme = None    # set by force_theme() for previews/tests


# ─── Host theme detection ─────────────────────────────────────────────────────

def current_theme():
    """``'dark'`` or ``'light'`` — whichever Revit is currently using.

    ``UIThemeManager`` only exists on Revit 2024+. Older hosts have no dark
    mode at all, so 'light' is the correct answer there, not a fallback we
    are settling for.
    """
    if _forced_theme is not None:
        return _forced_theme
    try:
        import clr
        clr.AddReference('RevitAPIUI')
        import Autodesk.Revit.UI as _RUI
        mgr = getattr(_RUI, 'UIThemeManager', None)
        if mgr is None:
            return 'light'
        if 'dark' in str(mgr.CurrentTheme).lower():
            return 'dark'
    except Exception:
        pass
    return 'light'


def is_dark(theme=None):
    """True when the surface should paint itself dark."""
    return (theme or current_theme()) == 'dark'


def force_theme(theme):
    """Pin the theme ('light' / 'dark'), or pass None to follow Revit again."""
    global _forced_theme
    _forced_theme = theme if theme in _PALETTES else None


# ─── Token lookup ─────────────────────────────────────────────────────────────

def palette(theme=None):
    """The whole ``{token: '#RRGGBB'}`` table for a theme."""
    return _PALETTES.get(theme or current_theme(), _LIGHT)


def hex_of(token, theme=None):
    """Hex string for a token. Unknown tokens fall back to the ink colour so a
    typo shows up as visibly wrong text instead of an invisible crash."""
    pal = palette(theme)
    return pal.get(token, pal['Ink'])


def rgb(token, theme=None):
    """``(r, g, b)`` for a token — for code that builds ``Color.FromRgb``."""
    h = hex_of(token, theme).lstrip('#')
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def brush(token, theme=None):
    """A frozen ``SolidColorBrush`` for a token, cached per theme.

    Frozen so the same instance can be shared across every bubble without WPF
    having to track it per-element; cached so a 500-message conversation does
    not allocate 500 identical brushes.
    """
    theme = theme or current_theme()
    key = (theme, token)
    cached = _brush_cache.get(key)
    if cached is not None:
        return cached
    from System.Windows.Media import SolidColorBrush, Color
    r, g, b = rgb(token, theme)
    br = SolidColorBrush(Color.FromRgb(r, g, b))
    try:
        br.Freeze()
    except Exception:
        pass
    _brush_cache[key] = br
    return br


def color(token, theme=None):
    """A ``System.Windows.Media.Color`` for a token (shadows, gradients)."""
    from System.Windows.Media import Color
    r, g, b = rgb(token, theme)
    return Color.FromRgb(r, g, b)


# ─── Resource injection ───────────────────────────────────────────────────────

def apply(element, theme=None):
    """Write every token into ``element.Resources`` as ``T3Theme<Token>``.

    Call once after the XAML loads, and again whenever the host theme changes:
    WPF re-evaluates every ``DynamicResource`` bound to these keys, so the
    whole window re-skins without being rebuilt. Returns the theme applied, or
    None if the element has no usable resource dictionary.
    """
    theme = theme or current_theme()
    try:
        res = element.Resources
    except Exception:
        return None
    for token in palette(theme):
        try:
            res[RESOURCE_PREFIX + token] = brush(token, theme)
        except Exception:
            pass
    return theme
