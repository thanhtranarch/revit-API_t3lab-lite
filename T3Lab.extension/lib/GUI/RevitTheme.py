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
   dialog, so it does not use the Lumina tool palette in
   ``.claude/rules/ui-design-standard.md``; it uses **Revit's own UI greys**
   plus the chat split (filled user bubble · bubble-less assistant text). That
   deviation is deliberate and scoped to this one surface — the copyright amber
   ``#F59E0B`` and the shared button styles are untouched, so the Lumina audit
   still passes. Everything else routes through the tokens below rather than
   being re-typed as a hex literal in 70 different places.

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
# Both palettes are Revit's own UI greys, not a chat product's palette. The
# assistant sits in a dockable pane between the Project Browser and Properties,
# so anything with a hue of its own — the warm "paper" cream this used to be,
# or a saturated brand accent — reads as a foreign app pasted over the host.
#
# Anchors, per theme, taken from Revit's own chrome:
#
#   Light   palette/ribbon chrome #F0F0F0 · list & field surfaces #FFFFFF
#           separators #D9D9D9 · control borders #C4C4C4 · label text black
#   Dark    (Revit 2024+) chrome #2E2E2E · palette content #383838
#           fields/menus #414141 · separators #4A4A4A · label text #E4E4E4
#
# The one colour in the window is Autodesk blue — the same blue Revit uses for
# its own highlights, so the send button, the active chip and the selection
# count stay identifiable without importing a second brand. It is tuned per
# theme (see 'Accent' in each palette), not one fixed hex.

_LIGHT = {
    # Surfaces — Revit light: grey chrome, white content
    'AppBg':          '#F0F0F0',   # pane backdrop = Revit's palette grey
    'ChatBg':         '#FFFFFF',   # transcript = a content surface (browser/canvas)
    'ComposerBg':     '#F0F0F0',   # composer strip reads as palette chrome
    'CardBg':         '#FFFFFF',   # composer card, popups, overlays = Revit fields
    'CardBorder':     '#C4C4C4',   # Revit control border
    'Divider':        '#D9D9D9',   # separators, footer rule
    'PaneEdge':       '#A6A6A6',   # outer edge when floating

    # Chat split — one speaker is filled, the other sits on the surface
    'UserBubbleBg':   '#DBE8F6',   # Revit's selected-row blue tint
    'UserBubbleText': '#000000',
    'BotText':        '#1C1C1C',   # assistant reply: text on the surface, no bubble

    # Type — Revit labels are black, not a warm near-black
    'Ink':            '#000000',   # headings, strong text
    'Muted':          '#5C5C5C',   # secondary text
    'Faint':          '#858585',   # placeholders, disclaimers, hints

    # Accent + status
    # Autodesk blue, one step deeper than the #0696D7 brand value: the accent
    # also paints 11.5px text (selection count, active chips), and the bright
    # value only reaches 2.9:1 on the #F0F0F0 composer bar. This clears 4.5:1
    # there and takes the white send arrow from 3.3:1 to 5.1:1.
    'Accent':         '#0A6FB3',   # send button, active states
    'AccentHover':    '#095C93',   # pressed/hover
    'AccentSoft':     '#D6E9F7',   # blue tint fill
    'Blue':           '#0A6FB3',   # links / info — Revit has one blue, not two
    'Success':        '#0B7A4A',
    'Danger':         '#C42B1C',
    'Amber':          '#F59E0B',   # copyright (Lumina standard, do not retint)

    # Controls
    'IconFg':         '#5C5C5C',   # line icons, Revit's own toolbar grey
    'IconFgHover':    '#1A1A1A',
    'IconHoverBg':    '#E4E4E4',
    'InputText':      '#000000',
    'InputCaret':     '#000000',
    'CodeBg':         '#F5F5F5',
    'CodeFg':         '#1C1C1C',
    'ScrollThumb':    '#C1C1C1',
    'SelectedBg':     '#E6E6E6',
    # Drop-shadow colour for the composer card and the popups. A shadow is what
    # lifts a popup off the surface behind it, and on both themes that means
    # black — what changes is the Opacity at each use, not the hue.
    'Shadow':         '#000000',
}

_DARK = {
    # Surfaces — Revit 2024+ dark chrome, content one step lighter than chrome
    'AppBg':          '#2E2E2E',
    'ChatBg':         '#383838',
    'ComposerBg':     '#2E2E2E',
    'CardBg':         '#414141',
    'CardBorder':     '#575757',
    'Divider':        '#4A4A4A',
    'PaneEdge':       '#1F1F1F',

    # Chat split
    'UserBubbleBg':   '#33475A',   # the same selected-row blue, dark theme
    'UserBubbleText': '#F5F5F5',
    'BotText':        '#E4E4E4',

    # Type — a step lighter than the light theme's mirror image: the surfaces
    # they sit on (#414141 cards) are lighter than #2E2E2E, and #A6A6A6 /
    # #828282 fell under 4.5:1 and 3:1 there.
    'Ink':            '#FFFFFF',
    'Muted':          '#B0B0B0',
    'Faint':          '#949494',

    # Accent + status — the brand value works as-is here: on dark chrome it is
    # the light end that needs care (a brighter fill would swallow the white
    # send arrow), so this keeps #0696D7 rather than lightening it further.
    'Accent':         '#0696D7',
    'AccentHover':    '#3AAEE6',   # lighter on dark: darkening vanishes there
    'AccentSoft':     '#1E3A4A',
    'Blue':           '#4FB3E8',   # link blue has to lift off #383838
    'Success':        '#3FBF87',
    'Danger':         '#F08880',   # error TEXT on #383838 — it is never a fill
    'Amber':          '#F59E0B',

    # Controls
    'IconFg':         '#A0A0A0',
    'IconFgHover':    '#FFFFFF',
    'IconHoverBg':    '#4A4A4A',
    'InputText':      '#F5F5F5',
    'InputCaret':     '#F5F5F5',
    'CodeBg':         '#262626',
    'CodeFg':         '#E4E4E4',
    'ScrollThumb':    '#5A5A5A',
    'SelectedBg':     '#454545',
    'Shadow':         '#000000',
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
    """Write every token into ``element.Resources``.

    Two keys per token:

    ``T3Theme<Token>``
        a ``SolidColorBrush`` — what almost everything binds to.
    ``T3Theme<Token>Color``
        the raw ``Color``. Some properties are typed ``Color``, not ``Brush``
        — ``DropShadowEffect.Color`` above all — and a brush bound there
        silently fails to apply. Without this the popup and composer shadows
        had to carry a hardcoded hex, which meant they kept their light-theme
        value on Revit's dark chrome and stopped separating the popup from
        what was behind it.

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
        try:
            res[RESOURCE_PREFIX + token + 'Color'] = color(token, theme)
        except Exception:
            pass
    return theme
