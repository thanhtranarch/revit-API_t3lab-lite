# -*- coding: utf-8 -*-
"""
CPython 3 test harness for the T3Lab Assistant's chat-surface UI.

Run:  python3 dev/test_assistant_ui.py
Exit code 0 = all pass. No external test framework — plain asserts, mirroring
dev/test_assistant_routing.py / dev/test_tool_registry.py conventions.

T3LabAssistant.xaml is UI-LOCKED (see .claude/CLAUDE.md): it deliberately
deviates from the Lumina standard, so dev/audit_ui.py skips it entirely and
dev/sync_wpf_styles.py must never touch it. That left the one window with the
most bespoke styling in the codebase with nothing checking it at all — which is
how a vertical-only scrollbar style ended up being used for horizontal
scrollbars, and how two fully-hardcoded styles sat unreferenced for months.

These tests hold the *rules the lock exists to protect*, not the visual design:
colours come from theme tokens, both palettes stay in step, the scrollbar
handles both orientations, and the auto-synced region stays untouched.
"""
from __future__ import unicode_literals

import ast
import io
import os
import re
import sys
import xml.etree.ElementTree as ET

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXT = os.path.join(REPO, 'T3Lab.extension')
XAML = os.path.join(EXT, 'lib', 'GUI', 'Tools', 'T3LabAssistant.xaml')
THEME = os.path.join(EXT, 'lib', 'GUI', 'RevitTheme.py')
SCRIPT = os.path.join(EXT, 'T3Lab.tab', 'Support.panel',
                      'T3LabAssistant.pushbutton', 'script.py')

FAILURES = []


def check(label, ok, detail=None):
    line = '  {}  {}'.format('ok  ' if ok else 'FAIL', label)
    if not ok and detail is not None:
        line += '   {!r}'.format(detail)
    print(line)
    if not ok:
        FAILURES.append(label)


def _read(path):
    with io.open(path, 'r', encoding='utf-8') as f:
        return f.read()


XAML_SRC = _read(XAML)
THEME_SRC = _read(THEME)
SCRIPT_SRC = _read(SCRIPT)


def _palettes():
    out = {}
    for node in ast.walk(ast.parse(THEME_SRC)):
        if (isinstance(node, ast.Assign)
                and isinstance(node.targets[0], ast.Name)
                and node.targets[0].id in ('_LIGHT', '_DARK')):
            out[node.targets[0].id] = ast.literal_eval(node.value)
    return out


# ─── theme tokens ─────────────────────────────────────────────────────────────

def test_palettes_stay_in_step():
    """A token in one palette and not the other is a hole that only shows up
    in the theme nobody was looking at."""
    print('[theme: palettes]')
    pals = _palettes()
    check('both palettes found', set(pals) == {'_LIGHT', '_DARK'})
    light, dark = set(pals['_LIGHT']), set(pals['_DARK'])
    check('light has no token dark lacks', not (light - dark),
          sorted(light - dark))
    check('dark has no token light lacks', not (dark - light),
          sorted(dark - light))
    check('every value is a hex colour',
          all(re.match(r'^#[0-9A-Fa-f]{6}$', v)
              for p in pals.values() for v in p.values()))


def test_every_token_reference_resolves():
    """A {DynamicResource T3ThemeTypo} silently paints nothing."""
    print('[theme: references]')
    tokens = set(_palettes()['_LIGHT'])

    # T3Theme<Token> is a brush; T3Theme<Token>Color is the raw Color that
    # DropShadowEffect.Color needs (a brush bound there silently fails).
    xaml_refs = set(re.findall(r'DynamicResource\s+T3Theme(\w+?)(?:Color)?\}',
                               XAML_SRC))
    check('every XAML token exists',
          not (xaml_refs - tokens), sorted(xaml_refs - tokens))

    py_refs = set(re.findall(
        r"_bind_(?:fg|bg|border|stroke)\([^,]+,\s*'(\w+)'\)", SCRIPT_SRC))
    py_refs |= set(re.findall(r"_tb\('(\w+)'\)", SCRIPT_SRC))
    py_refs |= set(re.findall(r"_trgb\('(\w+)'\)", SCRIPT_SRC))
    py_refs |= set(re.findall(r"_theme_color\('(\w+)'\)", SCRIPT_SRC))
    check('every Python token exists',
          not (py_refs - tokens), sorted(py_refs - tokens))


def test_apply_publishes_brush_and_colour():
    """DropShadowEffect.Color is typed Color, not Brush. Publishing only
    brushes is why the popup shadows had to carry a hardcoded hex."""
    print('[theme: apply]')
    check('apply writes the brush key',
          "res[RESOURCE_PREFIX + token] = brush(token, theme)" in THEME_SRC)
    check('apply writes the colour key',
          "res[RESOURCE_PREFIX + token + 'Color'] = color(token, theme)"
          in THEME_SRC)


def test_rendered_messages_follow_the_theme():
    """_tb() returns a frozen brush, so anything painted with it keeps its
    render-time colours forever. Code-built chat content must bind instead."""
    print('[theme: live binding]')
    check('the binding helper exists', 'def _bind_theme(' in SCRIPT_SRC)
    check('it uses SetResourceReference',
          'SetResourceReference' in SCRIPT_SRC)
    check('it falls back to a plain assignment',
          re.search(r'setattr\(element, prop_name, _tb\(token\)\)',
                    SCRIPT_SRC) is not None)
    for name in ('_bind_fg', '_bind_bg', '_bind_border', '_bind_stroke'):
        # once in the themed branch, once in the no-RevitTheme fallback
        check('{} defined in both import branches'.format(name),
              SCRIPT_SRC.count('def {}('.format(name)) == 2,
              SCRIPT_SRC.count('def {}('.format(name)))

    # No frozen brush may be assigned straight onto a visual property again.
    leaked = re.findall(
        r'^\s*[\w\.]+\.(?:Foreground|Background|BorderBrush|Stroke)'
        r'\s*=\s*_tb\(', SCRIPT_SRC, re.MULTILINE)
    check('no frozen brush assigned to a visual property', not leaked, leaked)


# ─── scrollbar ────────────────────────────────────────────────────────────────

def test_thin_scrollbar_handles_both_orientations():
    """The thin style was authored vertical-only — fixed Width, a centred 5px
    column, and IsDirectionReversed hardcoded true. Applied to a horizontal
    scrollbar (every fenced code block in a reply has one) that renders a
    vertical sliver that drags the wrong way."""
    print('[scrollbar: orientation]')
    check('a horizontal thumb style exists',
          'x:Key="RevitThinThumbH"' in XAML_SRC)
    check('a horizontal template exists',
          'x:Key="RevitThinScrollBarTemplateH"' in XAML_SRC)
    check('the style switches on Orientation',
          re.search(r'<Trigger Property="Orientation" Value="Horizontal">',
                    XAML_SRC) is not None)
    check('the vertical track is reversed',
          re.search(r'RevitThinScrollBarTemplateV.*?IsDirectionReversed="[Tt]rue"',
                    XAML_SRC, re.S) is not None)
    check('the horizontal track is NOT reversed',
          re.search(r'RevitThinScrollBarTemplateH.*?IsDirectionReversed="[Ff]alse"',
                    XAML_SRC, re.S) is not None)
    check('the horizontal thumb is sized by height',
          re.search(r'RevitThinThumbH.*?Height="5".*?VerticalAlignment="Center"',
                    XAML_SRC, re.S) is not None)


def test_thin_scrollbar_covers_the_whole_window():
    """The shared block ships an IMPLICIT ScrollBar style with hardcoded greys.
    Anything not explicitly overridden inherited it — including chat_input's
    own scrollbar and the ComboBox dropdowns, which then painted near-black
    thumbs in Revit's dark theme."""
    print('[scrollbar: coverage]')
    root_scope = XAML_SRC.split('</Window.Resources>', 1)[-1]
    check('an implicit ScrollBar style is declared outside Window.Resources',
          re.search(r'<Style TargetType="\{x:Type ScrollBar\}"\s+'
                    r'BasedOn="\{StaticResource RevitThinScrollBar\}"\s*/>',
                    root_scope) is not None)
    check('the composer still scrolls its own draft',
          'chat_input.VerticalScrollBarVisibility' in SCRIPT_SRC)


# ─── the lock's own rules ─────────────────────────────────────────────────────

def test_shared_style_block_is_untouched():
    print('[lock: shared styles]')
    check('the sync markers are both present',
          'T3LAB SHARED STYLES v2' in XAML_SRC
          and 'END T3LAB SHARED STYLES' in XAML_SRC)


def test_no_stray_hardcoded_colours():
    """Every colour goes through a token, with three deliberate exceptions:
    the copyright amber (a Lumina constant, audited elsewhere) and the two
    modal scrims, which are translucent black and read correctly over both a
    light and a dark surface."""
    print('[lock: hardcoded colours]')
    body = XAML_SRC.split('END T3LAB SHARED STYLES', 1)[-1]
    allowed = {'#F59E0B', '#CC18181B', '#7F18181B'}
    # The token fallback table itself is where hexes are supposed to live.
    body = re.sub(r'<SolidColorBrush x:Key="T3Theme\w+"[^/]*/>', '', body)
    body = re.sub(r'<Color x:Key="T3Theme\w+">[^<]*</Color>', '', body)
    stray = sorted(set(re.findall(r'"(#[0-9A-Fa-f]{6,8})"', body)) - allowed)
    check('no hardcoded colour outside the token table', not stray, stray)


def test_no_dead_styles():
    """A style with no reference is either a leftover or a wiring bug. Some are
    referenced only from Python via FindResource, so both files are searched.

    Only styles this window declares for itself are policed: the auto-synced
    shared block is copied verbatim into all 53 tool windows and is expected to
    carry styles any individual one does not use.
    """
    print('[lock: dead styles]')
    own = XAML_SRC.split('END T3LAB SHARED STYLES', 1)[-1]
    dead = []
    for key in re.findall(r'<Style x:Key="(\w+)"', own):
        uses = (XAML_SRC.count('StaticResource ' + key + '}')
                + XAML_SRC.count('BasedOn="{StaticResource ' + key + '}')
                + SCRIPT_SRC.count("'" + key + "'")
                + SCRIPT_SRC.count('"' + key + '"'))
        if uses == 0:
            dead.append(key)
    check('every declared style is referenced', not dead, dead)


def test_popups_fit_a_docked_pane():
    """A pane docked beside the Project Browser is routinely 300-340px, and the
    skills popup alone declares MinWidth 340."""
    print('[layout: narrow pane]')
    check('popup widths are clamped to the pane',
          'def _fit_popups_to_width(' in SCRIPT_SRC)
    check('it runs on every width change, not just the compact transition',
          re.search(r'if e\.WidthChanged:.*?_fit_popups_to_width',
                    SCRIPT_SRC, re.S) is not None)
    check('static DPs are not read off the instance',
          'FrameworkElement.MinWidthProperty' in SCRIPT_SRC)


def test_xaml_is_well_formed():
    print('[xaml: syntax]')
    try:
        ET.parse(XAML)
        ok, detail = True, None
    except ET.ParseError as ex:
        ok, detail = False, str(ex)
    check('T3LabAssistant.xaml parses', ok, detail)


def main():
    print('')
    for fn in (test_palettes_stay_in_step,
               test_every_token_reference_resolves,
               test_apply_publishes_brush_and_colour,
               test_rendered_messages_follow_the_theme,
               test_thin_scrollbar_handles_both_orientations,
               test_thin_scrollbar_covers_the_whole_window,
               test_shared_style_block_is_untouched,
               test_no_stray_hardcoded_colours,
               test_no_dead_styles,
               test_popups_fit_a_docked_pane,
               test_xaml_is_well_formed):
        fn()
        print('')

    if FAILURES:
        print('{} FAILURE(S): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('All assistant UI tests passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
