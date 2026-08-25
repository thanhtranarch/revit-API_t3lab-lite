# Making a T3Lab window indistinguishable from Revit

Research notes + the design that came out of them. Scope: how far a pyRevit WPF
window can go toward being *the host's own UI* rather than a good imitation.

The short version: Revit ships its live WPF palette inside `UIFramework.dll`, and
a tool running in-process can read it. `lib/GUI/RevitTheme.py` now does. Colour
is no longer copied — it is borrowed.

---

## 1. What Revit actually exposes

Verified by scanning the installed products on this machine
(`C:\Program Files\Autodesk\Revit <ver>\`). Reproduce with:

```bash
python3 -c "import re;b=open(r'C:\Program Files\Autodesk\Revit 2026\UIFramework.dll','rb').read();print('\n'.join(sorted({m.decode() for m in re.findall(rb'[ -~]{5,}',b) if b'Theme' in m})[:80]))"
```

### 1.1 `UIFramework.ApplicationTheme` — the live dictionary

A public class exposing WPF `ResourceDictionary` instances:

| Member | What it is |
|--------|-----------|
| `CurrentTheme` | the dictionary Revit is painting itself with right now |
| `FullLightTheme` | the light dictionary, explicitly |
| `FullDarkTheme` | the dark dictionary, explicitly |
| `ThemeDictionary` | ribbon/tab theme |

Inside it, ~150 keys covering exactly the controls a tool dialog has —
`CheckBox*`, `ComboBox*`, `ListBox*`, `ListHeaderBackground*`,
`TextBoxDisabledBackground*`, `DisabledForeground*`, `Chrome*`,
`TextPromptForeground*`, `DropDownListBox*` — each as both a `…Color` and a
`…Brush`.

**Key naming is version-dependent.** This is the whole compatibility story:

| Revit | `ApplicationTheme` | `Full{Light,Dark}Theme` | Suffixed keys (`…Brush_LightTheme`) | Result |
|-------|--------------------|-------------------------|-------------------------------------|--------|
| 2023 | yes | dark only | **no** — one unsuffixed set | partial; product has no dark mode anyway |
| 2024 | not verifiable (no `UIFramework.dll` in the install here) | — | — | bridge finds nothing, curated palette stands |
| 2025 | yes | yes | yes | full |
| 2026 | yes | yes | yes | full |

Revit 2023 also ships the spelling `CheckBoxBackGroundBrush` — capital **G**.
The candidate lists in `RevitTheme._HOST_KEYS` carry both spellings for that
reason; do not "fix" it.

### 1.2 `Autodesk.Revit.UI.UIThemeManager` — the current theme

`CurrentTheme` (Light/Dark), Revit 2024+. Already used by `current_theme()`.
2026 also carries a `ThemeChanged` event (`ThemeChangedEventArgs`,
`UIApplicationThemeChangedEventHelper`) — see §4 for why we do not use it.

### 1.3 `Themes/LightTheme.xbel`, `Themes/DarkTheme.xbel`

Plain XML on disk, ~227 named colours each. **Not** the general UI: these drive
the Material Browser / Appearance & Texture editors (`RWUXThemeSU2015.dll`).
Useful reading, wrong source for a tool dialog.

### 1.4 `Autodesk.UI.Themes.dll`

`ThemeColorGenerator`, `IThemedImageUriParser`, `GetThemeHSBColors`,
`IsCurrentThemeDark`. This is about **recolouring icons** for the active theme,
not about surfaces. Relevant later if T3Lab icons need to invert in dark mode;
not relevant to the palette.

### 1.5 `Autodesk.UI.Windows.ChildWindow` — the dialog base class

The base class every modern Revit dialog derives from. What it does **not** contain
is the finding: the assembly has no `generic.baml`, no control template, nothing
that re-skins `Window`. What it does contain is Win32:

```
WS_CAPTION   WS_SYSMENU   WS_MINIMIZEBOX   WS_MAXIMIZEBOX
WS_EX_DLGMODALFRAME   WS_EX_CONTEXTHELP   SC_CONTEXTHELP   WM_HELP
```

plus `RemoveWindowIcon`, `IsContextHelpButtonShown`, `CenterOnWindow`,
`SaveWindowSettings` / `WindowBounds` / `WindowMaximized`.

**Revit dialogs use the native Windows title bar.** They tune it through extended
styles — `WS_EX_DLGMODALFRAME` is the classic way to drop the caption icon,
`WS_EX_CONTEXTHELP` adds the `?` button — and never draw it themselves.

**T3Lab windows nonetheless draw their own, and this is a deliberate reversal.**
A native caption is painted from the *Windows* theme, not Revit's. Run Revit dark
on a light Windows and you get a white title bar above a dark window — observed in
Revit 2026, and the reason the rule flipped. Re-tinting a native caption needs
`DwmSetWindowAttribute(DWMWA_USE_IMMERSIVE_DARK_MODE)`, a P/Invoke IronPython
cannot declare. Owning the caption and painting it from the `T3Theme*` tokens is
the only pure-WPF way to make it follow the host. The trade is real and worth
naming: the caption is no longer pixel-identical to Revit's, but it is the only
version of it that is never the wrong colour.

`SaveWindowSettings` is worth stealing separately: Revit dialogs remember their
size and position between sessions. T3Lab tools do not.

### 1.6 The anatomy of a real dialog

`UIFramework.PdfImport.PdfImportDialog` (Revit 2026) read straight out of its BAML
string table — the complete structure, in order:

```
Autodesk.UI.Windows.ChildWindow          x:Name=AcPdfImportDialog
    Title = {STR_PdfImport_WindowTitle}          <- a localized resource, not a literal
    FocusedElement = PageSelectorContentControl

    Resources.MergedDictionaries
        /UIFrameworkRes;component/Resources/Strings.xaml
        PdfImportResources.xaml
        /UIFramework;component/Themes/ChildWindowControlStyles.xaml

    DockPanel                                     Style = ChildWindow_Panel_Margin
      MainTop      Dock=Top
        MainTop_FilenameLabel                     Style = ChildWindow_LeftLabelStyle
        MainTop_FilenameInput                     Style = ChildWindow_ReadOnlyTextBoxStyle
      MainBottom   Dock=Bottom      (Rows=1)
        MainBottomLeftStackPanel    Horizontal
          MainBottom_HelpLinkLabel  -> LearnMoreHyperlink
        MainBottomRightStackPanel                 Style = ChildWindow_ButtonStackPanelStyle
          MainBottom_OkButton                     Style = ChildWindow_DefaultButtonStyle
                                                  Content = {STR_OK_BUTTON}
          MainBottom_CancelButton                 Style = ChildWindow_CancelButtonStyle
                                                  Content = {STR_CANCEL_BUTTON}
      PageSelectorContentControl                  the content, fills the rest
```

Three things to take from it:

1. **Bottom-left is the help link, bottom-right is OK/Cancel.** That is the whole
   footer. There is no status bar, no toolbar, no branding strip. The T3Lab
   copyright goes in the help-link slot — the least foreign home it has here.
2. **Button labels come from `Strings.xaml`** (`STR_OK_BUTTON`, `STR_CANCEL_BUTTON`),
   so they are localized with the product. A hardcoded English "OK" is a small
   tell on a non-English Revit.
3. **`ChildWindowControlStyles.xaml` is the dialog style kit.** Its keys:
   `ChildWindow_ButtonStyle`, `_DefaultButtonStyle`, `_CancelButtonStyle`,
   `_ButtonStackPanelStyle`, `_Button_Width`, `_Button_Height`, `_LabelStyle`,
   `_LeftLabelStyle`, `_TopLabelStyle`, `_TextBoxStyle`, `_ReadOnlyTextBoxStyle`,
   `_TextOnlyComboBoxStyle`, `_Panel_Margin`, `_Control_Margin`, and the
   `_*_Margin_Thickness` set.

### 1.7 Adopting Revit's styles wholesale — evaluated, not done

Merging `ChildWindowControlStyles.xaml` would make our buttons and fields literally
Revit's, not a match for them. It is not wired up, deliberately:

- The palette bridge copies **colour values**, which cannot fail worse than being
  the wrong grey. Copying a `Style` copies a `ControlTemplate`, and a template that
  resolves a `DynamicResource` we did not also import renders as an unpainted
  control. That is a regression this environment cannot test for.
- The key set is version-dependent in the same way §1.1 describes.

What is prepared is the knowledge, not the wiring: the key names above are the
complete list, and the showcase already matches Revit's metrics (75x23 buttons,
22px fields, 11px gutter) by value. Enabling this needs one Revit session to
verify, not new research.

### 1.7b The dialog background is NOT the Windows dialog grey

`themes/applicationthemebase.baml` (40 colours) carries the one value everything
else is measured against, and it is not `#F0F0F0`:

| Key | Light | Dark |
|-----|-------|------|
| `ApplicationClientAreaBackgroundColor` | **#EEEEEE** | **#2E3440** |
| `ChromeBackgroundColor` | #D9D9D9 | #808080 |
| `ChromeBorderColor` | **#C1C1C1** | **#515C6B** |
| `ChromeForegroundColor` | #303030 | #F0F0F0 |
| `RibbonPanelBackgroundColor` | #F5F5F5 | — |

This corrects a real composition bug. The palette used `#F0F0F0`, the Windows
dialog constant, guessed rather than read. Against a field at `#F5F5F5` that is a
five-unit difference — invisible. Revit's actual client area is `#EEEEEE`, which
makes the stack read as a proper ramp:

```
#EEEEEE  client area  ( = PushButtonFaceColor: buttons sit AT the face and are
                        separated by their border, not by tone )
#F5F5F5  fields       ( LIGHTER than the face, not inset — the opposite of the
                        Windows convention, and easy to get backwards )
#FFFFFF  lists/grids
```

Dark follows the same key and drops a stray neutral grey (`#3C3C3C`) out of what
is otherwise a blue-grey family, so the dark window is now internally consistent
too.

### 1.7c What still will not decode

`applicationthemefull{light,dark}.baml` — the dictionaries holding the suffixed
`*_LightTheme` keys — merge other dictionaries by pack URI and throw on `Source`
even with an `Application` created and `ResourceAssembly` set. Their *values*
appear to be the same ones reachable through `themecolor*` and
`applicationthemebase`, so nothing is known to be missing; but the suffixed key
set has not been read directly and that is the remaining gap.

### 1.8 `AdWindows.dll`

Ribbon chrome (`/AdWindows;component/Themes/Button.xaml`, `CommonConstants.xaml`).
Also `ComponentManager.ApplicationWindow`, which is what pyRevit already uses to
own our windows (§4).

---

## 2. The design

`RevitTheme.hex_of()` is the single choke point. Every other entry point —
`rgb`, `brush`, `color`, `apply` — routes through it, so one lookup change
reaches XAML `{DynamicResource}`, code-built elements and the window's resource
dictionary at once.

```
hex_of(token, theme)
    ├─ host bridge:  UIFramework.ApplicationTheme → "<key>_LightTheme" → "<key>"
    └─ fallback:     the curated _LIGHT / _DARK table
```

Deliberate choices:

- **Read values, do not merge the dictionary.** Merging Revit's dictionary into
  `Window.Resources` would also pull in whatever implicit styles it carries, on
  a version we cannot test. Reading named brushes out of it cannot bleed.
- **Only `SolidColorBrush` and raw `Color` are accepted.** Gradient brushes have
  no `.Color` and drop out on their own; a gradient cannot stand in for a flat
  token. Fully transparent values are rejected too.
- **`CurrentTheme` is only used when it *is* the theme being asked for.**
  Previewing light while Revit runs dark would otherwise return the dark
  dictionary and paint a light window in dark greys.
- **Unmapped tokens keep the curated value.** 24 of 60 tokens have a genuine
  counterpart in Revit's chrome; the other 36 (accent ramp, status colours,
  progress track, tab strip) are ours to choose and always were.
- **Kill switch.** `RevitTheme.use_host_palette(False)` reverts to the curated
  palette wholesale if a mapping ever reads wrong on an untested version.

### Mapping

Source of truth is `_HOST_KEYS` in `lib/GUI/RevitTheme.py`. Summary:

| T3Lab token | Revit key |
|-------------|-----------|
| `DialogBg`, `AppBg` | `ChromeBackgroundBrush` |
| `Ink`, `GroupHeaderFg` | `ChromeForegroundBrush` |
| `Faint` | `TextPromptForegroundBrush` |
| `FieldBg` | `ComboBoxBackgroundBrush` → `CheckBoxBack{g,G}roundBrush` |
| `FieldBorder` | `ComboBoxBorderBrush` → `TextBoxBorderBrush` → `CheckBoxBorderBrush` |
| `IconFg` | `ComboBoxArrowBrush` |
| `ControlHoverBg/Border` | `CheckBoxRollOver{Background,Border}Brush` |
| `ControlPressedBg/Border` | `CheckBoxPressed{Background,Border}Brush` |
| `ControlDisabledBg` | `TextBoxDisabledBackgroundBrush` |
| `ControlDisabledFg` | `DisabledForegroundBrush` |
| `CardBg`, `CardBorder` | `DropDownListBox{Background,Border}Brush` |
| `HeaderBg` | `ListHeaderBackgroundBrush` |
| `RowHoverBg`, `RowSelectedBg`, `SelectedBg` | `ListBox{RollOverBackground,SelectBackground}Brush` |
| `GridLine`, `Divider`, `GroupBorder` | `ListBoxSeparatorBrush` → `RibbonPanelSeparatorBrush` |
| `TabStripBg` | `TabBarBackgroundFloatingBrush` |

The check box is the state source because it is the one control Revit themes for
every state — rollover and pressed, background and border.

### Font

`RevitTheme.host_font()` returns `SystemFonts.MessageFontFamily` /
`MessageFontSize` — the shell's message font, which is what Windows dialogs and
Revit use. The XAML still declares `Segoe UI` 12 so it renders standalone, and
the Python overrides it at load. This is what makes the window follow the user's
text-scaling setting instead of pinning 12px.

---

## 3. Verifying it in Revit

The status bar shows the count live:

```
Ready · palette 24/60 from Revit
```

`0/60` means the bridge found nothing — expected on Revit 2024 here, and on any
host where `UIFramework` did not load. Full detail:

```python
from GUI import RevitTheme
for token, (value, source) in sorted(RevitTheme.host_report().items()):
    print("%-22s %-9s %s" % (token, value, source))
```

---

## 4. Levers that are *not* colour

Colour was the biggest gap, not the only one.

**Owner window — already solved.** pyRevit's `forms.WPFWindow.setup_owner()`
does `WindowInteropHelper(self).Owner = AdWindows.ComponentManager.ApplicationWindow`,
so a tool window minimises with Revit, stays above it, and never gets lost
behind it. Anything that builds its own `Window` instead of subclassing
`WPFWindow` must do this by hand.

**Theme-change event — deliberately not used.** `UIApplication.ThemeChanged`
exists on 2026 but not 2023, and needs a `UIApplication` the dialog does not
hold. The showcase re-syncs on `Window.Activated` instead: flip Revit's theme,
click back into the tool, it follows. Same outcome, no version dependency.

**The largest remaining gap: `MessageBox`.** 304 `MessageBox.Show` call sites
across 24 files. A WPF `MessageBox` is a raw Win32 dialog — it ignores Revit's
dark theme completely, carries no Revit branding, and is the most obviously
foreign thing left in the toolset. `Autodesk.Revit.UI.TaskDialog` is Revit's own
and is already used in 127 places. Converting the rest is mechanical but wide;
it has not been done and is not part of this change.

**DPI.** Revit 2023+ is per-monitor DPI aware and WPF follows the host process
manifest. Nothing to do.

**Docking.** A `DockablePane` beats a floating window for anything panel-shaped —
see `AssistantPaneControl.py`. Out of scope here.

---

## 5. What is still ours

The bridge does not make the design disappear. Layout, metrics, the accent ramp,
status colours and the copyright amber are still T3Lab decisions, documented in
`.claude/rules/ui-design-standard.md` under the Revit-native track. What changed
is that the *surfaces* underneath them are no longer a guess.

---

## 6. Weight and stability — what was measured

A tool window is parsed from XAML on **every open**; pyRevit does not pre-compile
BAML. So file weight is not cosmetic. Numbers below are from this checkout.

| Finding | Measured | Severity |
|---------|----------|----------|
| **Virtualising list inside a `ScrollViewer`** | **37 of 54 tools** | the freeze |
| Shared style block duplicated per file | 1060 lines × 53 files = **56,180 of 81,776 total XAML lines (69%)** | open cost |
| `Hanken Grotesk` / `Inter` / `Manrope` | **not installed** (732 fonts on this machine) and **not shipped** in the repo | font fallback on every text element |
| `DropShadowEffect` | 2 per tool XAML, both from the shared block (`WPF_styles.xaml` lines 223, 295) | minor |
| `AllowsTransparency="True"` | 4 files, all on `<Popup>`, **none on `<Window>`** | not a problem — checked and cleared |

### The freeze

A `DataGrid`/`ListBox`/`ListView` inside a `ScrollViewer` is measured with infinite
height. UI virtualisation switches off and WPF realises **every** row at open. On a
model with thousands of views or sheets that is an unbounded allocation, and it
matches the "tool hangs / crashes on big models" reports. Thirty-seven tools do it.

The fix per tool is small — let the list own its scrolling and give it a bounded
row — but it is a layout change in 37 files and needs Revit to verify. **Not done
here.** The showcase is the reference implementation of the correct shape.

### The font

`.claude/rules/ui-design-standard.md` has specified `Hanken Grotesk` for the Lumina
track, and that font exists on neither this machine nor in the repo. Every Lumina
tool has therefore been rendering in a WPF fallback face, not in its own standard.
The Revit-native track uses `Segoe UI`, a system font: correct by definition, free
to resolve, and what Revit uses.

### The duplicated style block

`dev/sync_wpf_styles.py` copies a 1060-line block into all 53 tool XAMLs so they
stay identical. It succeeds at consistency and costs every tool 1060 lines of parse
for styles it mostly does not use. The alternative is one
`<ResourceDictionary Source=".../WPF_styles.xaml"/>` merge per tool: WPF caches a
merged dictionary per absolute URI, so the second tool opened in a session pays
nothing. That is a 53-file change and is **not done here**.

### What the showcase now costs

| | before | after |
|---|---|---|
| lines | 1423 | **721** |
| `ControlTemplate` | 30 | **19** |
| `Effect` | 0 | 0 |
| `WindowChrome` | yes | **none** — native caption |
| glyph-font icons | 13 | **0** (`Path` geometry instead) |
| list virtualisation | implicit | explicit, and not inside a `ScrollViewer` |

---

## 7. Previewing a tool XAML without Revit

`dev/preview_xaml.ps1` loads a tool XAML into WPF the way pyRevit does, applies
the palette, and writes a light and a dark PNG:

```bash
python3 dev/export_palette.py palette.json
powershell -STA -NoProfile -File dev/preview_xaml.ps1 -Xaml .claude/standard/UIStandardShowcase.xaml -Palette palette.json -OutDir .
```

It renders the **client area only** — the native title bar belongs to Windows —
and strips event handlers before parsing, so nothing is interactive. Off-screen
layout needs several measure/arrange passes before `DataGrid` star columns
settle; that is a property of the harness, not of the XAML, and a collapsed
`Width="*"` column in a single-pass render is not a bug in the file.

It has already earned its place: it caught a `Width="*"` column reading as
collapsed and a checked check box rendering empty, both before Revit was opened.

## 8. Control shapes verified against Revit itself

Corrections made after comparing against screenshots of the running product,
each of which overturned an assumption:

| Control | What was assumed | What Revit does |
|---------|------------------|-----------------|
| Footer buttons | `Apply · OK · Cancel`, square | **`OK · Cancel · Apply · Help`**, 4px radius; `Apply` starts disabled |
| Default button | outline — *correct* | outline, accent border. A filled button would have been wrong |
| Check box | white box, coloured tick | **accent-filled box, white tick**, 16px, 3px radius — and the checked state is an **image** (`CheckBoxCheckedImage_*`), not a brush |
| Search box | `Search:` label + plain field | **no label**; magnifier inside the field, prompt text carries the meaning, with a `Type [All v]` filter beside it |
| Progress bar | flat accent fill | **Aero green vertical gradient** — the stock Windows control |
| Caption icon | tool logo, via pyRevit `setup_icon()` | **no icon**. `ChildWindow` exposes `RemoveWindowIcon` and sets `WS_EX_DLGMODALFRAME` |

The check box is the important one architecturally: because Revit draws that
state from an image, matching it by picking a hex was never going to land.
`RevitTheme.host_resource('CheckBoxCheckedImage')` takes the host's own artwork
and publishes it as `T3ThemeCheckedGlyph`; the drawn stand-in only appears when
the host has none.

The caption icon is removed with `WindowStyle="ToolWindow"`, which also drops
the min/max buttons. That approach was ABANDONED: T3Lab tool windows keep all
three caption buttons by decision, and the native caption could not follow
Revit's theme anyway. Kept here only as a record of what was tried. What Revit's own modal
dialogs show. The fully faithful route is the Win32 one Revit takes
(`WS_EX_DLGMODALFRAME` + `WM_SETICON` on the HWND), which keeps a full-height
caption; IronPython cannot declare that P/Invoke directly, so it is not wired up.
