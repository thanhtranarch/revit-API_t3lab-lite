# UI Design Standard

Two tracks live in this repo. Read this section before touching any XAML — picking
the wrong one is the most common way a change gets reverted.

| Track | What it is | Reference file | Who uses it today |
|-------|-----------|----------------|-------------------|
| **Revit-native** | Revit's *own* dialog chrome: Segoe UI 12, square 1px hairlines, 23px buttons, colours resolved live from `GUI/RevitTheme.py` so the window follows Revit Light/Dark. | `.claude/standard/UIStandardShowcase.xaml` | `T3LabAssistant.xaml`; the target look for new tools |
| **Lumina** | The Deep Slate + Refined Blue brand system (rounded cards, Hanken Grotesk, filled buttons) described in the rest of this document. | `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` | ~50 existing tool XAMLs |

**The showcase is the Revit-native reference.** It was migrated on 2026-08-20 and no
longer demonstrates the Lumina palette. The Lumina rules below still govern every tool
XAML that has not been migrated, and `dev/audit_ui.py` still enforces them — it just
exempts the Revit-native files (`REVIT_THEMED` in that script) from the three
Lumina-only checks: the Segoe UI ban, the Hanken Grotesk root font, and the shared
style block.

Everything **outside** those three — copyright placement, `WindowChrome` form, Segoe
MDL2 chrome glyphs, the dot-notation `<Grid.RowDefinition/>` crash — applies to **both
tracks**, with no exceptions.

## Revit-Native Track

Use this for any new tool window. The goal is that a T3Lab dialog opened next to
"Manage Links" or "Visibility/Graphics" is indistinguishable from the host's own UI.

### Colour — never a hex literal, and preferably not even ours

Every colour comes from `T3Lab.extension/lib/GUI/RevitTheme.py`, bound as
`{DynamicResource T3Theme<Token>}`. The module reads
`Autodesk.Revit.UI.UIThemeManager.CurrentTheme` and carries a full Light and Dark
palette, so a themed window re-skins itself with no reload.

For 24 of the 60 tokens the value is not ours at all: the module reads the brush
**Revit itself is painting with**, out of `UIFramework.ApplicationTheme`, so a
T3Lab field and a Revit field agree to the pixel and keep agreeing across product
updates. Coverage depends on the host version (full on 2025/2026, partial on
2023); anything the host does not expose keeps the curated value. The status bar
of the showcase reports the live count. Details, mapping table and the kill
switch: **`docs/revit-native-ui.md`** — read it before changing `_HOST_KEYS`.

The Python side must call it once after loading the XAML, or every token falls back
to the literal light-theme brushes declared at the top of `Window.Resources`:

```python
from GUI import RevitTheme as _theme
...
forms.WPFWindow.__init__(self, XAML_FILE)
_theme.apply(self)          # writes T3Theme* brushes into Window.Resources
```

Tokens for tool dialogs (see the module for the full table and the exact values):

| Group | Tokens |
|-------|--------|
| Surfaces | `DialogBg`, `ChatBg`, `CardBg`, `CardBorder`, `Divider`, `PaneEdge` |
| Type | `Ink`, `Muted`, `Faint` |
| Push buttons | `ControlBg`, `ControlBorder`, `ControlHoverBg`, `ControlHoverBorder`, `ControlPressedBg`, `ControlPressedBorder`, `ControlDisabledBg/Border/Fg` |
| Fields | `FieldBg`, `FieldBorder`, `FieldFocusBorder` |
| Lists | `HeaderBg`, `HeaderFg`, `RowAltBg`, `RowHoverBg`, `RowSelectedBg`, `GridLine` |
| Tabs | `TabStripBg`, `TabActiveBg`, `TabHoverBg` |
| Accent / status | `Accent`, `AccentHover`, `AccentSoft`, `Success`, `Warning`, `Danger`, and the `*Soft` fills |

The only hex literal allowed in a Revit-native XAML is `#F59E0B`, the copyright
amber, which the T3Lab standard fixes.

### The window itself — we DO draw the caption

This is the rule that decides whether a dialog reads as Revit or as a redesign.

Revit's dialogs derive from `Autodesk.UI.Windows.ChildWindow`. That class does
**not** re-template `Window`: it adjusts Win32 styles (`WS_CAPTION`, `WS_SYSMENU`,
`WS_EX_DLGMODALFRAME`, `WS_EX_CONTEXTHELP`) and lets **Windows draw the title bar**.
There is no custom caption, no hand-drawn minimise/maximise/close, no glyph font in
the chrome. See `docs/revit-native-ui.md` §1 for how this was established.

**But we cannot use the native caption, and the reason is dark mode.** Windows
draws the caption from the *Windows* theme. On Revit-dark with Windows-light that
is a white title bar glued to the top of a dark window — the reported bug. There
is no pure-WPF way to re-tint a native caption; the Win32 route
(`DwmSetWindowAttribute` / `WS_EX_DLGMODALFRAME`) needs a P/Invoke IronPython
cannot declare.

So a T3Lab window owns its caption, and pays for that by having to paint it from
the same tokens as the body:

```xml
<Window WindowStyle="None" ResizeMode="CanResizeWithGrip" ShowInTaskbar="False">
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="56" ResizeBorderThickness="6"
                      GlassFrameThickness="0" CornerRadius="0"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>
```

Caption content is title + subtitle on the left, `Segoe MDL2` min/max/close on the
right (`&#xE921;` / `&#xE922;` / `&#xE8BB;`), close hovering `#E81123`. Every
brush is a `T3Theme*` token, so the caption follows Revit rather than Windows.
That is the whole point of owning it.

### Layout skeleton

Taken from a real dialog (`UIFramework.PdfImport.PdfImportDialog`, Revit 2026):

```
DockPanel  Margin="11"                       (Revit: ChildWindow_Panel_Margin)
├ Dock=Bottom   status / copyright left  ·  Apply · OK · Cancel right
├ Dock=Bottom   progress row (collapsed while idle)
├ Dock=Left     GroupBox stack — the options pane
└ (fill)        list pane: filter row · DataGrid · action strip under it
```

Buttons live **under the list** (the Manage Links idiom). Revit dialogs have no
icon toolbar; if a window has one, it is not a Revit dialog.

The footer button row is **OK · Cancel · Apply · Help**, in that order, right
aligned. `Apply` starts disabled and enables once something has actually
changed; that greyed-out Apply is part of how a Revit dialog reads. `Help` is a
button here, not a hyperlink.

### How a Revit dialog separates itself

Checked against **Apply View Template**, which is the ground truth here. The
separation is carried by **white surfaces and hairlines**, not by a stack of
greys:

| Part | What Revit does | Token |
|------|-----------------|-------|
| caption | **white — LIGHTER than the body.** Not a darker chrome band. | `ChromeBg` |
| body | one light grey, and that is the whole background | `DialogBg` |
| group box | **no fill**; a thin border with the label sitting on it | `GroupBorder` |
| lists, fields | **white** against the grey body — nearly all the separation | `ChatBg`, `FieldBg` |
| footer | no strip of its own, only a rule above the button row | `Divider` |
| selection | a **solid accent row with white text**, never a grey tint | `RowSelectedBg`, `OnAccent` |

**Pick the LIGHT members of Revit's set, not the heavy ones.** The palette holds
several greys for each role and the first pass took the darkest of each —
`ApplicationClientAreaBackgroundColor` #EEEEEE for the body,
`ScheduleHeaderColor` #E1E1E1 for the grid header, `AGridButtonBorderColor`
#BBBBBB for group frames, `ScheduleHeaderBorderColor` #CECECE for grid lines.
Those are tuned for the main window's chrome beside a dark drawing canvas; in a
dialog made mostly of white surfaces they read as grime. The lighter members of
the same palette are the right ones here:

| Role | Heavy (wrong here) | Light (use this) |
|------|--------------------|------------------|
| body | `#EEEEEE` ApplicationClientArea | **`#F5F5F5`** RibbonPanelBackgroundColor |
| field, button face | `#F5F5F5` / `#EEEEEE` | **`#FFFFFF`** ComboItemBgColor |
| grid header | `#E1E1E1` ScheduleHeaderColor | **`#F0F0F0`** DisabledBackgroundColor |
| grid lines, dividers | `#CECECE` | **`#ECECEC`** AGridBorderColor |
| group frame | `#BBBBBB` | **`#E5E5E5`** ComboDisabledBorderColor |
| card / pane edge | `#CCCCCC` / `#C1C1C1` | **`#DEDEDE`** HigOptionBarBtnBorderDisabled |

Dark is unaffected — the heavy values are correct there, because the surfaces
they sit on are dark.

Two traps, both fallen into here before the dialog was checked:

- `ChromeBackgroundColor` (`#D9D9D9`) is the **main window's** chrome. Using it
  for a dialog caption paints a dark band Revit does not have.
- `ScheduleHeaderSelectedColor` (`#D4D4D4`) is the schedule **header**, not a
  selected row. A selected row is the accent.

And one that follows from the accent row: a status column coloured red or amber
must yield to the selection, or it becomes red-on-blue. Put the `IsSelected`
trigger last so it wins.

Note also that the light stack is **not** the Windows convention: fields
(`#F5F5F5`) are *lighter* than the face they sit on, not inset. Getting that
backwards is what makes a Revit-native window look subtly wrong.

### Metrics

| Element | Value |
|---------|-------|
| Font | `Segoe UI` 12 on the `<Window>` root, overridden at load by `RevitTheme.host_font()` |
| Title bar | native — drawn by Windows, not by us |
| Gutter | one 11px margin, everywhere |
| Push button | 23px tall, 75px min width, **1px** border (the default action differs only in border *colour*, never thickness), `Padding="10,0"`, 4px apart |
| Text box / combo box | 22px tall, 1px border, `Padding="4,0"` |
| Check box / radio | 13 × 13, 2px radius. Checked is an accent-filled box with a white tick — and in Revit that state is an **image** (`CheckBoxCheckedImage_*`), which `RevitTheme` borrows. 16px is the Windows 11 size and reads oversized beside a 12px label. |
| List row | 22px · column header 23px |
| Corner radius | Exactly three values. **0** — surfaces, fields, group boxes, lists, popups, progress groove, window frame. **4** — anything clickable: push buttons, caption buttons, rail tiles; that is what Windows 11 draws, verified against Revit's own OK/Cancel/Apply/Help row. **2** — the check box only, because a 13px control cannot carry 4. The scroll thumb is rounded by nature. A fourth value means you are designing, not matching. |
| Container | `GroupBox` with a hairline frame and a notched caption — not a shadowed card |

### Things that break the illusion

- **A caption of our own.** The single loudest tell. Windows draws it.
- Rounded *surfaces*, drop shadows, or a colour-filled "primary" button.
  Windows marks the default action with a **border**, never a brand fill —
  Revit's own OK button is an outline. (Buttons themselves do carry the
  platform's 4px radius; cards and panels do not.)
- Gradients **we invented**. The progress bar is the exception and is not an
  exception at all: Revit uses the stock Windows control, which is the Aero
  **green** vertical gradient. A flat accent-coloured progress bar is the
  wrong one. Reproduce what the host draws; never add a gradient it lacks.
- An icon toolbar across the top. No Revit dialog has one.
- A display font, a type scale, or uppercase tracked micro-labels. Revit labels
  everything at one size and one weight; emphasis is colour.
- iOS-style toggle switches. Revit has none — use a check box.
- Filled rounded status pills in a table. Revit's own Manage Links and warning
  lists use plain text, coloured only when the row is a problem.

Aesthetics are not the casualty here. They come from the 11px gutter holding
everywhere, one type size, a single white list surface reading against the grey
dialog face, and colour spent only where it means something. That is the same
discipline that makes Revit's own dialogs look composed.

### Weight and stability

A tool window is parsed from XAML on **every open** — pyRevit does not
pre-compile BAML. Weight is a feature, and these rules are load-bearing:

- **No `Effect`.** `DropShadowEffect` forces an intermediate render surface per
  element. A Revit dialog has none.
- **Never wrap a `DataGrid` / `ListBox` / `ListView` in a `ScrollViewer`.** The
  outer viewer hands the list infinite height, UI virtualisation switches off, and
  WPF realises every row at once. On a real model that is the freeze — the single
  most common hang in this codebase. Let the list scroll itself, and set
  `EnableRowVirtualization`, `VirtualizingPanel.VirtualizationMode="Recycling"`,
  `ScrollViewer.CanContentScroll="True"` explicitly.
- **`Path` beats a glyph font** for arrows and ticks: no font resolution, no
  fallback scan, and it inherits the theme brush directly.
- **Only template a control when the dark theme needs it.** Every
  `ControlTemplate` is object graph rebuilt at each open.

The reference showcase is the yardstick: 721 lines, 19 templates, zero effects.

---

## Lumina Track (existing tools)

> Lumina replaces the old navy "Kinetix" palette and the teal/amber "Terra (v2)" palette.
> If you see old navy/teal/amber colors (#083D56, #0F766E, #F59E0B etc.) in a file
> (except where amber is specifically allowed for warning highlights/copyright),
> it has not been migrated — flag it.

Everything from here down describes the **Lumina** system (Deep Slate + Refined Blue).
It remains binding for the ~50 tool XAMLs that still use it. Since the showcase moved
to the Revit-native track, the live reference for Lumina styles is the master shared
style block, not the showcase.

## Design Reference Files
- **Revit-native reference**: `.claude/standard/UIStandardShowcase.xaml` (+ the mirrored copy in `T3Lab.extension/lib/GUI/Tools/`)
- **Revit theme palette**: `T3Lab.extension/lib/GUI/RevitTheme.py`
- **Master shared styles**: `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml`
- **Style sync script**: `dev/sync_wpf_styles.py` (see "Shared Styles" below)
- **All XAML files**: `T3Lab.extension/lib/GUI/Tools/`
- **XAML templates doc**: `.claude/docs/wpf-window-templates.md`
- **Python WPF pattern**: `.claude/docs/python-wpf-pattern.md`

## Two XAML Variants

The codebase has **two distinct XAML root patterns**. Pick the one that matches the use case — do not mix them inside a single file.

### Variant A — Standard Tool Window (default)
- Root is `<Window>` and the XAML defines its own chrome and title bar.
- Loaded from Python via `forms.WPFWindow.__init__(self, xaml_path)`.
- **Use this for every new tool**, on either track — the `<Window>` root and the chrome pattern are shared. Examples: `UIStandardShowcase.xaml` (Revit-native), `AutoDimension.xaml`, `DWGManagement.xaml` (Lumina).

### Variant B — Modal Dialog Content (rare, 3 existing files)
- Root is `<Grid>`; the XAML provides only the content. The hosting Python class creates a `Window` with `WindowStyle=NoStyle`, `AllowsTransparency=True`, then assigns the parsed Grid to `self.Content`.
- **Do not create new Variant B files** unless explicitly building a borderless modal hosted inside another window. Existing files: `FamilyLoader.xaml`, `FamilyLoaderCloud.xaml`, `ParameterSelector.xaml`.
- Variant B files use the same Lumina palette and the same shared style block (inside `Grid.Resources`).

The rest of this document targets **Variant A** unless otherwise noted.

## Color Palette (Lumina)

| Token            | Hex       | Usage                                                  |
|------------------|-----------|--------------------------------------------------------|
| Primary (slate)  | `#0F172A` | Brand title, primary buttons, window control glyphs    |
| Accent (blue)    | `#3B82F6` | Secondary active states, focus borders, primary actions|
| Success green    | `#10B981` | Success/confirm buttons (hover: `#059669`)              |
| Danger red       | `#EF4444` | Delete/destructive buttons (hover: `#DC2626`)           |
| Amber highlight  | `#F59E0B` | Warning labels, highlights, copyright (footer left)    |
| Ink text         | `#0F172A` | Headings, labels, main text                            |
| Muted text       | `#64748B` | Subtitles, secondary labels, disabled state text       |
| Faint text       | `#94A3B8` | Placeholders, disabled text                            |
| Surface          | `#F8FAFC` | Backgrounds, inputs background, status bar bg          |
| Surface hover    | `#F1F5F9` | Secondary button hover, menu hover                     |
| Surface press    | `#E2E8F0` | Button pressed bg, active selection tint               |
| Border           | `#E2E8F0` | Dividers, separators, connector lines                  |
| Input border     | `#CBD5E1` | TextBox / ComboBox / search-box borders                 |
| Selected tint    | `#E2E8F0` | DataGrid row selected                                   |
| Disabled bg      | `#E2E8F0` | Disabled button background                              |

### Migration map (Terra v2 → Lumina)

| Old (Terra v2)             | New (Lumina)                                              |
|----------------------------|-----------------------------------------------------------|
| `#0F766E` primary          | `#0F172A` (deep slate)                                    |
| `#115E59` primary hover    | `#1E293B`                                                 |
| `#0B4F4A` primary pressed  | `#0F172A`                                                 |
| `#F6F8F8` surface          | `#F8FAFC`                                                 |
| `#E6EDEC` surface hover    | `#F1F5F9`                                                 |
| `#D9EBE8` selected tint    | `#E2E8F0`                                                 |
| `#DDE5E7` border           | `#E2E8F0`                                                 |
| `#C7D2D4` input border     | `#CBD5E1`                                                 |
| `#15803D` success green    | `#10B981` (emerald)                                       |
| `#166534` success hover    | `#059669`                                                 |
| `#DC2626` danger red       | `#EF4444` (rose)                                          |
| `#B91C1C` danger hover     | `#DC2626`                                                 |
| `#F59E0B` (AccentButton)   | `#3B82F6` (refined blue AccentButton)                     |
| `#AEBFBC` disabled bg      | `#E2E8F0`                                                 |

## Typography

- **All text**: `Hanken Grotesk` (body, labels, headings; `Inter` is allowed as a secondary/fallback font). Revit-native files use `Segoe UI` instead — see the Revit-Native Track above.
- **Variant A** (`<Window>` root): add `FontFamily="Hanken Grotesk"` on the `<Window>` element — valid because `Window` inherits from `Control`.
- **Variant B** (`<Grid>` root): do **NOT** add `FontFamily` to the root `<Grid>` — `Grid` is a `Panel`, not a `Control`, and setting it causes a XAML crash. Set font per-element or via `Style` in resources instead.
- **Never** use `Manrope` or `Segoe UI` for body text. `Segoe MDL2 Assets` is allowed **only** for icon glyphs.

## Window Structure (every tool)

1. `<Window>` with `Background="White"`, `ResizeMode="CanResizeWithGrip"`, `FontFamily="Hanken Grotesk"` (or `Inter`)
2. **WindowChrome — multi-line form is required.** All five attributes below must be present together:
   ```xml
   <WindowChrome.WindowChrome>
       <WindowChrome CaptionHeight="64"
                     ResizeBorderThickness="5"
                     GlassFrameThickness="0"
                     CornerRadius="8"
                     UseAeroCaptionButtons="False"/>
   </WindowChrome.WindowChrome>
   ```
3. **Title bar** (Row 0, Height=64): White bg
   - Left: Tool Name (15px Bold `#0F172A`) + Subtitle (12.5px Italic `#64748B`), stacked vertically. No separate "T3Lab" wordmark — confirmed converged codebase pattern (0/50 files carry a standalone `T3Lab` label in the title bar).
   - Right: Min/Max/Close buttons (Segoe MDL2 Assets glyphs — see table below)
   - Bottom: Border 1px `#E2E8F0`
4. **Content area** — tool-specific. 16px horizontal margins; 10–12px vertical rhythm between sections.
5. **Status bar** (last row): `Background="#F8FAFC"`, `BorderBrush="#E2E8F0"`, `BorderThickness="0,1,0,0"`, `Padding="14,8"`
6. **Copyright** — inside the status bar (footer), **left-most element** — see "Copyright Notice" below

## Shared Styles (single source of truth)

Button styles are **not** hand-edited per file. The master copy lives in
`T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml`, and every tool XAML carries
an identical copy of the block between these markers inside `Window.Resources`
(Variant A) or `Grid.Resources` (Variant B):

```xml
<!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT (edit lib/GUI/Resources/WPF_styles.xaml, then run dev/sync_wpf_styles.py) ═══ -->
... shared styles ...
<!-- ═══ END T3LAB SHARED STYLES ═══ -->
```

- To change a button style: edit the master file, run `python dev/sync_wpf_styles.py` to propagate, or `python dev/sync_wpf_styles.py --check` to verify all files match.
- Window-specific styles go **outside** the markers.
- Never edit inside the markers by hand.

### Shared style keys

| Style Key         | Background  | Foreground | Hover Bg   | Pressed Bg | Border     | Padding | FontSize | Radius |
|-------------------|-------------|------------|------------|------------|------------|---------|----------|--------|
| `PrimaryButton`   | `#0F172A`   | White      | `#1E293B`  | `#0F172A`  | none       | `14,7`  | 12       | 6      |
| `SecondaryButton` | Transparent | `#0F172A`  | `#F1F5F9`  | `#E2E8F0`  | `#0F172A`  | `14,7`  | 12       | 6      |
| `SuccessButton`   | `#10B981`   | White      | `#059669`  | `#047857`  | none       | `14,7`  | 12       | 6      |
| `DangerButton`    | `#EF4444`   | White      | `#DC2626`  | `#B91C1C`  | none       | `14,7`  | 12       | 6      |
| `AccentButton`    | `#3B82F6`   | White      | `#2563EB`  | `#1D4ED8`  | none       | `14,7`  | 12       | 6      |
| `WinCtrlButton`   | Transparent | `#0F172A`  | `#F1F5F9`  | —          | none       | —       | —        | —      |
| `CloseButton`     | Transparent | `#0F172A`  | `#EF4444` (Foreground → White) | — | none | — | — | — |

All shared buttons use `FontWeight="SemiBold"`, `Cursor="Hand"`, disabled state background `#E2E8F0` and foreground `#94A3B8`.

## Window Control Buttons

The three chrome buttons (min/max/close) use **embedded `<TextBlock>` children**
with explicit `FontFamily` and `FontSize` but **NO `Foreground` attribute** —
the foreground is inherited from the button style so the close glyph can flip
to white on its red hover.

### Glyph table (Segoe MDL2 Assets, FontSize 10)

| Button   | Glyph      | Notes                                                  |
|----------|------------|--------------------------------------------------------|
| Minimize | `&#xE921;` | Do **not** substitute Unicode minus `&#x2212;`.        |
| Maximize | `&#xE922;` | Do **not** substitute Unicode white square `&#x25A1;`. |
| Close    | `&#xE8BB;` |                                                        |

### Canonical XAML
```xml
<StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top"
            WindowChrome.IsHitTestVisibleInChrome="True">
    <Button x:Name="btn_minimize" Style="{StaticResource WinCtrlButton}"
            Click="minimize_button_clicked" ToolTip="Minimize">
        <TextBlock Text="&#xE921;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
    </Button>
    <Button x:Name="btn_maximize" Style="{StaticResource WinCtrlButton}"
            Click="maximize_button_clicked" ToolTip="Maximize">
        <TextBlock Text="&#xE922;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
    </Button>
    <Button x:Name="btn_close" Style="{StaticResource CloseButton}"
            Click="close_button_clicked" ToolTip="Close">
        <TextBlock Text="&#xE8BB;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
    </Button>
</StackPanel>
```

## Inputs (TextBox / ComboBox / search boxes)

- `FontFamily="Inter"`, `FontSize="12"`, `Padding="6,4"`
- `BorderBrush="#CBD5E1"`, read-only fields `Background="#F8FAFC"`
- Focus state: `BorderBrush="#3B82F6"` with BorderThickness `2` (no glow)
- Placeholder text: `#94A3B8`

## DataGrid Style

> Ground truth is the values in this table. It used to cite the `sample_grid` DataGrid in `UIStandardShowcase.xaml`; that file is now Revit-native and its grid is styled from theme tokens, so it is no longer a Lumina reference.
> This section previously listed stale values (Header `Background="#F8FAFC"`, `Foreground="#0F172A"`,
> `FontWeight="SemiBold"`, `Height="34"`) that never matched the canonical showcase file — corrected below,
> same class of doc drift as the copyright-placement and title-bar-wordmark fixes above.

- `Background="White"`, `BorderBrush="#E2E8F0"`, `BorderThickness="1"`
- `AlternatingRowBackground="#F8FAFC"`, `FontFamily="Hanken Grotesk"` (or `"Inter"` as fallback), `FontSize="13.5"`, `Foreground="#27272A"`
- `GridLinesVisibility="Horizontal"`, `HorizontalGridLinesBrush="#F1F5F9"`
- Headers (`DataGrid.Resources` or `DataGrid.ColumnHeaderStyle`, both valid): `Background="#FFFFFF"`, `Foreground="#9A9AA2"`, `FontWeight="Bold"`, `FontSize="11"`, `Padding="10,8"`, `BorderBrush="#E2E8F0"`, `BorderThickness="0,0,0,1"`
- Row hover: `#F1F5F9`
- Row selected: `#E2E8F0`

```xml
<DataGrid AutoGenerateColumns="False"
          Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
          GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#F1F5F9"
          RowBackground="White" AlternatingRowBackground="#F8FAFC"
          FontFamily="Hanken Grotesk" FontSize="13.5" Foreground="#27272A">
    <DataGrid.Resources>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="#9A9AA2"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
        </Style>
    </DataGrid.Resources>
</DataGrid>
```

## Info / Tip Box

- `BorderBrush="#CBD5E1"`, `Background="#F8FAFC"`, `CornerRadius="4"`, `Padding="10,8"`
- Label: `Tip:` — `FontWeight="Bold"`, `Foreground="#3B82F6"`
- Body: `Foreground="#0F172A"`

## Progress Bar (long-running tasks)

- `Height="8"`, `Foreground="#3B82F6"`, `Background="#E2E8F0"`, `BorderThickness="0"`
- Inline Pause (secondary style) + Stop (`#EF4444`) buttons
- Panel `Visibility="Collapsed"` when idle

## Wizard-Style Navigation (multi-step tools)

When a tool has multiple steps (BatchOut / TileLayout / ExportManager):
- Use hidden `TabItem` with `Visibility="Collapsed"` — navigation driven by code-behind
- Add a **step progress bar** (Row 1) with numbered circles:
  - Active circle: `#3B82F6` bg, white number; active label `#1D4ED8` SemiBold
  - Optional 2px blue `#3B82F6` underline beneath the active step
  - Inactive circles: `#F1F5F9` bg, `#64748B` number; connector lines `#E2E8F0`
- Action bar contains Back (`SecondaryButton`) + Next (`SuccessButton`)
- Nav step icons (Segoe MDL2 Assets): Selection `&#xE14C;`, Format `&#xE1DC;`, Queue `&#xE914;`, Settings `&#xE713;`

## Copyright Notice

**Exactly one copyright TextBlock per file**, placed **inside the status bar (footer), as the LEFT-most element**. This rule spans BOTH tracks and is enforced by `dev/audit_ui.py`; `UIStandardShowcase.xaml` still demonstrates it:

```xml
<!-- Status bar (last row) — Left: Copyright first, then status beside it -->
<Border Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="0,1,0,0" Padding="14,8">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
            <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" VerticalAlignment="Center"/>
            <Border Width="1" Height="14" Background="#DCDCE0" Margin="12,0,12,0" VerticalAlignment="Center"/>
            <TextBlock x:Name="status_text" Text="Ready" FontSize="11" Foreground="#64748B" VerticalAlignment="Center"/>
        </StackPanel>
        <!-- Grid.Column="1": contextual info / counters (right side) -->
    </Grid>
</Border>
```

Rules:
- Text is verbatim `© Copyright by T3Lab`, `FontSize="11"`, `Foreground="#F59E0B"` (amber). Use the literal `©` character — never `&#169;`.
- It is always the **first (left-most) element of the footer**; the status text sits beside it, separated by a 1px vertical divider (`#DCDCE0`, 14px tall, `Margin="12,0,12,0"`).
- Do **not** right-align or center it, and do **not** float it as an overlay on top of the content area.
- **Variant B** (`<Grid>` root modal) with no status bar: place the same TextBlock as the last child of the root Grid with `HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="14,0,0,8" IsHitTestVisible="False" Panel.ZIndex="999" Grid.RowSpan="99" Grid.ColumnSpan="99"` (RowSpan/ColumnSpan required so the bottom-left anchor works regardless of row count).
- Item-template XAML files (root is a list-row `<Border>`/`<DataTemplate>`, e.g. `CadtoFloorLayerItem.xaml`) are **exempt** — no copyright there.
- Do **not** duplicate it.

## Common Deviations to Avoid

- Old Kinetix navy or Terra v2 teal palettes.
- Single-line `<WindowChrome>` tags (silently drops corner radius).
- Non-embedded text or wrong Segoe icons on control buttons.
- Hardcoded button background hexes.
- Copyright block right-aligned, centered, or floating as an overlay instead of sitting at the **left of the status bar (footer)**.
- **Dot-notation definitions**: Never use `<Grid.ColumnDefinition ... />` or `<Grid.RowDefinition ... />`. Always use standard `<ColumnDefinition ... />` and `<RowDefinition ... />` inside definition blocks. The dot-notation is parsed as an empty property element and will cause the runtime crash `Unexpected 'EMPTYPROPERTYELEMENT'`.
