# UI Design Standard — Lumina Architecture System

All pyRevit tool windows **MUST** follow the **Lumina Architecture System** design system
(Deep Slate + Refined Blue). The canonical reference is **`BulkFamilyExport.xaml`**.

> Lumina replaces the old navy "Kinetix" palette and the teal/amber "Terra (v2)" palette.
> If you see old navy/teal/amber colors (#083D56, #0F766E, #F59E0B etc.) in a file
> (except where amber is specifically allowed for warning highlights/copyright),
> it has not been migrated — flag it.

## Design Reference Files
- **Canonical UI (reference)**: `T3Lab.extension/lib/GUI/Tools/BulkFamilyExport.xaml`
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
- **Use this for every new tool.** Examples: `BulkFamilyExport.xaml`, `AutoDimension.xaml`, `DWGManagement.xaml`.

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
| Amber highlight  | `#F59E0B` | Warning labels, highlights, copyright overlay          |
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

- **All text**: `Inter` (body, labels, headings — all use Inter)
- **Root Window**: `FontFamily="Inter"` on `<Window>` element
- **Never** use `Manrope` or `Segoe UI` for body text. `Segoe MDL2 Assets` is allowed **only** for icon glyphs.

## Window Structure (every tool)

1. `<Window>` with `Background="White"`, `ResizeMode="CanResizeWithGrip"`, `FontFamily="Inter"`
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
   - Left: `T3Lab` (11px Bold `#0F172A`) + Tool Name (18px Bold `#0F172A`)
   - Below: Separator (1px `#E2E8F0`) + subtitle (10px Italic `#64748B`)
   - Right: Min/Max/Close buttons (Segoe MDL2 Assets glyphs — see table below)
   - Bottom: Border 1px `#E2E8F0`
4. **Content area** — tool-specific. 16px horizontal margins; 10–12px vertical rhythm between sections.
5. **Status bar** (last row): `Background="#F8FAFC"`, `BorderBrush="#E2E8F0"`, `BorderThickness="0,1,0,0"`, `Padding="14,8"`
6. **Copyright overlay** — always before closing root `<Grid>`

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

- `Background="White"`, `BorderBrush="#E2E8F0"`, `BorderThickness="1"`
- `AlternatingRowBackground="#F8FAFC"`, `FontFamily="Inter"`, `FontSize="12"`
- `GridLinesVisibility="Horizontal"`, `HorizontalGridLinesBrush="#F1F5F9"`
- Headers: `Background="#F8FAFC"`, `Foreground="#0F172A"`, `FontWeight="SemiBold"`, `Height="34"`, `BorderBrush="#E2E8F0"`
- Row hover: `#F1F5F9`
- Row selected: `#E2E8F0`

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

**Exactly one copyright TextBlock per file**, placed as the **last child of the root `<Grid>`** (immediately before the closing `</Grid>`), using this **verbatim snippet** — same for Variant A and Variant B:

```xml
<!-- Copyright added automatically -->
<TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999"/>
```

- Do **not** embed it inside the status bar, action bar, or any other content `<Border>` / `<StackPanel>` — it must float as an overlay.
- Do **not** add `Grid.Row`, `Grid.Column`, or `Grid.RowSpan` attributes.
- Do **not** use `&#169;` — use the literal `©` character.
- Do **not** duplicate it.

## Common Deviations to Avoid

- Old Kinetix navy or Terra v2 teal palettes.
- Single-line `<WindowChrome>` tags (silently drops corner radius).
- Non-embedded text or wrong Segoe icons on control buttons.
- Hardcoded button background hexes.
- Copyright block embedded inside templates instead of direct root Grid child.
