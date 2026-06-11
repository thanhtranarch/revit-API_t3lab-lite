# UI Design Standard — T3Lab Terra (v2)

All pyRevit tool windows **MUST** follow the **T3Lab Terra** design system
(Deep Teal + Amber). The canonical reference is **`BulkFamilyExport.xaml`**.

> v2 replaces the old navy "Kinetix" palette (`#083D56`). If you see navy/slate
> colors (`#083D56`, `#062A3C`, `#2C3E50`, `#546E7A`, `#7F8C8D`, `#27AE60`,
> `#D32F2F`, `#3498DB`, `#E2E6EA`, `#F8F9FA`) in a file, it has not been
> migrated — flag it.

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
- Variant B files use the same Terra palette and the same shared style block (inside `Grid.Resources`).

The rest of this document targets **Variant A** unless otherwise noted.

## Color Palette (Terra)

| Token            | Hex       | Usage                                                  |
|------------------|-----------|--------------------------------------------------------|
| Accent (teal)    | `#0F766E` | Primary buttons, brand, active states, progress fill    |
| Accent hover     | `#115E59` | Primary button hover, emphasized accent text            |
| Accent pressed   | `#0B4F4A` | Primary button pressed                                  |
| Amber highlight  | `#F59E0B` | AccentButton, active-step underline, badges, copyright  |
| Amber hover      | `#D97706` | AccentButton hover                                      |
| Ink text         | `#1C2B33` | Headings, labels, main text, window-control glyphs      |
| Muted text       | `#64748B` | Subtitles, status text, secondary labels                |
| Faint text       | `#94A3B8` | Placeholders, disabled text                             |
| Surface          | `#F6F8F8` | Backgrounds, cards, inputs, status bar, grid headers    |
| Surface hover    | `#E6EDEC` | Secondary button hover, inactive wizard circles         |
| Teal tint        | `#EFF6F5` | DataGrid row hover, tip box background                  |
| Selected tint    | `#D9EBE8` | DataGrid row selected                                   |
| Border           | `#DDE5E7` | Borders, separators, dividers, connector lines          |
| Input border     | `#C7D2D4` | TextBox / ComboBox / search-box borders                 |
| Success green    | `#15803D` | Success/confirm buttons (hover: `#166534`)              |
| Danger red       | `#DC2626` | Delete/destructive buttons (hover: `#B91C1C`)           |
| Disabled bg      | `#AEBFBC` | Disabled button background                              |

There is **no** separate "copyright blue" anymore — copyright uses amber `#F59E0B`.

### Migration map (old Kinetix → Terra)

| Old (Kinetix)              | New (Terra)                                              |
|----------------------------|----------------------------------------------------------|
| `#083D56` primary          | `#0F766E`                                                |
| `#062A3C` primary hover    | `#115E59`                                                |
| `#2C3E50` dark text        | `#1C2B33`                                                |
| `#7F8C8D` gray text        | `#64748B`                                                |
| `#F8F9FA` neutral bg       | `#F6F8F8`                                                |
| `#E2E6EA` hover/selected   | `#E6EDEC` (hover) / `#D9EBE8` (DataGrid selected)        |
| `#546E7A` (context!)       | borders/separators → `#DDE5E7`; input borders → `#C7D2D4`; disabled/placeholder text → `#94A3B8`; disabled button bg → `#AEBFBC` |
| `#27AE60` / `#1E8449`      | `#15803D` / `#166534`                                    |
| `#D32F2F` / `#B71C1C`      | `#DC2626` / `#B91C1C`                                    |
| `#E67E22` orange button    | `AccentButton` style (`#F59E0B`)                         |
| `#3498DB` copyright        | `#F59E0B`                                                |

## Typography

- **All text**: `Inter` (body, labels, headings — all use Inter)
- **Root Window**: `FontFamily="Inter"` on `<Window>` element
- **Never** use `Manrope` or `Segoe UI` for body text. `Segoe MDL2 Assets` is allowed **only** for icon glyphs.

## Window Structure (every tool)

1. `<Window>` with `Background="White"`, `ResizeMode="CanResizeWithGrip"`, `FontFamily="Inter"`
2. **WindowChrome — multi-line form is required.** All five attributes below must be present together. The single-line shortcut silently drops `CornerRadius` and `GlassFrameThickness` and breaks the rounded-corner shell:
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
   - Left: `T3Lab` (11px Bold `#0F766E`) + Tool Name (18px Bold `#1C2B33`)
   - Below: Separator (1px `#DDE5E7`) + subtitle (10px Italic `#64748B`)
   - Right: Min/Max/Close buttons (Segoe MDL2 Assets glyphs — see table below)
   - Bottom: Border 1px `#DDE5E7`
4. **Content area** — tool-specific. 16px horizontal margins; 10–12px vertical rhythm between sections.
5. **Status bar** (last row): `Background="#F6F8F8"`, `BorderBrush="#DDE5E7"`, `BorderThickness="0,1,0,0"`, `Padding="14,8"`
6. **Copyright overlay** — always before closing root `<Grid>`

## Shared Styles (single source of truth)

Button styles are **not** hand-edited per file anymore. The master copy lives in
`T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml`, and every tool XAML carries
an identical copy of the block between these markers inside `Window.Resources`
(Variant A) or `Grid.Resources` (Variant B):

```xml
<!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT (edit lib/GUI/Resources/WPF_styles.xaml, then run dev/sync_wpf_styles.py) ═══ -->
... shared styles ...
<!-- ═══ END T3LAB SHARED STYLES ═══ -->
```

- To change a button style: edit the master file, run `python3 dev/sync_wpf_styles.py` to propagate, or `python3 dev/sync_wpf_styles.py --check` to verify all files match.
- Window-specific styles (DataGrid cell styles, one-off toggles) go **outside** the markers.
- Never edit inside the markers by hand.

### Shared style keys

| Style Key         | Background  | Foreground | Hover Bg   | Pressed Bg | Border     | Padding | FontSize | Radius |
|-------------------|-------------|------------|------------|------------|------------|---------|----------|--------|
| `PrimaryButton`   | `#0F766E`   | White      | `#115E59`  | `#0B4F4A`  | none       | `14,7`  | 12       | 6      |
| `SecondaryButton` | `#F6F8F8`   | `#1C2B33`  | `#E6EDEC`  | `#DDE5E7`  | `#C7D2D4`  | `14,7`  | 12       | 6      |
| `SuccessButton`   | `#15803D`   | White      | `#166534`  | `#14532D`  | none       | `14,7`  | 12       | 6      |
| `DangerButton`    | `#DC2626`   | White      | `#B91C1C`  | `#991B1B`  | none       | `14,7`  | 12       | 6      |
| `AccentButton`    | `#F59E0B`   | White      | `#D97706`  | `#B45309`  | none       | `14,7`  | 12       | 6      |
| `WinCtrlButton`   | Transparent | `#1C2B33`  | `#E6EDEC`  | —          | none       | —       | —        | —      |
| `CloseButton`     | Transparent | `#1C2B33`  | `#DC2626` (Foreground → White) | — | none | — | — | — |

All shared buttons use `FontWeight="SemiBold"`, `Cursor="Hand"`, disabled state
`Background="#AEBFBC"`.

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
- `BorderBrush="#C7D2D4"`, read-only fields `Background="#F6F8F8"`
- Placeholder text: `#94A3B8`

## DataGrid Style

- `Background="White"`, `BorderBrush="#DDE5E7"`, `BorderThickness="1"`
- `AlternatingRowBackground="#F6F8F8"`, `FontFamily="Inter"`, `FontSize="12"`
- `GridLinesVisibility="Horizontal"`, `HorizontalGridLinesBrush="#EFF2F2"`
- Headers: `Background="#F6F8F8"`, `Foreground="#1C2B33"`, `FontWeight="SemiBold"`, `Height="34"`, `BorderBrush="#DDE5E7"`
- Row hover: `#EFF6F5`
- Row selected: `#D9EBE8`

## Info / Tip Box

- `BorderBrush="#BFDBD7"`, `Background="#EFF6F5"`, `CornerRadius="4"`, `Padding="10,8"`
- Label: `Tip:` — `FontWeight="Bold"`, `Foreground="#115E59"`
- Body: `Foreground="#1C2B33"`

## Progress Bar (long-running tasks)

- `Height="8"`, `Foreground="#0F766E"`, `Background="#E6EDEC"`, `BorderThickness="0"`
- Inline Pause (secondary style) + Stop (`#DC2626`) buttons
- Panel `Visibility="Collapsed"` when idle

## Wizard-Style Navigation (multi-step tools)

When a tool has multiple steps (BatchOut / TileLayout / ExportManager):
- Use hidden `TabItem` with `Visibility="Collapsed"` — navigation driven by code-behind
- Add a **step progress bar** (Row 1) with numbered circles:
  - Active circle: `#0F766E` bg, white number; active label `#115E59` SemiBold
  - Optional 2px amber `#F59E0B` underline beneath the active step
  - Inactive circles: `#E6EDEC` bg, `#64748B` number; connector lines `#DDE5E7`
- Action bar contains Back (`SecondaryButton`) + Next (`SuccessButton`)
- Nav step icons (Segoe MDL2 Assets): Selection `&#xE14C;`, Format `&#xE1DC;`, Queue `&#xE914;`, Settings `&#xE713;`

## Copyright Notice

**Exactly one copyright TextBlock per file**, placed as the **last child of the root `<Grid>`** (immediately before the closing `</Grid>`), using this **verbatim snippet** — same for Variant A and Variant B:

```xml
<!-- Copyright added automatically -->
<TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999"/>
```

- Do **not** embed it inside the status bar, action bar, or any other content `<Border>` / `<StackPanel>` — it must be a direct child of the root `<Grid>` so it floats as an overlay.
- Do **not** add `Grid.Row`, `Grid.Column`, or `Grid.RowSpan` — keep the snippet byte-for-byte identical across files.
- Do **not** use `&#169;` — use the literal `©` character.
- Do **not** duplicate it (some legacy files had two copyrights; only one is allowed).

## Common Deviations to Avoid

| Anti-pattern                                                            | Use instead                                                          |
|-------------------------------------------------------------------------|----------------------------------------------------------------------|
| Old Kinetix navy palette (`#083D56`, `#2C3E50`, `#546E7A`, …)           | Terra palette (see migration map above)                              |
| Single-line `<WindowChrome CaptionHeight="64" .../>`                    | Multi-line form with all 5 attributes (see Window Structure §2)      |
| `<Button Content="&#x25A1;" />` for maximize                            | `<Button><TextBlock Text="&#xE922;" FontFamily="Segoe MDL2 Assets" .../></Button>` |
| Explicit `Foreground` on window-control glyph TextBlocks                | No Foreground — inherit from `WinCtrlButton` / `CloseButton` style   |
| Hand-editing button styles inside the shared-styles markers             | Edit `lib/GUI/Resources/WPF_styles.xaml`, run `dev/sync_wpf_styles.py` |
| `FontFamily="Manrope"` or `FontFamily="Segoe UI"`                       | `FontFamily="Inter"` (only Segoe MDL2 Assets is allowed, for glyphs) |
| `<Image Source="…/T3Lab_logo.png"/>` in title bar                       | Logo was removed; do not add it back                                 |
| Hard-coded button colors (`Background="#FF0000"` etc.)                  | Use the shared style keys (incl. `AccentButton` for amber)           |
| Status bar with `Background="White"`                                    | Status bar background is `#F6F8F8`                                   |
| Copyright TextBlock embedded inside a status-bar / action-bar Border    | Place it as the last child of the root `<Grid>` (overlay form)       |
| Two copyright TextBlocks in the same file                               | Exactly one, using the verbatim snippet                              |

## Checklist for New Tools

- [ ] `Window.Background="White"`, `ResizeMode="CanResizeWithGrip"`, `FontFamily="Inter"`
- [ ] `WindowChrome` uses multi-line form with all 5 attributes (`CaptionHeight=64`, `ResizeBorderThickness=5`, `GlassFrameThickness=0`, `CornerRadius=8`, `UseAeroCaptionButtons=False`)
- [ ] Title bar: 64px, white, T3Lab brand (#0F766E) + tool name (#1C2B33) + subtitle (#64748B)
- [ ] Separator (1px #DDE5E7) between title and subtitle
- [ ] Min/Max/Close buttons use the canonical TextBlock pattern with Segoe MDL2 glyphs `&#xE921; &#xE922; &#xE8BB;` (FontSize=10, no Foreground attribute)
- [ ] Shared style block present between the AUTO-SYNCED markers, byte-identical to the master (`dev/sync_wpf_styles.py --check` passes)
- [ ] DataGrid with correct header/row styles (if applicable)
- [ ] Status bar: `#F6F8F8` background, `#DDE5E7` top border
- [ ] Copyright TextBlock overlay (amber `#F59E0B`) before closing `</Grid>`
- [ ] Font: `Inter` throughout (no Manrope, no Segoe UI)
- [ ] `minimize_button_clicked`, `maximize_button_clicked`, `close_button_clicked` implemented in Python
- [ ] No logo/image loading — logo was removed
- [ ] No old Kinetix colors anywhere in the file
- [ ] No deviation from the "Common Deviations to Avoid" table above

## Template References

See `.claude/docs/` for full XAML and Python templates:
- `wpf-window-templates.md` — Window structure, button styles, DataGrid, info box XAML
- `python-wpf-pattern.md` — Python WPF Window class pattern
