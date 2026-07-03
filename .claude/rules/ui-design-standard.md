# UI Design Standard — Lumina Architecture System

All pyRevit tool windows **MUST** follow the **Lumina Architecture System** design system
(Deep Slate + Refined Blue). The canonical reference is **`UIStandardShowcase.xaml`**.

> Lumina replaces the old navy "Kinetix" palette and the teal/amber "Terra (v2)" palette.
> If you see old navy/teal/amber colors (#083D56, #0F766E, #F59E0B etc.) in a file
> (except where amber is specifically allowed for warning highlights/copyright),
> it has not been migrated — flag it.

## Design Reference Files
- **Canonical UI (reference)**: `.claude/standard/UIStandardShowcase.xaml`
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
- **Use this for every new tool.** Examples: `UIStandardShowcase.xaml`, `AutoDimension.xaml`, `DWGManagement.xaml`.

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

- **All text**: `Hanken Grotesk` (body, labels, headings — as featured in the `UIStandardShowcase.xaml` benchmark. `Inter` is allowed as a secondary/fallback font).
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
   - Left: Tool Name (15px Bold `#0F172A`) + Subtitle (12.5px Italic `#64748B`), stacked vertically. No separate "T3Lab" wordmark — confirmed converged codebase pattern (0/50 files, including the canonical `UIStandardShowcase.xaml`, carry a standalone `T3Lab` label in the title bar).
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

> Ground truth is the `sample_grid` DataGrid in `UIStandardShowcase.xaml` (search `DataGrid standard style`).
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

**Exactly one copyright TextBlock per file**, placed **inside the status bar (footer), as the LEFT-most element** — matching the canonical `UIStandardShowcase.xaml`:

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
