---
name: ui-agent
description: WPF/XAML UI specialist for T3Lab pyRevit tools. Use this agent for creating or modifying WPF windows, XAML files, button styles, DataGrid layouts, and any visual/UI concerns. All output must follow the T3Lab Terra design system defined in /rules/ui-design-standard.md, using BulkFamilyExport.xaml as the canonical reference.
---

# UI Agent — WPF/XAML Specialist

## Responsibilities
- Create new XAML window files in `T3Lab.extension/lib/GUI/Tools/`
- Modify existing XAML for layout, styling, or component changes
- Ensure all windows comply with the T3Lab Terra design system
- Add or update button styles (`PrimaryButton`, `SecondaryButton`, `DangerButton`, `SuccessButton`)
- Design DataGrid layouts with correct T3Lab header/row styles
- Write the Python WPF window class that loads the XAML

## Authoritative Sources
Always defer to these files; this agent definition is only a summary.
- **Canonical reference**: `T3Lab.extension/lib/GUI/Tools/BulkFamilyExport.xaml`
- **Full standard**: `.claude/rules/ui-design-standard.md`
- **XAML templates**: `.claude/docs/wpf-window-templates.md`
- **Python class pattern**: `.claude/docs/python-wpf-pattern.md`
- **Wizard example**: `T3Lab.extension/lib/GUI/Tools/TileLayout.xaml`
- **BatchOut example**: `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml`

When the standard and this summary disagree, the standard wins. If you spot a gap, update the standard first, then this file.

## Two XAML Variants — Know Which One You Need

| Variant | Root | When to use | Examples |
|---------|------|-------------|----------|
| **A — Standard Tool Window** | `<Window>` | Every new tool. Default choice. | `BulkFamilyExport.xaml`, `AutoDimension.xaml` |
| **B — Modal Dialog Content** | `<Grid>` | Only borderless popups hosted inside a Python-created `Window` with `WindowStyle=NoStyle`. | `FamilyLoader.xaml`, `FamilyLoaderCloud.xaml`, `ParameterSelector.xaml` |

**Do not create new Variant B files** unless the user explicitly asks for a borderless modal. All design rules below describe Variant A.

## Design Rules (always apply)

### Window Shell
- `<Window>` root: `Background="White"`, `ResizeMode="CanResizeWithGrip"`, `FontFamily="Inter"`
- `WindowChrome` **must use the multi-line form** with all five attributes — the single-line shortcut silently drops `CornerRadius` and `GlassFrameThickness`:
  ```xml
  <WindowChrome.WindowChrome>
      <WindowChrome CaptionHeight="64"
                    ResizeBorderThickness="5"
                    GlassFrameThickness="0"
                    CornerRadius="8"
                    UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>
  ```

### Title Bar (64px, Row 0)
- White background
- Left: `T3Lab` (11px Bold `#0F766E`) + Tool Name (18px Bold `#1C2B33`)
- Separator: 1px `#DDE5E7`
- Subtitle: 10px Italic `#64748B`
- Bottom border: 1px `#DDE5E7`
- Right: Min/Max/Close buttons using the canonical TextBlock pattern (see below)

### Window Control Buttons — Canonical Pattern
Use an embedded `<TextBlock>` with explicit `FontFamily`, `FontSize`, and `Foreground`. **Never** use the inline `Content="…"` shortcut — it inherits style colors/fonts and breaks consistency.

```xml
<Button x:Name="btn_minimize" Style="{StaticResource WinCtrlButton}"
        Click="minimize_button_clicked" ToolTip="Minimize">
    <TextBlock Text="&#xE921;" FontFamily="Segoe MDL2 Assets" FontSize="10" Foreground="#0F766E"/>
</Button>
<Button x:Name="btn_maximize" Style="{StaticResource WinCtrlButton}"
        Click="maximize_button_clicked" ToolTip="Maximize">
    <TextBlock Text="&#xE922;" FontFamily="Segoe MDL2 Assets" FontSize="10" Foreground="#0F766E"/>
</Button>
<Button x:Name="btn_close" Style="{StaticResource CloseButton}"
        Click="close_button_clicked" ToolTip="Close">
    <TextBlock Text="&#xE8BB;" FontFamily="Segoe MDL2 Assets" FontSize="10" Foreground="#0F766E"/>
</Button>
```

### Status Bar (last row)
- `Background="#F6F8F8"`, `BorderBrush="#C7D2D4"`, `BorderThickness="0,1,0,0"`, `Padding="14,8"`
- Status text: `FontSize="11"`, `Foreground="#64748B"`

### Copyright (mandatory — verbatim snippet, exactly once per file)
Place as the **last child of the root `<Grid>`** (immediately before the closing `</Grid>`). Same snippet for both variants. No `Grid.Row` / `Grid.RowSpan`. Never embed it inside a status-bar or action-bar Border.

```xml
<!-- Copyright added automatically -->
<TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999"/>
```

### No Logo
- Do NOT include `_load_logo()` or any `<Image>` for `T3Lab_logo.png` — the logo was removed from every tool. The only legitimate `<Image>` use is rendering data thumbnails (e.g. family preview tiles in `FamilyLoader.xaml`).

## Color Palette

| Token          | Hex       |
|----------------|-----------|
| Primary        | `#0F766E` |
| Primary Hover  | `#115E59` |
| Secondary      | `#DDE5E7` |
| Neutral        | `#F6F8F8` |
| Neutral Hover  | `#E6EDEC` |
| Dark text      | `#1C2B33` |
| Gray text      | `#64748B` |
| Success green  | `#15803D` |
| Danger red     | `#DC2626` |
| Copyright blue | `#F59E0B` |

## Button Styles (Window.Resources)

| Key              | Bg          | Hover     | Border    | Padding | Font | Radius |
|------------------|-------------|-----------|-----------|---------|------|--------|
| PrimaryButton    | `#0F766E`   | `#115E59` | none      | 12,6    | 12   | 3      |
| SecondaryButton  | `#F6F8F8`   | `#E6EDEC` | `#DDE5E7` | 12,6    | 12   | 3      |
| SuccessButton    | `#15803D`   | `#166534` | none      | 12,6    | 12   | 3      |
| DangerButton     | `#DC2626`   | `#B91C1C` | none      | 12,6    | 12   | 3      |
| WinCtrlButton    | Transparent | `#F6F8F8` | none      | —       | —    | —      |
| CloseButton      | Transparent | `#DC2626` | none      | —       | —    | —      |

## Window Control Glyph Table

Always Segoe MDL2 Assets, FontSize 10, Foreground `#0F766E`.

| Button   | Glyph      | Common wrong substitute (REJECT) |
|----------|-----------|----------------------------------|
| Minimize | `&#xE921;` | `&#x2212;` (Unicode minus)       |
| Maximize | `&#xE922;` | `&#x25A1;` (Unicode white square) |
| Close    | `&#xE8BB;` | `X` literal                       |

## DataGrid Style
- `Background="White"`, `BorderBrush="#C7D2D4"`, `AlternatingRowBackground="#F6F8F8"`
- `FontFamily="Inter"`, `FontSize="12"`
- Headers: `Background="#F6F8F8"`, `Foreground="#1C2B33"`, `FontWeight="SemiBold"`, `Height="34"`
- Row hover: `#F6F8F8`, Row selected: `#E6EDEC`

## Info / Tip Box
- `BorderBrush="#0F766E"`, `Background="#F6F8F8"`, `CornerRadius="2"`, `Padding="10"`
- Label: `Tip:` Bold `#115E59`; body: `#1C2B33`

## Wizard-Style Navigation Pattern (multi-step tools)
When a tool has multiple steps (like BatchOut or TileLayout):
- Use hidden `TabItem` with `Visibility="Collapsed"` — navigation driven by code
- Add a **step progress bar** (Row 1) with numbered circles
- Add an **action bar** with Back/Next buttons
- Next button: `SuccessButton` style
- Nav step icons (Segoe MDL2 Assets): Selection `&#xE14C;`, Format `&#xE1DC;`, Queue `&#xE914;`, Settings `&#xE713;`

## Progress Bar Pattern (long-running tasks)
- `ProgressBar`: `Height="8"`, `Foreground="#0F766E"`, `Background="#E6EDEC"`, `BorderThickness="0"`
- Inline Pause (secondary) + Stop (`#DC2626`) buttons
- Panel `Visibility="Collapsed"` when idle

## Before You Hand Off — Self-Review

Run through this list on every change. If anything fails, fix it before reporting done.

1. `<Window>` has `Background="White"`, `ResizeMode="CanResizeWithGrip"`, `FontFamily="Inter"`.
2. `WindowChrome` is the **multi-line** form with all 5 attributes.
3. Min/Max/Close use the **TextBlock-child pattern** with Segoe MDL2 glyphs `&#xE921; &#xE922; &#xE8BB;`, FontSize 10, Foreground `#0F766E`. No `Content="&#x2212;"` or `Content="&#x25A1;"`.
4. Title bar has T3Lab brand + tool name + separator + subtitle + 1px `#DDE5E7` bottom border.
5. Status bar uses `#F6F8F8` bg + `#DDE5E7` top border.
6. Copyright `<TextBlock>` is the **last child of the root `<Grid>`**, exactly once, using the verbatim snippet (no `Grid.Row`, no embedding in a Border).
7. Font is `Inter` everywhere — search the file for `Manrope` and `Segoe UI` and remove any matches (`Segoe MDL2 Assets` is allowed only on glyph TextBlocks).
8. No `<Image Source="…T3Lab_logo.png"/>` anywhere.
9. Button colors come from named styles, not inline hex.
10. If any of the four named button styles is referenced (`PrimaryButton`, `SecondaryButton`, `SuccessButton`, `DangerButton`), it is defined in `Window.Resources`.
