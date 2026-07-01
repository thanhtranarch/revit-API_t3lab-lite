---
name: ui-agent
description: WPF/XAML UI specialist for T3Lab pyRevit tools. Use this agent for creating or modifying WPF windows, XAML files, button styles, DataGrid layouts, and any visual/UI concerns. All output must follow the T3Lab Lumina design system defined in /rules/ui-design-standard.md, using UIStandardShowcase.xaml as the canonical reference.
---

# UI Agent — WPF/XAML Specialist

## Responsibilities
- Create new XAML window files in `T3Lab.extension/lib/GUI/Tools/`
- Modify existing XAML for layout, styling, or component changes
- Ensure all windows comply with the T3Lab Lumina design system
- Write the Python WPF window class that loads the XAML

## Authoritative Sources
**Always read these files before writing any XAML.** This agent definition is only a summary — when there is any conflict, the source files win.

| Resource | Path |
|----------|------|
| **Canonical reference (read first)** | `.claude/standard/UIStandardShowcase.xaml` |
| Full design standard | `.claude/rules/ui-design-standard.md` |
| Shared styles master | `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` |
| XAML templates | `.claude/docs/wpf-window-templates.md` |
| Python class pattern | `.claude/docs/python-wpf-pattern.md` |

**Before writing any XAML**: read `UIStandardShowcase.xaml` and copy exact hex values, CornerRadius values, font sizes, and layout patterns from it. Do not rely on memory — the source file is always authoritative.

## Two XAML Variants

| Variant | Root | When to use | Examples |
|---------|------|-------------|----------|
| **A — Standard Tool Window** | `<Window>` | Every new tool. Default choice. | `UIStandardShowcase.xaml`, `AutoDimension.xaml` |
| **B — Modal Dialog Content** | `<Grid>` | Only borderless popups hosted inside a Python-created `Window` with `WindowStyle=NoStyle`. | `FamilyLoader.xaml`, `FamilyLoaderCloud.xaml`, `ParameterSelector.xaml` |

**Do not create new Variant B files** unless explicitly requested for a borderless modal.

## UI-Frozen Files — NEVER MODIFY
These files have finalized custom designs. Skip them entirely for any UI task:
- `T3Lab.extension/lib/GUI/Tools/DWGManagement.xaml`
- `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml`

---

## Design Rules (Variant A — Standard Tool Window)

### Window Shell
```xml
<Window Background="#E4E4E7"
        ResizeMode="CanResizeWithGrip"
        FontFamily="Hanken Grotesk"
        FontSize="14"
        WindowStartupLocation="CenterScreen">
```

- `Background="#E4E4E7"` — the gray application background (NOT `White`)
- Wrap root Grid in outer Border: `<Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7">`

### WindowChrome — Multi-line form required, all 5 attributes
```xml
<WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="64"
                  ResizeBorderThickness="5"
                  GlassFrameThickness="0"
                  CornerRadius="22"
                  UseAeroCaptionButtons="False"/>
</WindowChrome.WindowChrome>
```
- `CornerRadius="22"` — matches the outer Border. Never `8`.

### Root Grid Layout
```xml
<Grid.ColumnDefinitions>
    <ColumnDefinition Width="66"/>  <!-- Sidebar icon rail (omit if no sidebar) -->
    <ColumnDefinition Width="*"/>   <!-- Main content -->
</Grid.ColumnDefinitions>
<Grid.RowDefinitions>
    <RowDefinition Height="64"/>    <!-- Title bar -->
    <RowDefinition Height="*"/>     <!-- Content -->
    <RowDefinition Height="Auto"/>  <!-- Footer -->
</Grid.RowDefinitions>
```

### Title Bar (Row 0, Height=64)
```xml
<Grid Grid.Row="0" Grid.Column="1" Background="#F4F4F6">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <StackPanel Grid.Column="0" Margin="22,0,0,0" VerticalAlignment="Center"
                WindowChrome.IsHitTestVisibleInChrome="True">
        <TextBlock Text="Tool Name" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
        <TextBlock Text="Subtitle · Revit 2024–2026" FontSize="12.5" Foreground="#71717A" Margin="0,2,0,0"/>
    </StackPanel>

    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center"
                Margin="0,0,16,0" WindowChrome.IsHitTestVisibleInChrome="True">
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

    <Border Height="1" VerticalAlignment="Bottom" Background="#DCDCE0" Grid.ColumnSpan="3"/>
</Grid>
```

Key values:
- Background: `#F4F4F6` (NOT `White`)
- Bottom divider: `#DCDCE0` (NOT `#E2E8F0`)
- Title: FontSize=`15`, FontWeight=`Bold`, Foreground=`#18181B`
- Subtitle: FontSize=`12.5`, Foreground=`#71717A`, NO italic
- `WindowChrome.IsHitTestVisibleInChrome="True"` on both StackPanels

### Window Control Buttons
- TextBlock children with `FontFamily="Segoe MDL2 Assets"` `FontSize="10"` — NO `Foreground` attribute (inherited from style)
- Glyphs: Minimize `&#xE921;`, Maximize `&#xE922;`, Close `&#xE8BB;`
- NEVER use Unicode minus `&#x2212;`, white square `&#x25A1;`, or literal `X`

### Content Area (Row 1)
Wrap content in white card panels:
```xml
<Border Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="20" Padding="18">
    <!-- content -->
</Border>
```
Margin from grid edge: `Margin="18,18,18,10"`

### Footer / Status Bar (Row 2)

**Standard layout: Copyright + Status together on the LEFT, all action buttons on the RIGHT.**

```xml
<Border Grid.Row="2" Grid.Column="1"
        Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,16">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Left: Copyright, then Status right beside it -->
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
            <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" VerticalAlignment="Center"/>
            <Border Width="1" Height="14" Background="#DCDCE0" Margin="12,0,12,0" VerticalAlignment="Center"/>
            <TextBlock x:Name="status_text" Text="Ready" FontSize="13.5" Foreground="#71717A" FontWeight="SemiBold" VerticalAlignment="Center"/>
        </StackPanel>

        <!-- Right: all action buttons (omit the StackPanel entirely if the tool has none) -->
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="16,0,0,0">
            <Button Style="{StaticResource SecondaryButton}" Content="Cancel" Margin="0,0,8,0"/>
            <Button Style="{StaticResource PrimaryButton}" Content="Execute"/>
        </StackPanel>
    </Grid>
</Border>
```

Key values:
- Background: `#F4F4F6` (NOT `#F8FAFC`)
- BorderBrush: `#DCDCE0` (NOT `#E2E8F0`)
- Copyright: placed **inside the footer's left column, first**, immediately followed by a thin `#DCDCE0` divider and then `status_text` — NOT as a floating overlay on the root Grid, and NOT stacked vertically above/below status
- All buttons go in the right column, right-aligned

### Copyright Rule
Copyright is a **real element in the footer's left column** (first item, before status), NOT a floating overlay. Do NOT use `Panel.ZIndex="999"` floating pattern.

For simple tools with no status text, copyright can stand alone in the left column:
```xml
<StackPanel Grid.Column="0" VerticalAlignment="Center">
    <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B"/>
</StackPanel>
```

### Sidebar Icon Rail (Column 0, optional)
```xml
<Border Grid.Column="0" Grid.RowSpan="3" Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,0,1,0">
    <Grid Margin="0,16,0,16">
        <!-- Logo icon at top -->
        <Border Width="42" Height="42" CornerRadius="13" Background="White" BorderBrush="#DCDCE0" BorderThickness="1" HorizontalAlignment="Center" Margin="0,0,0,24" Grid.Row="0"/>
        <!-- Nav buttons using T3SidebarButton style -->
    </Grid>
</Border>
```

---

## Color Palette (UIStandardShowcase)

| Token | Hex | Usage |
|-------|-----|-------|
| App background | `#E4E4E7` | Window Background, outer Border bg |
| Component bg | `#FFFFFF` | Content cards, input focus |
| Panel bg | `#F4F4F6` | Title bar, sidebar, footer, input default bg |
| Input border | `#E6E6EA` | TextBox, ComboBox borders |
| Divider | `#DCDCE0` | Title bar bottom border, footer top border |
| Primary ink | `#18181B` | PrimaryButton bg, headings, active states |
| Primary text | `#27272A` | Body text, SecondaryButton text |
| Secondary text | `#3F3F46` | Paragraph text |
| Tertiary text | `#71717A` | Subtitles, placeholders |
| Disabled text | `#9A9AA2` | Disabled labels, meta text |
| Success | `#22A85C` | SuccessButton bg |
| Success hover | `#0F8A66` | SuccessButton hover |
| Success text | `#157038` | Success status text |
| Danger | `#D23B3B` | DangerButton bg |
| Danger hover | `#F87171` | DangerButton hover |
| Copyright amber | `#F59E0B` | Copyright TextBlock |
| Progress accent | `#C2410C` | ProgressBar, slider |

---

## Shared Styles

Copy the full block from `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` into `Window.Resources` between these markers:
```
<!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT ... ═══ -->
...
<!-- ═══ END T3LAB SHARED STYLES ═══ -->
```

**Never hand-edit inside the markers.** Never copy button hex values manually — always use the named style keys.

### Button Style Keys

| Key | Bg | Hover bg | Height | FontSize | CornerRadius |
|-----|----|----------|--------|----------|--------------|
| `PrimaryButton` | `#18181B` | `#000000` | 40 | 13.5 | 12 |
| `SecondaryButton` | `#FFFFFF` border `#DEDEE2` | `#FAFAFB` | 40 | 13.5 | 12 |
| `TertiaryButton` | `#ECECEF` | `#E2E2E6` | 40 | 13.5 | 12 |
| `SuccessButton` | `#22A85C` | `#0F8A66` | 40 | 13.5 | 12 |
| `DangerButton` | `#D23B3B` | `#F87171` | 40 | 13.5 | 12 |
| `GhostButton` | Transparent | `#EDEDEF` | 40 | 13.5 | 12 |
| `WinCtrlButton` | Transparent | `#E4E4E7` | 32 | — | 8 |
| `CloseButton` | Transparent | `#D23B3B` fg→White | 32 | — | 8 |
| `T3SidebarButton` | Transparent | `#F4F4F6` | 42 | — | 13 |

---

## Input Controls

### Text Input (T3InputField)
```xml
<TextBox Style="{StaticResource T3InputField}" Height="44"/>
```
- `Background="#F4F4F6"`, `BorderBrush="#E6E6EA"`, `CornerRadius="13"`, Height=44, FontSize=14, FontFamily=Hanken Grotesk
- Focus: `BorderBrush="#18181B"`, `Background="White"`

### ComboBox (T3ComboBox)
```xml
<ComboBox Style="{StaticResource T3ComboBox}" Height="44"/>
```
- Same pill shape, custom arrow, `CornerRadius="13"`

### Multi-line TextArea
```xml
<Border Background="#F4F4F6" BorderBrush="#E6E6EA" BorderThickness="1" CornerRadius="13">
    <TextBox Background="Transparent" BorderThickness="0"
             AcceptsReturn="True" TextWrapping="Wrap"
             VerticalScrollBarVisibility="Auto"
             FontFamily="Hanken Grotesk" FontSize="14"
             Padding="12,10" MinHeight="140" VerticalContentAlignment="Top"/>
</Border>
```

---

## Info / Alert Boxes (UIStandardShowcase pattern)

### Info Alert
```xml
<Border Background="#F4F4F6" BorderBrush="#E6E6EA" BorderThickness="1" CornerRadius="14" Padding="12,10">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Border Grid.Column="0" Width="20" Height="20" CornerRadius="10"
                Background="#18181B" VerticalAlignment="Top" Margin="0,1,8,0">
            <Path Data="M12 11 V16.5 M12 7.6 V7.62" Stroke="White" StrokeThickness="2.1"
                  StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                  Width="12" Height="12" Stretch="Uniform"
                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock Grid.Column="1" FontSize="13.5" Foreground="#3F3F46" TextWrapping="Wrap">
            <Run Text="Tip. " FontWeight="Bold" Foreground="#18181B"/>
            <Run Text="Your message here."/>
        </TextBlock>
    </Grid>
</Border>
```

### Success Alert
```xml
<Border Background="#EAF8F0" BorderBrush="#CDEBD9" BorderThickness="1" CornerRadius="14" Padding="12,10">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Border Grid.Column="0" Width="20" Height="20" CornerRadius="10"
                Background="#22A85C" VerticalAlignment="Top" Margin="0,1,8,0">
            <Path Data="M5 12 L9.5 16.5 L19 7" Stroke="White" StrokeThickness="2.6"
                  StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                  Width="11" Height="11" Stretch="Uniform"
                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <TextBlock Grid.Column="1" FontSize="13.5" Foreground="#1B7A45" TextWrapping="Wrap">
            <Run Text="Done. " FontWeight="Bold" Foreground="#157038"/>
            <Run Text="Operation completed successfully."/>
        </TextBlock>
    </Grid>
</Border>
```

---

## DataGrid Style
```xml
<DataGrid Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
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

---

## Field Labels (form sections)
```xml
<TextBlock Text="FIELD NAME" FontSize="11" FontWeight="Bold" Foreground="#9A9AA2" Margin="1,0,0,6"/>
```
All-caps, `#9A9AA2`, FontSize=11, FontWeight=Bold — matches UIStandardShowcase section headers.

---

## Progress Bar
```xml
<ProgressBar Value="75" Height="8" Background="#ECECEF" Foreground="#C2410C" BorderThickness="0"/>
```

---

## Self-Review Checklist (run before reporting done)

1. `<Window Background="#E4E4E7">` — NOT `White`
2. Outer `<Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7">` wraps root Grid
3. `WindowChrome CornerRadius="22"` — NOT `8`
4. Title bar `Background="#F4F4F6"` — NOT `White`
5. Title bar bottom `<Border Height="1" Background="#DCDCE0">` — NOT `#E2E8F0`
6. Window control buttons use TextBlock children with Segoe MDL2 Assets glyphs `&#xE921; &#xE922; &#xE8BB;` — no Foreground on TextBlock
7. `WindowChrome.IsHitTestVisibleInChrome="True"` on both title bar StackPanels
8. Footer `Background="#F4F4F6"` `BorderBrush="#DCDCE0"` — NOT `#F8FAFC`/`#E2E8F0`
9. Copyright `© Copyright by T3Lab` in amber `#F59E0B` placed **inside footer** — NOT floating overlay
10. Content area wrapped in white card `<Border Background="White" BorderBrush="#E2E8F0" CornerRadius="20">`
11. Inputs use `T3InputField` / `T3ComboBox` styles — NOT plain default WPF styles
12. Shared styles block present between `═══ T3LAB SHARED STYLES v2` markers
13. No `FontFamily="Manrope"` or `FontFamily="Segoe UI"` (body text) — only Hanken Grotesk / Inter allowed
14. No `<Image Source="…T3Lab_logo.png"/>` anywhere
15. No dot-notation `<Grid.ColumnDefinition>` — always `<ColumnDefinition>` inside `<Grid.ColumnDefinitions>`
