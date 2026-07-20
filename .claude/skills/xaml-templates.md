---
name: xaml-templates
description: XAML templates for T3Lab WPF windows. Contains complete snippets for window structure, title bar, footer, content cards, scrollbar styles, button styles, input controls, info alerts, and DataGrid. All templates match UIStandardShowcase.xaml exactly.
---

# XAML Templates — T3Lab Lumina System

> **IMPORTANT**: Always read `.claude/standard/UIStandardShowcase.xaml` before using these templates to verify current hex values and patterns. This file is a convenience summary — the Showcase is always authoritative. Never use absolute file paths (`C:\...`) in documentation or code.

---

## 1. Window Root + WindowChrome

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="T3Lab - Tool Name"
        Height="680" Width="1100"
        MinWidth="860" MinHeight="500"
        WindowStartupLocation="CenterScreen"
        Background="#E4E4E7"
        FontFamily="Hanken Grotesk"
        FontSize="14"
        ResizeMode="CanResizeWithGrip">

    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="22"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Window.Resources>
        <!-- Paste full shared styles block from T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml -->
        <!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT ═══ -->
        <!-- ═══ END T3LAB SHARED STYLES ═══ -->
    </Window.Resources>

    <!-- Outer border separates window from white Revit canvas -->
    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22"
            ClipToBounds="True" Background="#E4E4E7">
        <Grid>
            ...
        </Grid>
    </Border>
</Window>
```

Key values:
- `Background="#E4E4E7"` on Window (NOT `White`)
- `CornerRadius="22"` on WindowChrome (NOT `8`)
- Outer `<Border BorderBrush="#A1A1AA" ... CornerRadius="22">` is required

---

## 2. Root Grid Layout

### With Sidebar
```xml
<Grid.ColumnDefinitions>
    <ColumnDefinition Width="66"/>  <!-- Sidebar icon rail -->
    <ColumnDefinition Width="*"/>   <!-- Main content -->
</Grid.ColumnDefinitions>
<Grid.RowDefinitions>
    <RowDefinition Height="64"/>    <!-- Title bar -->
    <RowDefinition Height="*"/>     <!-- Content -->
    <RowDefinition Height="Auto"/>  <!-- Footer -->
</Grid.RowDefinitions>
```

### Without Sidebar (simple tool)
```xml
<Grid.RowDefinitions>
    <RowDefinition Height="64"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
</Grid.RowDefinitions>
```

---

## 3. Title Bar (Row 0, Height=64)

```xml
<Grid Grid.Row="0" Grid.Column="1" Background="#F4F4F6">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <!-- Title + Subtitle (two lines, no separator, no italic) -->
    <StackPanel Grid.Column="0" Margin="22,0,0,0" VerticalAlignment="Center"
                WindowChrome.IsHitTestVisibleInChrome="True">
        <TextBlock Text="Tool Name" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
        <TextBlock Text="Subtitle · Revit 2024–2026" FontSize="12.5" Foreground="#71717A" Margin="0,2,0,0"/>
    </StackPanel>

    <!-- Window Controls -->
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

    <!-- Bottom divider -->
    <Border Height="1" VerticalAlignment="Bottom" Background="#DCDCE0" Grid.ColumnSpan="3"/>
</Grid>
```

Key values:
- Background: `#F4F4F6` (NOT `White`)
- Bottom divider: `#DCDCE0` (NOT `#E2E8F0`)
- Title: FontSize=15, FontWeight=Bold, Foreground=`#18181B`
- Subtitle: FontSize=12.5, Foreground=`#71717A` (NO italic, NO separator)
- `WindowChrome.IsHitTestVisibleInChrome="True"` on BOTH StackPanels

---

## 4. Sidebar Icon Rail (Column 0, optional)

```xml
<Border Grid.Column="0" Grid.RowSpan="3"
        Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,0,1,0">
    <Grid Margin="0,16,0,16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Logo -->
            <RowDefinition Height="Auto"/> <!-- Button 1 (Active) -->
            <RowDefinition Height="Auto"/> <!-- Button 2 -->
            <RowDefinition Height="*"/>    <!-- Spacer -->
        </Grid.RowDefinitions>

        <!-- Logo Icon -->
        <Border Grid.Row="0" Width="42" Height="42" CornerRadius="13"
                Background="White" BorderBrush="#DCDCE0" BorderThickness="1"
                HorizontalAlignment="Center" Margin="0,0,0,24">
            <Path Data="M12 4 L20 8 L12 12 L4 8 Z M4 12 L12 16 L20 12 M4 16 L12 20 L20 16"
                  Stroke="#18181B" StrokeThickness="1.8" StrokeLineJoin="Round"
                  Width="20" Height="20" Stretch="Uniform"
                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>

        <!-- Active nav button -->
        <Border Grid.Row="1" Width="42" Height="42" CornerRadius="13"
                Background="#18181B" HorizontalAlignment="Center" Margin="0,0,0,8">
            <!-- Icon path here, Stroke="White" -->
        </Border>

        <!-- Inactive nav button -->
        <Button Grid.Row="2" Style="{StaticResource T3SidebarButton}" Margin="0,0,0,8">
            <!-- Icon path here, Stroke="#71717A" -->
        </Button>
    </Grid>
</Border>
```

---

## 5. Content Area (Row 1)

### Standard content card
```xml
<Border Grid.Row="1" Grid.Column="1"
        Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
        CornerRadius="20" Padding="18" Margin="18,18,18,10">
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel>
            <!-- Content here -->
        </StackPanel>
    </ScrollViewer>
</Border>
```

### Two-column layout
```xml
<Grid Grid.Row="1" Grid.Column="1" Margin="18,18,18,10">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="320"/>
        <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <Border Grid.Column="0" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
            CornerRadius="20" Padding="18" Margin="0,0,16,0">
        <!-- Left panel -->
    </Border>

    <Border Grid.Column="1" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
            CornerRadius="20" Padding="18">
        <!-- Right panel -->
    </Border>
</Grid>
```

---

## 6. Footer / Status Bar (Row 2)

### Full status bar (with status indicator + copyright)
```xml
<Border Grid.Row="2" Grid.Column="1"
        Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,16">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Left: action buttons -->
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
            <Button Style="{StaticResource PrimaryButton}" Content="Run" Margin="0,0,8,0" Padding="16,10"/>
            <Button Style="{StaticResource SecondaryButton}" Content="Configure" Padding="16,9"/>
        </StackPanel>

        <!-- Right: status + copyright -->
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="16,0,0,0">
            <Ellipse Width="8" Height="8" Fill="#22A85C" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock Text="System status: " FontSize="13.5" Foreground="#27272A" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <TextBlock x:Name="status_text" Text="Ready" FontSize="13.5" Foreground="#157038" FontWeight="Bold" Margin="0,0,16,0" VerticalAlignment="Center"/>
            <Border Width="1" Height="18" Background="#DEDEE2" Margin="0,0,16,0" VerticalAlignment="Center"/>
            <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                <TextBlock Text="© 2026 T3Lab · v2.4" FontSize="11" Foreground="#9A9AA2"/>
                <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Border>
```

### Simple footer (dialog-style, cancel + confirm)
```xml
<Border Grid.Row="2"
        Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,14">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Left: status + copyright stacked -->
        <StackPanel Grid.Column="0" VerticalAlignment="Center">
            <TextBlock x:Name="status_text" Text="Ready to send"
                       FontSize="13.5" Foreground="#71717A" FontWeight="SemiBold"/>
            <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0"/>
        </StackPanel>

        <!-- Right: action buttons -->
        <StackPanel Grid.Column="1" Orientation="Horizontal">
            <Button x:Name="btn_cancel" Style="{StaticResource SecondaryButton}"
                    Content="Cancel" Width="100" Margin="0,0,8,0" Click="close_button_clicked"/>
            <Button x:Name="btn_send" Style="{StaticResource SuccessButton}"
                    Content="Send" Width="140" Click="send_clicked"/>
        </StackPanel>
    </Grid>
</Border>
```

Key values:
- Background: `#F4F4F6` (NOT `#F8FAFC`)
- BorderBrush: `#DCDCE0` (NOT `#E2E8F0`)
- Copyright `© Copyright by T3Lab` in amber `#F59E0B` is placed **inside the footer** — never as a floating overlay

---

## 7. Ultra-Thin Scrollbar (auto-applied to all ScrollBars)

```xml
<Style TargetType="{x:Type Thumb}" x:Key="ScrollBarThumbStyle">
    <Setter Property="OverridesDefaultStyle" Value="true"/>
    <Setter Property="IsTabStop" Value="false"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="{x:Type Thumb}">
                <Border x:Name="thumbBorder" Background="#A1A1AA" CornerRadius="2" Margin="0"/>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="true">
                        <Setter TargetName="thumbBorder" Property="Background" Value="#71717A"/>
                    </Trigger>
                    <Trigger Property="IsDragging" Value="true">
                        <Setter TargetName="thumbBorder" Property="Background" Value="#18181B"/>
                    </Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
</Style>

<Style TargetType="{x:Type ScrollBar}">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderBrush" Value="Transparent"/>
    <Setter Property="MinWidth" Value="0"/>
    <Setter Property="MinHeight" Value="0"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="{x:Type ScrollBar}">
                <Grid x:Name="Bg" Background="Transparent">
                    <Track x:Name="PART_Track">
                        <Track.Thumb>
                            <Thumb Style="{StaticResource ScrollBarThumbStyle}"/>
                        </Track.Thumb>
                    </Track>
                </Grid>
                <ControlTemplate.Triggers>
                    <Trigger Property="Orientation" Value="Vertical">
                        <Setter TargetName="PART_Track" Property="IsDirectionReversed" Value="true"/>
                    </Trigger>
                    <Trigger Property="Orientation" Value="Horizontal">
                        <Setter TargetName="PART_Track" Property="IsDirectionReversed" Value="false"/>
                    </Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="Orientation" Value="Vertical">
            <Setter Property="Width" Value="4"/>
            <Setter Property="Height" Value="Auto"/>
        </Trigger>
        <Trigger Property="Orientation" Value="Horizontal">
            <Setter Property="Width" Value="Auto"/>
            <Setter Property="Height" Value="4"/>
        </Trigger>
    </Style.Triggers>
</Style>
```

---

## 8. Button Styles

All buttons come from the shared styles block. Key named styles:

| Key | Bg | Hover | Notes |
|-----|----|-------|-------|
| `PrimaryButton` | `#18181B` | `#000000` | Height=40, FontSize=13.5, CornerRadius=12 |
| `SecondaryButton` | `#FFFFFF` border `#DEDEE2` | `#FAFAFB` | Same sizing |
| `TertiaryButton` | `#ECECEF` | `#E2E2E6` | Same sizing |
| `SuccessButton` | `#22A85C` | `#0F8A66` | Same sizing |
| `DangerButton` | `#D23B3B` | `#F87171` | Same sizing |
| `GhostButton` | Transparent | `#EDEDEF` | Same sizing |
| `WinCtrlButton` | Transparent | `#E4E4E7` | Width=40, Height=32 |
| `CloseButton` | Transparent | `#D23B3B` fg→White | Same as WinCtrlButton |
| `T3SidebarButton` | Transparent | `#F4F4F6` | Width=42, Height=42, CornerRadius=13 |

---

## 9. Input Controls

### Text Input (T3InputField)
```xml
<TextBox Style="{StaticResource T3InputField}"/>
<!-- Background=#F4F4F6, BorderBrush=#E6E6EA, CornerRadius=13, Height=44, FontSize=14 -->
<!-- Focus: BorderBrush=#18181B, Background=White -->
```

### ComboBox (T3ComboBox)
```xml
<ComboBox Style="{StaticResource T3ComboBox}">
    <ComboBoxItem Content="Option A" IsSelected="True"/>
    <ComboBoxItem Content="Option B"/>
</ComboBox>
<!-- Height=44, CornerRadius=13, custom arrow, same #F4F4F6/#E6E6EA colors -->
```

### Multi-line TextArea
```xml
<Border Background="#F4F4F6" BorderBrush="#E6E6EA" BorderThickness="1" CornerRadius="13">
    <TextBox Background="Transparent" BorderThickness="0"
             AcceptsReturn="True" TextWrapping="Wrap"
             VerticalScrollBarVisibility="Auto"
             FontFamily="Hanken Grotesk" FontSize="14" Foreground="#27272A"
             Padding="12,10" MinHeight="140" VerticalContentAlignment="Top"/>
</Border>
```

### Field Labels (always ALL-CAPS)
```xml
<TextBlock Text="FIELD NAME" FontSize="11" FontWeight="Bold" Foreground="#9A9AA2" Margin="1,0,0,6"/>
```

---

## 10. Info / Alert Boxes

### Info Alert (dark circle icon)
```xml
<Border Background="#F4F4F6" BorderBrush="#E6E6EA" BorderThickness="1"
        CornerRadius="14" Padding="12,10" Margin="0,0,0,16">
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

### Success Alert (green circle icon)
```xml
<Border Background="#EAF8F0" BorderBrush="#CDEBD9" BorderThickness="1"
        CornerRadius="14" Padding="12,10" Margin="0,0,0,16">
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

## 11. Summary Cards (dashboard-style)

```xml
<Border Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
        CornerRadius="20" Padding="19,17">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0">
            <TextBlock Text="CARD TITLE" FontSize="11" FontWeight="Bold" Foreground="#9A9AA2"/>
            <Border Width="30" Height="30" CornerRadius="10" Background="#F4F4F6" HorizontalAlignment="Right">
                <!-- Icon here -->
            </Border>
        </Grid>
        <TextBlock Grid.Row="1" Text="245" FontSize="35" FontWeight="Bold" Foreground="#18181B" Margin="0,8,0,0"/>
        <TextBlock Grid.Row="2" Text="Subtitle text" FontSize="12.5" Foreground="#71717A" Margin="0,4,0,0"/>
    </Grid>
</Border>
```

Dark accent card (for highlight metrics):
```xml
<Border Background="#18181B" BorderThickness="0" CornerRadius="20" Padding="19,17">
    <!-- Same structure, text White and "#9A9AA2" for label -->
</Border>
```

---

## 12. DataGrid

```xml
<DataGrid AutoGenerateColumns="False" IsReadOnly="True"
          SelectionMode="Extended" SelectionUnit="FullRow" CanUserSortColumns="True"
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
    <DataGrid.Columns>
        <DataGridTextColumn Header="NAME" Binding="{Binding Name}" Width="*"/>
    </DataGrid.Columns>
</DataGrid>
```

---

## 13. Progress Bar + Pause/Stop (batch tools)

For any tool that loops over many items (export, create, rename…), use the shared
progress panel + `ProgressPauseMixin` pattern. Canonical live example: `AutoJoin.xaml`
+ `AutoJoin.pushbutton/script.py`.

**Python side** — inherit the mixin (`T3Lab.extension/lib/GUI/ProgressPauseMixin.py`):

```python
from GUI.ProgressPauseMixin import ProgressPauseMixin

class MyToolWindow(forms.WPFWindow, ProgressPauseMixin):
    # Override PP_BAR / PP_PAUSE / PP_STOP / PP_PANEL / PP_STATUS
    # class attrs if the XAML uses different x:Name values.

    def run_clicked(self, sender, e):
        items = collect_items()
        self.begin_progress(len(items), disable=[self.btn_run])
        for i, item in enumerate(items):
            if not self.step_progress(i, "Processing {}...".format(item)):
                break                      # user pressed Stop
            process(item)                  # one item per step
        cancelled = self.is_cancelled      # read BEFORE end_progress()
        self.end_progress()
```

Rules: modal windows (`ShowDialog`) only — never pump the dispatcher inside a
modeless ExternalEvent `Execute()`. Prefer per-item/chunk transactions
(`TransactionGroup`) so pause never holds a transaction open.

**XAML side** — place inside the status-bar row, above the copyright line
(`Collapsed` when idle):

```xml
<!-- Progress panel: bar + Pause + Stop (hidden when idle) -->
<Grid x:Name="progress_panel" Visibility="Collapsed" Margin="0,0,0,6">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="10"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="6"/>
        <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <ProgressBar x:Name="pb_run" Grid.Column="0" Style="{StaticResource T3ProgressBar}"
                 Minimum="0" Maximum="100" Value="0" VerticalAlignment="Center"/>

    <!-- Pause / Resume — the mixin toggles btn_pause_icon/btn_pause_label
         (Pause &#xE769; ↔ Play &#xE768;) when these x:Names are present -->
    <Button x:Name="btn_pause" Grid.Column="2" Style="{StaticResource SecondaryButton}"
            Height="28" Padding="14,0" FontSize="12"
            Click="pause_resume_clicked">
        <StackPanel Orientation="Horizontal">
            <TextBlock x:Name="btn_pause_icon" Text="&#xE769;" FontFamily="Segoe MDL2 Assets" FontSize="10" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBlock x:Name="btn_pause_label" Text="Pause" VerticalAlignment="Center"/>
        </StackPanel>
    </Button>

    <!-- Stop -->
    <Button x:Name="btn_stop" Grid.Column="4" Style="{StaticResource DangerButton}"
            Height="28" Padding="14,0" FontSize="12"
            Click="stop_clicked">
        <StackPanel Orientation="Horizontal">
            <TextBlock Text="&#xE71A;" FontFamily="Segoe MDL2 Assets" FontSize="10" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBlock Text="Stop" VerticalAlignment="Center"/>
        </StackPanel>
    </Button>
</Grid>
```

Icon rule: button icons are **Segoe MDL2 Assets glyphs** in an embedded TextBlock
(no `Foreground` — it inherits from the button style), never color emoji
(💾/📂/▶). Common glyphs: Save `&#xE74E;`, Open `&#xE8E5;`, Play `&#xE768;`,
Pause `&#xE769;`, Stop `&#xE71A;`, Add `&#xE710;`, Remove/X `&#xE711;`,
Switch `&#xE8AB;`.

Plain bar (no pause — quick loads only):

```xml
<ProgressBar Height="8" Style="{StaticResource T3ProgressBar}"/>
```

---

## 14. Tab Selector (pill-style)

```xml
<Border Background="#F4F4F6" CornerRadius="18" Padding="4" Height="36" HorizontalAlignment="Left">
    <StackPanel Orientation="Horizontal">
        <!-- Active tab -->
        <Border Background="White" CornerRadius="14" Padding="12,6">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="Active Tab" FontSize="13.5" FontWeight="SemiBold" Foreground="#18181B"/>
                <Border Background="#18181B" CornerRadius="9" MinWidth="18" Height="18"
                        Margin="6,0,0,0" VerticalAlignment="Center">
                    <TextBlock Text="9" FontSize="11" FontWeight="Bold" Foreground="White"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </StackPanel>
        </Border>
        <!-- Inactive tab -->
        <Border Background="Transparent" Padding="12,6">
            <TextBlock Text="Inactive Tab" FontSize="13.5" Foreground="#71717A"/>
        </Border>
    </StackPanel>
</Border>
```
