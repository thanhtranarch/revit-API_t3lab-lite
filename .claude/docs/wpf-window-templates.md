# WPF Window Templates

> **Reference**: All templates below match `BulkFamilyExport.xaml` — the canonical UI standard.

## Window Structure

Every tool window must include:

### 1. Window Root + WindowChrome
```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="T3Lab - Tool Name"
        Width="1100" Height="680"
        MinWidth="860" MinHeight="500"
        Background="White"
        ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterScreen"
        FontFamily="Inter">

    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="8"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>
```

### 2. Title Bar Row (64px, white)
```xml
<Grid Grid.Row="0" Height="64" Background="White">
    <StackPanel Orientation="Vertical" Margin="16,0,0,0" VerticalAlignment="Center"
                WindowChrome.IsHitTestVisibleInChrome="True">
        <StackPanel Orientation="Horizontal">
            <TextBlock Text="T3Lab" FontSize="11" FontWeight="Bold" Foreground="#0F172A"
                       VerticalAlignment="Bottom" Margin="0,0,6,3"/>
            <TextBlock Text="Tool Name" FontSize="18" FontWeight="Bold"
                       Foreground="#0F172A"/>
        </StackPanel>
        <Separator Height="1" Background="#E2E8F0" Margin="0,2,0,2"/>
        <TextBlock Text="Short description of the tool"
                   FontSize="10" Foreground="#64748B" FontStyle="Italic"/>
    </StackPanel>

    <!-- Right: Window control buttons -->
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

    <Border Height="1" VerticalAlignment="Bottom" Background="#E2E8F0"/>
</Grid>
```

### 3. Status Bar Row
```xml
<Border Grid.Row="N" Background="#F8FAFC" BorderBrush="#E2E8F0" BorderThickness="0,1,0,0"
        Padding="14,8">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="status_text" Grid.Column="0" FontSize="11"
                   Foreground="#64748B" Text="Ready"/>
        <TextBlock x:Name="count_text" Grid.Column="1" FontSize="11"
                   Foreground="#E2E8F0" Text="0 items"/>
    </Grid>
</Border>
```

### 4. Copyright (always before closing root Grid)
```xml
    <!-- Copyright added automatically -->
    <TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999"/>
</Grid>
```

---

## Button Styles

Define these as `Window.Resources`:

```xml
<!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT (edit lib/GUI/Resources/WPF_styles.xaml, then run dev/sync_wpf_styles.py) ═══ -->
<!-- PRIMARY BUTTON - solid deep slate, white text -->
<Style x:Key="PrimaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#0F172A"/>
    <Setter Property="Foreground"      Value="White"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="12"/>
    <Setter Property="FontFamily"      Value="Inter"/>
    <Setter Property="FontWeight"      Value="SemiBold"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}" CornerRadius="6"
                        Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#1E293B"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#0F172A"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#E2E8F0"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SECONDARY BUTTON - ghost style, border and text deep slate -->
<Style x:Key="SecondaryButton" TargetType="Button">
    <Setter Property="Background"      Value="Transparent"/>
    <Setter Property="Foreground"      Value="#0F172A"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="12"/>
    <Setter Property="FontFamily"      Value="Inter"/>
    <Setter Property="FontWeight"      Value="SemiBold"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush"     Value="#0F172A"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        BorderThickness="{TemplateBinding BorderThickness}"
                        CornerRadius="6" Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#F1F5F9"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#E2E8F0"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SUCCESS BUTTON - emerald green -->
<Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#10B981"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#059669"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#047857"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#E2E8F0"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- DANGER BUTTON - rose red -->
<Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#EF4444"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#DC2626"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#B91C1C"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#E2E8F0"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- ACCENT BUTTON - refined blue -->
<Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#3B82F6"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#2563EB"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#1D4ED8"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#E2E8F0"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- WINDOW CONTROL BUTTON -->
<Style x:Key="WinCtrlButton" TargetType="Button">
    <Setter Property="Width"           Value="40"/>
    <Setter Property="Height"          Value="32"/>
    <Setter Property="Background"      Value="Transparent"/>
    <Setter Property="Foreground"      Value="#0F172A"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="{TemplateBinding Background}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#F1F5F9"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- CLOSE BUTTON - rose red on hover, glyph flips to white -->
<Style x:Key="CloseButton" TargetType="Button" BasedOn="{StaticResource WinCtrlButton}">
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#EF4444"/>
            <Setter Property="Foreground" Value="White"/>
        </Trigger>
    </Style.Triggers>
</Style>
<!-- ═══ END T3LAB SHARED STYLES ═══ -->
```

---

## DataGrid Style

```xml
<DataGrid Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
          AlternatingRowBackground="#F8FAFC" FontFamily="Inter" FontSize="12"
          HeadersVisibility="Column" GridLinesVisibility="Horizontal"
          HorizontalGridLinesBrush="#F1F5F9">
    <DataGrid.ColumnHeaderStyle>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"      Value="#F8FAFC"/>
            <Setter Property="Foreground"      Value="#0F172A"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="8,6"/>
            <Setter Property="BorderBrush"     Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Height"          Value="34"/>
        </Style>
    </DataGrid.ColumnHeaderStyle>
    <DataGrid.RowStyle>
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#F1F5F9"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#E2E8F0"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </DataGrid.RowStyle>
</DataGrid>
```

---

## Info / Tip Box

```xml
<Border BorderBrush="#CBD5E1" BorderThickness="1" Background="#F8FAFC"
        CornerRadius="4" Padding="10">
    <StackPanel Orientation="Horizontal">
        <TextBlock Text="Tip:" FontWeight="Bold" Foreground="#3B82F6" Margin="0,0,5,0"/>
        <TextBlock Text="Your message here." Foreground="#0F172A" TextWrapping="Wrap"/>
    </StackPanel>
</Border>
```

---

## Progress Bar (for long-running tasks)

```xml
<!-- Place inside Status Bar, Visibility="Collapsed" when idle -->
<Grid x:Name="progress_panel" Visibility="Collapsed" Margin="0,0,0,6">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="8"/>
        <ColumnDefinition Width="64"/>
        <ColumnDefinition Width="4"/>
        <ColumnDefinition Width="52"/>
    </Grid.ColumnDefinitions>

    <ProgressBar x:Name="pb_export" Grid.Column="0"
                 Height="8" Minimum="0" Maximum="100" Value="0"
                 Foreground="#3B82F6" Background="#E2E8F0"
                 BorderThickness="0" VerticalAlignment="Center"/>

    <!-- Pause / Resume button -->
    <Button x:Name="btn_pause_export" Grid.Column="2"
            Content="⏸  Pause" Height="22" FontSize="10" FontFamily="Inter"
            Click="pause_resume_clicked" Cursor="Hand"
            BorderThickness="1" BorderBrush="#CBD5E1" VerticalAlignment="Center">
        <!-- Use inline secondary-like style -->
    </Button>

    <!-- Stop button -->
    <Button x:Name="btn_stop_export" Grid.Column="4"
            Content="■  Stop" Height="22" FontSize="10" FontFamily="Inter"
            Click="stop_export_clicked" Cursor="Hand"
            BorderThickness="0" VerticalAlignment="Center">
        <!-- Use inline danger-like style: Bg=#EF4444, hover=#DC2626 -->
    </Button>
</Grid>
```
