# WPF Window Templates

> **IMPORTANT**: All templates below align with `.claude/standard/UIStandardShowcase.xaml` — the definitive master visual standard. Always read that file before writing XAML to confirm current hex values and patterns. Never use absolute file paths (`C:\...`) — always write relative paths so the codebase remains portable.

---

## Variant A — Standard Tool Window

Every new tool window uses this structure. Includes outer Border wrapper, multi-line WindowChrome, sidebar icon rail, title bar, content area white card, and footer with copyright inside.

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="T3Lab - Tool Name"
        Width="1100" Height="680"
        MinWidth="860" MinHeight="500"
        Background="#E4E4E7"
        ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterScreen"
        FontFamily="Hanken Grotesk"
        FontSize="14">

    <!-- Multi-line WindowChrome — all 5 attributes required -->
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="22"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Window.Resources>
        <!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT
             (edit T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml,
              then run dev/sync_wpf_styles.py) ═══ -->
        <!-- paste full shared styles block here -->
        <!-- ═══ END T3LAB SHARED STYLES ═══ -->

        <!-- Window-specific styles go here, outside the markers -->
    </Window.Resources>

    <!-- Outer border separates window from white Revit canvas -->
    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22"
            ClipToBounds="True" Background="#E4E4E7">
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="66"/>  <!-- Col 0: Sidebar icon rail -->
                <ColumnDefinition Width="*"/>   <!-- Col 1: Main content -->
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="64"/>    <!-- Row 0: Title bar -->
                <RowDefinition Height="*"/>     <!-- Row 1: Content -->
                <RowDefinition Height="Auto"/>  <!-- Row 2: Footer -->
            </Grid.RowDefinitions>

            <!-- ═══ SIDEBAR (Col 0, all rows) ═══ -->
            <Border Grid.Column="0" Grid.RowSpan="3"
                    Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,0,1,0">
                <Grid Margin="0,16,0,16">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/> <!-- Logo -->
                        <RowDefinition Height="Auto"/> <!-- Nav 1 (active) -->
                        <RowDefinition Height="Auto"/> <!-- Nav 2 -->
                        <RowDefinition Height="Auto"/> <!-- Nav 3 -->
                        <RowDefinition Height="*"/>    <!-- Spacer -->
                    </Grid.RowDefinitions>

                    <!-- Logo -->
                    <Border Grid.Row="0" Width="42" Height="42" CornerRadius="13"
                            Background="White" BorderBrush="#DCDCE0" BorderThickness="1"
                            HorizontalAlignment="Center" Margin="0,0,0,24">
                        <Path Data="M12 4 L20 8 L12 12 L4 8 Z M4 12 L12 16 L20 12 M4 16 L12 20 L20 16"
                              Stroke="#18181B" StrokeThickness="1.8" StrokeLineJoin="Round"
                              Width="20" Height="20" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>

                    <!-- Active nav button (solid dark background) -->
                    <Border Grid.Row="1" Width="42" Height="42" CornerRadius="13"
                            Background="#18181B" HorizontalAlignment="Center" Margin="0,0,0,8">
                        <!-- Replace with tool-specific icon path, Stroke="White" -->
                        <Path Data="M4 4 h7 v7 h-7 z M13 4 h7 v7 h-7 z M4 13 h7 v7 h-7 z M13 13 h7 v7 h-7 z"
                              Stroke="White" StrokeThickness="1.8" StrokeLineJoin="Round"
                              Width="19" Height="19" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>

                    <!-- Inactive nav buttons -->
                    <Button Grid.Row="2" Style="{StaticResource T3SidebarButton}" Margin="0,0,0,8">
                        <!-- Replace with tool-specific icon path, Stroke="#71717A" -->
                    </Button>
                </Grid>
            </Border>

            <!-- ═══ TITLE BAR (Row 0, Col 1) ═══ -->
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

            <!-- ═══ CONTENT AREA (Row 1, Col 1) ═══ -->
            <Grid Grid.Row="1" Grid.Column="1" Margin="18,18,18,10">
                <!-- Example two-panel layout -->
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="320"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Left white card -->
                <Border Grid.Column="0" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
                        CornerRadius="20" Padding="18" Margin="0,0,16,0">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <!-- Controls here -->
                        </StackPanel>
                    </ScrollViewer>
                </Border>

                <!-- Right white card -->
                <Border Grid.Column="1" Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
                        CornerRadius="20" Padding="14">
                    <!-- DataGrid or other content here -->
                </Border>
            </Grid>

            <!-- ═══ FOOTER (Row 2, Col 1) ═══ -->
            <Border Grid.Row="2" Grid.Column="1"
                    Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,16">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left: action buttons -->
                    <StackPanel Grid.Column="0" Orientation="Horizontal"
                                HorizontalAlignment="Left" VerticalAlignment="Center">
                        <Button Style="{StaticResource PrimaryButton}" Content="Execute"
                                Margin="0,0,8,0" Padding="16,10"/>
                        <Button Style="{StaticResource SecondaryButton}" Content="Configure"
                                Padding="16,9"/>
                    </StackPanel>

                    <!-- Right: status indicator + copyright -->
                    <StackPanel Grid.Column="1" Orientation="Horizontal"
                                HorizontalAlignment="Right" VerticalAlignment="Center" Margin="16,0,0,0">
                        <Ellipse Width="8" Height="8" Fill="#22A85C"
                                 VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock Text="System status: " FontSize="13.5" Foreground="#27272A"
                                   FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <TextBlock x:Name="status_text" Text="Ready" FontSize="13.5"
                                   Foreground="#157038" FontWeight="Bold"
                                   Margin="0,0,16,0" VerticalAlignment="Center"/>
                        <Border Width="1" Height="18" Background="#DEDEE2"
                                Margin="0,0,16,0" VerticalAlignment="Center"/>
                        <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                            <TextBlock Text="© 2026 T3Lab · v2.4" FontSize="11" Foreground="#9A9AA2"/>
                            <TextBlock Text="© Copyright by T3Lab" FontSize="11"
                                       Foreground="#F59E0B" Margin="0,2,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </Border>

        </Grid>
    </Border>
</Window>
```

---

## Variant A — Simple Dialog Window (no sidebar)

For smaller tools (feedback, settings, pickers) that don't need a sidebar:

```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="T3Lab - Tool Name"
        Width="600" Height="640"
        MinWidth="520" MinHeight="560"
        Background="#E4E4E7"
        ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterScreen"
        FontFamily="Hanken Grotesk">

    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="22"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Window.Resources>
        <!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT ═══ -->
        <!-- paste full shared styles block -->
        <!-- ═══ END T3LAB SHARED STYLES ═══ -->

        <!-- Window-specific field label style -->
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#9A9AA2"/>
            <Setter Property="Margin" Value="1,0,0,6"/>
        </Style>
    </Window.Resources>

    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22"
            ClipToBounds="True" Background="#E4E4E7">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="64"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Title Bar -->
            <Grid Grid.Row="0" Background="#F4F4F6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="22,0,0,0" VerticalAlignment="Center">
                    <TextBlock Text="Tool Name" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
                    <TextBlock Text="Tool subtitle" FontSize="12.5" Foreground="#71717A" Margin="0,2,0,0"/>
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

            <!-- Content (white card) -->
            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="18,18,18,0">
                <Border Background="White" BorderBrush="#E2E8F0" BorderThickness="1"
                        CornerRadius="20" Padding="18" Margin="0,0,0,18">
                    <StackPanel>
                        <!-- Controls here -->
                    </StackPanel>
                </Border>
            </ScrollViewer>

            <!-- Footer -->
            <Border Grid.Row="2"
                    Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,14">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Status + copyright stacked -->
                    <StackPanel Grid.Column="0" VerticalAlignment="Center">
                        <TextBlock x:Name="status_text" Text="Ready"
                                   FontSize="13.5" Foreground="#71717A" FontWeight="SemiBold"/>
                        <TextBlock Text="© Copyright by T3Lab"
                                   FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0"/>
                    </StackPanel>

                    <!-- Action buttons -->
                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button x:Name="btn_cancel" Style="{StaticResource SecondaryButton}"
                                Content="Cancel" Width="100" Margin="0,0,8,0"
                                Click="close_button_clicked"/>
                        <Button x:Name="btn_ok" Style="{StaticResource SuccessButton}"
                                Content="Apply" Width="120"
                                Click="apply_clicked"/>
                    </StackPanel>
                </Grid>
            </Border>

        </Grid>
    </Border>
</Window>
```

---

## Variant B — Modal Dialog Content

Use **only** when parsing a `<Grid>` directly into a borderless Python-hosted `Window` with `WindowStyle=NoStyle`. Do not create new Variant B files unless explicitly required.

```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      Width="480" Height="320"
      Background="#E4E4E7"
      FontFamily="Hanken Grotesk">

    <Grid.Resources>
        <!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT ═══ -->
        <!-- paste full shared styles block -->
        <!-- ═══ END T3LAB SHARED STYLES ═══ -->
    </Grid.Resources>

    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22"
            ClipToBounds="True" Background="#E4E4E7">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Header -->
                <RowDefinition Height="*"/>    <!-- Content -->
                <RowDefinition Height="Auto"/> <!-- Footer -->
            </Grid.RowDefinitions>

            <!-- Header -->
            <Border Grid.Row="0" Background="#F4F4F6" Padding="18,14"
                    BorderBrush="#DCDCE0" BorderThickness="0,0,0,1">
                <TextBlock Text="Dialog Title" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
            </Border>

            <!-- Content (white card) -->
            <Border Grid.Row="1" Background="White" Margin="16" CornerRadius="16" Padding="16">
                <!-- Inputs here -->
            </Border>

            <!-- Footer -->
            <Border Grid.Row="2" Background="#F4F4F6" Padding="16,12"
                    BorderBrush="#DCDCE0" BorderThickness="0,1,0,0">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <StackPanel Orientation="Vertical" VerticalAlignment="Center" Margin="0,0,16,0">
                        <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B"/>
                    </StackPanel>
                    <Button Style="{StaticResource SecondaryButton}" Content="Cancel"
                            Width="90" Margin="0,0,8,0"/>
                    <Button Style="{StaticResource PrimaryButton}" Content="Select" Width="100"/>
                </StackPanel>
            </Border>
        </Grid>
    </Border>
</Grid>
```

---

## Critical Rules Summary

| Rule | Correct | Wrong |
|------|---------|-------|
| Window Background | `#E4E4E7` | `White` |
| WindowChrome CornerRadius | `22` | `8` |
| Outer Border CornerRadius | `22` | missing |
| Title bar Background | `#F4F4F6` | `White` |
| Title bar bottom border | `#DCDCE0` | `#E2E8F0` |
| Footer Background | `#F4F4F6` | `#F8FAFC` |
| Footer top border | `#DCDCE0` | `#E2E8F0` |
| Copyright placement | Inside footer column | Floating overlay on root Grid |
| PrimaryButton bg | `#18181B` | `#0F172A` |
| SuccessButton bg | `#22A85C` | `#10B981` |
| DangerButton bg | `#D23B3B` | `#EF4444` |
| Button CornerRadius | `12` | `6` |
| Button Height | `40` | `34` |
| Button FontSize | `13.5` | `12` |
| Input bg | `#F4F4F6` | `White` |
| Input border | `#E6E6EA` | `#CBD5E1` |
| Input CornerRadius | `13` | `4` or `6` |
| Field labels | ALL-CAPS `#9A9AA2` 11px Bold | lowercase `#64748B` |
