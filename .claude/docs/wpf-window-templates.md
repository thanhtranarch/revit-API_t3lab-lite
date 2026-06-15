# WPF Window Templates

> **IMPORTANT**:
> All templates below align with [UIStandardShowcase.xaml](.claude/standard/UIStandardShowcase.xaml) — the definitive master visual standard. 
> To ensure the codebase is portable and works across different developer machines and directory locations, **never use absolute file paths** (such as `file:///C:/...` or `D:\...`) in documentation, code references, or agent prompts. Always write paths relatively (e.g. `.claude/standard/UIStandardShowcase.xaml`).

## Variant A — Standard Tool Window Template

Every new standard tool window should follow this layout. It includes the multi-line `WindowChrome` definition, standard window control button text blocks, and the copyright overlay:

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

    <!-- Custom Title Bar Chrome (Multiline tag is required) -->
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="22"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Window.Resources>
        <!-- Paste shared styles here from T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml -->
    </Window.Resources>

    <!-- Outer border to separate window from white Revit worksheets -->
    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7">
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="66"/>     <!-- Left Sidebar Icon Rail -->
                <ColumnDefinition Width="*"/>      <!-- Right main content area -->
            </Grid.ColumnDefinitions>

            <Grid.RowDefinitions>
                <RowDefinition Height="64"/>       <!-- Row 0: Title Header / Top Bar -->
                <RowDefinition Height="*"/>        <!-- Row 1: Content Area -->
                <RowDefinition Height="Auto"/>     <!-- Row 2: Footer / Status Bar -->
            </Grid.RowDefinitions>

            <!-- ═══ HEADER ROW (Row 0, Column 1) ═══ -->
            <Grid Grid.Row="0" Grid.Column="1" Background="#F4F4F6">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Title & Subtitle -->
                <StackPanel Grid.Column="0" Orientation="Vertical" Margin="22,0,0,0" VerticalAlignment="Center"
                            WindowChrome.IsHitTestVisibleInChrome="True">
                    <TextBlock Text="Tool Name" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
                    <TextBlock Text="Lumina Suite · Revit 2024–2026" FontSize="12.5" Foreground="#71717A" Margin="0,2,0,0"/>
                </StackPanel>

                <!-- Window controls -->
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,16,0">
                    <Button x:Name="btn_minimize" WindowChrome.IsHitTestVisibleInChrome="True" Style="{StaticResource WinCtrlButton}" Click="minimize_button_clicked" ToolTip="Minimize">
                        <TextBlock Text="&#xE921;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
                    </Button>
                    <Button x:Name="btn_maximize" WindowChrome.IsHitTestVisibleInChrome="True" Style="{StaticResource WinCtrlButton}" Click="maximize_button_clicked" ToolTip="Maximize">
                        <TextBlock Text="&#xE922;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
                    </Button>
                    <Button x:Name="btn_close" WindowChrome.IsHitTestVisibleInChrome="True" Style="{StaticResource CloseButton}" Click="close_button_clicked" ToolTip="Close">
                        <TextBlock Text="&#xE8BB;" FontFamily="Segoe MDL2 Assets" FontSize="10"/>
                    </Button>
                </StackPanel>
                
                <Border Height="1" VerticalAlignment="Bottom" Background="#DCDCE0" Grid.ColumnSpan="3"/>
            </Grid>

            <!-- ═══ CONTENT AREA (Row 1, Column 1) ═══ -->
            <Grid Grid.Row="1" Grid.Column="1" Margin="18,18,18,10">
                <!-- Add your controls here -->
            </Grid>

            <!-- ═══ FOOTER ROW (Row 2, Column 1) ═══ -->
            <Border Grid.Row="2" Grid.Column="1" Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,16">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left: Action buttons -->
                    <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
                        <Button Style="{StaticResource PrimaryButton}" Content="Execute" Margin="0,0,8,0" Padding="16,10"/>
                        <Button Style="{StaticResource SecondaryButton}" Content="Cancel" Padding="16,9"/>
                    </StackPanel>

                    <!-- Right: System status and copyright in separate vertical lines -->
                    <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="16,0,0,0">
                        <Ellipse Width="8" Height="8" Fill="#22A85C" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBlock Text="System status: " FontSize="13.5" Foreground="#27272A" FontWeight="SemiBold" VerticalAlignment="Center"/>
                        <TextBlock Text="Ready" FontSize="13.5" Foreground="#157038" FontWeight="Bold" Margin="0,0,16,0" VerticalAlignment="Center"/>
                        <Border Width="1" Height="18" Background="#DEDEE2" Margin="0,0,16,0" VerticalAlignment="Center"/>
                        <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                            <TextBlock Text="© 2026 T3Lab · v2.4" FontSize="11" Foreground="#9A9AA2" HorizontalAlignment="Left"/>
                            <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
```

---

## Variant B — Modal Dialog Content Template

Use this format **only** when parsing a Grid directly to the content of a borderless modal window generated inside Python code.

```xml
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      Width="480" Height="320"
      Background="#FFFFFF"
      FontFamily="Hanken Grotesk"
      FontSize="14">

    <Grid.Resources>
        <!-- Paste shared styles here from T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml -->
    </Grid.Resources>

    <!-- Outer border to separate dialog from drawing environment -->
    <Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="16" Background="#FFFFFF">
        <Grid Margin="16">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/> <!-- Header -->
                <RowDefinition Height="*"/>    <!-- Content -->
                <RowDefinition Height="Auto"/> <!-- Footer -->
            </Grid.RowDefinitions>

            <!-- Header -->
            <TextBlock Grid.Row="0" Text="Select Parameter" FontSize="15" FontWeight="Bold" Foreground="#18181B" Margin="0,0,0,12"/>

            <!-- Content -->
            <StackPanel Grid.Row="1" VerticalAlignment="Center">
                <!-- Inputs go here -->
            </StackPanel>

            <!-- Footer -->
            <Grid Grid.Row="2" Margin="0,16,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button Style="{StaticResource SecondaryButton}" Content="Cancel" Margin="0,0,8,0"/>
                    <Button Style="{StaticResource PrimaryButton}" Content="Select"/>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Grid>
```
