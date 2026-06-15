---
name: xaml-templates
description: XAML templates for T3Lab WPF windows. Contains complete snippets for window structure, title bar, status bar, scrollbar custom styles, all button styles, DataGrid styling, inputs, and info boxes. All templates follow the definitive UIStandardShowcase.xaml light theme.
---

# XAML Templates (Lumina System)

> **IMPORTANT**:
> Always refer to the master standard [UIStandardShowcase.xaml](.claude/standard/UIStandardShowcase.xaml) inside the repository workspace for definitive styling, sizing, and colors. Do not hardcode absolute file paths (`C:\...` or `D:\...`) in documentation or agent instructions. Always write relative paths (e.g. `.claude/standard/UIStandardShowcase.xaml`) so the codebase remains portable.

## Window Structure

### 1. Window Root + WindowChrome
```xml
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="T3Lab - Tool Title"
        Height="680" Width="1100"
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
```

### 2. Title Bar Row (64px, background #F4F4F6)
```xml
<Grid Grid.Row="0" Grid.Column="1" Background="#F4F4F6">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/> <!-- Title and Subtitle -->
        <ColumnDefinition Width="*"/>    <!-- Spacer or Search Bar -->
        <ColumnDefinition Width="Auto"/> <!-- Right Actions and Window Controls -->
    </Grid.ColumnDefinitions>

    <!-- Title and Subtitle -->
    <StackPanel Grid.Column="0" Orientation="Vertical" Margin="22,0,0,0" VerticalAlignment="Center">
        <TextBlock Text="UI Standard Showcase" FontSize="15" FontWeight="Bold" Foreground="#18181B"/>
        <TextBlock Text="Lumina Compliance Reviewer · Revit 2024–2026" FontSize="12.5" Foreground="#71717A" Margin="0,2,0,0"/>
    </StackPanel>

    <!-- Right Controls -->
    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,16,0">
        <!-- Min / Max / Close Window Control Buttons -->
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
```

### 3. Status Bar Row (Footer)
Always use two explicit columns in the Grid to prevent the action buttons and status elements from overlapping on small screen sizes.
```xml
<Border Grid.Row="3" Grid.Column="1" Background="#F4F4F6" BorderBrush="#DCDCE0" BorderThickness="0,1,0,0" Padding="20,16">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Left: Action buttons stack -->
        <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
            <Button Style="{StaticResource PrimaryButton}" Content="Run check" Margin="0,0,8,0" Padding="16,10"/>
            <Button Style="{StaticResource SecondaryButton}" Content="Configure" Margin="0,0,8,0" Padding="16,9"/>
        </StackPanel>

        <!-- Right: Status info stack (Status aligned horizontally, copyright on 2 vertical lines) -->
        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="16,0,0,0">
            <Ellipse Width="8" Height="8" Fill="#22A85C" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBlock Text="System status: " FontSize="13.5" Foreground="#27272A" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <TextBlock Text="Fully compliant" FontSize="13.5" Foreground="#157038" FontWeight="Bold" Margin="0,0,16,0" VerticalAlignment="Center"/>
            <Border Width="1" Height="18" Background="#DEDEE2" Margin="0,0,16,0" VerticalAlignment="Center"/>
            <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                <TextBlock Text="© 2026 T3Lab · v2.4" FontSize="11" Foreground="#9A9AA2" HorizontalAlignment="Left"/>
                <TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0" HorizontalAlignment="Left"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Border>
```

---

## ScrollBar Custom Style (Ultra-Thin)

Implicitly styled for all scrollbars (targets `ScrollBar` and `Thumb`).
Always override `MinWidth` and `MinHeight` to `0` to allow them to shrink below system defaults.

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

## Button Styles

Define these in `Window.Resources`:

```xml
<!-- PRIMARY BUTTON - solid deep slate, white text -->
<Style x:Key="PrimaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#18181B"/>
    <Setter Property="Foreground"      Value="White"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="13.5"/>
    <Setter Property="FontFamily"      Value="Hanken Grotesk"/>
    <Setter Property="FontWeight"      Value="SemiBold"/>
    <Setter Property="Height"          Value="40"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}" CornerRadius="12"
                        Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#000000"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#18181B"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#F4F4F6"/>
            <Setter Property="Foreground" Value="#9A9AA2"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SECONDARY BUTTON - white background, border and text slate -->
<Style x:Key="SecondaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#FFFFFF"/>
    <Setter Property="Foreground"      Value="#27272A"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="13.5"/>
    <Setter Property="FontFamily"      Value="Hanken Grotesk"/>
    <Setter Property="FontWeight"      Value="SemiBold"/>
    <Setter Property="Height"          Value="40"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush"     Value="#DEDEE2"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        BorderThickness="{TemplateBinding BorderThickness}"
                        CornerRadius="12" Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#FAFAFB"/>
            <Setter Property="BorderBrush" Value="#CFCFD5"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#EDEDEF"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Foreground" Value="#9A9AA2"/>
            <Setter Property="BorderBrush" Value="#E6E6EA"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- TERTIARY BUTTON - soft gray background, dark text -->
<Style x:Key="TertiaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#ECECEF"/>
    <Setter Property="Foreground"      Value="#27272A"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="13.5"/>
    <Setter Property="FontFamily"      Value="Hanken Grotesk"/>
    <Setter Property="FontWeight"      Value="SemiBold"/>
    <Setter Property="Height"          Value="40"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}" CornerRadius="12"
                        Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#E2E2E6"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#ECECEF"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#F4F4F6"/>
            <Setter Property="Foreground" Value="#9A9AA2"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SUCCESS BUTTON - emerald green -->
<Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#22A85C"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#0F8A66"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#22A85C"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- DANGER BUTTON - rose red -->
<Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#D23B3B"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#F87171"/>
            <Setter Property="Foreground" Value="#1F1F23"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
            <Setter Property="Background" Value="#D23B3B"/>
            <Setter Property="Foreground" Value="White"/>
        </Trigger>
    </Style.Triggers>
</Style>
