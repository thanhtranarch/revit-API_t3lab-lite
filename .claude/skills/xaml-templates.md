---
name: xaml-templates
description: XAML templates for T3Lab WPF windows. Contains complete snippets for window structure, title bar, toolbar, status bar, all button styles, DataGrid styling, and info boxes. All templates follow the BatchOut white light theme.
---

# XAML Templates

## Window Structure

Every tool window must include:

### 1. WindowChrome (custom title bar)
```xml
<Window Background="White" ResizeMode="CanResizeWithGrip" ...>
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="64"
                      ResizeBorderThickness="5"
                      GlassFrameThickness="0"
                      CornerRadius="0"
                      UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>
```

### 2. Title Bar Row (64px, white)
```xml
<Grid Height="64" Background="White">
    <!-- Left: Logo + "T3Lab" (blue) + Tool Name (dark) + Italic subtitle (gray) -->
    <StackPanel Orientation="Horizontal" Margin="16,0,0,0" VerticalAlignment="Center"
                WindowChrome.IsHitTestVisibleInChrome="True">
        <Image x:Name="logo_image" Width="40" Height="40" Margin="0,0,10,0"/>
        <StackPanel VerticalAlignment="Center">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="T3Lab"     FontSize="11" FontWeight="Bold"  Foreground="#3498DB" .../>
                <TextBlock Text="Tool Name" FontSize="17" FontWeight="Bold"  Foreground="#2C3E50" .../>
            </StackPanel>
            <Separator Height="1" Background="#BDC3C7" Margin="0,3"/>
            <TextBlock Text="Short description" FontSize="10" Foreground="#7F8C8D" FontStyle="Italic"/>
        </StackPanel>
    </StackPanel>

    <!-- Right: Window control buttons (Minimize / Maximize / Close) -->
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top"
                WindowChrome.IsHitTestVisibleInChrome="True">
        <Button x:Name="btn_minimize" Style="{StaticResource WinCtrlButton}" Click="minimize_button_clicked"/>
        <Button x:Name="btn_maximize" Style="{StaticResource WinCtrlButton}" Click="maximize_button_clicked"/>
        <Button x:Name="btn_close"    Style="{StaticResource CloseButton}"   Click="close_button_clicked"/>
    </StackPanel>

    <Border Height="1" VerticalAlignment="Bottom" Background="#BDC3C7"/>
</Grid>
```

### 3. Toolbar Row
```xml
<Border Background="White" BorderBrush="#BDC3C7" BorderThickness="0,0,0,1" Padding="12,8">
    <WrapPanel>
        <!-- Primary actions first, then secondary, separated by thin rectangles -->
        <Rectangle Width="1" Fill="#BDC3C7" Margin="0,2,12,2"/> <!-- group separator -->
    </WrapPanel>
</Border>
```

### 4. Status Bar Row
```xml
<Border Background="#FAFAFA" BorderBrush="#BDC3C7" BorderThickness="0,1,0,0" Padding="14,6">
    <Grid>
        <TextBlock x:Name="status_text" FontSize="11" Foreground="#7F8C8D"/>
        <TextBlock Grid.Column="1"      FontSize="11" Foreground="#BDC3C7" HorizontalAlignment="Right"/>
    </Grid>
</Border>
```

---

## Button Styles

Define these as `Window.Resources`:

```xml
<!-- PRIMARY - blue, white text -->
<Style x:Key="PrimaryButton" TargetType="Button">
    <Setter Property="Background"   Value="#3498DB"/>
    <Setter Property="Foreground"   Value="White"/>
    <Setter Property="Padding"      Value="12,6"/>
    <Setter Property="FontSize"     Value="12"/>
    <Setter Property="FontFamily"   Value="Segoe UI"/>
    <Setter Property="Cursor"       Value="Hand"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}" CornerRadius="3"
                        Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#2980B9"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#BDC3C7"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SECONDARY - light gray, dark text -->
<Style x:Key="SecondaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#ECF0F1"/>
    <Setter Property="Foreground"      Value="#2C3E50"/>
    <Setter Property="Padding"         Value="12,6"/>
    <Setter Property="FontSize"        Value="12"/>
    <Setter Property="FontFamily"      Value="Segoe UI"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush"     Value="#BDC3C7"/>
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}"
                        BorderBrush="{TemplateBinding BorderBrush}"
                        BorderThickness="{TemplateBinding BorderThickness}"
                        CornerRadius="3" Padding="{TemplateBinding Padding}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#D5DBDB"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- DANGER - red (delete/destructive) -->
<Style x:Key="DangerButton"  TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#E74C3C"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#C0392B"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SUCCESS - green (apply/confirm) -->
<Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#27AE60"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#1E8449"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- WINDOW CONTROL - transparent, ECF0F1 on hover -->
<Style x:Key="WinCtrlButton" TargetType="Button">
    <Setter Property="Width"  Value="40"/>
    <Setter Property="Height" Value="32"/>
    <Setter Property="Background"      Value="Transparent"/>
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
            <Setter Property="Background" Value="#ECF0F1"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- CLOSE BUTTON - red on hover -->
<Style x:Key="CloseButton" TargetType="Button" BasedOn="{StaticResource WinCtrlButton}">
    <Setter Property="Template">
        <Setter.Value>
            <ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="{TemplateBinding Background}">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="bd" Property="Background" Value="#E74C3C"/>
                    </Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate>
        </Setter.Value>
    </Setter>
</Style>
```

---

## DataGrid Style

```xml
<DataGrid Background="White" BorderBrush="#BDC3C7" BorderThickness="1"
          AlternatingRowBackground="#FAFAFA" FontFamily="Segoe UI" FontSize="12">
    <DataGrid.ColumnHeaderStyle>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"   Value="#ECF0F1"/>
            <Setter Property="Foreground"   Value="#2C3E50"/>
            <Setter Property="FontWeight"   Value="SemiBold"/>
            <Setter Property="Padding"      Value="8,6"/>
            <Setter Property="BorderBrush"  Value="#BDC3C7"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Height"       Value="34"/>
        </Style>
    </DataGrid.ColumnHeaderStyle>
    <DataGrid.RowStyle>
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#EBF5FB"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#D6EAF8"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </DataGrid.RowStyle>
</DataGrid>
```

---

## Info / Tip Box

```xml
<Border BorderBrush="#3498DB" BorderThickness="1" Background="#E8F4F8"
        CornerRadius="2" Padding="10">
    <StackPanel Orientation="Horizontal">
        <TextBlock Text="Tip:" FontWeight="Bold" Foreground="#2980B9" Margin="0,0,5,0"/>
        <TextBlock Text="Your message here." Foreground="#2C3E50"/>
    </StackPanel>
</Border>
```
