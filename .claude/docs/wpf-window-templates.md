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
            <TextBlock Text="T3Lab" FontSize="11" FontWeight="Bold" Foreground="#0F766E"
                       VerticalAlignment="Bottom" Margin="0,0,6,3"/>
            <TextBlock Text="Tool Name" FontSize="18" FontWeight="Bold"
                       Foreground="#1C2B33"/>
        </StackPanel>
        <Separator Height="1" Background="#DDE5E7" Margin="0,2,0,2"/>
        <TextBlock Text="Short description of the tool"
                   FontSize="10" Foreground="#64748B" FontStyle="Italic"/>
    </StackPanel>

    <!-- Right: Window control buttons -->
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top"
                WindowChrome.IsHitTestVisibleInChrome="True">
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
    </StackPanel>

    <Border Height="1" VerticalAlignment="Bottom" Background="#DDE5E7"/>
</Grid>
```

### 3. Status Bar Row
```xml
<Border Grid.Row="N" Background="#F6F8F8" BorderBrush="#C7D2D4" BorderThickness="0,1,0,0"
        Padding="14,8">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="status_text" Grid.Column="0" FontSize="11"
                   Foreground="#64748B" Text="Ready"/>
        <TextBlock x:Name="count_text" Grid.Column="1" FontSize="11"
                   Foreground="#DDE5E7" Text="0 items"/>
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
<!-- PRIMARY BUTTON - blue, white text -->
<Style x:Key="PrimaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#0F766E"/>
    <Setter Property="Foreground"      Value="White"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="12"/>
    <Setter Property="FontFamily"      Value="Inter"/>
    <Setter Property="Cursor"          Value="Hand"/>
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
            <Setter Property="Background" Value="#115E59"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
            <Setter Property="Background" Value="#DDE5E7"/>
            <Setter Property="Cursor"     Value="Arrow"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SECONDARY BUTTON - light gray, dark text -->
<Style x:Key="SecondaryButton" TargetType="Button">
    <Setter Property="Background"      Value="#F6F8F8"/>
    <Setter Property="Foreground"      Value="#1C2B33"/>
    <Setter Property="Padding"         Value="14,7"/>
    <Setter Property="FontSize"        Value="12"/>
    <Setter Property="FontFamily"      Value="Inter"/>
    <Setter Property="Cursor"          Value="Hand"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush"     Value="#DDE5E7"/>
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
            <Setter Property="Background" Value="#E6EDEC"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- SUCCESS BUTTON - green -->
<Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#15803D"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#166534"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- DANGER BUTTON - red (delete/destructive) -->
<Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
    <Setter Property="Background" Value="#DC2626"/>
    <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
            <Setter Property="Background" Value="#B91C1C"/>
        </Trigger>
    </Style.Triggers>
</Style>

<!-- WINDOW CONTROL BUTTON -->
<Style x:Key="WinCtrlButton" TargetType="Button">
    <Setter Property="Width"           Value="40"/>
    <Setter Property="Height"          Value="32"/>
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
            <Setter Property="Background" Value="#F6F8F8"/>
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
                        <Setter TargetName="bd" Property="Background" Value="#DC2626"/>
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
<DataGrid Background="White" BorderBrush="#C7D2D4" BorderThickness="1"
          AlternatingRowBackground="#F6F8F8" FontFamily="Inter" FontSize="12"
          HeadersVisibility="Column" GridLinesVisibility="Horizontal"
          HorizontalGridLinesBrush="#F6F8F8">
    <DataGrid.ColumnHeaderStyle>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"      Value="#F6F8F8"/>
            <Setter Property="Foreground"      Value="#1C2B33"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="8,6"/>
            <Setter Property="BorderBrush"     Value="#DDE5E7"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Height"          Value="34"/>
        </Style>
    </DataGrid.ColumnHeaderStyle>
    <DataGrid.RowStyle>
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#F6F8F8"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#E6EDEC"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </DataGrid.RowStyle>
</DataGrid>
```

---

## Info / Tip Box

```xml
<Border BorderBrush="#0F766E" BorderThickness="1" Background="#F6F8F8"
        CornerRadius="2" Padding="10">
    <StackPanel Orientation="Horizontal">
        <TextBlock Text="Tip:" FontWeight="Bold" Foreground="#115E59" Margin="0,0,5,0"/>
        <TextBlock Text="Your message here." Foreground="#1C2B33" TextWrapping="Wrap"/>
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
                 Foreground="#0F766E" Background="#E6EDEC"
                 BorderThickness="0" VerticalAlignment="Center"/>

    <!-- Pause / Resume button -->
    <Button x:Name="btn_pause_export" Grid.Column="2"
            Content="⏸  Pause" Height="22" FontSize="10" FontFamily="Inter"
            Click="pause_resume_clicked" Cursor="Hand"
            BorderThickness="1" BorderBrush="#C7D2D4" VerticalAlignment="Center">
        <!-- Use inline secondary-like style -->
    </Button>

    <!-- Stop button -->
    <Button x:Name="btn_stop_export" Grid.Column="4"
            Content="■  Stop" Height="22" FontSize="10" FontFamily="Inter"
            Click="stop_export_clicked" Cursor="Hand"
            BorderThickness="0" VerticalAlignment="Center">
        <!-- Use inline danger-like style: Bg=#DC2626, hover=#B91C1C -->
    </Button>
</Grid>
```
