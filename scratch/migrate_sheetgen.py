# -*- coding: utf-8 -*-
import io

xaml_content = '''<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        mc:Ignorable="d"
        Title="Create Room Plan"
        Width="1100" Height="720" MinWidth="1000" MinHeight="620"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None" AllowsTransparency="False"
        Background="{DynamicResource T3.Canvas}"
        FontFamily="Segoe UI"
        FontSize="13"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">

    <WindowChrome.WindowChrome>
        <shell:WindowChrome CaptionHeight="40"
                            ResizeBorderThickness="6"
                            CornerRadius="0"
                            GlassFrameThickness="0"
                            UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Window.Resources>
        <ResourceDictionary>
<!-- ═══ T3 STYLES — SINH TỰ ĐỘNG, ĐỪNG SỬA Ở ĐÂY. Sửa `pyRevit UI Design System/T3Lab.Styles.xaml` rồi chạy `python3 dev/sync_t3_styles.py` ═══ -->
<!-- ═══ HẾT T3 STYLES ═══ -->
        </ResourceDictionary>
    </Window.Resources>

    <!-- ═══ Outer Window Border ═══ -->
    <Border BorderBrush="{StaticResource T3.BorderStrong}" BorderThickness="1.5"
            CornerRadius="8" ClipToBounds="True" Background="{DynamicResource T3.Canvas}">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="40"/>      <!-- Row 0: Full-width Title Bar -->
                <RowDefinition Height="*"/>       <!-- Row 1: Main Content & Sidebar -->
                <RowDefinition Height="48"/>      <!-- Row 2: Full-width Footer Bar -->
            </Grid.RowDefinitions>

            <!-- ══════════════════ ROW 0: FULL-WIDTH TITLE BAR ══════════════════ -->
            <Border Grid.Row="0" Style="{StaticResource T3.TitleBar}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Create Room Plan" Style="{StaticResource T3.Title}"/>
                        <Border Width="1" Height="14" Margin="12,0"
                                Background="{StaticResource T3.Border}" VerticalAlignment="Center"/>
                        <TextBlock x:Name="lbl_subtitle" Text="Generate plan views and sheet layouts from selected rooms"
                                   Style="{StaticResource T3.Body.Secondary}"/>
                    </StackPanel>
                    <StackPanel Grid.Column="1" Orientation="Horizontal"
                                shell:WindowChrome.IsHitTestVisibleInChrome="True">
                        <Button x:Name="btn_minimize" Style="{StaticResource T3.WinCtrl}" ToolTip="Minimize"
                                Click="minimize_button_clicked">
                            <TextBlock Text="&#xE921;" FontFamily="Segoe MDL2 Assets" FontSize="11"/>
                        </Button>
                        <Button x:Name="btn_maximize" Style="{StaticResource T3.WinCtrl}" ToolTip="Maximize"
                                Click="maximize_button_clicked">
                            <TextBlock Text="&#xE922;" FontFamily="Segoe MDL2 Assets" FontSize="11"/>
                        </Button>
                        <Button x:Name="btn_close" Style="{StaticResource T3.WinClose}" ToolTip="Close"
                                IsCancel="True" Click="close_button_clicked">
                            <TextBlock Text="&#xE8BB;" FontFamily="Segoe MDL2 Assets" FontSize="11"/>
                        </Button>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- ══════════════════ ROW 1: WORKSPACE WITH REFINED SIDEBAR ══════════════════ -->
            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="56"/>    <!-- Col 0: Refined Sidebar Rail -->
                    <ColumnDefinition Width="*"/>     <!-- Col 1: Main Content Area -->
                </Grid.ColumnDefinitions>

                <!-- ── Col 0: Left Navigation Icon Rail ── -->
                <Border Grid.Column="0"
                        Background="{StaticResource T3.SurfaceSunken}"
                        BorderBrush="{StaticResource T3.Border}" BorderThickness="0,0,1,0">
                    <Grid Margin="0,12,0,12">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/> <!-- Nav 1: Rooms -->
                            <RowDefinition Height="Auto"/> <!-- Nav 2: Layout -->
                            <RowDefinition Height="*"/>    <!-- Spacer -->
                        </Grid.RowDefinitions>

                        <!-- Nav 1: Rooms -->
                        <ToggleButton x:Name="nav_rooms" Grid.Row="0"
                                      IsChecked="True"
                                      ToolTip="Room List &amp; Settings" Click="nav_toggle_clicked"
                                      HorizontalAlignment="Center" Margin="0,0,0,8">
                            <ToggleButton.Style>
                                <Style TargetType="ToggleButton" BasedOn="{StaticResource T3.Rail.Tile}">
                                    <Setter Property="Template">
                                        <Setter.Value>
                                            <ControlTemplate TargetType="ToggleButton">
                                                <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="8">
                                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsChecked" Value="True">
                                                        <Setter TargetName="border" Property="Background" Value="{StaticResource T3.Ink}"/>
                                                        <Setter Property="Foreground" Value="{StaticResource T3.Surface}"/>
                                                    </Trigger>
                                                    <MultiTrigger>
                                                        <MultiTrigger.Conditions>
                                                            <Condition Property="IsChecked" Value="False"/>
                                                            <Condition Property="IsMouseOver" Value="True"/>
                                                        </MultiTrigger.Conditions>
                                                        <Setter TargetName="border" Property="Background" Value="{StaticResource T3.Border}"/>
                                                        <Setter Property="Foreground" Value="{StaticResource T3.Ink}"/>
                                                    </MultiTrigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Setter.Value>
                                    </Setter>
                                </Style>
                            </ToggleButton.Style>
                            <Path Data="M3 3 h18 v18 H3 z M3 12 h18 M12 3 v9"
                                  Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                                  StrokeThickness="1.6" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                                  Width="18" Height="18" Stretch="Uniform"
                                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </ToggleButton>

                        <!-- Nav 2: Sheet Layout -->
                        <ToggleButton x:Name="nav_layout" Grid.Row="1"
                                      ToolTip="Sheet Layout &amp; Mockup" Click="nav_toggle_clicked"
                                      HorizontalAlignment="Center" Margin="0,0,0,8">
                            <ToggleButton.Style>
                                <Style TargetType="ToggleButton" BasedOn="{StaticResource T3.Rail.Tile}">
                                    <Setter Property="Template">
                                        <Setter.Value>
                                            <ControlTemplate TargetType="ToggleButton">
                                                <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="8">
                                                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsChecked" Value="True">
                                                        <Setter TargetName="border" Property="Background" Value="{StaticResource T3.Ink}"/>
                                                        <Setter Property="Foreground" Value="{StaticResource T3.Surface}"/>
                                                    </Trigger>
                                                    <MultiTrigger>
                                                        <MultiTrigger.Conditions>
                                                            <Condition Property="IsChecked" Value="False"/>
                                                            <Condition Property="IsMouseOver" Value="True"/>
                                                        </MultiTrigger.Conditions>
                                                        <Setter TargetName="border" Property="Background" Value="{StaticResource T3.Border}"/>
                                                        <Setter Property="Foreground" Value="{StaticResource T3.Ink}"/>
                                                    </MultiTrigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Setter.Value>
                                    </Setter>
                                </Style>
                            </ToggleButton.Style>
                            <Path Data="M4 2 h16 v20 H4 z M4 8 h16 M4 13 h7 v6 H4 z M13 13 h7 v6 h-7 z"
                                  Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                                  StrokeThickness="1.6" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                                  Width="18" Height="18" Stretch="Uniform"
                                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </ToggleButton>
                    </Grid>
                </Border>

                <!-- ── Col 1: Main View TabControl ── -->
                <TabControl x:Name="main_tab_control" Grid.Column="1"
                            Margin="0" Background="Transparent" BorderThickness="0">
                    <TabControl.ItemContainerStyle>
                        <Style TargetType="TabItem" BasedOn="{StaticResource T3.TabItem.Hidden}"/>
                    </TabControl.ItemContainerStyle>

                    <!-- ══════════ TAB 0: ROOM LIST & VIEW SETTINGS ══════════ -->
                    <TabItem Header="Rooms">
                        <Grid Margin="16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>     <!-- Room list DataGrid -->
                                <ColumnDefinition Width="16"/>    <!-- Column Gutter -->
                                <ColumnDefinition Width="280"/>   <!-- Settings panel -->
                            </Grid.ColumnDefinitions>

                            <!-- Left: Room List with Toolbar -->
                            <Grid Grid.Column="0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/> <!-- Toolbar -->
                                    <RowDefinition Height="*"/>    <!-- DataGrid -->
                                </Grid.RowDefinitions>

                                <!-- Toolbar -->
                                <Border Grid.Row="0" Style="{StaticResource T3.Panel}" Padding="12,8" Margin="0,0,0,12">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="Auto"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>

                                        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                                            <Button x:Name="btn_select_all" Content="Select All"
                                                    Style="{StaticResource T3.Button.Ghost}" Margin="0,0,8,0"
                                                    Click="select_all_clicked"/>
                                            <Button x:Name="btn_select_none" Content="Select None"
                                                    Style="{StaticResource T3.Button.Ghost}" Margin="0,0,12,0"
                                                    Click="select_none_clicked"/>
                                            <Border Width="1" Height="14" Background="{StaticResource T3.Border}"
                                                    Margin="0,0,12,0" VerticalAlignment="Center"/>
                                            <TextBox x:Name="txt_search" Style="{StaticResource T3.Search}"
                                                     Width="200" Tag="Search rooms..."
                                                     TextChanged="search_changed"/>
                                        </StackPanel>
                                    </Grid>
                                </Border>

                                <!-- DataGrid -->
                                <Border Grid.Row="1" Style="{StaticResource T3.Panel}" Padding="16">
                                    <Grid>
                                        <DataGrid x:Name="room_datagrid" Style="{StaticResource T3.DataGrid}"
                                                  AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                                                  SelectionMode="Extended" IsReadOnly="False"
                                                  HorizontalScrollBarVisibility="Disabled"
                                                  EnableRowVirtualization="True"
                                                  SelectionChanged="room_selection_changed"
                                                  PreviewMouseLeftButtonDown="room_row_clicked">
                                            <DataGrid.Columns>
                                                <DataGridCheckBoxColumn Header="" Binding="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                                                        ElementStyle="{StaticResource T3.CheckBox}"
                                                                        EditingElementStyle="{StaticResource T3.CheckBox}"
                                                                        IsReadOnly="False" Width="38"/>
                                                <DataGridTextColumn Header="NUMBER" Binding="{Binding Number}" IsReadOnly="True" Width="80"
                                                                    ElementStyle="{StaticResource T3.Cell.Mono}"/>
                                                <DataGridTextColumn Header="ROOM NAME" Binding="{Binding Name}" IsReadOnly="True" Width="*"/>
                                                <DataGridTextColumn Header="TYPE" Binding="{Binding RoomType}" IsReadOnly="True" Width="90"
                                                                    ElementStyle="{StaticResource T3.Cell.Muted}"/>
                                                <DataGridTextColumn Header="LEVEL" Binding="{Binding Level}" IsReadOnly="True" Width="90"
                                                                    ElementStyle="{StaticResource T3.Cell.Muted}"/>
                                                <DataGridTextColumn Header="PLANS" Binding="{Binding FloorPlanCount}" IsReadOnly="True" Width="70"
                                                                    ElementStyle="{StaticResource T3.Cell.Mono}"/>
                                                <DataGridTextColumn Header="QTY" Binding="{Binding GenQty, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                                                    IsReadOnly="False" Width="60"
                                                                    ElementStyle="{StaticResource T3.Cell.Mono}"/>
                                            </DataGrid.Columns>
                                        </DataGrid>

                                        <!-- Empty State -->
                                        <TextBlock x:Name="room_grid_empty" Style="{StaticResource T3.Empty}"
                                                   Text="No rooms found in active document" Visibility="Collapsed"/>
                                    </Grid>
                                </Border>
                            </Grid>

                            <!-- Right: Settings Panel -->
                            <Border Grid.Column="2" Style="{StaticResource T3.Panel}" Padding="0" ClipToBounds="True">
                                <ScrollViewer x:Name="settings_scrollviewer"
                                              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                              Padding="16,16,12,16">
                                    <StackPanel>
                                        <!-- Plan Type Card -->
                                        <TextBlock Text="PLAN TYPE" Style="{StaticResource T3.Label}" Margin="0,0,0,8"/>
                                        <CheckBox x:Name="chk_floor_plan" Content="Floor Plan"
                                                  Style="{StaticResource T3.CheckBox}" IsChecked="True"
                                                  Margin="0,0,0,8" Checked="mockup_setting_changed" Unchecked="mockup_setting_changed"/>
                                        <CheckBox x:Name="chk_ceiling_plan" Content="Ceiling Plan"
                                                  Style="{StaticResource T3.CheckBox}"
                                                  Margin="0,0,0,8" Checked="mockup_setting_changed" Unchecked="mockup_setting_changed"/>
                                        <CheckBox x:Name="chk_elevations" Content="Interior Elevations"
                                                  Style="{StaticResource T3.CheckBox}"
                                                  Margin="0,0,0,16" Checked="mockup_setting_changed" Unchecked="mockup_setting_changed"/>

                                        <!-- View Templates Card -->
                                        <TextBlock Text="VIEW TEMPLATES" Style="{StaticResource T3.Label}" Margin="0,0,0,8"/>
                                        <TextBlock Text="FLOOR PLAN" Style="{StaticResource T3.Caption}" Margin="0,0,0,4"/>
                                        <ComboBox x:Name="cmb_plan_template" Style="{StaticResource T3.ComboBox}"
                                                  ItemContainerStyle="{StaticResource T3.ComboBoxItem}"
                                                  Margin="0,0,0,12" SelectionChanged="mockup_setting_changed"/>

                                        <TextBlock Text="CEILING PLAN" Style="{StaticResource T3.Caption}" Margin="0,0,0,4"/>
                                        <ComboBox x:Name="cmb_rcp_template" Style="{StaticResource T3.ComboBox}"
                                                  ItemContainerStyle="{StaticResource T3.ComboBoxItem}"
                                                  Margin="0,0,0,12" SelectionChanged="mockup_setting_changed"/>

                                        <TextBlock Text="ELEVATIONS" Style="{StaticResource T3.Caption}" Margin="0,0,0,4"/>
                                        <ComboBox x:Name="cmb_elev_template" Style="{StaticResource T3.ComboBox}"
                                                  ItemContainerStyle="{StaticResource T3.ComboBoxItem}"
                                                  Margin="0,0,0,16" SelectionChanged="mockup_setting_changed"/>

                                        <!-- Options Card -->
                                        <TextBlock Text="OPTIONS" Style="{StaticResource T3.Label}" Margin="0,0,0,8"/>
                                        <CheckBox x:Name="chk_cropbox_visible" Content="Show Crop Box"
                                                  Style="{StaticResource T3.CheckBox}" IsChecked="True" Margin="0,0,0,8"/>

                                        <Grid Margin="0,0,0,16">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="Auto"/>
                                                <ColumnDefinition Width="8"/>
                                                <ColumnDefinition Width="*"/>
                                            </Grid.ColumnDefinitions>
                                            <TextBlock Grid.Column="0" Text="OFFSET (M)" Style="{StaticResource T3.Caption}"
                                                       VerticalAlignment="Center"/>
                                            <TextBox x:Name="txt_offset" Grid.Column="2" Text="1.0"
                                                     Style="{StaticResource T3.TextBox.Mono}"
                                                     TextChanged="offset_changed"/>
                                        </Grid>

                                        <!-- Tip Callout -->
                                        <Border Style="{StaticResource T3.Callout}">
                                            <StackPanel>
                                                <TextBlock Text="Tips:" Style="{StaticResource T3.BodyStrong}" Margin="0,0,0,4"/>
                                                <TextBlock Text="Select rooms and choose plan types. Interior Elevations creates views for all boundary walls."
                                                           Style="{StaticResource T3.Body.Secondary}" TextWrapping="Wrap"/>
                                            </StackPanel>
                                        </Border>
                                    </StackPanel>
                                </ScrollViewer>
                            </Border>
                        </Grid>
                    </TabItem>

                    <!-- ══════════ TAB 1: SHEET LAYOUT & MOCKUP ══════════ -->
                    <TabItem Header="Layout">
                        <Grid Margin="16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>     <!-- Visual Mockup area -->
                                <ColumnDefinition Width="16"/>    <!-- Column Gutter -->
                                <ColumnDefinition Width="280"/>   <!-- Settings panel -->
                            </Grid.ColumnDefinitions>

                            <!-- Left: Sheet Mockup Area -->
                            <Border Grid.Column="0" Style="{StaticResource T3.Panel}" Padding="16" ClipToBounds="True">
                                <Grid>
                                    <!-- Fallback Text when all view types are unchecked -->
                                    <TextBlock Text="Select view types (Floor Plan, Ceiling Plan, or Elevations) to preview sheet layouts"
                                               Style="{StaticResource T3.Empty}" Margin="32">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock" BasedOn="{StaticResource T3.Empty}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <MultiDataTrigger>
                                                        <MultiDataTrigger.Conditions>
                                                            <Condition Binding="{Binding ElementName=chk_floor_plan, Path=IsChecked}" Value="False"/>
                                                            <Condition Binding="{Binding ElementName=chk_ceiling_plan, Path=IsChecked}" Value="False"/>
                                                            <Condition Binding="{Binding ElementName=chk_elevations, Path=IsChecked}" Value="False"/>
                                                        </MultiDataTrigger.Conditions>
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </MultiDataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>

                                    <Viewbox Stretch="Uniform">
                                        <Grid Width="640" Height="420">
                                            <!-- Combined Layout Mockup (One Sheet per Room) -->
                                            <Border x:Name="CombinedSheetMockup" Width="580" Height="380" Background="{StaticResource T3.Surface}"
                                                    BorderBrush="{StaticResource T3.BorderStrong}" BorderThickness="1" CornerRadius="8">
                                                <Border.Style>
                                                    <Style TargetType="Border">
                                                        <Setter Property="Visibility" Value="Collapsed"/>
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding ElementName=rdo_layout_combined, Path=IsChecked}" Value="True">
                                                                <Setter Property="Visibility" Value="Visible"/>
                                                            </DataTrigger>
                                                            <MultiDataTrigger>
                                                                <MultiDataTrigger.Conditions>
                                                                    <Condition Binding="{Binding ElementName=chk_floor_plan, Path=IsChecked}" Value="False"/>
                                                                    <Condition Binding="{Binding ElementName=chk_ceiling_plan, Path=IsChecked}" Value="False"/>
                                                                    <Condition Binding="{Binding ElementName=chk_elevations, Path=IsChecked}" Value="False"/>
                                                                </MultiDataTrigger.Conditions>
                                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                            </MultiDataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </Border.Style>
                                                <Grid>
                                                    <Canvas x:Name="combined_canvas" Margin="0" Background="Transparent" ClipToBounds="True"/>
                                                </Grid>
                                            </Border>

                                            <!-- Separate Layout Mockup (Plans & Elevations Sheets) -->
                                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                                                <StackPanel.Style>
                                                    <Style TargetType="StackPanel">
                                                        <Setter Property="Visibility" Value="Collapsed"/>
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding ElementName=rdo_layout_separate, Path=IsChecked}" Value="True">
                                                                <Setter Property="Visibility" Value="Visible"/>
                                                            </DataTrigger>
                                                            <MultiDataTrigger>
                                                                <MultiDataTrigger.Conditions>
                                                                    <Condition Binding="{Binding ElementName=chk_floor_plan, Path=IsChecked}" Value="False"/>
                                                                    <Condition Binding="{Binding ElementName=chk_ceiling_plan, Path=IsChecked}" Value="False"/>
                                                                    <Condition Binding="{Binding ElementName=chk_elevations, Path=IsChecked}" Value="False"/>
                                                                </MultiDataTrigger.Conditions>
                                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                            </MultiDataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </StackPanel.Style>

                                                <!-- Sheet 1: Plans Sheet -->
                                                <Border x:Name="PlansSheetMockup" Width="290" Height="200" Background="{StaticResource T3.Surface}"
                                                        BorderBrush="{StaticResource T3.BorderStrong}" BorderThickness="1" CornerRadius="8" Margin="8">
                                                    <Border.Style>
                                                        <Style TargetType="Border">
                                                            <Setter Property="Visibility" Value="Collapsed"/>
                                                            <Style.Triggers>
                                                                <DataTrigger Binding="{Binding ElementName=chk_floor_plan, Path=IsChecked}" Value="True">
                                                                    <Setter Property="Visibility" Value="Visible"/>
                                                                </DataTrigger>
                                                                <DataTrigger Binding="{Binding ElementName=chk_ceiling_plan, Path=IsChecked}" Value="True">
                                                                    <Setter Property="Visibility" Value="Visible"/>
                                                                </DataTrigger>
                                                            </Style.Triggers>
                                                        </Style>
                                                    </Border.Style>
                                                    <Grid>
                                                        <Canvas x:Name="plans_canvas" Margin="0" Background="Transparent" ClipToBounds="True"/>
                                                    </Grid>
                                                </Border>

                                                <!-- Sheet 2: Elevations Sheet -->
                                                <Border x:Name="ElevationsSheetMockup" Width="290" Height="200" Background="{StaticResource T3.Surface}"
                                                        BorderBrush="{StaticResource T3.BorderStrong}" BorderThickness="1" CornerRadius="8" Margin="8">
                                                    <Border.Style>
                                                        <Style TargetType="Border">
                                                            <Setter Property="Visibility" Value="Collapsed"/>
                                                            <Style.Triggers>
                                                                <DataTrigger Binding="{Binding ElementName=chk_elevations, Path=IsChecked}" Value="True">
                                                                    <Setter Property="Visibility" Value="Visible"/>
                                                                </DataTrigger>
                                                            </Style.Triggers>
                                                        </Style>
                                                    </Border.Style>
                                                    <Grid>
                                                        <Canvas x:Name="elevations_canvas" Margin="0" Background="Transparent" ClipToBounds="True"/>
                                                    </Grid>
                                                </Border>
                                            </StackPanel>
                                        </Grid>
                                    </Viewbox>
                                </Grid>
                            </Border>

                            <!-- Right: Sheet Settings & Preview Panel -->
                            <Border Grid.Column="2" Style="{StaticResource T3.Panel}" Padding="0" ClipToBounds="True">
                                <ScrollViewer x:Name="layout_scrollviewer"
                                              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"
                                              Padding="16,16,12,16">
                                    <StackPanel>
                                        <!-- Sheet Layout Card -->
                                        <TextBlock Text="SHEET LAYOUT" Style="{StaticResource T3.Label}" Margin="0,0,0,8"/>
                                        <CheckBox x:Name="chk_layout_on_sheet" Content="Place views on sheets"
                                                  Style="{StaticResource T3.CheckBox}" Margin="0,0,0,12"
                                                  Checked="layout_toggle_changed" Unchecked="layout_toggle_changed"/>

                                        <StackPanel x:Name="pnl_layout_options" IsEnabled="False" Opacity="0.45">
                                            <TextBlock Text="TITLE BLOCK" Style="{StaticResource T3.Caption}" Margin="0,0,0,4"/>
                                            <ComboBox x:Name="cmb_titleblock" Style="{StaticResource T3.ComboBox}"
                                                      ItemContainerStyle="{StaticResource T3.ComboBoxItem}"
                                                      Margin="0,0,0,12" SelectionChanged="mockup_setting_changed"/>

                                            <TextBlock Text="HEADER STRIP" Style="{StaticResource T3.Caption}" Margin="0,0,0,4"/>
                                            <ComboBox x:Name="cmb_strip_side" Style="{StaticResource T3.ComboBox}"
                                                      ItemContainerStyle="{StaticResource T3.ComboBoxItem}"
                                                      Margin="0,0,0,8" SelectionChanged="mockup_setting_changed">
                                                <ComboBoxItem Content="Right (vertical)"/>
                                                <ComboBoxItem Content="Bottom (horizontal)"/>
                                                <ComboBoxItem Content="None"/>
                                            </ComboBox>

                                            <Grid Margin="0,0,0,12">
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="Auto"/>
                                                    <ColumnDefinition Width="8"/>
                                                    <ColumnDefinition Width="*"/>
                                                </Grid.ColumnDefinitions>
                                                <TextBlock Grid.Column="0" Text="HEADER SIZE (MM)" Style="{StaticResource T3.Caption}"
                                                           VerticalAlignment="Center"/>
                                                <TextBox x:Name="txt_strip_mm" Grid.Column="2" Text="70"
                                                         Style="{StaticResource T3.TextBox.Mono}"
                                                         TextChanged="mockup_setting_changed"/>
                                            </Grid>

                                            <TextBlock Text="LAYOUT MODE" Style="{StaticResource T3.Caption}" Margin="0,0,0,6"/>
                                            <RadioButton x:Name="rdo_layout_combined" Content="Combined (Plan + Elevations)"
                                                         GroupName="LayoutMode" IsChecked="True"
                                                         Margin="0,0,0,6" Checked="mockup_setting_changed"/>
                                            <RadioButton x:Name="rdo_layout_separate" Content="Separate Sheets"
                                                         GroupName="LayoutMode" Margin="0,0,0,16"
                                                         Checked="mockup_setting_changed"/>
                                        </StackPanel>

                                        <!-- Generated Sheets Preview Card -->
                                        <TextBlock Text="GENERATED SHEETS" Style="{StaticResource T3.Label}" Margin="0,0,0,8"/>
                                        <TextBlock Text="SELECT SHEET TO OPEN" Style="{StaticResource T3.Caption}" Margin="0,0,0,4"/>
                                        <ComboBox x:Name="cmb_generated_sheets" Style="{StaticResource T3.ComboBox}"
                                                  ItemContainerStyle="{StaticResource T3.ComboBoxItem}"
                                                  IsEnabled="False" Margin="0,0,0,8"/>
                                        <Button x:Name="btn_open_sheet" Content="Open Selected Sheet"
                                                Style="{StaticResource T3.Button.Secondary}" IsEnabled="False"
                                                Margin="0,0,0,16" Click="open_sheet_clicked"/>

                                        <!-- Tip Callout -->
                                        <Border Style="{StaticResource T3.Callout}">
                                            <StackPanel>
                                                <TextBlock Text="Tips:" Style="{StaticResource T3.BodyStrong}" Margin="0,0,0,4"/>
                                                <TextBlock Text="Views are automatically placed on sheets using the selected Title Block."
                                                           Style="{StaticResource T3.Body.Secondary}" TextWrapping="Wrap"/>
                                            </StackPanel>
                                        </Border>
                                    </StackPanel>
                                </ScrollViewer>
                            </Border>
                        </Grid>
                    </TabItem>
                </TabControl>
            </Grid>

            <!-- ══════════════════ ROW 2: FULL-WIDTH FOOTER BAR ══════════════════ -->
            <Border Grid.Row="2" Style="{StaticResource T3.FooterBar}">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Style="{StaticResource T3.Copyright}"/>
                        <Border Width="1" Height="16" Margin="12,0"
                                Background="{StaticResource T3.Border}" VerticalAlignment="Center"/>
                        <Ellipse Style="{StaticResource T3.Dot}" Fill="{StaticResource T3.Success.Accent}"/>
                        <TextBlock x:Name="status_text" Text="Ready"
                                   Style="{StaticResource T3.Body.Secondary}" Margin="0,0,8,0"/>
                        <Border Width="1" Height="12" Background="{StaticResource T3.Border}"
                                Margin="0,0,8,0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="status_count" Text="0 rooms"
                                   Style="{StaticResource T3.Caption}"/>
                    </StackPanel>

                    <!-- Progress panel: bar + Pause + Stop (ProgressPauseMixin) -->
                    <Grid x:Name="sg_progress_panel" Grid.Column="1" Visibility="Collapsed" VerticalAlignment="Center" Margin="0,0,12,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="150"/>
                            <ColumnDefinition Width="8"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="8"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <ProgressBar x:Name="sg_pb" Grid.Column="0" Style="{StaticResource T3.ProgressBar}"
                                     Minimum="0" Maximum="100" Value="0" VerticalAlignment="Center"/>
                        <Button x:Name="sg_btn_pause" Grid.Column="2" Style="{StaticResource T3.Button.Secondary}"
                                Height="28" Padding="8,0" Click="pause_resume_clicked">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock x:Name="sg_btn_pause_icon" Text="&#xE769;" FontFamily="Segoe MDL2 Assets"
                                           FontSize="11" VerticalAlignment="Center" Margin="0,0,4,0"/>
                                <TextBlock x:Name="sg_btn_pause_label" Text="Pause" Style="{StaticResource T3.Caption}"
                                           VerticalAlignment="Center"/>
                            </StackPanel>
                        </Button>
                        <Button x:Name="sg_btn_stop" Grid.Column="4" Style="{StaticResource T3.Button.Danger}"
                                Height="28" Padding="8,0" Click="stop_clicked">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE71A;" FontFamily="Segoe MDL2 Assets" FontSize="11"
                                           VerticalAlignment="Center" Margin="0,0,4,0"/>
                                <TextBlock Text="Stop" Style="{StaticResource T3.Caption}" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Button>
                    </Grid>

                    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                        <Button x:Name="btn_cancel" Content="Cancel"
                                Style="{StaticResource T3.Button.Secondary}" Margin="0,0,8,0"
                                Click="close_button_clicked"/>
                        <Button x:Name="btn_create" Content="Create Plans"
                                Style="{StaticResource T3.Button.Primary}" IsDefault="True"
                                Click="create_plans_clicked"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>
    </Border>
</Window>
'''

with io.open(r"T3Lab.extension\lib\GUI\Tools\SheetGen.xaml", "w", encoding="utf-8") as f:
    f.write(xaml_content)
print("Migrated SheetGen.xaml to T3 UI Standard")
