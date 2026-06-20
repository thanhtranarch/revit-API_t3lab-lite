# -*- coding: utf-8 -*-
import io
import re

xaml_path = r"d:\01. T3Lab\02 Revit Tools\t3lab-revit-api\T3Lab.extension\lib\GUI\Tools\AnnotationManager.xaml"

with io.open(xaml_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Replace RowDefinitions using regex
row_def_pattern = re.compile(
    r"(<Grid.RowDefinitions>\s*<RowDefinition[^>]*/>\s*<!-- Logo -->\s*<RowDefinition[^>]*/>\s*<!-- Tab 1 -->\s*<RowDefinition[^>]*/>\s*<!-- Tab 2 -->\s*<RowDefinition[^>]*/>\s*<!-- Tab 3 -->\s*<RowDefinition[^>]*/>\s*<!-- Spacer -->\s*</Grid.RowDefinitions>)",
    re.MULTILINE
)

new_rows = """<Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/> <!-- Logo -->
                        <RowDefinition Height="Auto"/> <!-- Tab 1: Dim -->
                        <RowDefinition Height="Auto"/> <!-- Tab 2: Txt -->
                        <RowDefinition Height="Auto"/> <!-- Tab 3: Tag Checker -->
                        <RowDefinition Height="Auto"/> <!-- Tab 4: DimText -->
                        <RowDefinition Height="Auto"/> <!-- Tab 5: Utilities -->
                        <RowDefinition Height="Auto"/> <!-- Tab 6: Settings -->
                        <RowDefinition Height="*"/>    <!-- Spacer -->
                    </Grid.RowDefinitions>"""

if row_def_pattern.search(content):
    content = row_def_pattern.sub(new_rows, content)
    print("Row definitions matched and replaced.")
else:
    # Try a looser match
    loose_pattern = re.compile(r"<Grid.RowDefinitions>.*?<!-- Logo -->.*?<!-- Spacer -->.*?</Grid.RowDefinitions>", re.DOTALL)
    if loose_pattern.search(content):
        content = loose_pattern.sub(new_rows, content)
        print("Row definitions matched loosely and replaced.")
    else:
        print("Row definitions NOT matched.")

# 2. Replace ToggleButtons
toggle_pattern = re.compile(r"<!-- Tab: Dimensions -->.*?<!-- Tab: Settings -->.*?</ToggleButton>", re.DOTALL)

new_toggles = """<!-- Tab: Dimensions -->
                    <ToggleButton x:Name="nav_dim" Grid.Row="1"
                                  Style="{StaticResource T3SidebarToggle}"
                                  Margin="0,0,0,8" IsChecked="True"
                                  ToolTip="Dimensions" Click="nav_dimensions_checked">
                        <Path Data="M3 20 H21 M7 20 V8 M17 20 V8 M7 8 H17 M7 14 H17"
                               Stroke="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>
                    <!-- Tab: Text Notes -->
                    <ToggleButton x:Name="nav_txt" Grid.Row="2"
                                  Style="{StaticResource T3SidebarToggle}"
                                  Margin="0,0,0,8"
                                  ToolTip="Text Notes" Click="nav_textnotes_checked">
                        <Path Data="M4 7 h16 M4 12 h16 M4 17 h10 M12 4 H4 V6"
                               Stroke="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>
                    <!-- Tab: Tag Checker -->
                    <ToggleButton x:Name="nav_tag" Grid.Row="3"
                                  Style="{StaticResource T3SidebarToggle}"
                                  Margin="0,0,0,8"
                                  ToolTip="Tag Checker" Click="nav_tag_checked">
                        <Path Data="M9 11 l3 3 l8 -8 M20 12 v6 a2 2 0 0 1 -2 2 H4 a2 2 0 0 1 -2 -2 V6 a2 2 0 0 1 2 -2 h9"
                               Stroke="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>
                    <!-- Tab: DimText -->
                    <ToggleButton x:Name="nav_dimtext" Grid.Row="4"
                                  Style="{StaticResource T3SidebarToggle}"
                                  Margin="0,0,0,8"
                                  ToolTip="Dim Text" Click="nav_dimtext_checked">
                        <Path Data="M5 4 h14 M12 4 v16 M9 20 h6 M4 8 h4 M16 8 h4"
                               Stroke="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>
                    <!-- Tab: Utilities -->
                    <ToggleButton x:Name="nav_utils" Grid.Row="5"
                                  Style="{StaticResource T3SidebarToggle}"
                                  Margin="0,0,0,8"
                                  ToolTip="Utilities" Click="nav_utils_checked">
                        <Path Data="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"
                               Stroke="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>
                    <!-- Tab: Settings -->
                    <ToggleButton x:Name="nav_settings" Grid.Row="6"
                                  Style="{StaticResource T3SidebarToggle}"
                                  Margin="0,0,0,8"
                                  ToolTip="Settings" Click="nav_settings_checked">
                        <Path Data="M12 15 A3 3 0 1 0 12 9 A3 3 0 0 0 12 15 M12 4 V6 M12 18 V20 M4.2 6.2 L5.6 7.6 M18.4 17.8 L19.8 19.2 M4 12 H6 M18 12 H20"
                               Stroke="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=ToggleButton}}"
                               StrokeThickness="1.8" StrokeLineJoin="Round"
                               StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                               Width="17" Height="17" Stretch="Uniform"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ToggleButton>"""

if toggle_pattern.search(content):
    content = toggle_pattern.sub(new_toggles, content)
    print("ToggleButtons matched and replaced.")
else:
    print("ToggleButtons NOT matched.")

# 3. Add TabItems before Settings TabItem (the last TabItem)
# Let's search for the last TabItem or the one that contains Grid Margin="24,24,24,24"
tab_pattern = re.compile(r"(<TabItem>\s*<Grid Margin=\"24,24,24,24\">\s*<Grid.ColumnDefinitions>\s*<ColumnDefinition Width=\"\*\"/>\s*<ColumnDefinition Width=\"\*\"/>\s*</Grid.ColumnDefinitions>)", re.DOTALL)

new_tabs = """<!-- 3. TAG CHECKER TAB -->
            <TabItem>
                <Grid x:Name="grid_tag_checker" Background="White"/>
            </TabItem>

            <!-- 4. DIMTEXT TAB -->
            <TabItem>
                <Grid x:Name="grid_dimtext" Background="White"/>
            </TabItem>

            <!-- 5. UTILITIES TAB -->
            <TabItem>
                <Grid Background="White" Margin="24">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>     <!-- Header -->
                        <RowDefinition Height="*"/>        <!-- Cards Grid -->
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Margin="0,0,0,24">
                        <TextBlock Text="Annotation Utilities" FontSize="18" FontWeight="Bold" Foreground="#0F172A"/>
                        <TextBlock Text="Helper tools for renumbering, case formatting, and document syncing" FontSize="11" Foreground="#64748B" Margin="0,2,0,0"/>
                    </StackPanel>

                    <!-- Cards Row -->
                    <UniformGrid Grid.Row="1" Columns="3" VerticalAlignment="Top">
                        <!-- Card 1: Copy Annotation -->
                        <Border BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,16,0" Background="#F8FAFC">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <StackPanel Grid.Row="0">
                                    <TextBlock Text="Copy Annotations" FontSize="14" FontWeight="Bold" Foreground="#0F172A"/>
                                    <TextBlock Text="Transfer dimensions, text notes, tags, and detail components from this model to another open Revit document." 
                                               FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,6,0,16"/>
                                </StackPanel>
                                <Button Grid.Row="2" x:Name="btn_util_copy_anno" Style="{StaticResource PrimaryButton}" Height="36" Content="Launch Copier"/>
                            </Grid>
                        </Border>

                        <!-- Card 2: Renumber Along Spline -->
                        <Border BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,16,0" Background="#F8FAFC">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <StackPanel Grid.Row="0">
                                    <TextBlock Text="Renumber Along Spline" FontSize="14" FontWeight="Bold" Foreground="#0F172A"/>
                                    <TextBlock Text="Renumber any category of elements (such as doors, rooms, equipment) sequentially along a selected line/spline path." 
                                               FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,6,0,16"/>
                                </StackPanel>
                                <Button Grid.Row="2" x:Name="btn_util_renumber_spline" Style="{StaticResource PrimaryButton}" Height="36" Content="Start Renumbering"/>
                            </Grid>
                        </Border>

                        <!-- Card 3: Upper Case All -->
                        <Border BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="8" Padding="16" Background="#F8FAFC">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <StackPanel Grid.Row="0">
                                    <TextBlock Text="Uppercase All Texts" FontSize="14" FontWeight="Bold" Foreground="#0F172A"/>
                                    <TextBlock Text="Convert view names, sheet names, title block parameter strings, text notes, and dimension overrides to UPPERCASE." 
                                               FontSize="11" Foreground="#64748B" TextWrapping="Wrap" Margin="0,6,0,16"/>
                                </StackPanel>
                                <Button Grid.Row="2" x:Name="btn_util_upper_all" Style="{StaticResource AccentButton}" Height="36" Content="Run Uppercase Converter"/>
                            </Grid>
                        </Border>
                    </UniformGrid>
                </Grid>
            </TabItem>

            \\1"""

# Note: we use lambda or group reference to preserve the matched pattern and prepend
if tab_pattern.search(content):
    content = tab_pattern.sub(new_tabs.replace("\\1", r"\1"), content)
    print("TabItems inserted successfully.")
else:
    print("TabItems NOT matched.")

with io.open(xaml_path, "w", encoding="utf-8") as f:
    f.write(content)

print("XAML updated successfully!")
