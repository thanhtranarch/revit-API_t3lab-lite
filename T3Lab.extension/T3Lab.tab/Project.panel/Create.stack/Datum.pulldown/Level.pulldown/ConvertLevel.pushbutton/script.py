# -*- coding: utf-8 -*-
"""
Level Swap v1.0 - DQT
Swap levels between 3D and 2D extents in Revit views.
Control level bubble visibility. Batch operations across views.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Level\nSwap"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Swap levels between 3D and 2D extents. Control bubbles. Batch operations."

# =====================================================
# IMPORTS
# =====================================================
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, View, ViewType, \
    DatumEnds, DatumExtentType, ElementId, Level
from Autodesk.Revit.UI import *
from pyrevit import revit, HOST_APP

import sys
import os
import System
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, TextWrapping, Visibility
from System.Windows import RoutedEventArgs, MessageBox as WPFMessageBox, MessageBoxButton, MessageBoxResult, MessageBoxImage
from System.Windows.Controls import *
from System.Windows.Media import SolidColorBrush, BrushConverter
from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader

# =====================================================
# REVIT CONTEXT
# =====================================================
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# =====================================================
# UNIT HELPERS
# =====================================================
def internal_to_mm(value):
    """Convert Revit internal units (feet) to millimeters"""
    return value * 304.8

def format_elevation(value_feet):
    """Format elevation from internal feet to display string"""
    mm = internal_to_mm(value_feet)
    if abs(mm) < 0.1:
        return "0 mm"
    return "{:,.0f} mm".format(mm)

# =====================================================
# XAML UI DEFINITION
# =====================================================
MAIN_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Level Swap - By DQT" Height="720" Width="1100"
        WindowStartupLocation="CenterScreen" Background="#F8FAFC"
        MinHeight="600" MinWidth="900">
    <Grid Margin="0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#0F172A" Padding="12" BorderBrush="#CBD5E1" BorderThickness="0,0,0,2">
            <StackPanel>
                <TextBlock Text="LEVEL SWAP" 
                           FontSize="20" FontWeight="Bold" 
                           HorizontalAlignment="Center"
                           Foreground="White"/>
                <TextBlock Text="Swap Levels between 3D and 2D Extents | Control Bubble Visibility" 
                           FontSize="11" 
                           HorizontalAlignment="Center"
                           Foreground="White"
                           Margin="0,3,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Summary Cards -->
        <Border Grid.Row="1" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="10" Margin="15,10,15,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                    <TextBlock x:Name="txtTotalLevels" Text="0" FontSize="22" FontWeight="Bold" Foreground="#0F172A" HorizontalAlignment="Center"/>
                    <TextBlock Text="Total Levels" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                </StackPanel>
                
                <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                    <TextBlock x:Name="txtIs3D" Text="0" FontSize="22" FontWeight="Bold" Foreground="#10B981" HorizontalAlignment="Center"/>
                    <TextBlock Text="3D Extents" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                </StackPanel>
                
                <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                    <TextBlock x:Name="txtIs2D" Text="0" FontSize="22" FontWeight="Bold" Foreground="#2196F3" HorizontalAlignment="Center"/>
                    <TextBlock Text="2D Extents" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                </StackPanel>
                
                <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                    <TextBlock x:Name="txtSelectedLevels" Text="0" FontSize="22" FontWeight="Bold" Foreground="#F59E0B" HorizontalAlignment="Center"/>
                    <TextBlock Text="Selected" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                </StackPanel>
                
                <StackPanel Grid.Column="4" HorizontalAlignment="Center">
                    <TextBlock x:Name="txtViewCount" Text="0" FontSize="22" FontWeight="Bold" Foreground="#9C27B0" HorizontalAlignment="Center"/>
                    <TextBlock Text="Views" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Toolbar: View Selector + Search + Filter -->
        <Border Grid.Row="2" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="8" Margin="15,8,15,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="View:" VerticalAlignment="Center" Margin="0,0,5,0" FontWeight="SemiBold" Foreground="#FFFFFF"/>
                <ComboBox x:Name="cmbView" Grid.Column="1" Width="300" Height="28" Margin="0,0,10,0"/>
                
                <TextBlock Grid.Column="2" Text="Filter:" VerticalAlignment="Center" Margin="0,0,5,0" FontWeight="SemiBold" Foreground="#FFFFFF"/>
                <ComboBox x:Name="cmbFilter" Grid.Column="3" Width="120" Height="28" Margin="0,0,10,0">
                    <ComboBoxItem Content="All" IsSelected="True"/>
                    <ComboBoxItem Content="3D Only"/>
                    <ComboBoxItem Content="2D Only"/>
                </ComboBox>
                
                <TextBlock Grid.Column="4" Text="Search:" VerticalAlignment="Center" Margin="10,0,5,0" FontWeight="SemiBold" Foreground="#FFFFFF"/>
                <TextBox x:Name="txtSearch" Grid.Column="5" Width="180" Height="28" VerticalContentAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Main DataGrid -->
        <Border Grid.Row="3" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Margin="15,8,15,0" Padding="0">
            <DataGrid x:Name="dgLevels" 
                      AutoGenerateColumns="False" 
                      CanUserAddRows="False" 
                      CanUserDeleteRows="False"
                      SelectionMode="Extended"
                      GridLinesVisibility="Horizontal"
                      HeadersVisibility="Column"
                      BorderThickness="0"
                      Background="White"
                      RowBackground="White"
                      AlternatingRowBackground="#F1F5F9"
                      IsReadOnly="False"
                      CanUserSortColumns="True"
                      HorizontalGridLinesBrush="#E8E8E8">
                <DataGrid.Columns>
                    <DataGridCheckBoxColumn Binding="{Binding IsChecked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" 
                                            Width="40" 
                                            Header="&#x2611;"/>
                    <DataGridTextColumn Binding="{Binding LevelName}" 
                                        Header="Level Name" Width="180" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding Elevation}" 
                                        Header="Elevation" Width="120" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding ExtentStatus}" 
                                        Header="Extent Type" Width="100" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding BubbleStart}" 
                                        Header="Bubble Left" Width="100" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding BubbleEnd}" 
                                        Header="Bubble Right" Width="100" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding IsStructural}" 
                                        Header="Structural" Width="80" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding IsBuildingStory}" 
                                        Header="Story" Width="80" IsReadOnly="True"/>
                    <DataGridTextColumn Binding="{Binding ElementId}" 
                                        Header="Element ID" Width="90" IsReadOnly="True"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>
        
        <!-- Action Buttons -->
        <Border Grid.Row="4" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="8" Margin="15,8,15,0">
            <StackPanel>
                <!-- Row 1: Selection + Extent Swap -->
                <Grid>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                        <Button x:Name="btnSelectAll" Content="Select All" Padding="12,5" Margin="2" Background="White" BorderBrush="#CBD5E1"/>
                        <Button x:Name="btnSelectNone" Content="Clear" Padding="12,5" Margin="2" Background="White" BorderBrush="#CBD5E1"/>
                        <Button x:Name="btnSelect3D" Content="Select 3D" Padding="12,5" Margin="2" Background="White" BorderBrush="#CBD5E1"/>
                        <Button x:Name="btnSelect2D" Content="Select 2D" Padding="12,5" Margin="2" Background="White" BorderBrush="#CBD5E1"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                        <Button x:Name="btnSwapTo2D" Content="&#x21E8; Swap to 2D" Padding="16,6" Margin="3" Background="#2196F3" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
                        <Button x:Name="btnSwapTo3D" Content="&#x21E8; Swap to 3D" Padding="16,6" Margin="3" Background="#10B981" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
                        <Button x:Name="btnToggle" Content="&#x21C4; Toggle" Padding="16,6" Margin="3" Background="#0F172A" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="btnRefresh" Content="Refresh" Padding="12,5" Margin="2" Background="White" BorderBrush="#CBD5E1"/>
                    </StackPanel>
                </Grid>
                
                <!-- Separator -->
                <Border BorderBrush="#E8E8E8" BorderThickness="0,1,0,0" Margin="0,6,0,6"/>
                
                <!-- Row 2: Bubble Controls -->
                <Grid>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Left">
                        <TextBlock Text="Bubbles:" VerticalAlignment="Center" Margin="2,0,8,0" FontWeight="SemiBold" Foreground="#FFFFFF" FontSize="11"/>
                        <Button x:Name="btnBubbleStartOn" Content="&#x25C9; Left ON" Padding="10,5" Margin="2" Background="#E8F5E9" BorderBrush="#10B981" Foreground="#2E7D32" FontWeight="SemiBold" FontSize="11"/>
                        <Button x:Name="btnBubbleStartOff" Content="&#x25CB; Left OFF" Padding="10,5" Margin="2" Background="#FFEBEE" BorderBrush="#EF4444" Foreground="#C62828" FontWeight="SemiBold" FontSize="11"/>
                        <Border BorderBrush="#CBD5E1" BorderThickness="1,0,0,0" Margin="6,2,6,2"/>
                        <Button x:Name="btnBubbleEndOn" Content="&#x25C9; Right ON" Padding="10,5" Margin="2" Background="#E8F5E9" BorderBrush="#10B981" Foreground="#2E7D32" FontWeight="SemiBold" FontSize="11"/>
                        <Button x:Name="btnBubbleEndOff" Content="&#x25CB; Right OFF" Padding="10,5" Margin="2" Background="#FFEBEE" BorderBrush="#EF4444" Foreground="#C62828" FontWeight="SemiBold" FontSize="11"/>
                        <Border BorderBrush="#CBD5E1" BorderThickness="1,0,0,0" Margin="6,2,6,2"/>
                        <Button x:Name="btnBubbleAllOn" Content="&#x25C9; All ON" Padding="10,5" Margin="2" Background="#E8F5E9" BorderBrush="#10B981" Foreground="#2E7D32" FontWeight="Bold" FontSize="11"/>
                        <Button x:Name="btnBubbleAllOff" Content="&#x25CB; All OFF" Padding="10,5" Margin="2" Background="#FFEBEE" BorderBrush="#EF4444" Foreground="#C62828" FontWeight="Bold" FontSize="11"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>
        
        <!-- Footer -->
        <Grid Grid.Row="5" Margin="15,8,15,10">
            <TextBlock Text="Ctrl+Click: multi-select | Checkbox: batch select | Double-click row: select in Revit" FontSize="10" Foreground="#888" VerticalAlignment="Center"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <TextBlock Text="pyDQT v1.0" FontSize="9" Foreground="#AAA" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <Button x:Name="btnClose" Content="Close" Padding="15,5" Background="White" BorderBrush="#CBD5E1"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"""

# =====================================================
# LEVEL ITEM DATA CLASS
# =====================================================
class LevelItem(System.Object):
    """Data class representing a Level's state in a specific view"""
    
    def __init__(self, level, view):
        self._level = level
        self._view = view
        self._is_checked = False
        self._level_name = level.Name if level.Name else "Unnamed"
        self._element_id = str(level.Id.IntegerValue)
        self._elevation = format_elevation(level.Elevation)
        self._elevation_value = level.Elevation
        
        # Determine 3D or 2D extent
        try:
            datum_extent_type = level.GetDatumExtentTypeInView(DatumEnds.End0, view)
            self._is_3d = (datum_extent_type == DatumExtentType.Model)
        except:
            self._is_3d = True
        
        # Get bubble visibility (Left = End0, Right = End1)
        try:
            self._bubble_start = "Visible" if level.IsBubbleVisibleInView(DatumEnds.End0, view) else "Hidden"
        except:
            self._bubble_start = "N/A"
        
        try:
            self._bubble_end = "Visible" if level.IsBubbleVisibleInView(DatumEnds.End1, view) else "Hidden"
        except:
            self._bubble_end = "N/A"
        
        # Get structural flag
        try:
            struct_param = level.get_Parameter(DB.BuiltInParameter.LEVEL_IS_STRUCTURAL)
            if struct_param and struct_param.HasValue:
                self._is_structural = "Yes" if struct_param.AsInteger() == 1 else "No"
            else:
                self._is_structural = "N/A"
        except:
            self._is_structural = "N/A"
        
        # Get building story flag
        try:
            story_param = level.get_Parameter(DB.BuiltInParameter.LEVEL_IS_BUILDING_STORY)
            if story_param and story_param.HasValue:
                self._is_building_story = "Yes" if story_param.AsInteger() == 1 else "No"
            else:
                self._is_building_story = "N/A"
        except:
            self._is_building_story = "N/A"
    
    # --- Properties for WPF Binding ---
    @property
    def IsChecked(self):
        return self._is_checked
    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
    
    @property
    def LevelName(self):
        return self._level_name
    
    @property
    def Elevation(self):
        return self._elevation
    
    @property
    def ElevationValue(self):
        return self._elevation_value
    
    @property
    def ExtentStatus(self):
        return "3D" if self._is_3d else "2D"
    
    @property
    def BubbleStart(self):
        return self._bubble_start
    
    @property
    def BubbleEnd(self):
        return self._bubble_end
    
    @property
    def IsStructural(self):
        return self._is_structural
    
    @property
    def IsBuildingStory(self):
        return self._is_building_story
    
    @property
    def ElementId(self):
        return self._element_id
    
    @property
    def Is3D(self):
        return self._is_3d
    
    @property
    def LevelElement(self):
        return self._level
    
    @property
    def View(self):
        return self._view


# =====================================================
# MAIN WINDOW CLASS
# =====================================================
class LevelSwapWindow(object):
    """Main window for Level Swap tool"""
    
    def __init__(self):
        # Parse XAML
        xml_reader = XmlReader.Create(StringReader(MAIN_XAML))
        self.window = XamlReader.Load(xml_reader)
        
        # Get UI elements - Summary
        self.txt_total_levels = self.window.FindName("txtTotalLevels")
        self.txt_is_3d = self.window.FindName("txtIs3D")
        self.txt_is_2d = self.window.FindName("txtIs2D")
        self.txt_selected_levels = self.window.FindName("txtSelectedLevels")
        self.txt_view_count = self.window.FindName("txtViewCount")
        
        # Toolbar
        self.cmb_view = self.window.FindName("cmbView")
        self.cmb_filter = self.window.FindName("cmbFilter")
        self.txt_search = self.window.FindName("txtSearch")
        
        # DataGrid
        self.dg_levels = self.window.FindName("dgLevels")
        
        # Selection buttons
        self.btn_select_all = self.window.FindName("btnSelectAll")
        self.btn_select_none = self.window.FindName("btnSelectNone")
        self.btn_select_3d = self.window.FindName("btnSelect3D")
        self.btn_select_2d = self.window.FindName("btnSelect2D")
        
        # Swap buttons
        self.btn_swap_to_2d = self.window.FindName("btnSwapTo2D")
        self.btn_swap_to_3d = self.window.FindName("btnSwapTo3D")
        self.btn_toggle = self.window.FindName("btnToggle")
        self.btn_refresh = self.window.FindName("btnRefresh")
        self.btn_close = self.window.FindName("btnClose")
        
        # Bubble control buttons
        self.btn_bubble_start_on = self.window.FindName("btnBubbleStartOn")
        self.btn_bubble_start_off = self.window.FindName("btnBubbleStartOff")
        self.btn_bubble_end_on = self.window.FindName("btnBubbleEndOn")
        self.btn_bubble_end_off = self.window.FindName("btnBubbleEndOff")
        self.btn_bubble_all_on = self.window.FindName("btnBubbleAllOn")
        self.btn_bubble_all_off = self.window.FindName("btnBubbleAllOff")
        
        # Data storage
        self.all_level_items = []
        self.views_dict = {}
        
        # Wire events
        self._wire_events()
        
        # Initialize data
        self._load_views()
        self._load_levels()
    
    def _wire_events(self):
        """Connect UI events to handlers"""
        self.btn_select_all.Click += self._on_select_all
        self.btn_select_none.Click += self._on_select_none
        self.btn_select_3d.Click += self._on_select_3d
        self.btn_select_2d.Click += self._on_select_2d
        self.btn_swap_to_2d.Click += self._on_swap_to_2d
        self.btn_swap_to_3d.Click += self._on_swap_to_3d
        self.btn_toggle.Click += self._on_toggle
        self.btn_refresh.Click += self._on_refresh
        self.btn_close.Click += self._on_close
        self.cmb_view.SelectionChanged += self._on_view_changed
        self.cmb_filter.SelectionChanged += self._on_filter_changed
        self.txt_search.TextChanged += self._on_search_changed
        self.dg_levels.MouseDoubleClick += self._on_row_double_click
        
        # Bubble control events
        self.btn_bubble_start_on.Click += self._on_bubble_start_on
        self.btn_bubble_start_off.Click += self._on_bubble_start_off
        self.btn_bubble_end_on.Click += self._on_bubble_end_on
        self.btn_bubble_end_off.Click += self._on_bubble_end_off
        self.btn_bubble_all_on.Click += self._on_bubble_all_on
        self.btn_bubble_all_off.Click += self._on_bubble_all_off
    
    # =====================================================
    # DATA LOADING
    # =====================================================
    def _load_views(self):
        """Load views where levels are visible"""
        self.views_dict.clear()
        self.cmb_view.Items.Clear()
        
        active_view = doc.ActiveView
        
        collector = FilteredElementCollector(doc).OfClass(View)
        valid_view_types = [
            ViewType.Elevation, ViewType.Section, ViewType.ThreeD,
            ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.AreaPlan,
            ViewType.EngineeringPlan, ViewType.Detail
        ]
        
        views = []
        for v in collector:
            if v.IsTemplate:
                continue
            if v.ViewType in valid_view_types:
                views.append(v)
        
        views.sort(key=lambda v: v.Name)
        
        # Add active view first
        if active_view and not active_view.IsTemplate:
            active_name = "[Active] " + active_view.Name
            self.cmb_view.Items.Add(active_name)
            self.views_dict[active_name] = active_view
        
        for v in views:
            if v.Id != active_view.Id:
                name = v.Name
                if name in self.views_dict:
                    name = "{} (ID:{})".format(name, v.Id.IntegerValue)
                self.cmb_view.Items.Add(name)
                self.views_dict[name] = v
        
        if self.cmb_view.Items.Count > 0:
            self.cmb_view.SelectedIndex = 0
        
        self.txt_view_count.Text = str(len(views))
    
    def _load_levels(self):
        """Load all levels for the selected view"""
        self.all_level_items = []
        
        view = self._get_selected_view()
        if view is None:
            self._update_display()
            return
        
        collector = FilteredElementCollector(doc).OfClass(Level)
        levels = list(collector)
        
        for level in levels:
            try:
                item = LevelItem(level, view)
                self.all_level_items.append(item)
            except Exception as e:
                print("Error loading level {}: {}".format(level.Name, str(e)))
        
        # Sort by elevation ascending
        self.all_level_items.sort(key=lambda x: x.ElevationValue)
        
        self._update_display()
    
    def _get_selected_view(self):
        if self.cmb_view.SelectedItem is None:
            return None
        selected_name = str(self.cmb_view.SelectedItem)
        return self.views_dict.get(selected_name, None)
    
    def _update_display(self):
        """Apply filters and update DataGrid and summary cards"""
        search_text = self.txt_search.Text.lower().strip() if self.txt_search.Text else ""
        filter_idx = self.cmb_filter.SelectedIndex if self.cmb_filter.SelectedIndex >= 0 else 0
        
        filtered = []
        for item in self.all_level_items:
            if search_text and search_text not in item.LevelName.lower():
                continue
            if filter_idx == 1 and not item.Is3D:
                continue
            if filter_idx == 2 and item.Is3D:
                continue
            filtered.append(item)
        
        self.dg_levels.ItemsSource = System.Array[System.Object](filtered)
        
        total_3d = sum(1 for i in self.all_level_items if i.Is3D)
        total_2d = sum(1 for i in self.all_level_items if not i.Is3D)
        checked = sum(1 for i in self.all_level_items if i.IsChecked)
        
        self.txt_total_levels.Text = str(len(self.all_level_items))
        self.txt_is_3d.Text = str(total_3d)
        self.txt_is_2d.Text = str(total_2d)
        self.txt_selected_levels.Text = str(checked)
    
    # =====================================================
    # SELECTION HANDLERS
    # =====================================================
    def _on_select_all(self, sender, args):
        for item in self.dg_levels.ItemsSource:
            item.IsChecked = True
        self._update_display()
    
    def _on_select_none(self, sender, args):
        for item in self.all_level_items:
            item.IsChecked = False
        self._update_display()
    
    def _on_select_3d(self, sender, args):
        for item in self.all_level_items:
            item.IsChecked = item.Is3D
        self._update_display()
    
    def _on_select_2d(self, sender, args):
        for item in self.all_level_items:
            item.IsChecked = not item.Is3D
        self._update_display()
    
    # =====================================================
    # SWAP OPERATIONS
    # =====================================================
    def _get_checked_items(self):
        return [i for i in self.all_level_items if i.IsChecked]
    
    def _on_swap_to_2d(self, sender, args):
        """Swap selected levels from 3D to 2D"""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No levels selected.\nPlease check the levels you want to swap.",
                              "Level Swap", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        targets = [i for i in checked if i.Is3D]
        if not targets:
            WPFMessageBox.Show("No 3D levels found in selection.\nAll selected levels are already 2D.",
                              "Level Swap", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        result = WPFMessageBox.Show(
            "Swap {} level(s) from 3D to 2D in view '{}'?\n\nThis will change the extent type to view-specific (2D).".format(
                len(targets), view.Name),
            "Confirm Swap to 2D", MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - Swap Levels to 2D")
        t.Start()
        try:
            for item in targets:
                try:
                    level = item.LevelElement
                    level.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.ViewSpecific)
                    level.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.ViewSpecific)
                    success += 1
                except Exception as e:
                    errors.append("{}: {}".format(item.LevelName, str(e)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        self._load_levels()
        
        msg = "Successfully swapped {} level(s) to 2D.".format(success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Level Swap Complete", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_swap_to_3d(self, sender, args):
        """Swap selected levels from 2D to 3D"""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No levels selected.\nPlease check the levels you want to swap.",
                              "Level Swap", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        targets = [i for i in checked if not i.Is3D]
        if not targets:
            WPFMessageBox.Show("No 2D levels found in selection.\nAll selected levels are already 3D.",
                              "Level Swap", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        result = WPFMessageBox.Show(
            "Swap {} level(s) from 2D to 3D in view '{}'?\n\nThis will change the extent type to model-wide (3D).\nNote: Level extents will reset to model extents.".format(
                len(targets), view.Name),
            "Confirm Swap to 3D", MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - Swap Levels to 3D")
        t.Start()
        try:
            for item in targets:
                try:
                    level = item.LevelElement
                    level.SetDatumExtentType(DatumEnds.End0, view, DatumExtentType.Model)
                    level.SetDatumExtentType(DatumEnds.End1, view, DatumExtentType.Model)
                    success += 1
                except Exception as e:
                    errors.append("{}: {}".format(item.LevelName, str(e)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        self._load_levels()
        
        msg = "Successfully swapped {} level(s) to 3D.".format(success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Level Swap Complete", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_toggle(self, sender, args):
        """Toggle selected levels: 3D becomes 2D, 2D becomes 3D"""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No levels selected.\nPlease check the levels you want to toggle.",
                              "Level Swap", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        count_3d_to_2d = sum(1 for i in checked if i.Is3D)
        count_2d_to_3d = sum(1 for i in checked if not i.Is3D)
        
        result = WPFMessageBox.Show(
            "Toggle {} level(s) in view '{}'?\n\n  {} level(s): 3D -> 2D\n  {} level(s): 2D -> 3D".format(
                len(checked), view.Name, count_3d_to_2d, count_2d_to_3d),
            "Confirm Toggle", MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - Toggle Level Extents")
        t.Start()
        try:
            for item in checked:
                try:
                    level = item.LevelElement
                    if item.Is3D:
                        new_type = DatumExtentType.ViewSpecific
                    else:
                        new_type = DatumExtentType.Model
                    
                    level.SetDatumExtentType(DatumEnds.End0, view, new_type)
                    level.SetDatumExtentType(DatumEnds.End1, view, new_type)
                    success += 1
                except Exception as e:
                    errors.append("{}: {}".format(item.LevelName, str(e)))
            t.Commit()
        except Exception as e:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        self._load_levels()
        
        msg = "Successfully toggled {} level(s).".format(success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Toggle Complete", MessageBoxButton.OK, MessageBoxImage.Information)
    
    # =====================================================
    # BUBBLE CONTROL OPERATIONS
    # =====================================================
    def _set_bubble_visibility(self, end, visible):
        """Set bubble visibility for checked levels at specified end."""
        checked = self._get_checked_items()
        if not checked:
            WPFMessageBox.Show("No levels selected.\nPlease check the levels you want to modify.",
                              "Bubble Control", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        view = self._get_selected_view()
        if view is None:
            return
        
        if end == DatumEnds.End0:
            end_name = "Left"
        elif end == DatumEnds.End1:
            end_name = "Right"
        else:
            end_name = "Left + Right"
        
        action = "Show" if visible else "Hide"
        
        success = 0
        errors = []
        
        t = Transaction(doc, "DQT - {} {} Level Bubble".format(action, end_name))
        t.Start()
        try:
            for item in checked:
                try:
                    level = item.LevelElement
                    ends_to_set = []
                    if end is None:
                        ends_to_set = [DatumEnds.End0, DatumEnds.End1]
                    else:
                        ends_to_set = [end]
                    
                    for e in ends_to_set:
                        if visible:
                            level.ShowBubbleInView(e, view)
                        else:
                            level.HideBubbleInView(e, view)
                    success += 1
                except Exception as ex:
                    errors.append("{}: {}".format(item.LevelName, str(ex)))
            t.Commit()
        except Exception as ex:
            t.RollBack()
            WPFMessageBox.Show("Transaction failed: {}".format(str(ex)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        self._load_levels()
        
        msg = "{} {} bubble for {} level(s).".format(action, end_name, success)
        if errors:
            msg += "\n\nErrors ({}):\n{}".format(len(errors), "\n".join(errors[:10]))
        WPFMessageBox.Show(msg, "Bubble Control", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_bubble_start_on(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End0, True)
    
    def _on_bubble_start_off(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End0, False)
    
    def _on_bubble_end_on(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End1, True)
    
    def _on_bubble_end_off(self, sender, args):
        self._set_bubble_visibility(DatumEnds.End1, False)
    
    def _on_bubble_all_on(self, sender, args):
        self._set_bubble_visibility(None, True)
    
    def _on_bubble_all_off(self, sender, args):
        self._set_bubble_visibility(None, False)
    
    # =====================================================
    # UI EVENT HANDLERS
    # =====================================================
    def _on_view_changed(self, sender, args):
        self._load_levels()
    
    def _on_filter_changed(self, sender, args):
        self._update_display()
    
    def _on_search_changed(self, sender, args):
        self._update_display()
    
    def _on_refresh(self, sender, args):
        self._load_views()
        self._load_levels()
    
    def _on_close(self, sender, args):
        self.window.Close()
    
    def _on_row_double_click(self, sender, args):
        """Select level in Revit on double-click"""
        if self.dg_levels.SelectedItem is not None:
            try:
                item = self.dg_levels.SelectedItem
                level = item.LevelElement
                element_id = level.Id
                uidoc.Selection.SetElementIds(System.Collections.Generic.List[ElementId]([element_id]))
                uidoc.ShowElements(element_id)
            except Exception as e:
                print("Error selecting level: {}".format(str(e)))
    
    def show(self):
        self.window.ShowDialog()


# =====================================================
# MAIN ENTRY POINT
# =====================================================
try:
    window = LevelSwapWindow()
    window.show()
except Exception as e:
    from pyrevit import forms
    forms.alert("Error launching Level Swap:\n{}".format(str(e)), title="Level Swap Error")