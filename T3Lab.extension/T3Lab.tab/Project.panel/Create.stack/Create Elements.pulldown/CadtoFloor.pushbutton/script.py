# -*- coding: utf-8 -*-
"""
CAD to Floor / Part v1.3 - DQT
Creates Revit Floor or Part (DirectShape) elements from linked/imported AutoCAD DWG geometry.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "CAD to\nFloor"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Create Revit Floors or Parts from linked/imported AutoCAD DWG layers."

# ==============================================================================
# IMPORTS - CRITICAL: Use aliased Revit DB import to avoid Grid conflict with WPF
# ==============================================================================
import clr
import os
import sys
import math

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xml')

# CRITICAL: Do NOT use "from Autodesk.Revit.DB import *" - conflicts with WPF Grid
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as RUI

from System import EventHandler
from System.Collections.Generic import List
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxResult, MessageBoxImage
from System.Windows import Visibility, Thickness, HorizontalAlignment, VerticalAlignment, TextAlignment, FontWeights
from System.Windows.Controls import *
from System.Windows.Media import BrushConverter
from System.Windows.Markup import XamlReader

# PyRevit
from pyrevit import revit

# ==============================================================================
# CONSTANTS
# ==============================================================================
MODE_FLOOR = "floor"
MODE_PART = "part"

# ==============================================================================
# XAML UI
# ==============================================================================
XAML_MAIN = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CAD to Floor / Part v1.3 - DQT"
        Width="820" Height="750"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResizeWithGrip"
        Background="#F5F5F5">
    <Window.Resources>
        <Style x:Key="ActionButton" TargetType="Button">
            <Setter Property="Height" Value="32"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="StyledComboBox" TargetType="ComboBox">
            <Setter Property="Height" Value="30"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="SearchBox" TargetType="TextBox">
            <Setter Property="Height" Value="30"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="ModeRadio" TargetType="RadioButton">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,0,20,0"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- HEADER -->
        <Border Grid.Row="0" Background="#0F172A" Padding="16,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="CAD to Floor / Part" FontSize="20" FontWeight="Bold" Foreground="#2D2D2D"/>
                    <TextBlock Text="Create Revit Floors or Parts from AutoCAD DWG layers" FontSize="11" Foreground="#FFFFFF" Margin="0,2,0,0"/>
                    <TextBlock Text="Copyright (c) 2025 Dang Quoc Truong (DQT). All rights reserved." 
                               FontSize="9" Foreground="#8B7355" Margin="0,4,0,0" FontStyle="Italic"/>
                </StackPanel>
                <TextBlock Grid.Column="1" VerticalAlignment="Bottom" HorizontalAlignment="Right"
                           FontSize="9" Foreground="#8B7355" Text="v1.3"/>
            </Grid>
        </Border>
        
        <!-- MODE SELECTION -->
        <Border Grid.Row="1" Background="White" Margin="12,8,12,0" Padding="12,10" CornerRadius="6"
                BorderBrush="#E8E8E8" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Create Mode:" VerticalAlignment="Center" 
                           FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <RadioButton x:Name="rb_floor" Content="Floor" GroupName="mode" 
                                 IsChecked="True" Style="{StaticResource ModeRadio}"/>
                    <RadioButton x:Name="rb_part" Content="Part (DirectShape)" GroupName="mode" 
                                 Style="{StaticResource ModeRadio}"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- SETTINGS -->
        <Border Grid.Row="2" Background="White" Margin="12,8,12,0" Padding="12" CornerRadius="6"
                BorderBrush="#E8E8E8" BorderThickness="1">
            <StackPanel>
                <!-- CAD File + Pick + Scan -->
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="CAD File:" VerticalAlignment="Center" 
                               FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                    <ComboBox x:Name="cmb_cad_files" Grid.Column="1" Style="{StaticResource StyledComboBox}"/>
                    <Button x:Name="btn_pick_cad" Grid.Column="3" Content="Pick"
                            Style="{StaticResource ActionButton}" Background="#0F172A" 
                            Foreground="#2D2D2D" Width="55" ToolTip="Pick a CAD instance from view"/>
                    <Button x:Name="btn_scan" Grid.Column="5" Content="Scan Layers"
                            Style="{StaticResource ActionButton}" Background="#C89650" 
                            Foreground="White" Width="95" ToolTip="Scan selected CAD file for layers"/>
                </Grid>
                
                <!-- Level -->
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Level:" VerticalAlignment="Center" 
                               FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                    <ComboBox x:Name="cmb_levels" Grid.Column="1" Style="{StaticResource StyledComboBox}"/>
                </Grid>
                
                <!-- Floor Type (Floor mode) -->
                <Grid x:Name="pnl_floor_type_row" Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Floor Type:" VerticalAlignment="Center" 
                               FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                    <ComboBox x:Name="cmb_floor_types" Grid.Column="1" Style="{StaticResource StyledComboBox}"/>
                </Grid>
                
                <!-- Part Category (Part mode) -->
                <Grid x:Name="pnl_part_row1" Margin="0,8,0,0" Visibility="Collapsed">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Category:" VerticalAlignment="Center" 
                               FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                    <ComboBox x:Name="cmb_ds_category" Grid.Column="1" Style="{StaticResource StyledComboBox}"/>
                </Grid>
                
                <!-- Part Thickness (Part mode) -->
                <Grid x:Name="pnl_part_row2" Margin="0,8,0,0" Visibility="Collapsed">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="120"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Thickness:" VerticalAlignment="Center" 
                               FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                    <TextBox x:Name="txt_thickness" Grid.Column="1" Text="200" Style="{StaticResource SearchBox}"/>
                    <TextBlock Grid.Column="2" Text="  mm" VerticalAlignment="Center" FontSize="12" Foreground="#888"/>
                </Grid>
                
                <!-- Offset + Structural -->
                <Grid Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="120"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Offset (mm):" VerticalAlignment="Center" 
                               FontWeight="SemiBold" FontSize="12" Foreground="#2D2D2D"/>
                    <TextBox x:Name="txt_offset" Grid.Column="1" Text="0" Style="{StaticResource SearchBox}"/>
                    <TextBlock Grid.Column="2" Text="  mm" VerticalAlignment="Center" FontSize="12" Foreground="#888"/>
                    <CheckBox x:Name="chk_structural" Grid.Column="4" Content="Structural Floor" 
                              VerticalAlignment="Center" FontSize="12" IsChecked="False"/>
                </Grid>
            </StackPanel>
        </Border>
        
        <!-- SUMMARY CARDS -->
        <Border Grid.Row="3" Margin="12,8,12,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" Background="White" CornerRadius="6" Padding="10,8"
                        BorderBrush="#E8E8E8" BorderThickness="1">
                    <StackPanel HorizontalAlignment="Center">
                        <TextBlock x:Name="txt_total_layers" Text="0" FontSize="20" FontWeight="Bold" 
                                   Foreground="#C89650" HorizontalAlignment="Center"/>
                        <TextBlock Text="Total Layers" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                    </StackPanel>
                </Border>
                <Border Grid.Column="2" Background="White" CornerRadius="6" Padding="10,8"
                        BorderBrush="#E8E8E8" BorderThickness="1">
                    <StackPanel HorizontalAlignment="Center">
                        <TextBlock x:Name="txt_selected_layers" Text="0" FontSize="20" FontWeight="Bold" 
                                   Foreground="#2196F3" HorizontalAlignment="Center"/>
                        <TextBlock Text="Selected" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                    </StackPanel>
                </Border>
                <Border Grid.Column="4" Background="White" CornerRadius="6" Padding="10,8"
                        BorderBrush="#E8E8E8" BorderThickness="1">
                    <StackPanel HorizontalAlignment="Center">
                        <TextBlock x:Name="txt_total_curves" Text="0" FontSize="20" FontWeight="Bold" 
                                   Foreground="#10B981" HorizontalAlignment="Center"/>
                        <TextBlock Text="Closed Loops" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                    </StackPanel>
                </Border>
                <Border Grid.Column="6" Background="White" CornerRadius="6" Padding="10,8"
                        BorderBrush="#E8E8E8" BorderThickness="1">
                    <StackPanel HorizontalAlignment="Center">
                        <TextBlock x:Name="txt_elements_created" Text="0" FontSize="20" FontWeight="Bold" 
                                   Foreground="#0F172A" HorizontalAlignment="Center"/>
                        <TextBlock x:Name="lbl_created" Text="Created" FontSize="10" Foreground="#888" HorizontalAlignment="Center"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
        
        <!-- LAYER LIST -->
        <Border Grid.Row="4" Background="White" Margin="12,8,12,0" Padding="0" CornerRadius="6"
                BorderBrush="#E8E8E8" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Grid Grid.Row="0" Margin="10,10,10,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="8"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="txt_search" Grid.Column="0" Style="{StaticResource SearchBox}" 
                             ToolTip="Search layers by name..."/>
                    <Button x:Name="btn_select_all" Grid.Column="2" Content="Select All"
                            Style="{StaticResource ActionButton}" Background="#0F172A" Foreground="#2D2D2D" Width="80"/>
                    <Button x:Name="btn_select_none" Grid.Column="4" Content="Select None"
                            Style="{StaticResource ActionButton}" Background="#EEEEEE" Foreground="#555" Width="85"/>
                </Grid>
                <Grid Grid.Row="1" Margin="10,8,10,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="35"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="100"/>
                        <ColumnDefinition Width="80"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="1" Text="Layer Name" FontSize="10" FontWeight="SemiBold" Foreground="#999"/>
                    <TextBlock Grid.Column="2" Text="Closed Loops" FontSize="10" FontWeight="SemiBold" 
                               Foreground="#999" TextAlignment="Center"/>
                    <TextBlock Grid.Column="3" Text="All Curves" FontSize="10" FontWeight="SemiBold" 
                               Foreground="#999" TextAlignment="Center"/>
                </Grid>
                <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" Margin="10,4,10,10">
                    <StackPanel x:Name="pnl_layers"/>
                </ScrollViewer>
            </Grid>
        </Border>
        
        <!-- STATUS -->
        <Border Grid.Row="5" Margin="12,8,12,0">
            <TextBlock x:Name="txt_status" Text="Select a CAD file and click 'Scan Layers' to begin." 
                       FontSize="11" Foreground="#888" VerticalAlignment="Center"/>
        </Border>
        
        <!-- ACTION BUTTONS -->
        <Border Grid.Row="6" Background="White" Padding="12,10" Margin="0,8,0,0"
                BorderBrush="#E8E8E8" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="btn_refresh" Grid.Column="0" Content="Refresh"
                        Style="{StaticResource ActionButton}" Background="#EEEEEE" 
                        Foreground="#555" Width="90"/>
                <Button x:Name="btn_create" Grid.Column="2" Content="Create Floors"
                        Style="{StaticResource ActionButton}" Background="#C89650" 
                        Foreground="White" Width="140" IsEnabled="False"/>
                <Button x:Name="btn_close" Grid.Column="4" Content="Close"
                        Style="{StaticResource ActionButton}" Background="#EEEEEE" 
                        Foreground="#555" Width="90"/>
            </Grid>
        </Border>
        
        <!-- COPYRIGHT FOOTER -->
        <Border Grid.Row="7" Background="#0F172A" Padding="8,6">
            <TextBlock Text="Copyright (c) 2025 Dang Quoc Truong (DQT). All rights reserved." 
                       FontSize="9" Foreground="#FFFFFF" HorizontalAlignment="Center" FontStyle="Italic"/>
        </Border>
    </Grid>
</Window>
"""

XAML_LAYER_ITEM = """
<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Background="Transparent" Padding="4,6" Margin="0,1"
        CornerRadius="6" BorderBrush="Transparent" BorderThickness="1"
        Cursor="Hand">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="35"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="100"/>
            <ColumnDefinition Width="80"/>
        </Grid.ColumnDefinitions>
        <CheckBox x:Name="chk" Grid.Column="0" VerticalAlignment="Center" IsChecked="False"/>
        <TextBlock x:Name="txt_name" Grid.Column="1" VerticalAlignment="Center" FontSize="12"
                   Foreground="#2D2D2D" TextTrimming="CharacterEllipsis"/>
        <TextBlock x:Name="txt_closed" Grid.Column="2" VerticalAlignment="Center" FontSize="12"
                   Foreground="#10B981" FontWeight="SemiBold" TextAlignment="Center"/>
        <TextBlock x:Name="txt_total" Grid.Column="3" VerticalAlignment="Center" FontSize="12"
                   Foreground="#888" TextAlignment="Center"/>
    </Grid>
</Border>
"""


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def load_xaml_from_string(xaml_string):
    """Load XAML from string"""
    string_reader = StringReader(xaml_string)
    xml_reader = XmlReader.Create(string_reader)
    return XamlReader.Load(xml_reader)


def mm_to_feet(mm):
    return mm / 304.8


def feet_to_mm(feet):
    return feet * 304.8


def safe_bool(nullable_bool):
    """Safely convert Nullable[Boolean] to Python bool"""
    try:
        if nullable_bool is None:
            return False
        return bool(nullable_bool)
    except Exception:
        return False


def get_cad_instances(doc):
    """Get all linked/imported CAD instances"""
    cad_instances = []
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance)
        for inst in collector:
            try:
                name = "Unknown"
                type_elem = doc.GetElement(inst.GetTypeId())
                if type_elem:
                    param = type_elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    if param and param.AsString():
                        name = param.AsString()
                    elif hasattr(type_elem, 'Name'):
                        name = type_elem.Name
                
                is_linked = False
                try:
                    is_linked = inst.IsLinked
                except Exception:
                    pass
                
                cad_instances.append({
                    'element': inst,
                    'id': inst.Id,
                    'name': name,
                    'is_linked': is_linked
                })
            except Exception:
                pass
    except Exception:
        pass
    return cad_instances


def get_cad_layer_geometry(doc, cad_instance):
    """Extract geometry from CAD instance, organized by layer"""
    layers = {}
    
    try:
        opts = DB.Options()
        geom_elem = cad_instance.get_Geometry(opts)
        if geom_elem is None:
            return layers
        
        for geom_obj in geom_elem:
            if isinstance(geom_obj, DB.GeometryInstance):
                sub_geom = geom_obj.GetInstanceGeometry()
                if sub_geom:
                    for sub_obj in sub_geom:
                        layer_name = "Default"
                        try:
                            style = doc.GetElement(sub_obj.GraphicsStyleId)
                            if style:
                                cat = style.GraphicsStyleCategory
                                if cat:
                                    layer_name = cat.Name
                        except Exception:
                            pass
                        _process_geom_obj(sub_obj, layers, layer_name)
            else:
                _process_geom_obj(geom_obj, layers, "Default")
    except Exception:
        pass
    
    return layers


def _process_geom_obj(geom_obj, layers, layer_name):
    """Process a single geometry object and add to layer dict"""
    if layer_name not in layers:
        layers[layer_name] = {
            'curves': [],
            'closed_loops': [],
            'all_curves_count': 0
        }
    
    layer = layers[layer_name]
    
    try:
        if isinstance(geom_obj, DB.PolyLine):
            coords = geom_obj.GetCoordinates()
            layer['all_curves_count'] += 1
            if coords.Count >= 4:  # Need at least 4 points (3 segments + close)
                first = coords[0]
                last = coords[coords.Count - 1]
                if first.DistanceTo(last) < 0.01:
                    try:
                        curve_loop = DB.CurveLoop()
                        valid = True
                        for i in range(coords.Count - 1):
                            p1 = coords[i]
                            p2 = coords[i + 1]
                            dist = p1.DistanceTo(p2)
                            if dist > 0.001:
                                line = DB.Line.CreateBound(p1, p2)
                                curve_loop.Append(line)
                            elif dist > 0.0001:
                                # Too short segment - skip but continue
                                pass
                        if not curve_loop.IsOpen():
                            layer['closed_loops'].append(curve_loop)
                    except Exception:
                        pass
        
        elif isinstance(geom_obj, DB.Line):
            layer['all_curves_count'] += 1
            layer['curves'].append(geom_obj)
        
        elif isinstance(geom_obj, DB.Arc):
            layer['all_curves_count'] += 1
            layer['curves'].append(geom_obj)
        
        elif isinstance(geom_obj, DB.Curve):
            layer['all_curves_count'] += 1
            layer['curves'].append(geom_obj)
    except Exception:
        pass


def try_build_loops_from_curves(curves):
    """Try to build closed CurveLoops from individual curves by endpoint matching"""
    if not curves:
        return []
    
    closed_loops = []
    used = set()
    tolerance = 0.01
    
    for start_idx in range(len(curves)):
        if start_idx in used:
            continue
        
        try:
            loop_curves = [curves[start_idx]]
            used_in_loop = {start_idx}
            current_end = curves[start_idx].GetEndPoint(1)
            start_point = curves[start_idx].GetEndPoint(0)
            
            max_iter = len(curves)
            it = 0
            
            while it < max_iter:
                it += 1
                found = False
                
                if current_end.DistanceTo(start_point) < tolerance and len(loop_curves) >= 3:
                    try:
                        cl = DB.CurveLoop()
                        for c in loop_curves:
                            cl.Append(c)
                        if not cl.IsOpen():
                            closed_loops.append(cl)
                            used.update(used_in_loop)
                    except Exception:
                        pass
                    break
                
                for j in range(len(curves)):
                    if j in used or j in used_in_loop:
                        continue
                    try:
                        c = curves[j]
                        p0 = c.GetEndPoint(0)
                        p1 = c.GetEndPoint(1)
                        
                        if current_end.DistanceTo(p0) < tolerance:
                            loop_curves.append(c)
                            used_in_loop.add(j)
                            current_end = p1
                            found = True
                            break
                        elif current_end.DistanceTo(p1) < tolerance:
                            rev = c.CreateReversed()
                            loop_curves.append(rev)
                            used_in_loop.add(j)
                            current_end = rev.GetEndPoint(1)
                            found = True
                            break
                    except Exception:
                        pass
                
                if not found:
                    break
        except Exception:
            pass
    
    return closed_loops


def get_floor_types(doc):
    """Get all floor types"""
    floor_types = []
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.FloorType)
        for ft in collector:
            try:
                name_param = ft.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
                type_name = name_param.AsString() if name_param and name_param.AsString() else "Unknown"
                family_name = ""
                try:
                    family_name = ft.FamilyName
                except Exception:
                    pass
                display = "{}: {}".format(family_name, type_name) if family_name else type_name
                floor_types.append({
                    'id': ft.Id,
                    'name': display,
                    'element': ft
                })
            except Exception:
                pass
    except Exception:
        pass
    floor_types.sort(key=lambda x: x['name'])
    return floor_types


def get_levels(doc):
    """Get all levels sorted by elevation"""
    levels = []
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.Level)
        for lvl in collector:
            try:
                levels.append({
                    'id': lvl.Id,
                    'name': lvl.Name,
                    'elevation': lvl.Elevation,
                    'element': lvl
                })
            except Exception:
                pass
    except Exception:
        pass
    levels.sort(key=lambda x: x['elevation'])
    return levels


def get_ds_categories():
    """Get common categories for DirectShape"""
    return [
        {"name": "Floors", "bic": DB.BuiltInCategory.OST_Floors},
        {"name": "Generic Models", "bic": DB.BuiltInCategory.OST_GenericModel},
        {"name": "Mass", "bic": DB.BuiltInCategory.OST_Mass},
        {"name": "Structural Foundations", "bic": DB.BuiltInCategory.OST_StructuralFoundation},
        {"name": "Walls", "bic": DB.BuiltInCategory.OST_Walls},
        {"name": "Roofs", "bic": DB.BuiltInCategory.OST_Roofs},
        {"name": "Ceilings", "bic": DB.BuiltInCategory.OST_Ceilings},
        {"name": "Site", "bic": DB.BuiltInCategory.OST_Site},
    ]


def create_floor_from_loop(doc, curve_loop, floor_type_id, level_id, offset_mm=0, is_structural=False):
    """Create a Floor element from a CurveLoop"""
    try:
        # Try newer API first (Revit 2022+)
        try:
            loop_list = List[DB.CurveLoop]()
            loop_list.Add(curve_loop)
            floor = DB.Floor.Create(doc, loop_list, floor_type_id, level_id)
            if floor and offset_mm != 0:
                param = floor.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                if param and not param.IsReadOnly:
                    param.Set(mm_to_feet(offset_mm))
            return floor
        except Exception:
            pass
        
        # Fallback to legacy API
        try:
            curve_array = DB.CurveArray()
            for curve in curve_loop:
                curve_array.Append(curve)
            floor_type = doc.GetElement(floor_type_id)
            level = doc.GetElement(level_id)
            normal = DB.XYZ.BasisZ
            floor = doc.Create.NewFloor(curve_array, floor_type, level, is_structural, normal)
            if floor and offset_mm != 0:
                param = floor.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                if param and not param.IsReadOnly:
                    param.Set(mm_to_feet(offset_mm))
            return floor
        except Exception:
            pass
    except Exception:
        pass
    return None


def create_part_from_loop(doc, curve_loop, category_bic, level_id, thickness_mm=200, offset_mm=0):
    """Create a DirectShape (Part) by extruding a CurveLoop vertically"""
    try:
        thickness_ft = mm_to_feet(thickness_mm)
        offset_ft = mm_to_feet(offset_mm)
        
        if curve_loop.IsOpen():
            return None
        
        # Build profile
        profile = List[DB.CurveLoop]()
        profile.Add(curve_loop)
        
        # Extrude upward
        solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            profile, DB.XYZ.BasisZ, thickness_ft)
        
        if solid is None:
            return None
        
        try:
            vol = solid.Volume
            if vol < 0.0001:
                return None
        except Exception:
            return None
        
        # Apply offset
        if abs(offset_ft) > 0.0001:
            try:
                move_tf = DB.Transform.CreateTranslation(DB.XYZ(0, 0, offset_ft))
                moved = DB.SolidUtils.CreateTransformed(solid, move_tf)
                # CreateTransformed may return a list or single solid
                if moved is not None:
                    solid = moved
            except Exception:
                pass
        
        # Create DirectShape
        cat_id = DB.ElementId(category_bic)
        ds = DB.DirectShape.CreateElement(doc, cat_id)
        
        geom_list = List[DB.GeometryObject]()
        geom_list.Add(solid)
        ds.SetShape(geom_list)
        ds.Name = "DQT Part"
        
        return ds
    except Exception:
        pass
    return None


# ==============================================================================
# LAYER DATA CLASS
# ==============================================================================
class LayerData(object):
    def __init__(self, name, closed_loops, all_curves_count):
        self.name = name
        self.closed_loops = closed_loops
        self.all_curves_count = all_curves_count
        self.closed_count = len(closed_loops)
        self.is_selected = False


# ==============================================================================
# MAIN WINDOW CLASS
# ==============================================================================
class CADtoFloorWindow(object):
    
    def __init__(self):
        self.window = load_xaml_from_string(XAML_MAIN)
        
        # Get UI elements
        self.rb_floor = self.window.FindName("rb_floor")
        self.rb_part = self.window.FindName("rb_part")
        self.cmb_cad_files = self.window.FindName("cmb_cad_files")
        self.btn_pick_cad = self.window.FindName("btn_pick_cad")
        self.btn_scan = self.window.FindName("btn_scan")
        self.cmb_levels = self.window.FindName("cmb_levels")
        self.pnl_floor_type_row = self.window.FindName("pnl_floor_type_row")
        self.cmb_floor_types = self.window.FindName("cmb_floor_types")
        self.pnl_part_row1 = self.window.FindName("pnl_part_row1")
        self.pnl_part_row2 = self.window.FindName("pnl_part_row2")
        self.cmb_ds_category = self.window.FindName("cmb_ds_category")
        self.txt_thickness = self.window.FindName("txt_thickness")
        self.txt_offset = self.window.FindName("txt_offset")
        self.chk_structural = self.window.FindName("chk_structural")
        self.txt_total_layers = self.window.FindName("txt_total_layers")
        self.txt_selected_layers = self.window.FindName("txt_selected_layers")
        self.txt_total_curves = self.window.FindName("txt_total_curves")
        self.txt_elements_created = self.window.FindName("txt_elements_created")
        self.lbl_created = self.window.FindName("lbl_created")
        self.txt_search = self.window.FindName("txt_search")
        self.btn_select_all = self.window.FindName("btn_select_all")
        self.btn_select_none = self.window.FindName("btn_select_none")
        self.pnl_layers = self.window.FindName("pnl_layers")
        self.txt_status = self.window.FindName("txt_status")
        self.btn_refresh = self.window.FindName("btn_refresh")
        self.btn_create = self.window.FindName("btn_create")
        self.btn_close = self.window.FindName("btn_close")
        
        # Data
        self.doc = revit.doc
        self.uidoc = revit.uidoc
        self.cad_instances = []
        self.floor_types = []
        self.levels = []
        self.ds_categories = get_ds_categories()
        self.layer_data = []
        self.layer_checkboxes = []
        self.elements_created_count = 0
        self.current_mode = MODE_FLOOR
        self._pick_element_id = None  # For pick workflow
        
        # Events - NOTE: NO auto-scan on cmb_cad_files.SelectionChanged (too slow)
        self.btn_pick_cad.Click += self._on_pick_cad
        self.btn_scan.Click += self._on_scan_layers
        self.txt_search.TextChanged += self._on_search_changed
        self.btn_select_all.Click += self._on_select_all
        self.btn_select_none.Click += self._on_select_none
        self.btn_refresh.Click += self._on_refresh
        self.btn_create.Click += self._on_create_elements
        self.btn_close.Click += self._on_close
        self.rb_floor.Checked += self._on_mode_changed
        self.rb_part.Checked += self._on_mode_changed
        
        # Load data
        self._load_cad_files()
        self._load_floor_types()
        self._load_levels()
        self._load_ds_categories()
        self._update_mode_ui()
    
    def _update_status(self, msg):
        self.txt_status.Text = str(msg)
    
    def _update_mode_ui(self):
        if safe_bool(self.rb_floor.IsChecked):
            self.current_mode = MODE_FLOOR
            self.pnl_floor_type_row.Visibility = Visibility.Visible
            self.pnl_part_row1.Visibility = Visibility.Collapsed
            self.pnl_part_row2.Visibility = Visibility.Collapsed
            self.chk_structural.Visibility = Visibility.Visible
            self.btn_create.Content = "Create Floors"
            self.lbl_created.Text = "Floors Created"
        else:
            self.current_mode = MODE_PART
            self.pnl_floor_type_row.Visibility = Visibility.Collapsed
            self.pnl_part_row1.Visibility = Visibility.Visible
            self.pnl_part_row2.Visibility = Visibility.Visible
            self.chk_structural.Visibility = Visibility.Collapsed
            self.btn_create.Content = "Create Parts"
            self.lbl_created.Text = "Parts Created"
    
    def _update_summary(self):
        total_layers = len(self.layer_data)
        selected = 0
        for _, ld in self.layer_checkboxes:
            if ld.is_selected:
                selected += 1
        total_closed = 0
        for ld in self.layer_data:
            total_closed += ld.closed_count
        
        self.txt_total_layers.Text = str(total_layers)
        self.txt_selected_layers.Text = str(selected)
        self.txt_total_curves.Text = str(total_closed)
        self.txt_elements_created.Text = str(self.elements_created_count)
        self.btn_create.IsEnabled = selected > 0
    
    def _load_cad_files(self):
        self.cad_instances = get_cad_instances(self.doc)
        self.cmb_cad_files.Items.Clear()
        
        if not self.cad_instances:
            self.cmb_cad_files.Items.Add("No CAD files found in model")
            self.cmb_cad_files.SelectedIndex = 0
            self.cmb_cad_files.IsEnabled = False
            self.btn_scan.IsEnabled = False
            self._update_status("No linked/imported CAD files found.")
            return
        
        for cad in self.cad_instances:
            prefix = "[Linked]" if cad['is_linked'] else "[Imported]"
            display = "{} {} (ID: {})".format(prefix, cad['name'], str(cad['id']))
            self.cmb_cad_files.Items.Add(display)
        
        self.cmb_cad_files.IsEnabled = True
        self.btn_scan.IsEnabled = True
        self.cmb_cad_files.SelectedIndex = 0
        self._update_status("{} CAD file(s) found. Select one and click 'Scan Layers'.".format(len(self.cad_instances)))
    
    def _load_floor_types(self):
        self.floor_types = get_floor_types(self.doc)
        self.cmb_floor_types.Items.Clear()
        for ft in self.floor_types:
            self.cmb_floor_types.Items.Add(ft['name'])
        if self.floor_types:
            self.cmb_floor_types.SelectedIndex = 0
    
    def _load_levels(self):
        self.levels = get_levels(self.doc)
        self.cmb_levels.Items.Clear()
        for lvl in self.levels:
            elev_mm = str(int(round(feet_to_mm(lvl['elevation']))))
            display = "{} ({} mm)".format(lvl['name'], elev_mm)
            self.cmb_levels.Items.Add(display)
        if self.levels:
            self.cmb_levels.SelectedIndex = 0
    
    def _load_ds_categories(self):
        self.cmb_ds_category.Items.Clear()
        for cat in self.ds_categories:
            self.cmb_ds_category.Items.Add(cat['name'])
        if self.ds_categories:
            self.cmb_ds_category.SelectedIndex = 0
    
    def _scan_layers(self, cad_instance):
        self.layer_data = []
        self._update_status("Scanning CAD layers...")
        
        try:
            layers_geom = get_cad_layer_geometry(self.doc, cad_instance)
            
            for layer_name, data in sorted(layers_geom.items()):
                closed_loops = list(data.get('closed_loops', []))
                all_curves_count = data.get('all_curves_count', 0)
                
                individual_curves = data.get('curves', [])
                if individual_curves:
                    extra_loops = try_build_loops_from_curves(individual_curves)
                    closed_loops.extend(extra_loops)
                
                ld = LayerData(layer_name, closed_loops, all_curves_count)
                self.layer_data.append(ld)
        except Exception as ex:
            self._update_status("Error scanning: {}".format(str(ex)))
        
        self._render_layers()
        self._update_summary()
        
        total_closed = 0
        for ld in self.layer_data:
            total_closed += ld.closed_count
        self._update_status("Found {} layers with {} closed loops.".format(
            len(self.layer_data), total_closed))
    
    def _render_layers(self, filter_text=""):
        self.pnl_layers.Children.Clear()
        self.layer_checkboxes = []
        
        filter_lower = filter_text.lower().strip()
        
        for ld in self.layer_data:
            if filter_lower and filter_lower not in ld.name.lower():
                continue
            
            try:
                border = load_xaml_from_string(XAML_LAYER_ITEM)
                chk = border.FindName("chk")
                txt_name = border.FindName("txt_name")
                txt_closed = border.FindName("txt_closed")
                txt_total = border.FindName("txt_total")
                
                txt_name.Text = ld.name
                txt_closed.Text = str(ld.closed_count)
                txt_total.Text = str(ld.all_curves_count)
                chk.IsChecked = ld.is_selected
                
                if ld.closed_count > 0:
                    txt_name.FontWeight = FontWeights.SemiBold
                    txt_closed.Foreground = BrushConverter().ConvertFromString("#10B981")
                else:
                    txt_closed.Foreground = BrushConverter().ConvertFromString("#CCC")
                
                # Checkbox toggle handler
                def make_handler(layer_data, checkbox):
                    def handler(sender, args):
                        layer_data.is_selected = safe_bool(checkbox.IsChecked)
                        self._update_summary()
                    return handler
                
                chk.Checked += make_handler(ld, chk)
                chk.Unchecked += make_handler(ld, chk)
                
                # Border click -> toggle checkbox
                def make_border_handler(checkbox):
                    def handler(sender, args):
                        try:
                            src = args.OriginalSource
                            if src != checkbox:
                                checkbox.IsChecked = not safe_bool(checkbox.IsChecked)
                        except Exception:
                            pass
                    return handler
                
                border.MouseLeftButtonUp += make_border_handler(chk)
                
                # Hover
                def make_enter(brd):
                    def handler(s, e):
                        try:
                            brd.Background = BrushConverter().ConvertFromString("#FFF5E0")
                            brd.BorderBrush = BrushConverter().ConvertFromString("#0F172A")
                        except Exception:
                            pass
                    return handler
                
                def make_leave(brd):
                    def handler(s, e):
                        try:
                            brd.Background = BrushConverter().ConvertFromString("Transparent")
                            brd.BorderBrush = BrushConverter().ConvertFromString("Transparent")
                        except Exception:
                            pass
                    return handler
                
                border.MouseEnter += make_enter(border)
                border.MouseLeave += make_leave(border)
                
                self.pnl_layers.Children.Add(border)
                self.layer_checkboxes.append((chk, ld))
            except Exception:
                pass
        
        self._update_summary()
    
    # ========== EVENT HANDLERS ==========
    
    def _on_mode_changed(self, sender, args):
        self._update_mode_ui()
    
    def _on_pick_cad(self, sender, args):
        """Pick CAD from viewport - close dialog, pick, then reopen"""
        try:
            # Store that we want to pick, then close dialog
            self._pick_element_id = "PICK_REQUESTED"
            self.window.Close()
        except Exception as ex:
            self._update_status("Error: {}".format(str(ex)))
    
    def _on_scan_layers(self, sender, args):
        """User clicked Scan Layers button"""
        idx = self.cmb_cad_files.SelectedIndex
        if idx < 0 or idx >= len(self.cad_instances):
            MessageBox.Show("Please select a CAD file first.", "CAD to Floor/Part",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        cad = self.cad_instances[idx]
        self._scan_layers(cad['element'])
    
    def _on_search_changed(self, sender, args):
        self._render_layers(self.txt_search.Text)
    
    def _on_select_all(self, sender, args):
        for chk, ld in self.layer_checkboxes:
            chk.IsChecked = True
            ld.is_selected = True
        self._update_summary()
    
    def _on_select_none(self, sender, args):
        for chk, ld in self.layer_checkboxes:
            chk.IsChecked = False
            ld.is_selected = False
        self._update_summary()
    
    def _on_refresh(self, sender, args):
        self.elements_created_count = 0
        self._load_cad_files()
        self._load_floor_types()
        self._load_levels()
        self._load_ds_categories()
        self.layer_data = []
        self.layer_checkboxes = []
        self.pnl_layers.Children.Clear()
        self._update_summary()
        self._update_status("Refreshed.")
    
    def _on_create_elements(self, sender, args):
        selected_layers = []
        for _, ld in self.layer_checkboxes:
            if ld.is_selected:
                selected_layers.append(ld)
        
        if not selected_layers:
            MessageBox.Show("No layers selected.", "CAD to Floor/Part", 
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        total_loops = 0
        for ld in selected_layers:
            total_loops += ld.closed_count
        
        if total_loops == 0:
            MessageBox.Show(
                "Selected layers have no closed loops.\nOnly closed polylines can be converted.",
                "CAD to Floor/Part", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        lvl_idx = self.cmb_levels.SelectedIndex
        if lvl_idx < 0 or lvl_idx >= len(self.levels):
            MessageBox.Show("Please select a Level.", "CAD to Floor/Part",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        level_id = self.levels[lvl_idx]['id']
        
        try:
            offset_mm = float(self.txt_offset.Text)
        except (ValueError, TypeError):
            offset_mm = 0.0
        
        if self.current_mode == MODE_FLOOR:
            self._create_floors(selected_layers, total_loops, level_id, offset_mm)
        else:
            self._create_parts(selected_layers, total_loops, level_id, offset_mm)
    
    def _create_floors(self, selected_layers, total_loops, level_id, offset_mm):
        ft_idx = self.cmb_floor_types.SelectedIndex
        if ft_idx < 0 or ft_idx >= len(self.floor_types):
            MessageBox.Show("Please select a Floor Type.", "CAD to Floor",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        floor_type_id = self.floor_types[ft_idx]['id']
        is_structural = safe_bool(self.chk_structural.IsChecked)
        
        msg = "Create FLOORS from {} layer(s) with {} loop(s)?\n\n".format(
            len(selected_layers), total_loops)
        msg += "Floor Type: {}\n".format(self.floor_types[ft_idx]['name'])
        msg += "Level: {}\n".format(self.levels[self.cmb_levels.SelectedIndex]['name'])
        msg += "Offset: {} mm".format(offset_mm)
        
        result = MessageBox.Show(msg, "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question)
        if result != MessageBoxResult.Yes:
            return
        
        self._update_status("Creating floors...")
        created = 0
        failed = 0
        
        t = DB.Transaction(self.doc, "DQT - CAD to Floor")
        t.Start()
        
        try:
            for ld in selected_layers:
                for loop in ld.closed_loops:
                    try:
                        floor = create_floor_from_loop(
                            self.doc, loop, floor_type_id, level_id, offset_mm, is_structural)
                        if floor:
                            created += 1
                        else:
                            failed += 1
                    except Exception:
                        failed += 1
            
            if created > 0:
                t.Commit()
                self.elements_created_count += created
                self._update_summary()
                self._update_status("Created {} floor(s). {} failed.".format(created, failed))
                MessageBox.Show("Created: {} floor(s)\nFailed: {}".format(created, failed),
                              "Result", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                t.RollBack()
                self._update_status("No floors created.")
                MessageBox.Show("Failed to create any floors.", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
        except Exception as ex:
            try:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
            except Exception:
                pass
            self._update_status("Error: {}".format(str(ex)))
    
    def _create_parts(self, selected_layers, total_loops, level_id, offset_mm):
        cat_idx = self.cmb_ds_category.SelectedIndex
        if cat_idx < 0 or cat_idx >= len(self.ds_categories):
            MessageBox.Show("Please select a Category.", "CAD to Part",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        category_bic = self.ds_categories[cat_idx]['bic']
        
        try:
            thickness_mm = float(self.txt_thickness.Text)
            if thickness_mm <= 0:
                MessageBox.Show("Thickness must be > 0.", "CAD to Part",
                              MessageBoxButton.OK, MessageBoxImage.Warning)
                return
        except (ValueError, TypeError):
            MessageBox.Show("Invalid thickness.", "CAD to Part",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        msg = "Create PARTS from {} layer(s) with {} loop(s)?\n\n".format(
            len(selected_layers), total_loops)
        msg += "Category: {}\n".format(self.ds_categories[cat_idx]['name'])
        msg += "Thickness: {} mm\n".format(thickness_mm)
        msg += "Offset: {} mm".format(offset_mm)
        
        result = MessageBox.Show(msg, "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question)
        if result != MessageBoxResult.Yes:
            return
        
        self._update_status("Creating parts...")
        created = 0
        failed = 0
        
        t = DB.Transaction(self.doc, "DQT - CAD to Part")
        t.Start()
        
        try:
            for ld in selected_layers:
                for loop in ld.closed_loops:
                    try:
                        ds = create_part_from_loop(
                            self.doc, loop, category_bic, level_id, thickness_mm, offset_mm)
                        if ds:
                            created += 1
                        else:
                            failed += 1
                    except Exception:
                        failed += 1
            
            if created > 0:
                t.Commit()
                self.elements_created_count += created
                self._update_summary()
                self._update_status("Created {} part(s). {} failed.".format(created, failed))
                MessageBox.Show("Created: {} part(s)\nFailed: {}".format(created, failed),
                              "Result", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                t.RollBack()
                self._update_status("No parts created.")
                MessageBox.Show("Failed to create any parts.", "Error",
                              MessageBoxButton.OK, MessageBoxImage.Error)
        except Exception as ex:
            try:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
            except Exception:
                pass
            self._update_status("Error: {}".format(str(ex)))
    
    def _on_close(self, sender, args):
        self._pick_element_id = None
        self.window.Close()
    
    def show(self):
        self.window.ShowDialog()


# ==============================================================================
# PICK WORKFLOW HELPER
# ==============================================================================
def do_pick_cad(uidoc, doc):
    """Pick a CAD ImportInstance from the Revit view. Returns ElementId or None."""
    try:
        from Autodesk.Revit.UI.Selection import ObjectType
        ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick a CAD (DWG) instance")
        if ref:
            elem = doc.GetElement(ref.ElementId)
            if elem and isinstance(elem, DB.ImportInstance):
                return ref.ElementId
    except Exception:
        pass
    return None


# ==============================================================================
# ENTRY POINT
# ==============================================================================
try:
    doc = revit.doc
    uidoc = revit.uidoc
    
    picked_cad_id = None
    
    while True:
        win = CADtoFloorWindow()
        
        # If we have a picked CAD ID from previous iteration, select it and auto-scan
        if picked_cad_id is not None:
            for i, cad in enumerate(win.cad_instances):
                if cad['id'] == picked_cad_id:
                    win.cmb_cad_files.SelectedIndex = i
                    win._update_status("Picked: {}. Scanning layers...".format(win.cmb_cad_files.Items[i]))
                    # Auto-scan since user explicitly picked this CAD
                    win._scan_layers(cad['element'])
                    break
            picked_cad_id = None
        
        win.show()
        
        # After dialog closes, check if Pick was requested
        if win._pick_element_id == "PICK_REQUESTED":
            # Do the pick outside the dialog
            new_id = do_pick_cad(uidoc, doc)
            if new_id is not None:
                picked_cad_id = new_id
            # Loop back to reopen dialog
            continue
        else:
            # Normal close - exit loop
            break

except Exception as ex:
    from pyrevit import forms
    forms.alert(
        "Error launching CAD to Floor/Part:\n{}".format(str(ex)),
        title="CAD to Floor/Part - DQT",
        exitscript=True
    )