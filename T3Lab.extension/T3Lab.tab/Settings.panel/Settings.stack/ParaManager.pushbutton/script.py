# -*- coding: utf-8 -*-
"""
Parameter Manager (Standalone)
Manage Parameters with beautiful UI - Sheet Manager Style

Compatible with Revit 2024 / 2025 / 2026

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""
__title__ = "Parameter\nManage"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Manage and view Parameters - By DQT"

import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import (Window, Thickness, GridLength, GridUnitType,
                            HorizontalAlignment, VerticalAlignment, FontWeights,
                            MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult)
from System.Windows.Controls import (StackPanel, TextBlock, TextBox, Button,
                                      ComboBox, ComboBoxItem, DataGrid, Orientation,
                                      DataGridTextColumn,
                                      ScrollViewer, TabControl, TabItem, CheckBox,
                                      ListBox, ListBoxItem, SelectionMode)
# Import WPF Grid with alias BEFORE Revit DB wildcard import
import System.Windows.Controls as WPFControls
WPFGrid = WPFControls.Grid
WPFRowDefinition = WPFControls.RowDefinition
WPFColumnDefinition = WPFControls.ColumnDefinition
WPFBorder = WPFControls.Border

from System.Windows.Media import SolidColorBrush, BrushConverter
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Reflection import BindingFlags

import re
import codecs
import os

from pyrevit import revit, DB, forms, script

doc = revit.doc
app = doc.Application


# ============================================================================
# VERSION DETECTION
# ============================================================================

def _get_revit_version():
    """Get Revit version as integer (2024, 2025, 2026, etc.)"""
    try:
        return int(app.VersionNumber)
    except:
        return 2024

REVIT_VERSION = _get_revit_version()


def _eid_int(element_id):
    """Get integer value from ElementId - compatible across versions.
    Revit 2024: .IntegerValue
    Revit 2025+: .Value (IntegerValue deprecated)
    """
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


# ============================================================================
# CONFIGURATION
# ============================================================================

class Config(object):
    """Color scheme and settings"""
    PRIMARY_COLOR = "#0F172A"
    SECONDARY_COLOR = "#E5B85C"
    BACKGROUND_COLOR = "#F8FAFC"
    BORDER_COLOR = "#CBD5E1"
    TEXT_DARK = "#5D4E37"
    TEXT_LIGHT = "#888888"
    SUCCESS_COLOR = "#10B981"
    WARNING_COLOR = "#F59E0B"
    ERROR_COLOR = "#FF6B6B"
    ROW_ALT_COLOR = "#F1F5F9"
    WHITE = "#FFFFFF"
    BLUE_COLOR = "#5DADE2"
    
    _converter = BrushConverter()
    
    @staticmethod
    def hex_to_brush(hex_color):
        """Convert hex string to SolidColorBrush - safe for IronPython.
        Color.FromArgb/FromRgb is unreliable in IronPython; BrushConverter is safe."""
        try:
            return Config._converter.ConvertFromString(hex_color)
        except:
            return SolidColorBrush()


# ============================================================================
# PARAMETER GROUP HELPERS (Version-compatible)
# ============================================================================

def _get_group_name_from_definition(definition):
    """Get parameter group display name - compatible across Revit versions.
    
    Revit 2024: definition.ParameterGroup -> BuiltInParameterGroup enum
    Revit 2025: definition.ParameterGroup still works (deprecated)
    Revit 2026: definition.ParameterGroup REMOVED -> use GetGroupTypeId()
    """
    # Try new API first (Revit 2025+/2026)
    try:
        if hasattr(definition, 'GetGroupTypeId'):
            group_type_id = definition.GetGroupTypeId()
            if group_type_id and hasattr(group_type_id, 'TypeId'):
                type_id_str = group_type_id.TypeId
                if type_id_str:
                    # Extract readable name from TypeId string
                    # e.g. "autodesk.parameter.group.general-2.0.0" -> "General"
                    parts = type_id_str.split(':')
                    if len(parts) > 0:
                        last_part = parts[-1] if len(parts) > 1 else parts[0]
                        # Try LabelUtils for nice display name
                        try:
                            label = DB.LabelUtils.GetLabelForGroup(group_type_id)
                            if label:
                                return label, group_type_id
                        except:
                            pass
                        # Fallback: parse TypeId
                        name = last_part.split('-')[0]
                        name = name.replace('autodesk.parameter.group.', '')
                        name = name.replace('.', ' ').title()
                        return name, group_type_id
    except:
        pass
    
    # Fallback to old API (Revit 2024/2025)
    try:
        param_group_enum = definition.ParameterGroup
        group_name = str(param_group_enum).replace('PG_', '').replace('_', ' ').title()
        return group_name, param_group_enum
    except:
        pass
    
    return "Unknown", None


def get_all_parameter_groups():
    """Get all available parameter groups - compatible across versions.
    
    Returns list of (group_id, group_display_name) tuples.
    group_id is BuiltInParameterGroup enum (2024) or ForgeTypeId (2026).
    """
    groups = []
    
    # Try new ForgeTypeId-based groups (Revit 2026+)
    if REVIT_VERSION >= 2026:
        try:
            # Use GroupTypeId static properties
            group_type_id_type = clr.GetClrType(DB.GroupTypeId) if hasattr(DB, 'GroupTypeId') else None
            if group_type_id_type:
                props = group_type_id_type.GetProperties(
                    BindingFlags.Public | BindingFlags.Static
                )
                for prop in props:
                    try:
                        forge_id = prop.GetValue(None)
                        if forge_id:
                            try:
                                label = DB.LabelUtils.GetLabelForGroup(forge_id)
                            except:
                                label = prop.Name.replace('_', ' ').title()
                            groups.append((forge_id, label))
                    except:
                        continue
            
            if groups:
                return sorted(groups, key=lambda x: x[1])
        except:
            pass
    
    # Fallback: BuiltInParameterGroup enum (Revit 2024/2025)
    try:
        bipg_type = getattr(DB, 'BuiltInParameterGroup', None)
        if bipg_type:
            for group in DB.BuiltInParameterGroup.GetValues(bipg_type):
                if group != DB.BuiltInParameterGroup.INVALID:
                    group_name = str(group).replace('PG_', '').replace('_', ' ').title()
                    groups.append((group, group_name))
    except:
        pass
    
    return sorted(groups, key=lambda x: x[1])


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_shared_parameter_guids():
    """Get all shared parameter GUIDs from document"""
    shared_guids = {}
    
    try:
        collector = DB.FilteredElementCollector(doc)
        shared_param_elements = collector.OfClass(DB.SharedParameterElement)
        
        for elem in shared_param_elements:
            guid = elem.GuidValue
            name = elem.GetDefinition().Name
            if guid:
                shared_guids[str(guid)] = name
        
        try:
            def_file = app.OpenSharedParameterFile()
            if def_file:
                for group in def_file.Groups:
                    for definition in group.Definitions:
                        if hasattr(definition, 'GUID'):
                            guid_str = str(definition.GUID)
                            if guid_str not in shared_guids:
                                shared_guids[guid_str] = definition.Name
        except:
            pass
            
    except:
        pass
    
    return shared_guids


def get_parameter_data_type(definition):
    """Get readable parameter data type name.
    
    Revit 2024: definition.GetDataType() works; fallback to ParameterType
    Revit 2025: same as 2024
    Revit 2026: ParameterType REMOVED -> GetDataType() only
    """
    try:
        if hasattr(definition, 'GetDataType'):
            spec_type_id = definition.GetDataType()
            
            if hasattr(spec_type_id, 'TypeId'):
                type_id_str = spec_type_id.TypeId
            else:
                type_id_str = str(spec_type_id)
            
            if type_id_str:
                type_id_lower = type_id_str.lower()
                
                # Try LabelUtils first (best display name)
                try:
                    label = DB.LabelUtils.GetLabelForSpec(spec_type_id)
                    if label:
                        return label
                except:
                    pass
                
                if 'autodesk.spec' in type_id_lower:
                    parts = type_id_str.split(':')
                    if len(parts) > 1:
                        type_part = parts[-1]
                        type_name = type_part.split('-')[0]
                        type_name = type_name.replace('spec.', '')
                        
                        type_map = {
                            'string': 'Text',
                            'int': 'Integer',
                            'integer': 'Integer',
                            'double': 'Number',
                            'number': 'Number',
                            'length': 'Length',
                            'area': 'Area',
                            'volume': 'Volume',
                            'angle': 'Angle',
                            'url': 'URL',
                            'material': 'Material',
                            'yesno': 'Yes/No',
                            'bool': 'Yes/No',
                            'boolean': 'Yes/No',
                            'multilinetext': 'Multiline Text',
                            'familytype': 'Family Type',
                            'image': 'Image',
                        }
                        
                        type_name_lower = type_name.lower()
                        for key, value in type_map.items():
                            if key in type_name_lower:
                                return value
                        
                        return type_name.replace('.', ' ').title()
        
        # Fallback: ParameterType (Revit 2024/2025 only - removed in 2026)
        if REVIT_VERSION < 2026:
            try:
                if hasattr(definition, 'ParameterType'):
                    param_type = definition.ParameterType
                    return str(param_type)
            except:
                pass
        
        return "Unknown"
        
    except:
        return "Unknown"


def is_shared_parameter(param_name):
    """Check if parameter is shared"""
    shared_guids_dict = get_shared_parameter_guids()
    shared_param_names = set(shared_guids_dict.values())
    return param_name in shared_param_names


def get_all_categories():
    """Get all available categories in document"""
    categories = []
    
    try:
        for cat in doc.Settings.Categories:
            if cat.AllowsBoundParameters and cat.CategoryType == DB.CategoryType.Model:
                categories.append((cat, cat.Name))
    except:
        pass
    
    return sorted(categories, key=lambda x: x[1])


def delete_parameter(param_name):
    """Delete a parameter"""
    try:
        t = DB.Transaction(doc, "DQT - Delete Parameter")
        t.Start()
        
        try:
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            
            found = False
            while iterator.MoveNext():
                definition = iterator.Key
                if definition.Name == param_name:
                    param_bindings.Remove(definition)
                    found = True
                    break
            
            if found:
                t.Commit()
                return True, "Deleted"
            else:
                t.RollBack()
                return False, "Not found"
                
        except Exception as e:
            t.RollBack()
            return False, str(e)
            
    except Exception as e:
        return False, str(e)


def update_parameter_group(param_name, new_group_id):
    """Update parameter group using version-compatible approach.
    
    new_group_id: BuiltInParameterGroup (2024) or ForgeTypeId (2026)
    """
    try:
        t = DB.Transaction(doc, "DQT - Update Parameter Group")
        t.Start()
        
        try:
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            
            definition = None
            
            while iterator.MoveNext():
                defn = iterator.Key
                if defn.Name == param_name:
                    definition = defn
                    break
            
            if not definition:
                t.RollBack()
                return False, "Parameter not found"
            
            # Method 1: SetGroupTypeId (Revit 2025+/2026)
            if hasattr(definition, 'SetGroupTypeId'):
                try:
                    # If new_group_id is already a ForgeTypeId, use directly
                    if hasattr(new_group_id, 'TypeId'):
                        definition.SetGroupTypeId(new_group_id)
                        t.Commit()
                        return True, "Updated"
                    else:
                        # It's a BuiltInParameterGroup enum - need to convert
                        # Try getting ForgeTypeId from BuiltInParameterGroup
                        try:
                            if hasattr(DB, 'ParameterUtils') and hasattr(DB.ParameterUtils, 'GetParameterGroupTypeId'):
                                forge_id = DB.ParameterUtils.GetParameterGroupTypeId(new_group_id)
                                definition.SetGroupTypeId(forge_id)
                                t.Commit()
                                return True, "Updated"
                        except:
                            pass
                except:
                    pass
            
            # Method 2: Direct ParameterGroup property (Revit 2024/2025)
            if REVIT_VERSION < 2026:
                try:
                    if hasattr(definition, 'ParameterGroup'):
                        # ParameterGroup property is read-only on InternalDefinition
                        # Try Reflection approach
                        pass
                except:
                    pass
            
            # Method 3: Reflection (last resort)
            definition_type = definition.GetType()
            
            field_names = [
                'm_builtInParameterGroup',
                'm_parameterGroup', 
                'parameterGroup',
                'm_group',
                'group',
                '_builtInParameterGroup',
                '_parameterGroup'
            ]
            
            for field_name in field_names:
                try:
                    field = definition_type.GetField(
                        field_name,
                        BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public
                    )
                    
                    if field:
                        field.SetValue(definition, new_group_id)
                        t.Commit()
                        return True, "Updated"
                except:
                    continue
            
            # Try property via Reflection
            try:
                prop = definition_type.GetProperty(
                    "ParameterGroup",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic
                )
                
                if prop and prop.CanWrite:
                    prop.SetValue(definition, new_group_id, None)
                    t.Commit()
                    return True, "Updated"
            except:
                pass
            
            t.RollBack()
            return False, "Cannot modify group (API limitation)"
                
        except Exception as e:
            t.RollBack()
            return False, str(e)
            
    except Exception as e:
        return False, str(e)


def update_parameter_categories(param_name, new_categories):
    """Update parameter categories"""
    
    if doc.IsModifiable:
        return False, "Document already in transaction"
    
    try:
        is_shared = is_shared_parameter(param_name)
        
        t = DB.Transaction(doc, "DQT - Update Categories")
        t.Start()
        
        try:
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            
            found = False
            old_definition = None
            old_binding = None
            
            while iterator.MoveNext():
                definition = iterator.Key
                binding = iterator.Current
                
                if definition.Name == param_name:
                    old_definition = definition
                    old_binding = binding
                    found = True
                    break
            
            if not found:
                t.RollBack()
                return False, "Parameter not found"
            
            new_cat_set = app.Create.NewCategorySet()
            for cat in new_categories:
                new_cat_set.Insert(cat)
            
            is_instance = isinstance(old_binding, DB.InstanceBinding)
            
            if is_instance:
                new_binding = app.Create.NewInstanceBinding(new_cat_set)
            else:
                new_binding = app.Create.NewTypeBinding(new_cat_set)
            
            success = param_bindings.ReInsert(old_definition, new_binding)
            
            if success:
                t.Commit()
                return True, "Categories updated"
            else:
                t.RollBack()
                
                if not is_shared:
                    return False, "Project parameters: Use Revit UI to change categories"
                else:
                    return False, "Failed to update categories"
                
        except Exception as e:
            t.RollBack()
            return False, str(e)
            
    except Exception as e:
        return False, str(e)


def update_parameter_binding(param_name, new_binding_type):
    """Update parameter binding (Instance/Type)"""
    
    if doc.IsModifiable:
        return False, "Document already in transaction"
    
    try:
        is_shared = is_shared_parameter(param_name)
        
        if not is_shared:
            return False, "Project parameters cannot change binding via API"
        
        t = DB.Transaction(doc, "DQT - Update Binding")
        t.Start()
        
        try:
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            
            found = False
            old_definition = None
            old_binding = None
            
            while iterator.MoveNext():
                definition = iterator.Key
                binding = iterator.Current
                
                if definition.Name == param_name:
                    old_definition = definition
                    old_binding = binding
                    found = True
                    break
            
            if not found:
                t.RollBack()
                return False, "Parameter not found"
            
            categories = old_binding.Categories
            is_instance = isinstance(old_binding, DB.InstanceBinding)
            wants_instance = (new_binding_type == "Instance")
            
            if is_instance == wants_instance:
                t.RollBack()
                return True, "Already {}".format(new_binding_type)
            
            if wants_instance:
                new_binding = app.Create.NewInstanceBinding(categories)
            else:
                new_binding = app.Create.NewTypeBinding(categories)
            
            success = param_bindings.ReInsert(old_definition, new_binding)
            
            if success:
                t.Commit()
                return True, "Updated"
            else:
                t.RollBack()
                return False, "ReInsert failed"
                
        except Exception as e:
            t.RollBack()
            return False, str(e)
            
    except Exception as e:
        return False, str(e)


# ============================================================================
# PARAMETER ITEM (Data Model)
# ============================================================================

class ParameterItem(INotifyPropertyChanged):
    """Wrapper class for Parameter with WPF binding support"""
    
    def __init__(self, name, param_type, data_type, group, group_id, binding, categories, category_objects=None):
        self._property_changed_handlers = []
        self._is_selected = False
        
        self._name = name
        self._param_type = param_type  # Shared or Project
        self._data_type = data_type
        self._group = group
        self._group_id = group_id  # BuiltInParameterGroup (2024) or ForgeTypeId (2026)
        self._binding = binding  # Instance or Type
        self._categories = categories
        self._category_objects = category_objects or []
        self._is_shared = param_type == "Shared"
    
    # Properties
    @property
    def name(self):
        return self._name
    
    @property
    def is_selected(self):
        return self._is_selected
    
    @is_selected.setter
    def is_selected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("is_selected")
    
    @property
    def param_type(self):
        return self._param_type
    
    @property
    def data_type(self):
        return self._data_type
    
    @property
    def group(self):
        return self._group
    
    @group.setter
    def group(self, value):
        self._group = value
        self.OnPropertyChanged("group")
    
    @property
    def group_id(self):
        return self._group_id
    
    @group_id.setter
    def group_id(self, value):
        self._group_id = value
    
    @property
    def binding(self):
        return self._binding
    
    @binding.setter
    def binding(self, value):
        self._binding = value
        self.OnPropertyChanged("binding")
    
    @property
    def categories(self):
        return self._categories
    
    @categories.setter
    def categories(self, value):
        self._categories = value
        self.OnPropertyChanged("categories")
    
    @property
    def category_objects(self):
        return self._category_objects
    
    @category_objects.setter
    def category_objects(self, value):
        self._category_objects = value
    
    @property
    def category_count(self):
        if self._categories == "N/A":
            return 0
        return len(self._categories.split(','))
    
    @property
    def is_shared(self):
        return self._is_shared
    
    @property
    def selected_display(self):
        """Display text for selection column"""
        return u"\u2611" if self._is_selected else u"\u2610"
    
    # INotifyPropertyChanged
    def add_PropertyChanged(self, handler):
        self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def OnPropertyChanged(self, prop_name):
        for handler in self._property_changed_handlers:
            try:
                handler(self, PropertyChangedEventArgs(prop_name))
            except:
                pass


# ============================================================================
# EDIT GROUP DIALOG
# ============================================================================

class EditGroupDialog(Window):
    """Dialog to edit parameter group"""
    
    def __init__(self, item, all_groups):
        self.item = item
        self.all_groups = all_groups
        self.result = False
        self.selected_group = None
        self.selected_group_id = None
        
        self._build_ui()
    
    def _build_ui(self):
        self.Title = "Edit Parameter Group"
        self.Width = 450
        self.Height = 500
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = Config.hex_to_brush(Config.BACKGROUND_COLOR)
        
        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)
        
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        
        # Header
        header = WPFBorder()
        header.Background = Config.hex_to_brush(Config.PRIMARY_COLOR)
        header.CornerRadius = System.Windows.CornerRadius(4)
        header.Padding = Thickness(10, 8, 10, 8)
        header.Margin = Thickness(0, 0, 0, 10)
        
        header_text = TextBlock()
        header_text.Text = "Edit Group"
        header_text.FontSize = 16
        header_text.FontWeight = FontWeights.Bold
        
        header.Child = header_text
        WPFGrid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        # Parameter info
        info_panel = StackPanel()
        info_panel.Margin = Thickness(0, 0, 0, 10)
        
        info1 = TextBlock()
        info1.Text = "Parameter: {}".format(self.item.name)
        info1.FontWeight = FontWeights.SemiBold
        info1.Margin = Thickness(0, 0, 0, 5)
        info_panel.Children.Add(info1)
        
        info2 = TextBlock()
        info2.Text = "Current Group: {}".format(self.item.group)
        info2.Foreground = Config.hex_to_brush(Config.TEXT_LIGHT)
        info_panel.Children.Add(info2)
        
        WPFGrid.SetRow(info_panel, 1)
        main_grid.Children.Add(info_panel)
        
        # Group list
        list_panel = StackPanel()
        
        label = TextBlock()
        label.Text = "Select New Group:"
        label.FontWeight = FontWeights.SemiBold
        label.Margin = Thickness(0, 0, 0, 5)
        list_panel.Children.Add(label)
        
        self.group_listbox = ListBox()
        self.group_listbox.Height = 280
        self.group_listbox.Background = Config.hex_to_brush(Config.WHITE)
        
        for group_id, group_name in self.all_groups:
            lb_item = ListBoxItem()
            lb_item.Content = group_name
            lb_item.Tag = group_id
            
            if group_name == self.item.group:
                lb_item.IsSelected = True
            
            self.group_listbox.Items.Add(lb_item)
        
        list_panel.Children.Add(self.group_listbox)
        
        WPFGrid.SetRow(list_panel, 2)
        main_grid.Children.Add(list_panel)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 15, 0, 0)
        
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 100
        btn_apply.Height = 35
        btn_apply.Margin = Thickness(0, 0, 10, 0)
        btn_apply.Background = Config.hex_to_brush(Config.SUCCESS_COLOR)
        btn_apply.Foreground = Config.hex_to_brush(Config.WHITE)
        btn_apply.Click += self._on_apply
        btn_panel.Children.Add(btn_apply)
        
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 100
        btn_cancel.Height = 35
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_cancel)
        
        WPFGrid.SetRow(btn_panel, 3)
        main_grid.Children.Add(btn_panel)
        
        self.Content = main_grid
    
    def _on_apply(self, sender, args):
        if not self.group_listbox.SelectedItem:
            MessageBox.Show("Please select a group!", "Warning",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        selected_item = self.group_listbox.SelectedItem
        self.selected_group = selected_item.Content
        self.selected_group_id = selected_item.Tag
        
        if self.selected_group == self.item.group:
            MessageBox.Show("No change - same group selected.", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Apply change
        success, msg = update_parameter_group(self.item.name, self.selected_group_id)
        
        if success:
            self.item.group = self.selected_group
            self.item.group_id = self.selected_group_id
            self.result = True
            MessageBox.Show("Group updated successfully!", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            self.Close()
        else:
            MessageBox.Show("Failed to update group:\n\n{}".format(msg), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# EDIT CATEGORIES DIALOG
# ============================================================================

class EditCategoriesDialog(Window):
    """Dialog to edit parameter categories"""
    
    def __init__(self, item, all_categories):
        self.item = item
        self.all_categories = all_categories
        self.result = False
        
        self._build_ui()
    
    def _build_ui(self):
        self.Title = "Edit Parameter Categories"
        self.Width = 450
        self.Height = 550
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = Config.hex_to_brush(Config.BACKGROUND_COLOR)
        
        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)
        
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        
        # Header
        header = WPFBorder()
        header.Background = Config.hex_to_brush(Config.BLUE_COLOR)
        header.CornerRadius = System.Windows.CornerRadius(4)
        header.Padding = Thickness(10, 8, 10, 8)
        header.Margin = Thickness(0, 0, 0, 10)
        
        header_text = TextBlock()
        header_text.Text = "Edit Categories"
        header_text.FontSize = 16
        header_text.FontWeight = FontWeights.Bold
        header_text.Foreground = Config.hex_to_brush(Config.WHITE)
        
        header.Child = header_text
        WPFGrid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        # Parameter info
        info_panel = StackPanel()
        info_panel.Margin = Thickness(0, 0, 0, 10)
        
        info1 = TextBlock()
        info1.Text = "Parameter: {}".format(self.item.name)
        info1.FontWeight = FontWeights.SemiBold
        info1.Margin = Thickness(0, 0, 0, 5)
        info_panel.Children.Add(info1)
        
        info2 = TextBlock()
        info2.Text = "Type: {} | Binding: {}".format(self.item.param_type, self.item.binding)
        info2.Foreground = Config.hex_to_brush(Config.TEXT_LIGHT)
        info_panel.Children.Add(info2)
        
        WPFGrid.SetRow(info_panel, 1)
        main_grid.Children.Add(info_panel)
        
        # Quick select buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.Margin = Thickness(0, 0, 0, 10)
        
        btn_all = Button()
        btn_all.Content = "Select All"
        btn_all.Padding = Thickness(10, 4, 10, 4)
        btn_all.Margin = Thickness(0, 0, 5, 0)
        btn_all.Click += self._on_select_all
        btn_panel.Children.Add(btn_all)
        
        btn_none = Button()
        btn_none.Content = "Select None"
        btn_none.Padding = Thickness(10, 4, 10, 4)
        btn_none.Click += self._on_select_none
        btn_panel.Children.Add(btn_none)
        
        WPFGrid.SetRow(btn_panel, 2)
        main_grid.Children.Add(btn_panel)
        
        # Category list
        list_panel = StackPanel()
        
        label = TextBlock()
        label.Text = "Select Categories:"
        label.FontWeight = FontWeights.SemiBold
        label.Margin = Thickness(0, 0, 0, 5)
        list_panel.Children.Add(label)
        
        # Get current category names
        current_cat_names = []
        if self.item.categories != "N/A":
            current_cat_names = [c.strip() for c in self.item.categories.split(',')]
        
        self.category_listbox = ListBox()
        self.category_listbox.Height = 280
        self.category_listbox.Background = Config.hex_to_brush(Config.WHITE)
        self.category_listbox.SelectionMode = SelectionMode.Multiple
        
        for cat_obj, cat_name in self.all_categories:
            lb_item = ListBoxItem()
            lb_item.Content = cat_name
            lb_item.Tag = cat_obj
            
            if cat_name in current_cat_names:
                lb_item.IsSelected = True
            
            self.category_listbox.Items.Add(lb_item)
        
        list_panel.Children.Add(self.category_listbox)
        
        WPFGrid.SetRow(list_panel, 3)
        main_grid.Children.Add(list_panel)
        
        # Buttons
        btn_panel2 = StackPanel()
        btn_panel2.Orientation = Orientation.Horizontal
        btn_panel2.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel2.Margin = Thickness(0, 15, 0, 0)
        
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 100
        btn_apply.Height = 35
        btn_apply.Margin = Thickness(0, 0, 10, 0)
        btn_apply.Background = Config.hex_to_brush(Config.SUCCESS_COLOR)
        btn_apply.Foreground = Config.hex_to_brush(Config.WHITE)
        btn_apply.Click += self._on_apply
        btn_panel2.Children.Add(btn_apply)
        
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 100
        btn_cancel.Height = 35
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel2.Children.Add(btn_cancel)
        
        WPFGrid.SetRow(btn_panel2, 4)
        main_grid.Children.Add(btn_panel2)
        
        self.Content = main_grid
    
    def _on_select_all(self, sender, args):
        for lb_item in self.category_listbox.Items:
            lb_item.IsSelected = True
    
    def _on_select_none(self, sender, args):
        for lb_item in self.category_listbox.Items:
            lb_item.IsSelected = False
    
    def _on_apply(self, sender, args):
        selected_items = [lb_item for lb_item in self.category_listbox.Items if lb_item.IsSelected]
        
        if not selected_items:
            MessageBox.Show("Please select at least one category!", "Warning",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        selected_cats = [lb_item.Tag for lb_item in selected_items]
        selected_names = [lb_item.Content for lb_item in selected_items]
        
        # Apply change
        success, msg = update_parameter_categories(self.item.name, selected_cats)
        
        if success:
            self.item.categories = ', '.join(selected_names)
            self.item.category_objects = selected_cats
            self.result = True
            MessageBox.Show("Categories updated successfully!", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            self.Close()
        else:
            MessageBox.Show("Failed to update categories:\n\n{}".format(msg), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# EDIT BINDING DIALOG
# ============================================================================

class EditBindingDialog(Window):
    """Dialog to edit parameter binding (Instance/Type)"""
    
    def __init__(self, item):
        self.item = item
        self.result = False
        
        self._build_ui()
    
    def _build_ui(self):
        self.Title = "Edit Parameter Binding"
        self.Width = 400
        self.Height = 280
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = Config.hex_to_brush(Config.BACKGROUND_COLOR)
        
        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)
        
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        
        # Header
        header = WPFBorder()
        header.Background = Config.hex_to_brush(Config.WARNING_COLOR)
        header.CornerRadius = System.Windows.CornerRadius(4)
        header.Padding = Thickness(10, 8, 10, 8)
        header.Margin = Thickness(0, 0, 0, 10)
        
        header_text = TextBlock()
        header_text.Text = "Edit Binding"
        header_text.FontSize = 16
        header_text.FontWeight = FontWeights.Bold
        header_text.Foreground = Config.hex_to_brush(Config.WHITE)
        
        header.Child = header_text
        WPFGrid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        # Parameter info
        info_panel = StackPanel()
        info_panel.Margin = Thickness(0, 0, 0, 15)
        
        info1 = TextBlock()
        info1.Text = "Parameter: {}".format(self.item.name)
        info1.FontWeight = FontWeights.SemiBold
        info1.Margin = Thickness(0, 0, 0, 5)
        info_panel.Children.Add(info1)
        
        info2 = TextBlock()
        info2.Text = "Type: {} | Current Binding: {}".format(self.item.param_type, self.item.binding)
        info2.Foreground = Config.hex_to_brush(Config.TEXT_LIGHT)
        info_panel.Children.Add(info2)
        
        if not self.item.is_shared:
            warning = TextBlock()
            warning.Text = u"\u26A0 Project parameters cannot change binding via API"
            warning.Foreground = Config.hex_to_brush(Config.ERROR_COLOR)
            warning.Margin = Thickness(0, 10, 0, 0)
            warning.TextWrapping = System.Windows.TextWrapping.Wrap
            info_panel.Children.Add(warning)
        
        WPFGrid.SetRow(info_panel, 1)
        main_grid.Children.Add(info_panel)
        
        # Binding selection
        select_panel = StackPanel()
        select_panel.Margin = Thickness(0, 0, 0, 15)
        
        label = TextBlock()
        label.Text = "Select Binding:"
        label.FontWeight = FontWeights.SemiBold
        label.Margin = Thickness(0, 0, 0, 5)
        select_panel.Children.Add(label)
        
        self.binding_combo = ComboBox()
        self.binding_combo.Padding = Thickness(10, 5, 10, 5)
        self.binding_combo.IsEnabled = self.item.is_shared
        
        for opt in ["Instance", "Type"]:
            cb_item = ComboBoxItem()
            cb_item.Content = opt
            self.binding_combo.Items.Add(cb_item)
        
        self.binding_combo.SelectedIndex = 0 if self.item.binding == "Instance" else 1
        
        select_panel.Children.Add(self.binding_combo)
        
        WPFGrid.SetRow(select_panel, 2)
        main_grid.Children.Add(select_panel)
        
        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 100
        btn_apply.Height = 35
        btn_apply.Margin = Thickness(0, 0, 10, 0)
        btn_apply.Background = Config.hex_to_brush(Config.SUCCESS_COLOR)
        btn_apply.Foreground = Config.hex_to_brush(Config.WHITE)
        btn_apply.IsEnabled = self.item.is_shared
        btn_apply.Click += self._on_apply
        btn_panel.Children.Add(btn_apply)
        
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 100
        btn_cancel.Height = 35
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_cancel)
        
        WPFGrid.SetRow(btn_panel, 3)
        main_grid.Children.Add(btn_panel)
        
        self.Content = main_grid
    
    def _on_apply(self, sender, args):
        new_binding = self.binding_combo.SelectedItem.Content
        
        if new_binding == self.item.binding:
            MessageBox.Show("No change - same binding selected.", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Apply change
        success, msg = update_parameter_binding(self.item.name, new_binding)
        
        if success:
            self.item.binding = new_binding
            self.result = True
            MessageBox.Show("Binding updated successfully!", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            self.Close()
        else:
            MessageBox.Show("Failed to update binding:\n\n{}".format(msg), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# MAIN WINDOW
# ============================================================================

class ParameterManagerWindow(Window):
    """Parameter Manager with Sheet Manager style UI"""
    
    def __init__(self):
        self.all_items = []
        self.filtered_items = []
        self.all_groups = get_all_parameter_groups()
        self.all_categories = get_all_categories()
        
        self.txt_total = None
        self.txt_selected = None
        self.txt_shared = None
        self.txt_project = None
        self.search_box = None
        self.filter_type_combo = None
        self.filter_binding_combo = None
        self.filter_datatype_combo = None
        self.data_grid = None
        
        self._build_ui()
        self._load_data()
    
    def _build_ui(self):
        self.Title = "Parameter Manager v1.1 (Revit {})".format(REVIT_VERSION)
        self.Width = 1250
        self.Height = 700
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = Config.hex_to_brush(Config.BACKGROUND_COLOR)
        
        main_grid = WPFGrid()
        main_grid.Margin = Thickness(12)
        
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        
        header = self._create_header()
        WPFGrid.SetRow(header, 0)
        main_grid.Children.Add(header)
        
        cards = self._create_summary_cards()
        WPFGrid.SetRow(cards, 1)
        main_grid.Children.Add(cards)
        
        content = self._create_main_content()
        WPFGrid.SetRow(content, 2)
        main_grid.Children.Add(content)
        
        buttons = self._create_action_buttons()
        WPFGrid.SetRow(buttons, 3)
        main_grid.Children.Add(buttons)
        
        tips = self._create_tips()
        WPFGrid.SetRow(tips, 4)
        main_grid.Children.Add(tips)
        
        footer = self._create_footer()
        WPFGrid.SetRow(footer, 5)
        main_grid.Children.Add(footer)
        
        self.Content = main_grid
    
    def _create_header(self):
        border = WPFBorder()
        border.Background = Config.hex_to_brush(Config.PRIMARY_COLOR)
        border.CornerRadius = System.Windows.CornerRadius(5)
        border.Padding = Thickness(12, 8, 12, 8)
        border.Margin = Thickness(0, 0, 0, 10)
        
        panel = StackPanel()
        
        title = TextBlock()
        title.Text = "Parameter Manager v1.1"
        title.FontSize = 17
        title.FontWeight = FontWeights.Bold
        panel.Children.Add(title)
        
        subtitle = TextBlock()
        subtitle.Text = "by Dang Quoc Truong (DQT) | Revit {}".format(REVIT_VERSION)
        subtitle.FontSize = 10
        subtitle.Foreground = Config.hex_to_brush(Config.TEXT_DARK)
        subtitle.Margin = Thickness(0, 2, 0, 0)
        panel.Children.Add(subtitle)
        
        border.Child = panel
        return border
    
    def _create_summary_cards(self):
        grid = WPFGrid()
        grid.Margin = Thickness(0, 0, 0, 10)
        
        for i in range(4):
            grid.ColumnDefinitions.Add(WPFColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        card1, self.txt_total = self._create_card("TOTAL", "0", 0)
        WPFGrid.SetColumn(card1, 0)
        grid.Children.Add(card1)
        
        card2, self.txt_selected = self._create_card("SELECTED", "0", 1, Config.SECONDARY_COLOR)
        WPFGrid.SetColumn(card2, 1)
        grid.Children.Add(card2)
        
        card3, self.txt_shared = self._create_card("SHARED", "0", 2, Config.SUCCESS_COLOR)
        WPFGrid.SetColumn(card3, 2)
        grid.Children.Add(card3)
        
        card4, self.txt_project = self._create_card("PROJECT", "0", 3, Config.WARNING_COLOR)
        WPFGrid.SetColumn(card4, 3)
        grid.Children.Add(card4)
        
        return grid
    
    def _create_card(self, label, value, margin_index, value_color=None):
        border = WPFBorder()
        border.Background = Config.hex_to_brush(Config.WHITE)
        border.BorderBrush = Config.hex_to_brush(Config.BORDER_COLOR)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(10, 6, 10, 6)
        
        if margin_index == 0:
            border.Margin = Thickness(0, 0, 4, 0)
        elif margin_index == 3:
            border.Margin = Thickness(4, 0, 0, 0)
        else:
            border.Margin = Thickness(4, 0, 4, 0)
        
        panel = StackPanel()
        
        label_text = TextBlock()
        label_text.Text = label
        label_text.FontSize = 9
        label_text.Foreground = Config.hex_to_brush(Config.TEXT_LIGHT)
        panel.Children.Add(label_text)
        
        value_text = TextBlock()
        value_text.Text = value
        value_text.FontSize = 22
        value_text.FontWeight = FontWeights.Bold
        
        if value_color:
            value_text.Foreground = Config.hex_to_brush(value_color)
        
        panel.Children.Add(value_text)
        border.Child = panel
        
        return border, value_text
    
    def _create_main_content(self):
        grid = WPFGrid()
        
        grid.ColumnDefinitions.Add(WPFColumnDefinition(Width=GridLength(200)))
        grid.ColumnDefinitions.Add(WPFColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        
        left_panel = self._create_left_panel()
        WPFGrid.SetColumn(left_panel, 0)
        grid.Children.Add(left_panel)
        
        self.data_grid = self._create_datagrid()
        WPFGrid.SetColumn(self.data_grid, 1)
        grid.Children.Add(self.data_grid)
        
        return grid
    
    def _create_left_panel(self):
        border = WPFBorder()
        border.Background = Config.hex_to_brush(Config.WHITE)
        border.BorderBrush = Config.hex_to_brush(Config.BORDER_COLOR)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 0, 8, 0)
        
        panel = StackPanel()
        
        # Search
        label1 = TextBlock()
        label1.Text = "SEARCH"
        label1.FontSize = 9
        label1.FontWeight = FontWeights.SemiBold
        label1.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label1)
        
        self.search_box = TextBox()
        self.search_box.Padding = Thickness(6, 4, 6, 4)
        self.search_box.Margin = Thickness(0, 0, 0, 10)
        self.search_box.TextChanged += self._on_filter_changed
        panel.Children.Add(self.search_box)
        
        # Filter by Type (Shared/Project)
        label2 = TextBlock()
        label2.Text = "FILTER BY TYPE"
        label2.FontSize = 9
        label2.FontWeight = FontWeights.SemiBold
        label2.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label2)
        
        self.filter_type_combo = ComboBox()
        self.filter_type_combo.Padding = Thickness(6, 4, 6, 4)
        self.filter_type_combo.Margin = Thickness(0, 0, 0, 10)
        
        for item_text in ["All Parameters", "Shared Only", "Project Only"]:
            cb_item = ComboBoxItem()
            cb_item.Content = item_text
            self.filter_type_combo.Items.Add(cb_item)
        
        self.filter_type_combo.SelectedIndex = 0
        self.filter_type_combo.SelectionChanged += self._on_filter_changed
        panel.Children.Add(self.filter_type_combo)
        
        # Filter by Binding
        label3 = TextBlock()
        label3.Text = "FILTER BY BINDING"
        label3.FontSize = 9
        label3.FontWeight = FontWeights.SemiBold
        label3.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label3)
        
        self.filter_binding_combo = ComboBox()
        self.filter_binding_combo.Padding = Thickness(6, 4, 6, 4)
        self.filter_binding_combo.Margin = Thickness(0, 0, 0, 10)
        
        for item_text in ["All Bindings", "Instance Only", "Type Only"]:
            cb_item = ComboBoxItem()
            cb_item.Content = item_text
            self.filter_binding_combo.Items.Add(cb_item)
        
        self.filter_binding_combo.SelectedIndex = 0
        self.filter_binding_combo.SelectionChanged += self._on_filter_changed
        panel.Children.Add(self.filter_binding_combo)
        
        # Filter by Data Type
        label4 = TextBlock()
        label4.Text = "FILTER BY DATA TYPE"
        label4.FontSize = 9
        label4.FontWeight = FontWeights.SemiBold
        label4.Margin = Thickness(0, 0, 0, 4)
        panel.Children.Add(label4)
        
        self.filter_datatype_combo = ComboBox()
        self.filter_datatype_combo.Padding = Thickness(6, 4, 6, 4)
        self.filter_datatype_combo.Margin = Thickness(0, 0, 0, 10)
        
        cb_item = ComboBoxItem()
        cb_item.Content = "All Data Types"
        self.filter_datatype_combo.Items.Add(cb_item)
        self.filter_datatype_combo.SelectedIndex = 0
        self.filter_datatype_combo.SelectionChanged += self._on_filter_changed
        panel.Children.Add(self.filter_datatype_combo)
        
        # Quick Select
        label5 = TextBlock()
        label5.Text = "QUICK SELECT"
        label5.FontSize = 9
        label5.FontWeight = FontWeights.SemiBold
        label5.Margin = Thickness(0, 10, 0, 4)
        panel.Children.Add(label5)
        
        btn_all = Button()
        btn_all.Content = "Select All"
        btn_all.Padding = Thickness(8, 4, 8, 4)
        btn_all.Margin = Thickness(0, 0, 0, 4)
        btn_all.Background = Config.hex_to_brush(Config.WHITE)
        btn_all.Click += self._on_select_all
        panel.Children.Add(btn_all)
        
        btn_clear = Button()
        btn_clear.Content = "Clear All"
        btn_clear.Padding = Thickness(8, 4, 8, 4)
        btn_clear.Margin = Thickness(0, 0, 0, 4)
        btn_clear.Background = Config.hex_to_brush(Config.WHITE)
        btn_clear.Click += self._on_clear_all
        panel.Children.Add(btn_clear)
        
        btn_shared = Button()
        btn_shared.Content = "Select Shared"
        btn_shared.Padding = Thickness(8, 4, 8, 4)
        btn_shared.Margin = Thickness(0, 0, 0, 4)
        btn_shared.Background = Config.hex_to_brush(Config.WHITE)
        btn_shared.Click += self._on_select_shared
        panel.Children.Add(btn_shared)
        
        btn_project = Button()
        btn_project.Content = "Select Project"
        btn_project.Padding = Thickness(8, 4, 8, 4)
        btn_project.Background = Config.hex_to_brush(Config.WHITE)
        btn_project.Click += self._on_select_project
        panel.Children.Add(btn_project)
        
        border.Child = panel
        return border
    
    def _create_datagrid(self):
        """Create DataGrid - uses text columns only for IronPython compatibility.
        
        NOTE: DataGridCheckBoxColumn is unsafe in IronPython edit mode.
        Using DataGrid row selection (SelectionMode=Extended) instead of checkboxes.
        """
        grid = DataGrid()
        grid.AutoGenerateColumns = False
        grid.IsReadOnly = True
        grid.SelectionMode = WPFControls.DataGridSelectionMode.Extended
        grid.SelectionUnit = WPFControls.DataGridSelectionUnit.FullRow
        grid.CanUserSortColumns = True
        grid.Background = Config.hex_to_brush(Config.WHITE)
        grid.BorderBrush = Config.hex_to_brush(Config.BORDER_COLOR)
        grid.GridLinesVisibility = WPFControls.DataGridGridLinesVisibility.Horizontal
        grid.HorizontalGridLinesBrush = Config.hex_to_brush("#EEEEEE")
        grid.RowBackground = Config.hex_to_brush(Config.WHITE)
        grid.AlternatingRowBackground = Config.hex_to_brush(Config.ROW_ALT_COLOR)
        
        # Use text column for selection display instead of CheckBox column
        # (DataGridCheckBoxColumn is unsafe in IronPython)
        col_check = DataGridTextColumn()
        col_check.Header = u"Sel"
        col_check.Binding = Binding("selected_display")
        col_check.Width = DataGridLength(35)
        grid.Columns.Add(col_check)
        
        col_name = DataGridTextColumn()
        col_name.Header = "Parameter Name"
        col_name.Binding = Binding("name")
        col_name.Width = DataGridLength(200)
        grid.Columns.Add(col_name)
        
        col_type = DataGridTextColumn()
        col_type.Header = "Type"
        col_type.Binding = Binding("param_type")
        col_type.Width = DataGridLength(70)
        grid.Columns.Add(col_type)
        
        col_datatype = DataGridTextColumn()
        col_datatype.Header = "Data Type"
        col_datatype.Binding = Binding("data_type")
        col_datatype.Width = DataGridLength(90)
        grid.Columns.Add(col_datatype)
        
        col_group = DataGridTextColumn()
        col_group.Header = "Group"
        col_group.Binding = Binding("group")
        col_group.Width = DataGridLength(150)
        grid.Columns.Add(col_group)
        
        col_binding = DataGridTextColumn()
        col_binding.Header = "Binding"
        col_binding.Binding = Binding("binding")
        col_binding.Width = DataGridLength(70)
        grid.Columns.Add(col_binding)
        
        col_cats = DataGridTextColumn()
        col_cats.Header = "Categories"
        col_cats.Binding = Binding("categories")
        col_cats.Width = DataGridLength(300)
        grid.Columns.Add(col_cats)
        
        grid.SelectionChanged += self._on_selection_changed
        
        return grid
    
    def _create_action_buttons(self):
        border = WPFBorder()
        border.Background = Config.hex_to_brush(Config.WHITE)
        border.BorderBrush = Config.hex_to_brush(Config.BORDER_COLOR)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = System.Windows.CornerRadius(4)
        border.Padding = Thickness(8)
        border.Margin = Thickness(0, 10, 0, 0)
        
        grid = WPFGrid()
        
        center_panel = StackPanel()
        center_panel.Orientation = Orientation.Horizontal
        center_panel.HorizontalAlignment = HorizontalAlignment.Center
        
        primary = Config.hex_to_brush(Config.PRIMARY_COLOR)
        blue = Config.hex_to_brush(Config.BLUE_COLOR)
        orange = Config.hex_to_brush(Config.WARNING_COLOR)
        red = Config.hex_to_brush(Config.ERROR_COLOR)
        white = Config.hex_to_brush(Config.WHITE)
        
        btn_edit_group = self._create_button("Edit Group", primary)
        btn_edit_group.Click += self._on_edit_group
        center_panel.Children.Add(btn_edit_group)
        
        btn_edit_cats = self._create_button("Edit Categories", blue, white_fg=True)
        btn_edit_cats.Click += self._on_edit_categories
        center_panel.Children.Add(btn_edit_cats)
        
        btn_edit_binding = self._create_button("Edit Binding", orange, white_fg=True)
        btn_edit_binding.Click += self._on_edit_binding
        center_panel.Children.Add(btn_edit_binding)
        
        btn_export = self._create_button("Export CSV", primary)
        btn_export.Click += self._on_export
        center_panel.Children.Add(btn_export)
        
        btn_delete = self._create_button("Delete", red, white_fg=True)
        btn_delete.Click += self._on_delete
        center_panel.Children.Add(btn_delete)
        
        grid.Children.Add(center_panel)
        
        btn_refresh = Button()
        btn_refresh.Content = "Refresh"
        btn_refresh.Padding = Thickness(10, 5, 10, 5)
        btn_refresh.Margin = Thickness(2)
        btn_refresh.Background = white
        btn_refresh.HorizontalAlignment = HorizontalAlignment.Right
        btn_refresh.Click += self._on_refresh
        grid.Children.Add(btn_refresh)
        
        border.Child = grid
        return border
    
    def _create_button(self, text, bg_brush, white_fg=False):
        btn = Button()
        btn.Content = text
        btn.Padding = Thickness(10, 5, 10, 5)
        btn.Margin = Thickness(2)
        btn.Background = bg_brush
        
        if white_fg:
            btn.Foreground = Config.hex_to_brush(Config.WHITE)
        
        return btn
    
    def _create_tips(self):
        grid = WPFGrid()
        grid.Margin = Thickness(0, 8, 0, 0)
        
        tips = TextBlock()
        tips.Text = "Tip: Click rows to select (hold Ctrl/Shift for multi-select) | Edit Group/Categories/Binding for selected parameter"
        tips.FontSize = 10
        tips.Foreground = Config.hex_to_brush(Config.TEXT_LIGHT)
        grid.Children.Add(tips)
        
        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Padding = Thickness(15, 5, 15, 5)
        btn_close.Background = Config.hex_to_brush(Config.WHITE)
        btn_close.HorizontalAlignment = HorizontalAlignment.Right
        btn_close.Click += lambda s, e: self.Close()
        grid.Children.Add(btn_close)
        
        return grid
    
    def _create_footer(self):
        border = WPFBorder()
        border.Background = Config.hex_to_brush(Config.PRIMARY_COLOR)
        border.CornerRadius = System.Windows.CornerRadius(3)
        border.Padding = Thickness(8, 5, 8, 5)
        border.Margin = Thickness(0, 8, 0, 0)
        
        text = TextBlock()
        text.Text = u"Copyright by Dang Quoc Truong - DQT (c) 2026"
        text.FontSize = 10
        text.FontWeight = FontWeights.SemiBold
        text.HorizontalAlignment = HorizontalAlignment.Center
        text.Foreground = Config.hex_to_brush(Config.TEXT_DARK)
        
        border.Child = text
        return border
    
    # ========================================================================
    # DATA METHODS
    # ========================================================================
    
    def _load_data(self):
        self.all_items = []
        
        try:
            shared_guids_dict = get_shared_parameter_guids()
            shared_param_names = set(shared_guids_dict.values())
            
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            
            seen_names = {}
            data_types = set()
            
            while iterator.MoveNext():
                definition = iterator.Key
                binding = iterator.Current
                
                param_name = definition.Name
                
                is_shared = param_name in shared_param_names
                
                if param_name in seen_names:
                    if seen_names[param_name]:
                        param_type = "Project"
                    else:
                        param_type = "Shared" if is_shared else "Project"
                else:
                    seen_names[param_name] = is_shared
                    param_type = "Shared" if is_shared else "Project"
                
                data_type = get_parameter_data_type(definition)
                data_types.add(data_type)
                
                # Get parameter group (version-compatible)
                param_group, param_group_id = _get_group_name_from_definition(definition)
                
                if isinstance(binding, DB.InstanceBinding):
                    binding_type = "Instance"
                elif isinstance(binding, DB.TypeBinding):
                    binding_type = "Type"
                else:
                    binding_type = "Unknown"
                
                categories = []
                category_objects = []
                if hasattr(binding, 'Categories'):
                    for cat in binding.Categories:
                        categories.append(cat.Name)
                        category_objects.append(cat)
                
                item = ParameterItem(
                    param_name,
                    param_type,
                    data_type,
                    param_group,
                    param_group_id,
                    binding_type,
                    ', '.join(categories) if categories else 'N/A',
                    category_objects
                )
                self.all_items.append(item)
            
            self.all_items.sort(key=lambda x: x.name)
            
            # Populate data type filter
            self.filter_datatype_combo.Items.Clear()
            cb_item = ComboBoxItem()
            cb_item.Content = "All Data Types"
            self.filter_datatype_combo.Items.Add(cb_item)
            
            for dt in sorted(data_types):
                cb_item = ComboBoxItem()
                cb_item.Content = dt
                self.filter_datatype_combo.Items.Add(cb_item)
            
            self.filter_datatype_combo.SelectedIndex = 0
            
            self._apply_filters()
            self._update_stats()
            
        except Exception as ex:
            print("Error loading parameters: {}".format(str(ex)))
            MessageBox.Show("Failed to load parameters:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _apply_filters(self):
        search_text = ""
        if self.search_box and self.search_box.Text:
            search_text = self.search_box.Text.lower()
        
        type_filter = 0
        if self.filter_type_combo:
            type_filter = self.filter_type_combo.SelectedIndex
        
        binding_filter = 0
        if self.filter_binding_combo:
            binding_filter = self.filter_binding_combo.SelectedIndex
        
        datatype_filter = ""
        if self.filter_datatype_combo and self.filter_datatype_combo.SelectedIndex > 0:
            selected = self.filter_datatype_combo.SelectedItem
            if selected:
                datatype_filter = selected.Content
        
        self.filtered_items = []
        
        for item in self.all_items:
            if search_text and search_text not in item.name.lower():
                continue
            
            if type_filter == 1:  # Shared Only
                if not item.is_shared:
                    continue
            elif type_filter == 2:  # Project Only
                if item.is_shared:
                    continue
            
            if binding_filter == 1:  # Instance Only
                if item.binding != "Instance":
                    continue
            elif binding_filter == 2:  # Type Only
                if item.binding != "Type":
                    continue
            
            if datatype_filter and item.data_type != datatype_filter:
                continue
            
            self.filtered_items.append(item)
        
        self.data_grid.ItemsSource = ObservableCollection[object](self.filtered_items)
    
    def _update_stats(self):
        if self.txt_total:
            self.txt_total.Text = str(len(self.all_items))
        
        if self.txt_selected:
            selected = sum(1 for item in self.all_items if item.is_selected)
            self.txt_selected.Text = str(selected)
        
        if self.txt_shared:
            shared = sum(1 for item in self.all_items if item.is_shared)
            self.txt_shared.Text = str(shared)
        
        if self.txt_project:
            project = sum(1 for item in self.all_items if not item.is_shared)
            self.txt_project.Text = str(project)
    
    def _get_selected_items(self):
        """Get selected items from DataGrid row selection"""
        selected = []
        try:
            for item in self.data_grid.SelectedItems:
                selected.append(item)
        except:
            pass
        return selected
    
    # ========================================================================
    # EVENT HANDLERS
    # ========================================================================
    
    def _on_filter_changed(self, sender, args):
        self._apply_filters()
    
    def _on_selection_changed(self, sender, args):
        if self.txt_selected:
            count = 0
            try:
                count = self.data_grid.SelectedItems.Count
            except:
                pass
            self.txt_selected.Text = str(count)
        
        # Update is_selected flags on items
        try:
            selected_set = set()
            for sel_item in self.data_grid.SelectedItems:
                selected_set.add(id(sel_item))
            
            for item in self.filtered_items:
                item._is_selected = id(item) in selected_set
        except:
            pass
    
    def _on_select_all(self, sender, args):
        self.data_grid.SelectAll()
    
    def _on_clear_all(self, sender, args):
        self.data_grid.UnselectAll()
        for item in self.filtered_items:
            item._is_selected = False
        self._update_stats()
    
    def _on_select_shared(self, sender, args):
        self.data_grid.UnselectAll()
        for i, item in enumerate(self.filtered_items):
            if item.is_shared:
                try:
                    self.data_grid.SelectedItems.Add(item)
                except:
                    pass
        self._update_stats()
    
    def _on_select_project(self, sender, args):
        self.data_grid.UnselectAll()
        for i, item in enumerate(self.filtered_items):
            if not item.is_shared:
                try:
                    self.data_grid.SelectedItems.Add(item)
                except:
                    pass
        self._update_stats()
    
    def _on_refresh(self, sender, args):
        self._load_data()
        MessageBox.Show("Data refreshed!", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_edit_group(self, sender, args):
        """Edit parameter group"""
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select one parameter to edit group!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if len(selected) > 1:
            MessageBox.Show("Please select only ONE parameter to edit group!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        
        dialog = EditGroupDialog(item, self.all_groups)
        dialog.ShowDialog()
        
        if dialog.result:
            self.data_grid.Items.Refresh()
    
    def _on_edit_categories(self, sender, args):
        """Edit parameter categories"""
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select one parameter to edit categories!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if len(selected) > 1:
            MessageBox.Show("Please select only ONE parameter to edit categories!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        
        dialog = EditCategoriesDialog(item, self.all_categories)
        dialog.ShowDialog()
        
        if dialog.result:
            self.data_grid.Items.Refresh()
    
    def _on_edit_binding(self, sender, args):
        """Edit parameter binding"""
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select one parameter to edit binding!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if len(selected) > 1:
            MessageBox.Show("Please select only ONE parameter to edit binding!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        
        dialog = EditBindingDialog(item)
        dialog.ShowDialog()
        
        if dialog.result:
            self.data_grid.Items.Refresh()
    
    def _on_export(self, sender, args):
        """Export parameters to CSV"""
        save_path = forms.save_file(file_ext='csv', default_name='Revit_Parameters_DQT.csv')
        
        if save_path:
            try:
                with codecs.open(save_path, 'w', encoding='utf-8-sig') as csvfile:
                    csvfile.write('Parameter Name,Type,Data Type,Group,Binding,Categories\n')
                    
                    for item in self.all_items:
                        name = '"{}"'.format(item.name) if ',' in item.name else item.name
                        cats = '"{}"'.format(item.categories) if ',' in item.categories else item.categories
                        
                        line = '{},{},{},{},{},{}\n'.format(
                            name, item.param_type, item.data_type, item.group, item.binding, cats
                        )
                        csvfile.write(line)
                
                MessageBox.Show("Exported {} parameters to:\n{}".format(len(self.all_items), save_path),
                              "Export Success", MessageBoxButton.OK, MessageBoxImage.Information)
                
                try:
                    os.startfile(save_path)
                except:
                    pass
            
            except Exception as e:
                MessageBox.Show("Export error:\n{}".format(str(e)),
                              "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_delete(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one parameter to delete!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        result = MessageBox.Show(
            "Delete {} parameter(s)?\n\nThis will remove the parameters from the project.\nParameter values will be lost!".format(len(selected)),
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning
        )
        
        if result != MessageBoxResult.Yes:
            return
        
        success_count = 0
        error_count = 0
        
        for item in selected:
            success, msg = delete_parameter(item.name)
            if success:
                success_count += 1
            else:
                error_count += 1
                print("Error deleting {}: {}".format(item.name, msg))
        
        msg = "Deleted: {}".format(success_count)
        if error_count > 0:
            msg += "\nFailed: {}".format(error_count)
        
        MessageBox.Show(msg, "Result", MessageBoxButton.OK, MessageBoxImage.Information)
        self._load_data()


# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == '__main__':
    try:
        window = ParameterManagerWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        MessageBox.Show(
            "Error starting Parameter Manager:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )