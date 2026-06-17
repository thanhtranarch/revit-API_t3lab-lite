# -*- coding: utf-8 -*-
"""Parameter Manager — event handling for the Parameter Manager launcher window."""

import os
import codecs
import __builtin__

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
                                      DataGridTextColumn, ScrollViewer, ListBox,
                                      ListBoxItem, SelectionMode)
import System.Windows.Controls as WPFControls
WPFGrid = WPFControls.Grid
WPFRowDefinition = WPFControls.RowDefinition
WPFColumnDefinition = WPFControls.ColumnDefinition
WPFBorder = WPFControls.Border

from System.Windows.Media import BrushConverter, SolidColorBrush
from System.Windows.Data import Binding as WPFBinding
from System.Windows.Controls import DataGridLength
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Reflection import BindingFlags

from pyrevit import revit, DB, forms, script


_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ParameterManager.xaml')

_brush_converter = BrushConverter()

def _hex_brush(hex_color):
    try:
        return _brush_converter.ConvertFromString(hex_color)
    except:
        return SolidColorBrush()


# ============================================================================
# VERSION HELPERS
# ============================================================================

def _eid_int(element_id):
    """ElementId integer value — compatible across Revit 2024/2025/2026."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def _get_revit_version():
    try:
        return int(revit.doc.Application.VersionNumber)
    except:
        return 2024


# ============================================================================
# MODULE-LEVEL HELPER FUNCTIONS (used by dialog classes and window methods)
# ============================================================================

def _get_group_name_from_definition(definition):
    """Return (group_display_name, group_id) — compatible across versions."""
    try:
        if hasattr(definition, 'GetGroupTypeId'):
            group_type_id = definition.GetGroupTypeId()
            if group_type_id and hasattr(group_type_id, 'TypeId'):
                try:
                    label = DB.LabelUtils.GetLabelForGroup(group_type_id)
                    if label:
                        return label, group_type_id
                except:
                    pass
                type_id_str = group_type_id.TypeId
                if type_id_str:
                    parts = type_id_str.split(':')
                    name = parts[-1] if len(parts) > 1 else parts[0]
                    name = name.split('-')[0]
                    name = name.replace('autodesk.parameter.group.', '')
                    name = name.replace('.', ' ').title()
                    return name, group_type_id
    except:
        pass

    try:
        param_group_enum = definition.ParameterGroup
        group_name = str(param_group_enum).replace('PG_', '').replace('_', ' ').title()
        return group_name, param_group_enum
    except:
        pass

    return "Unknown", None


def _get_parameter_data_type(definition):
    """Return readable data type string — compatible across versions."""
    try:
        if hasattr(definition, 'GetDataType'):
            spec_type_id = definition.GetDataType()
            try:
                label = DB.LabelUtils.GetLabelForSpec(spec_type_id)
                if label:
                    return label
            except:
                pass

            type_id_str = spec_type_id.TypeId if hasattr(spec_type_id, 'TypeId') else str(spec_type_id)
            if type_id_str and 'autodesk.spec' in type_id_str.lower():
                parts = type_id_str.split(':')
                if len(parts) > 1:
                    type_part = parts[-1].split('-')[0].replace('spec.', '')
                    type_map = {
                        'string': 'Text', 'int': 'Integer', 'integer': 'Integer',
                        'double': 'Number', 'number': 'Number', 'length': 'Length',
                        'area': 'Area', 'volume': 'Volume', 'angle': 'Angle',
                        'url': 'URL', 'material': 'Material', 'yesno': 'Yes/No',
                        'bool': 'Yes/No', 'boolean': 'Yes/No',
                        'multilinetext': 'Multiline Text', 'familytype': 'Family Type',
                        'image': 'Image',
                    }
                    for key, val in type_map.items():
                        if key in type_part.lower():
                            return val
                    return type_part.replace('.', ' ').title()

        rev_version = _get_revit_version()
        if rev_version < 2026:
            try:
                if hasattr(definition, 'ParameterType'):
                    return str(definition.ParameterType)
            except:
                pass
    except:
        pass
    return "Unknown"


def _get_shared_parameter_guids():
    """Return set of shared parameter names found in the document."""
    shared_names = set()
    try:
        doc = revit.doc
        collector = DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement)
        for elem in collector:
            try:
                shared_names.add(elem.GetDefinition().Name)
            except:
                pass
        try:
            def_file = doc.Application.OpenSharedParameterFile()
            if def_file:
                for grp in def_file.Groups:
                    for defn in grp.Definitions:
                        shared_names.add(defn.Name)
        except:
            pass
    except:
        pass
    return shared_names


def _get_all_parameter_groups():
    """Return sorted list of (group_id, display_name) tuples."""
    groups = []
    rev_version = _get_revit_version()

    if rev_version >= 2026:
        try:
            group_type_id_type = clr.GetClrType(DB.GroupTypeId) if hasattr(DB, 'GroupTypeId') else None
            if group_type_id_type:
                props = group_type_id_type.GetProperties(BindingFlags.Public | BindingFlags.Static)
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

    try:
        bipg_type = getattr(DB, 'BuiltInParameterGroup', None)
        if bipg_type:
            for group in DB.BuiltInParameterGroup.GetValues(bipg_type):
                if group != DB.BuiltInParameterGroup.INVALID:
                    name = str(group).replace('PG_', '').replace('_', ' ').title()
                    groups.append((group, name))
    except:
        pass

    return sorted(groups, key=lambda x: x[1])


def _get_all_categories():
    """Return sorted list of (category_object, category_name) tuples."""
    cats = []
    try:
        for cat in revit.doc.Settings.Categories:
            if cat.AllowsBoundParameters and cat.CategoryType == DB.CategoryType.Model:
                cats.append((cat, cat.Name))
    except:
        pass
    return sorted(cats, key=lambda x: x[1])


def _delete_parameter(param_name):
    doc = revit.doc
    try:
        t = DB.Transaction(doc, "T3Lab: Delete Parameter")
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
        except Exception as ex:
            t.RollBack()
            return False, str(ex)
    except Exception as ex:
        return False, str(ex)


def _update_parameter_group(param_name, new_group_id):
    doc = revit.doc
    try:
        t = DB.Transaction(doc, "T3Lab: Update Parameter Group")
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

            if hasattr(definition, 'SetGroupTypeId'):
                try:
                    if hasattr(new_group_id, 'TypeId'):
                        definition.SetGroupTypeId(new_group_id)
                        t.Commit()
                        return True, "Updated"
                    else:
                        if hasattr(DB, 'ParameterUtils') and hasattr(DB.ParameterUtils, 'GetParameterGroupTypeId'):
                            forge_id = DB.ParameterUtils.GetParameterGroupTypeId(new_group_id)
                            definition.SetGroupTypeId(forge_id)
                            t.Commit()
                            return True, "Updated"
                except:
                    pass

            definition_type = definition.GetType()
            field_names = [
                'm_builtInParameterGroup', 'm_parameterGroup', 'parameterGroup',
                'm_group', 'group', '_builtInParameterGroup', '_parameterGroup'
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
        except Exception as ex:
            t.RollBack()
            return False, str(ex)
    except Exception as ex:
        return False, str(ex)


def _update_parameter_categories(param_name, new_categories):
    doc = revit.doc
    app = doc.Application
    if doc.IsModifiable:
        return False, "Document already in transaction"
    try:
        t = DB.Transaction(doc, "T3Lab: Update Parameter Categories")
        t.Start()
        try:
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            old_definition = None
            old_binding = None
            while iterator.MoveNext():
                defn = iterator.Key
                binding = iterator.Current
                if defn.Name == param_name:
                    old_definition = defn
                    old_binding = binding
                    break
            if not old_definition:
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
                return False, "Failed to update categories"
        except Exception as ex:
            t.RollBack()
            return False, str(ex)
    except Exception as ex:
        return False, str(ex)


def _update_parameter_binding(param_name, new_binding_type):
    doc = revit.doc
    app = doc.Application
    if doc.IsModifiable:
        return False, "Document already in transaction"
    try:
        t = DB.Transaction(doc, "T3Lab: Update Parameter Binding")
        t.Start()
        try:
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            old_definition = None
            old_binding = None
            while iterator.MoveNext():
                defn = iterator.Key
                binding = iterator.Current
                if defn.Name == param_name:
                    old_definition = defn
                    old_binding = binding
                    break
            if not old_definition:
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
        except Exception as ex:
            t.RollBack()
            return False, str(ex)
    except Exception as ex:
        return False, str(ex)


# ============================================================================
# PARAMETER DATA MODEL
# ============================================================================

class ParameterItem(INotifyPropertyChanged):
    """WPF-bindable parameter data model."""

    def __init__(self, name, param_type, data_type, group, group_id,
                 binding, categories, category_objects=None):
        self._handlers = []
        self._name = name
        self._param_type = param_type
        self._data_type = data_type
        self._group = group
        self._group_id = group_id
        self._binding = binding
        self._categories = categories
        self._category_objects = category_objects or []
        self._is_shared = (param_type == "Shared")

    # INotifyPropertyChanged
    def add_PropertyChanged(self, handler):
        self._handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._handlers:
            self._handlers.remove(handler)

    def OnPropertyChanged(self, prop_name):
        for h in self._handlers:
            try:
                h(self, PropertyChangedEventArgs(prop_name))
            except:
                pass

    @property
    def name(self):
        return self._name

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
    def is_shared(self):
        return self._is_shared


# ============================================================================
# EDIT DIALOGS (pure Python WPF windows — no XAML needed)
# ============================================================================

class EditGroupDialog(Window):
    def __init__(self, item, all_groups):
        self.item = item
        self.all_groups = all_groups
        self.result = False
        self._build_ui()

    def _build_ui(self):
        self.Title = "Edit Parameter Group"
        self.Width = 450
        self.Height = 500
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = _hex_brush("#F8FAFC")

        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        hdr = WPFBorder()
        hdr.Background = _hex_brush("#0F172A")
        hdr.Padding = Thickness(10, 8, 10, 8)
        hdr.Margin = Thickness(0, 0, 0, 10)
        hdr_text = TextBlock()
        hdr_text.Text = "Edit Parameter Group: {}".format(self.item.name)
        hdr_text.FontSize = 14
        hdr_text.FontWeight = FontWeights.SemiBold
        hdr_text.Foreground = _hex_brush("#FFFFFF")
        hdr.Child = hdr_text
        WPFGrid.SetRow(hdr, 0)
        main_grid.Children.Add(hdr)

        info = TextBlock()
        info.Text = "Current Group: {}".format(self.item.group)
        info.Foreground = _hex_brush("#64748B")
        info.Margin = Thickness(0, 0, 0, 8)
        WPFGrid.SetRow(info, 1)
        main_grid.Children.Add(info)

        list_panel = StackPanel()
        lbl = TextBlock()
        lbl.Text = "Select New Group:"
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Margin = Thickness(0, 0, 0, 4)
        list_panel.Children.Add(lbl)
        self.group_listbox = ListBox()
        self.group_listbox.Height = 300
        for group_id, group_name in self.all_groups:
            lb = ListBoxItem()
            lb.Content = group_name
            lb.Tag = group_id
            if group_name == self.item.group:
                lb.IsSelected = True
            self.group_listbox.Items.Add(lb)
        list_panel.Children.Add(self.group_listbox)
        WPFGrid.SetRow(list_panel, 2)
        main_grid.Children.Add(list_panel)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 10, 0, 0)
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 90
        btn_apply.Height = 32
        btn_apply.Margin = Thickness(0, 0, 8, 0)
        btn_apply.Background = _hex_brush("#10B981")
        btn_apply.Foreground = _hex_brush("#FFFFFF")
        btn_apply.Click += self._on_apply
        btn_panel.Children.Add(btn_apply)
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 90
        btn_cancel.Height = 32
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_cancel)
        WPFGrid.SetRow(btn_panel, 3)
        main_grid.Children.Add(btn_panel)
        self.Content = main_grid

    def _on_apply(self, sender, args):
        if not self.group_listbox.SelectedItem:
            MessageBox.Show("Please select a group.", "Warning",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        sel = self.group_listbox.SelectedItem
        new_name = sel.Content
        new_id = sel.Tag
        if new_name == self.item.group:
            MessageBox.Show("No change — same group selected.", "Info",
                            MessageBoxButton.OK, MessageBoxImage.Information)
            return
        success, msg = _update_parameter_group(self.item.name, new_id)
        if success:
            self.item.group = new_name
            self.item.group_id = new_id
            self.result = True
            MessageBox.Show("Group updated successfully.", "Success",
                            MessageBoxButton.OK, MessageBoxImage.Information)
            self.Close()
        else:
            MessageBox.Show("Failed to update group:\n\n{}".format(msg), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)


class EditCategoriesDialog(Window):
    def __init__(self, item, all_categories):
        self.item = item
        self.all_categories = all_categories
        self.result = False
        self._build_ui()

    def _build_ui(self):
        self.Title = "Edit Parameter Categories"
        self.Width = 450
        self.Height = 560
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = _hex_brush("#F8FAFC")

        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Star)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        hdr = WPFBorder()
        hdr.Background = _hex_brush("#3B82F6")
        hdr.Padding = Thickness(10, 8, 10, 8)
        hdr.Margin = Thickness(0, 0, 0, 10)
        hdr_text = TextBlock()
        hdr_text.Text = "Edit Categories: {}".format(self.item.name)
        hdr_text.FontSize = 14
        hdr_text.FontWeight = FontWeights.SemiBold
        hdr_text.Foreground = _hex_brush("#FFFFFF")
        hdr.Child = hdr_text
        WPFGrid.SetRow(hdr, 0)
        main_grid.Children.Add(hdr)

        info = TextBlock()
        info.Text = "Type: {} | Binding: {}".format(self.item.param_type, self.item.binding)
        info.Foreground = _hex_brush("#64748B")
        info.Margin = Thickness(0, 0, 0, 6)
        WPFGrid.SetRow(info, 1)
        main_grid.Children.Add(info)

        quick_panel = StackPanel()
        quick_panel.Orientation = Orientation.Horizontal
        quick_panel.Margin = Thickness(0, 0, 0, 8)
        btn_all = Button()
        btn_all.Content = "Select All"
        btn_all.Padding = Thickness(8, 4, 8, 4)
        btn_all.Margin = Thickness(0, 0, 6, 0)
        btn_all.Click += self._on_select_all
        quick_panel.Children.Add(btn_all)
        btn_none = Button()
        btn_none.Content = "Select None"
        btn_none.Padding = Thickness(8, 4, 8, 4)
        btn_none.Click += self._on_select_none
        quick_panel.Children.Add(btn_none)
        WPFGrid.SetRow(quick_panel, 2)
        main_grid.Children.Add(quick_panel)

        list_panel = StackPanel()
        lbl = TextBlock()
        lbl.Text = "Select Categories:"
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Margin = Thickness(0, 0, 0, 4)
        list_panel.Children.Add(lbl)
        current_names = []
        if self.item.categories and self.item.categories != "N/A":
            current_names = [c.strip() for c in self.item.categories.split(',')]
        self.cat_listbox = ListBox()
        self.cat_listbox.Height = 300
        self.cat_listbox.SelectionMode = SelectionMode.Multiple
        for cat_obj, cat_name in self.all_categories:
            lb = ListBoxItem()
            lb.Content = cat_name
            lb.Tag = cat_obj
            if cat_name in current_names:
                lb.IsSelected = True
            self.cat_listbox.Items.Add(lb)
        list_panel.Children.Add(self.cat_listbox)
        WPFGrid.SetRow(list_panel, 3)
        main_grid.Children.Add(list_panel)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 10, 0, 0)
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 90
        btn_apply.Height = 32
        btn_apply.Margin = Thickness(0, 0, 8, 0)
        btn_apply.Background = _hex_brush("#10B981")
        btn_apply.Foreground = _hex_brush("#FFFFFF")
        btn_apply.Click += self._on_apply
        btn_panel.Children.Add(btn_apply)
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 90
        btn_cancel.Height = 32
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_cancel)
        WPFGrid.SetRow(btn_panel, 4)
        main_grid.Children.Add(btn_panel)
        self.Content = main_grid

    def _on_select_all(self, sender, args):
        for lb in self.cat_listbox.Items:
            lb.IsSelected = True

    def _on_select_none(self, sender, args):
        for lb in self.cat_listbox.Items:
            lb.IsSelected = False

    def _on_apply(self, sender, args):
        selected = [lb for lb in self.cat_listbox.Items if lb.IsSelected]
        if not selected:
            MessageBox.Show("Please select at least one category.", "Warning",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        cats = [lb.Tag for lb in selected]
        names = [lb.Content for lb in selected]
        success, msg = _update_parameter_categories(self.item.name, cats)
        if success:
            self.item.categories = ', '.join(names)
            self.item.category_objects = cats
            self.result = True
            MessageBox.Show("Categories updated successfully.", "Success",
                            MessageBoxButton.OK, MessageBoxImage.Information)
            self.Close()
        else:
            MessageBox.Show("Failed:\n\n{}".format(msg), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)


class EditBindingDialog(Window):
    def __init__(self, item):
        self.item = item
        self.result = False
        self._build_ui()

    def _build_ui(self):
        self.Title = "Edit Parameter Binding"
        self.Width = 400
        self.Height = 280
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = _hex_brush("#F8FAFC")

        main_grid = WPFGrid()
        main_grid.Margin = Thickness(15)
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))
        main_grid.RowDefinitions.Add(WPFRowDefinition(Height=GridLength(1, GridUnitType.Auto)))

        hdr = WPFBorder()
        hdr.Background = _hex_brush("#F59E0B")
        hdr.Padding = Thickness(10, 8, 10, 8)
        hdr.Margin = Thickness(0, 0, 0, 10)
        hdr_text = TextBlock()
        hdr_text.Text = "Edit Binding: {}".format(self.item.name)
        hdr_text.FontSize = 14
        hdr_text.FontWeight = FontWeights.SemiBold
        hdr_text.Foreground = _hex_brush("#FFFFFF")
        hdr.Child = hdr_text
        WPFGrid.SetRow(hdr, 0)
        main_grid.Children.Add(hdr)

        info_panel = StackPanel()
        info_panel.Margin = Thickness(0, 0, 0, 10)
        info1 = TextBlock()
        info1.Text = "Type: {} | Current Binding: {}".format(self.item.param_type, self.item.binding)
        info1.Foreground = _hex_brush("#64748B")
        info_panel.Children.Add(info1)
        if not self.item.is_shared:
            warn = TextBlock()
            warn.Text = u"⚠ Project parameters cannot change binding via API"
            warn.Foreground = _hex_brush("#EF4444")
            warn.Margin = Thickness(0, 6, 0, 0)
            warn.TextWrapping = System.Windows.TextWrapping.Wrap
            info_panel.Children.Add(warn)
        WPFGrid.SetRow(info_panel, 1)
        main_grid.Children.Add(info_panel)

        select_panel = StackPanel()
        select_panel.Margin = Thickness(0, 0, 0, 10)
        lbl = TextBlock()
        lbl.Text = "Select Binding:"
        lbl.FontWeight = FontWeights.SemiBold
        lbl.Margin = Thickness(0, 0, 0, 4)
        select_panel.Children.Add(lbl)
        self.binding_combo = ComboBox()
        self.binding_combo.Padding = Thickness(8, 4, 8, 4)
        self.binding_combo.IsEnabled = self.item.is_shared
        for opt in ["Instance", "Type"]:
            cb = ComboBoxItem()
            cb.Content = opt
            self.binding_combo.Items.Add(cb)
        self.binding_combo.SelectedIndex = 0 if self.item.binding == "Instance" else 1
        select_panel.Children.Add(self.binding_combo)
        WPFGrid.SetRow(select_panel, 2)
        main_grid.Children.Add(select_panel)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_apply = Button()
        btn_apply.Content = "Apply"
        btn_apply.Width = 90
        btn_apply.Height = 32
        btn_apply.Margin = Thickness(0, 0, 8, 0)
        btn_apply.Background = _hex_brush("#10B981")
        btn_apply.Foreground = _hex_brush("#FFFFFF")
        btn_apply.IsEnabled = self.item.is_shared
        btn_apply.Click += self._on_apply
        btn_panel.Children.Add(btn_apply)
        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 90
        btn_cancel.Height = 32
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_cancel)
        WPFGrid.SetRow(btn_panel, 3)
        main_grid.Children.Add(btn_panel)
        self.Content = main_grid

    def _on_apply(self, sender, args):
        new_binding = self.binding_combo.SelectedItem.Content
        if new_binding == self.item.binding:
            MessageBox.Show("No change — same binding selected.", "Info",
                            MessageBoxButton.OK, MessageBoxImage.Information)
            return
        success, msg = _update_parameter_binding(self.item.name, new_binding)
        if success:
            self.item.binding = new_binding
            self.result = True
            MessageBox.Show("Binding updated successfully.", "Success",
                            MessageBoxButton.OK, MessageBoxImage.Information)
            self.Close()
        else:
            MessageBox.Show("Failed:\n\n{}".format(msg), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# MAIN WINDOW
# ============================================================================

class ParameterManagerWindow(forms.WPFWindow):
    def __init__(self, script_dir, revit_obj):
        forms.WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._revit = revit_obj

        # Existing 3-tab buttons
        self.btn_transfer_params.Click += self._open_transfer_params
        self.btn_text_to_element.Click += self._open_text_to_element
        self.btn_values_to_region.Click += self._open_values_to_region

        # Window chrome buttons
        self.btn_minimize.Click += self._minimize
        self.btn_maximize.Click += self._maximize
        self.btn_close_chrome.Click += self._close_chrome
        self.PreviewKeyDown += self._on_key_down

        # Browse Parameters tab — populate ComboBox items from Python
        self._populate_filter_combos()

        # Browse Parameters tab event handlers
        self.txt_param_search.TextChanged += self._on_param_filter_changed
        self.cmb_param_type.SelectionChanged += self._on_param_filter_changed
        self.cmb_param_binding.SelectionChanged += self._on_param_filter_changed
        self.btn_param_refresh.Click += self._on_param_refresh
        self.btn_param_edit_group.Click += self._on_param_edit_group
        self.btn_param_edit_cats.Click += self._on_param_edit_cats
        self.btn_param_edit_binding.Click += self._on_param_edit_binding
        self.btn_param_export.Click += self._on_param_export
        self.btn_param_delete.Click += self._on_param_delete

        # State
        self._all_param_items = []
        self._all_groups = []
        self._all_categories = []

        # Load parameter data
        self._load_param_data()

    # ------------------------------------------------------------------
    # Existing launcher methods
    # ------------------------------------------------------------------

    def _launch(self, rel_path):
        script_path = os.path.normpath(os.path.join(self._script_dir, rel_path))
        self.Close()
        g = {'__name__': '__main__', '__file__': script_path,
             '__builtins__': __builtin__, '__revit__': self._revit}
        try:
            execfile(script_path, g)
        except Exception as ex:
            forms.alert("Error launching tool:\n{}".format(ex))

    def _open_transfer_params(self, sender, e):
        self._launch("../Transfer Para/script.py")

    def _open_text_to_element(self, sender, e):
        self._launch("../Text to element/script.py")

    def _open_values_to_region(self, sender, e):
        self._launch("../Values to Filled Region /script.py")

    def _minimize(self, sender, e):
        self.WindowState = System.Windows.WindowState.Minimized

    def _maximize(self, sender, e):
        if self.WindowState == System.Windows.WindowState.Maximized:
            self.WindowState = System.Windows.WindowState.Normal
        else:
            self.WindowState = System.Windows.WindowState.Maximized

    def _close_chrome(self, sender, e):
        self.Close()

    def _on_key_down(self, sender, e):
        import System.Windows.Input as WI
        if e.Key == WI.Key.Escape:
            self.Close()
        elif e.Key == WI.Key.F5:
            try:
                self._on_param_refresh(None, None)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Browse Parameters helpers
    # ------------------------------------------------------------------

    def _populate_filter_combos(self):
        for text in ["All", "Shared Only", "Project Only"]:
            cb = ComboBoxItem()
            cb.Content = text
            self.cmb_param_type.Items.Add(cb)
        self.cmb_param_type.SelectedIndex = 0

        for text in ["All", "Instance", "Type"]:
            cb = ComboBoxItem()
            cb.Content = text
            self.cmb_param_binding.Items.Add(cb)
        self.cmb_param_binding.SelectedIndex = 0

    def _set_status(self, text):
        try:
            self.txt_param_status.Text = text
        except:
            pass

    def _load_param_data(self):
        self._set_status("Loading parameters...")
        self._all_param_items = []
        try:
            doc = revit.doc
            shared_names = _get_shared_parameter_guids()
            param_bindings = doc.ParameterBindings
            iterator = param_bindings.ForwardIterator()
            while iterator.MoveNext():
                definition = iterator.Key
                binding = iterator.Current
                param_name = definition.Name
                is_shared = param_name in shared_names
                param_type = "Shared" if is_shared else "Project"
                data_type = _get_parameter_data_type(definition)
                group_name, group_id = _get_group_name_from_definition(definition)
                if isinstance(binding, DB.InstanceBinding):
                    binding_type = "Instance"
                elif isinstance(binding, DB.TypeBinding):
                    binding_type = "Type"
                else:
                    binding_type = "Unknown"
                categories = []
                cat_objects = []
                if hasattr(binding, 'Categories'):
                    for cat in binding.Categories:
                        categories.append(cat.Name)
                        cat_objects.append(cat)
                item = ParameterItem(
                    param_name, param_type, data_type, group_name, group_id,
                    binding_type,
                    ', '.join(categories) if categories else 'N/A',
                    cat_objects
                )
                self._all_param_items.append(item)
            self._all_param_items.sort(key=lambda x: x.name)
            # Cache groups and categories for dialog use
            self._all_groups = _get_all_parameter_groups()
            self._all_categories = _get_all_categories()
            self._apply_param_filters()
            self._set_status("Loaded {} parameters.".format(len(self._all_param_items)))
        except Exception as ex:
            self._set_status("Error loading parameters: {}".format(str(ex)))

    def _apply_param_filters(self):
        try:
            search_text = ""
            if self.txt_param_search and self.txt_param_search.Text:
                search_text = self.txt_param_search.Text.lower()

            type_idx = 0
            if self.cmb_param_type:
                type_idx = self.cmb_param_type.SelectedIndex

            binding_idx = 0
            if self.cmb_param_binding:
                binding_idx = self.cmb_param_binding.SelectedIndex

            filtered = []
            for item in self._all_param_items:
                if search_text and search_text not in item.name.lower():
                    continue
                if type_idx == 1 and not item.is_shared:
                    continue
                if type_idx == 2 and item.is_shared:
                    continue
                if binding_idx == 1 and item.binding != "Instance":
                    continue
                if binding_idx == 2 and item.binding != "Type":
                    continue
                filtered.append(item)

            self.dg_parameters.ItemsSource = ObservableCollection[object](filtered)
            self._set_status("Showing {} of {} parameters.".format(
                len(filtered), len(self._all_param_items)))
        except Exception as ex:
            self._set_status("Filter error: {}".format(str(ex)))

    def _get_selected_param(self):
        """Return the single selected ParameterItem or None."""
        try:
            sel = self.dg_parameters.SelectedItem
            return sel
        except:
            return None

    # ------------------------------------------------------------------
    # Browse Parameters event handlers
    # ------------------------------------------------------------------

    def _on_param_filter_changed(self, sender, e):
        self._apply_param_filters()

    def _on_param_refresh(self, sender, e):
        self._load_param_data()

    def _on_param_edit_group(self, sender, e):
        item = self._get_selected_param()
        if not item:
            MessageBox.Show("Please select a parameter first.", "Warning",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        dlg = EditGroupDialog(item, self._all_groups)
        dlg.ShowDialog()
        if dlg.result:
            self.dg_parameters.Items.Refresh()
            self._set_status("Group updated for '{}'.".format(item.name))

    def _on_param_edit_cats(self, sender, e):
        item = self._get_selected_param()
        if not item:
            MessageBox.Show("Please select a parameter first.", "Warning",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        dlg = EditCategoriesDialog(item, self._all_categories)
        dlg.ShowDialog()
        if dlg.result:
            self.dg_parameters.Items.Refresh()
            self._set_status("Categories updated for '{}'.".format(item.name))

    def _on_param_edit_binding(self, sender, e):
        item = self._get_selected_param()
        if not item:
            MessageBox.Show("Please select a parameter first.", "Warning",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        dlg = EditBindingDialog(item)
        dlg.ShowDialog()
        if dlg.result:
            self.dg_parameters.Items.Refresh()
            self._set_status("Binding updated for '{}'.".format(item.name))

    def _on_param_export(self, sender, e):
        save_path = forms.save_file(file_ext='csv', default_name='Revit_Parameters.csv')
        if not save_path:
            return
        try:
            with codecs.open(save_path, 'w', encoding='utf-8-sig') as f:
                f.write('Parameter Name,Type,Data Type,Group,Binding,Categories\n')
                for item in self._all_param_items:
                    name = '"{}"'.format(item.name) if ',' in item.name else item.name
                    cats = '"{}"'.format(item.categories) if ',' in item.categories else item.categories
                    f.write('{},{},{},{},{},{}\n'.format(
                        name, item.param_type, item.data_type,
                        item.group, item.binding, cats
                    ))
            self._set_status("Exported {} parameters to {}".format(
                len(self._all_param_items), save_path))
            MessageBox.Show("Exported {} parameters.".format(len(self._all_param_items)),
                            "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            MessageBox.Show("Export error:\n{}".format(str(ex)), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_param_delete(self, sender, e):
        item = self._get_selected_param()
        if not item:
            MessageBox.Show("Please select a parameter to delete.", "Warning",
                            MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        result = MessageBox.Show(
            "Delete parameter '{}'?\n\nAll parameter values will be lost!".format(item.name),
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning
        )
        if result != MessageBoxResult.Yes:
            return
        success, msg = _delete_parameter(item.name)
        if success:
            self._set_status("Deleted '{}'.".format(item.name))
            self._load_param_data()
        else:
            MessageBox.Show("Delete failed:\n\n{}".format(msg), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)


# ============================================================================
# PUBLIC API
# ============================================================================

def show_parameter_manager(script_dir, revit_obj):
    ParameterManagerWindow(script_dir, revit_obj).ShowDialog()
