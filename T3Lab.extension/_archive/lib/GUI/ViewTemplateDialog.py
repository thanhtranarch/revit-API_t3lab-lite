# -*- coding: utf-8 -*-
"""
View Template Manager Dialog
GUI classes for View Template Manager.
"""

import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

import System
from System.Windows import (Thickness, GridLength, GridUnitType,
                            HorizontalAlignment, VerticalAlignment, FontWeights,
                            MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult, WindowState)
from System.Windows.Controls import (RowDefinition, ColumnDefinition, Border,
                                      StackPanel, TextBlock, TextBox, Button,
                                      ComboBox, ComboBoxItem, DataGrid, Orientation,
                                      DataGridTextColumn, DataGridCheckBoxColumn,
                                      ScrollViewer, TabControl, TabItem, CheckBox)
from System.Windows.Media import SolidColorBrush, BrushConverter
from System.Windows.Data import Binding
from System.Windows.Controls import DataGridLength
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

# CRITICAL: Import WPF Grid BEFORE Revit wildcard import
from System.Windows.Controls import Grid as WPFGrid

import re

from pyrevit import revit, DB, forms

# Import Core/Execution functions
from core.view_template import (
    _eid_int,
    _eid_invalid_value,
    calculate_viewtemplate_usage,
    rename_template,
    batch_rename_templates,
    duplicate_templates,
    delete_templates
)

doc = revit.doc

# XAML Paths
GUI_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'ViewTemplateManager.xaml')
BATCH_XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'ViewTemplateBatchRename.xaml')


def _hex_to_brush(hex_color):
    """Convert hex color string to SolidColorBrush."""
    converter = BrushConverter()
    return converter.ConvertFromString(hex_color)


class Config(object):
    """Color scheme and settings"""
    PRIMARY_COLOR = "#0F172A"
    SECONDARY_COLOR = "#3B82F6"
    BACKGROUND_COLOR = "#F8FAFC"
    BORDER_COLOR = "#E2E8F0"
    TEXT_DARK = "#0F172A"
    TEXT_LIGHT = "#64748B"
    SUCCESS_COLOR = "#10B981"
    WARNING_COLOR = "#F59E0B"
    ERROR_COLOR = "#EF4444"
    ROW_ALT_COLOR = "#F8FAFC"
    WHITE = "#FFFFFF"
    BLUE_COLOR = "#3B82F6"
    GRID_LINE_COLOR = "#F1F5F9"


class ViewTemplateItem(INotifyPropertyChanged):
    """Wrapper class for View Template with WPF binding support"""
    
    def __init__(self, view_template):
        self._property_changed_handlers = []
        self._view_template = view_template
        self._is_selected = False
        
        self._name = view_template.Name
        self._id = _eid_int(view_template.Id)
        
        # Get view type
        try:
            self._view_type = str(view_template.ViewType)
        except:
            self._view_type = "Unknown"
        
        # Get scale
        try:
            scale_param = view_template.get_Parameter(DB.BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC)
            if scale_param and scale_param.HasValue:
                self._scale = scale_param.AsValueString()
            else:
                self._scale = "N/A"
        except:
            self._scale = "N/A"
        
        self._usage_count = 0
        self._usage_percentage = 0.0
    
    @property
    def element(self):
        return self._view_template
    
    @property
    def view_template(self):
        return self._view_template
    
    @property
    def id(self):
        return self._id
    
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
    def view_type(self):
        return self._view_type
    
    @property
    def scale(self):
        return self._scale
    
    @property
    def usage_count(self):
        return self._usage_count
    
    @usage_count.setter
    def usage_count(self, value):
        self._usage_count = value
        self.OnPropertyChanged("usage_count")
    
    @property
    def usage_percentage(self):
        return str(round(self._usage_percentage, 1)) + "%"
    
    @usage_percentage.setter
    def usage_percentage(self, value):
        self._usage_percentage = value
        self.OnPropertyChanged("usage_percentage")
    
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


class BatchRenameDialog(forms.WPFWindow):
    """Comprehensive Batch Rename Dialog with multiple tabs"""
    
    def __init__(self, items, parent_window):
        forms.WPFWindow.__init__(self, BATCH_XAML_FILE)
        self.items = items
        self.parent_window = parent_window
        
        # Bind events
        self.tab_control.SelectionChanged += lambda s, e: self._update_preview()
        
        self.prefix_box.TextChanged += lambda s, e: self._update_preview()
        self.suffix_box.TextChanged += lambda s, e: self._update_preview()
        
        self.find_box.TextChanged += lambda s, e: self._update_preview()
        self.replace_box.TextChanged += lambda s, e: self._update_preview()
        self.case_check.Checked += lambda s, e: self._update_preview()
        self.case_check.Unchecked += lambda s, e: self._update_preview()
        
        self.remove_numbers_check.Checked += lambda s, e: self._update_preview()
        self.remove_numbers_check.Unchecked += lambda s, e: self._update_preview()
        self.remove_special_check.Checked += lambda s, e: self._update_preview()
        self.remove_special_check.Unchecked += lambda s, e: self._update_preview()
        self.remove_spaces_check.Checked += lambda s, e: self._update_preview()
        self.remove_spaces_check.Unchecked += lambda s, e: self._update_preview()
        self.remove_custom_box.TextChanged += lambda s, e: self._update_preview()
        
        self.case_combo.SelectionChanged += lambda s, e: self._update_preview()
        
        self.numbering_check.Checked += lambda s, e: self._update_preview()
        self.numbering_check.Unchecked += lambda s, e: self._update_preview()
        self.start_number_box.TextChanged += lambda s, e: self._update_preview()
        self.padding_box.TextChanged += lambda s, e: self._update_preview()
        self.number_position_combo.SelectionChanged += lambda s, e: self._update_preview()
        self.number_separator_box.TextChanged += lambda s, e: self._update_preview()
        
        self.add_viewtype_check.Checked += lambda s, e: self._update_preview()
        self.add_viewtype_check.Unchecked += lambda s, e: self._update_preview()
        self.add_scale_check.Checked += lambda s, e: self._update_preview()
        self.add_scale_check.Unchecked += lambda s, e: self._update_preview()
        self.remove_current_name_check.Checked += lambda s, e: self._update_preview()
        self.remove_current_name_check.Unchecked += lambda s, e: self._update_preview()
        self.info_position_combo.SelectionChanged += lambda s, e: self._update_preview()
        self.info_bracket_combo.SelectionChanged += lambda s, e: self._update_preview()
        self.info_separator_box.TextChanged += lambda s, e: self._update_preview()
        
        self.apply_btn.Click += self._on_apply
        
        self.info_text.Text = "{} template(s) selected for renaming".format(len(self.items))
        self._update_preview()
        
    def close_button_clicked(self, sender, e):
        self.Close()
    
    def _sanitize_name(self, name):
        """Sanitize name for Revit view"""
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|', '{', '}', '[', ']', ';']
        for char in invalid_chars:
            name = name.replace(char, '')
        
        while '  ' in name:
            name = name.replace('  ', ' ')
        while '__' in name:
            name = name.replace('__', '_')
        
        return name.strip()
    
    def _apply_rules_to_name(self, name, index, item=None):
        """Apply all rename rules to a single name"""
        new_name = name
        selected_tab = self.tab_control.SelectedIndex
        
        # Tab 0: Prefix/Suffix
        if selected_tab == 0:
            prefix = self.prefix_box.Text if self.prefix_box.Text else ""
            suffix = self.suffix_box.Text if self.suffix_box.Text else ""
            new_name = prefix + new_name + suffix
        
        # Tab 1: Find/Replace
        elif selected_tab == 1:
            find_text = self.find_box.Text if self.find_box.Text else ""
            replace_text = self.replace_box.Text if self.replace_box.Text else ""
            
            if find_text:
                if self.case_check.IsChecked:
                    new_name = new_name.replace(find_text, replace_text)
                else:
                    pattern = re.compile(re.escape(find_text), re.IGNORECASE)
                    new_name = pattern.sub(replace_text, new_name)
        
        # Tab 2: Remove Characters
        elif selected_tab == 2:
            if self.remove_numbers_check.IsChecked:
                new_name = re.sub(r'[0-9]', '', new_name)
            
            if self.remove_special_check.IsChecked:
                new_name = re.sub(r'[!@#$%^&*()+=\[\]{};:\'",.<>?/\\|`~]', '', new_name)
            
            if self.remove_spaces_check.IsChecked:
                new_name = new_name.replace(' ', '')
            
            custom = self.remove_custom_box.Text if self.remove_custom_box.Text else ""
            if custom:
                for char in custom:
                    new_name = new_name.replace(char, '')
        
        # Tab 3: Case Change
        elif selected_tab == 3:
            case_option = self.case_combo.SelectedIndex
            if case_option == 1:
                new_name = new_name.upper()
            elif case_option == 2:
                new_name = new_name.lower()
            elif case_option == 3:
                new_name = new_name.title()
            elif case_option == 4:
                new_name = new_name.capitalize()
        
        # Tab 4: Numbering
        elif selected_tab == 4:
            if self.numbering_check.IsChecked:
                try:
                    start = int(self.start_number_box.Text) if self.start_number_box.Text else 1
                    padding = int(self.padding_box.Text) if self.padding_box.Text else 2
                except:
                    start = 1
                    padding = 2
                
                number = str(start + index).zfill(padding)
                separator = self.number_separator_box.Text if self.number_separator_box.Text else "_"
                position = self.number_position_combo.SelectedIndex
                
                if position == 0:
                    new_name = number + separator + new_name
                else:
                    new_name = new_name + separator + number
        
        # Tab 5: Template Info
        elif selected_tab == 5:
            if item:
                info_parts = []
                
                if self.add_viewtype_check.IsChecked:
                    info_parts.append(item.view_type)
                
                if self.add_scale_check.IsChecked:
                    scale = item.scale.replace(":", "-") if item.scale else "N-A"
                    info_parts.append(scale)
                
                if info_parts:
                    bracket_index = self.info_bracket_combo.SelectedIndex if self.info_bracket_combo else 0
                    brackets = [("", ""), ("(", ")"), ("_", "_")]
                    left_b, right_b = brackets[bracket_index] if bracket_index < len(brackets) else ("", "")
                    
                    info_str = "".join(["{}{}{}".format(left_b, part, right_b) for part in info_parts])
                    
                    separator = self.info_separator_box.Text if self.info_separator_box.Text else "_"
                    
                    if self.remove_current_name_check.IsChecked:
                        new_name = info_str
                    else:
                        position = self.info_position_combo.SelectedIndex if self.info_position_combo else 1
                        
                        if position == 0:
                            new_name = info_str + separator + new_name
                        else:
                            new_name = new_name + separator + info_str
        
        return self._sanitize_name(new_name)
    
    def _update_preview(self, sender=None, args=None):
        if not self.preview_text:
            return
        
        preview_lines = []
        changed_count = 0
        
        for i, item in enumerate(self.items[:15]):
            old_name = item.name
            new_name = self._apply_rules_to_name(old_name, i, item)
            
            if old_name != new_name:
                preview_lines.append("{} -> {}".format(old_name, new_name))
                changed_count += 1
            else:
                preview_lines.append("{} (no change)".format(old_name))
        
        if len(self.items) > 15:
            preview_lines.append("... and {} more items".format(len(self.items) - 15))
        
        if changed_count > 0:
            self.preview_text.Text = "\n".join(preview_lines)
            self.preview_text.Foreground = _hex_to_brush("#0F172A")
        else:
            self.preview_text.Text = "No changes will be made with current settings"
            self.preview_text.Foreground = _hex_to_brush(Config.TEXT_LIGHT)
    
    def _on_apply(self, sender, args):
        changes = []
        for i, item in enumerate(self.items):
            old_name = item.name
            new_name = self._apply_rules_to_name(old_name, i, item)
            if old_name != new_name and new_name:
                changes.append((item.view_template, new_name))
        
        if not changes:
            MessageBox.Show("No changes to apply!", "Info",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        try:
            success_count = batch_rename_templates(doc, changes)
            MessageBox.Show("Successfully renamed: {}".format(success_count), "Result",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            
            self.Close()
            self.parent_window._load_data()
            
        except Exception as ex:
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


class ViewTemplateManagerWindow(forms.WPFWindow):
    """View Template Manager with Sheet Manager style UI"""
    
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML_FILE)
        self.all_items = []
        self.filtered_items = []
        
        # Bind events
        self.search_box.TextChanged += self._on_filter_changed
        self.filter_combo.SelectionChanged += self._on_filter_changed
        self.viewtype_combo.SelectionChanged += self._on_filter_changed
        
        self.btn_all.Click += self._on_select_all
        self.btn_clear.Click += self._on_clear_all
        self.btn_unused.Click += self._on_select_unused
        
        self.btn_rename.Click += self._on_rename
        self.btn_batch.Click += self._on_batch_rename
        self.btn_duplicate.Click += self._on_duplicate
        self.btn_delete.Click += self._on_delete
        
        self.btn_refresh.Click += self._on_refresh
        self.btn_close_action.Click += lambda s, e: self.Close()
        
        self.data_grid.SelectionChanged += self._on_selection_changed
        
        self._load_data()
        
    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()
    
    def _load_data(self):
        self.all_items = []
        
        try:
            collector = DB.FilteredElementCollector(doc).OfClass(DB.View)
            
            view_types = set()
            
            for view in collector:
                try:
                    if view.IsTemplate:
                        item = ViewTemplateItem(view)
                        self.all_items.append(item)
                        view_types.add(item.view_type)
                except:
                    pass
            
            self.all_items.sort(key=lambda x: x.name)
            
            # Calculate usage
            calculate_viewtemplate_usage(doc, self.all_items)
            
            # Populate view type filter
            self.viewtype_combo.Items.Clear()
            item = ComboBoxItem()
            item.Content = "All View Types"
            self.viewtype_combo.Items.Add(item)
            
            for vt in sorted(view_types):
                item = ComboBoxItem()
                item.Content = vt
                self.viewtype_combo.Items.Add(item)
            
            self.viewtype_combo.SelectedIndex = 0
            
            self._apply_filters()
            self._update_stats()
            
        except Exception as ex:
            print("Error loading view templates: {}".format(str(ex)))
            MessageBox.Show("Failed to load view templates:\n\n{}".format(str(ex)),
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _apply_filters(self):
        search_text = ""
        if self.search_box and self.search_box.Text:
            search_text = self.search_box.Text.lower()
        
        filter_index = 0
        if self.filter_combo:
            filter_index = self.filter_combo.SelectedIndex
        
        viewtype_filter = ""
        if self.viewtype_combo and self.viewtype_combo.SelectedIndex > 0:
            selected = self.viewtype_combo.SelectedItem
            if selected:
                viewtype_filter = selected.Content
        
        self.filtered_items = []
        
        for item in self.all_items:
            if search_text and search_text not in item.name.lower():
                continue
            
            if filter_index == 1:  # In Use Only
                if item.usage_count == 0:
                    continue
            elif filter_index == 2:  # Unused Only
                if item.usage_count > 0:
                    continue
            
            if viewtype_filter and item.view_type != viewtype_filter:
                continue
            
            self.filtered_items.append(item)
        
        self.data_grid.ItemsSource = ObservableCollection[object](self.filtered_items)
    
    def _update_stats(self):
        if self.txt_total:
            self.txt_total.Text = str(len(self.all_items))
        
        if self.txt_selected:
            self.txt_selected.Text = str(self.data_grid.SelectedItems.Count)
        
        if self.txt_used:
            used = sum(1 for item in self.all_items if item.usage_count > 0)
            self.txt_used.Text = str(used)
        
        if self.txt_unused:
            unused = sum(1 for item in self.all_items if item.usage_count == 0)
            self.txt_unused.Text = str(unused)
    
    def _get_selected_items(self):
        try:
            return list(self.data_grid.SelectedItems)
        except:
            return []
    
    def _on_filter_changed(self, sender, args):
        self._apply_filters()
    
    def _on_selection_changed(self, sender, args):
        if self.txt_selected:
            self.txt_selected.Text = str(self.data_grid.SelectedItems.Count)
    
    def _on_select_all(self, sender, args):
        self.data_grid.SelectAll()
    
    def _on_clear_all(self, sender, args):
        self.data_grid.UnselectAll()
        self._update_stats()
    
    def _on_select_unused(self, sender, args):
        self.data_grid.UnselectAll()
        for item in self.filtered_items:
            if item.usage_count == 0:
                self.data_grid.SelectedItems.Add(item)
        self._update_stats()
    
    def _on_refresh(self, sender, args):
        self._load_data()
        MessageBox.Show("Data refreshed!", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
    
    def _on_rename(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select one template to rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if len(selected) > 1:
            MessageBox.Show("Please select only ONE template to rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        item = selected[0]
        
        new_name = forms.ask_for_string(
            default=item.name,
            prompt="Enter new name:",
            title="Rename View Template"
        )
        
        if not new_name or new_name == item.name:
            return
        
        # Sanitize
        invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
        for char in invalid_chars:
            new_name = new_name.replace(char, '')
        new_name = new_name.strip()
        
        if not new_name:
            return
        
        try:
            rename_template(doc, item.view_template, new_name)
            MessageBox.Show("Renamed successfully!", "Success",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            self._load_data()
            
        except Exception as ex:
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_batch_rename(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one template to batch rename!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        dialog = BatchRenameDialog(selected, self)
        dialog.ShowDialog()
    
    def _on_duplicate(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one template to duplicate!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        try:
            templates = [item.view_template for item in selected]
            success_count = duplicate_templates(doc, templates)
            
            MessageBox.Show("Duplicated: {}".format(success_count), "Result",
                          MessageBoxButton.OK, MessageBoxImage.Information)
            self._load_data()
            
        except Exception as ex:
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_delete(self, sender, args):
        selected = self._get_selected_items()
        
        if not selected:
            MessageBox.Show("Please select at least one template to delete!",
                          "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        # Check usage - only unused can be deleted
        in_use = [item for item in selected if item.usage_count > 0]
        can_delete = [item for item in selected if item.usage_count == 0]
        
        if in_use:
            msg = "WARNING: {} template(s) are IN USE and CANNOT be deleted:\n\n".format(len(in_use))
            for item in in_use[:5]:
                msg += "  - '{}': {} views\n".format(item.name, item.usage_count)
            if len(in_use) > 5:
                msg += "  ... and {} more\n".format(len(in_use) - 5)
            
            MessageBox.Show(msg, "Templates In Use",
                          MessageBoxButton.OK, MessageBoxImage.Warning)
        
        if not can_delete:
            MessageBox.Show("No deletable templates selected.\nOnly UNUSED templates can be deleted.",
                          "Info", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        result = MessageBox.Show(
            "Delete {} unused template(s)?".format(len(can_delete)),
            "Confirm Delete",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question
        )
        
        if result != MessageBoxResult.Yes:
            return
        
        try:
            templates_to_delete = [item.view_template for item in can_delete]
            success_count, error_count = delete_templates(doc, templates_to_delete)
            
            msg = "Deleted: {}".format(success_count)
            if error_count > 0:
                msg += "\nFailed: {}".format(error_count)
            
            MessageBox.Show(msg, "Result", MessageBoxButton.OK, MessageBoxImage.Information)
            self._load_data()
            
        except Exception as ex:
            MessageBox.Show("Error: {}".format(str(ex)), "Error",
                          MessageBoxButton.OK, MessageBoxImage.Error)


def show_view_template_manager():
    """Launch the View Template Manager Dialog"""
    try:
        window = ViewTemplateManagerWindow()
        window.ShowDialog()
    except Exception as e:
        print("\nFATAL ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()
        
        MessageBox.Show(
            "Error starting View Template Manager:\n\n{}".format(str(e)),
            "Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error
        )
