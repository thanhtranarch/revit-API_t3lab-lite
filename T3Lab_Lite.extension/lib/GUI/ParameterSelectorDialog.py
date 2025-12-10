# -*- coding: utf-8 -*-
"""Parameter Selector Dialog for Custom Filenames.

This dialog allows users to select from all available sheet/view parameters
to build custom filename patterns.
"""

import os
import sys
import clr
from collections import OrderedDict

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System')

from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.Windows.Controls import ListBox
from System.Collections.ObjectModel import ObservableCollection
from System import Uri, UriKind

from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    BuiltInParameter, ViewSheet, View,
    ParameterType, StorageType
)


class ParameterItem:
    """Represents a parameter item that can be selected."""

    def __init__(self, name, display_name, is_builtin=True, builtin_param=None):
        self.Name = name
        self.DisplayName = display_name
        self.IsBuiltIn = is_builtin
        self.BuiltInParameter = builtin_param

    def __str__(self):
        return self.DisplayName

    def __repr__(self):
        return "ParameterItem({})".format(self.DisplayName)


class ParameterSelectorDialog(Window):
    """Dialog for selecting parameters to build custom filenames."""

    def __init__(self, doc, element_type='sheet'):
        """Initialize the parameter selector dialog.

        Args:
            doc: Revit document
            element_type: 'sheet' or 'view' - determines which parameters to load
        """
        self.doc = doc
        self.element_type = element_type
        self.selected_result = None
        self.field_separator = '-'

        # Load XAML
        xaml_file = os.path.join(os.path.dirname(__file__), 'ParameterSelector.xaml')
        try:
            with open(xaml_file, 'r') as f:
                xaml_content = f.read()
            self.ui = XamlReader.Parse(xaml_content)
        except Exception as e:
            forms.alert("Error loading XAML: {}".format(str(e)))
            return

        # Set up window
        self.Title = self.ui.Title
        self.Width = self.ui.Width
        self.Height = self.ui.Height
        self.Content = self.ui.Content
        self.WindowStartupLocation = self.ui.WindowStartupLocation

        # Get controls
        self.list_available = self.ui.FindName('list_available')
        self.list_selected = self.ui.FindName('list_selected')
        self.txt_field_separator = self.ui.FindName('txt_field_separator')
        self.chk_field_separator = self.ui.FindName('chk_field_separator')
        self.chk_include_project_params = self.ui.FindName('chk_include_project_params')
        self.txt_custom_field = self.ui.FindName('txt_custom_field')
        self.txt_custom_separator = self.ui.FindName('txt_custom_separator')

        # Set up event handlers
        self.ui.FindName('button_close').Click += self.button_close
        self.ui.FindName('button_add_parameter').Click += self.button_add_parameter
        self.ui.FindName('button_remove_parameter').Click += self.button_remove_parameter
        self.ui.FindName('button_move_up').Click += self.button_move_up
        self.ui.FindName('button_move_down').Click += self.button_move_down
        self.ui.FindName('button_refresh').Click += self.button_refresh
        self.ui.FindName('button_add_custom_field').Click += self.button_add_custom_field
        self.ui.FindName('button_add_custom_separator').Click += self.button_add_custom_separator
        self.ui.FindName('button_ok').Click += self.button_ok
        self.ui.FindName('button_cancel').Click += self.button_cancel
        self.ui.FindName('button_preview').Click += self.button_preview

        # Set up drag handler for header
        self.ui.MouseDown += self.header_drag

        # Initialize parameter collections
        self.available_params = ObservableCollection[ParameterItem]()
        self.selected_params = ObservableCollection[ParameterItem]()

        self.list_available.ItemsSource = self.available_params
        self.list_selected.ItemsSource = self.selected_params

        # Load parameters
        self.load_parameters()

    def load_parameters(self):
        """Load all available parameters from sheets or views."""
        self.available_params.Clear()

        # Get a sample element to extract parameters from
        sample_element = self._get_sample_element()
        if not sample_element:
            forms.alert("No {} found in the project.".format(self.element_type))
            return

        # Collect all parameters
        all_params = OrderedDict()

        # Add built-in parameters specific to sheets/views
        if self.element_type == 'sheet':
            self._add_sheet_builtin_params(all_params)
        else:
            self._add_view_builtin_params(all_params)

        # Add project information parameters if checkbox is checked
        if self.chk_include_project_params.IsChecked:
            self._add_project_info_params(all_params)

        # Extract all parameters from the sample element
        for param in sample_element.Parameters:
            try:
                param_name = param.Definition.Name
                if param_name not in all_params:
                    # Check if it's a built-in parameter
                    is_builtin = False
                    builtin_param = None
                    try:
                        builtin_param = param.Definition.BuiltInParameter
                        if builtin_param != BuiltInParameter.INVALID:
                            is_builtin = True
                    except:
                        pass

                    all_params[param_name] = ParameterItem(
                        param_name,
                        param_name,
                        is_builtin,
                        builtin_param
                    )
            except:
                continue

        # Add standard date/time placeholders
        all_params['Date'] = ParameterItem('Date', '{Date} - Current Date', False)
        all_params['Time'] = ParameterItem('Time', '{Time} - Current Time', False)

        # Add to available list
        for param in all_params.values():
            self.available_params.Add(param)

    def _get_sample_element(self):
        """Get a sample sheet or view to extract parameters from."""
        if self.element_type == 'sheet':
            sheets = DB.FilteredElementCollector(self.doc)\
                      .OfClass(DB.ViewSheet)\
                      .WhereElementIsNotElementType()\
                      .ToElements()
            return sheets[0] if sheets else None
        else:
            views = DB.FilteredElementCollector(self.doc)\
                     .OfClass(DB.View)\
                     .WhereElementIsNotElementType()\
                     .ToElements()
            for view in views:
                if not view.IsTemplate:
                    return view
            return None

    def _add_sheet_builtin_params(self, all_params):
        """Add sheet-specific built-in parameters."""
        sheet_params = [
            (BuiltInParameter.SHEET_NUMBER, 'Sheet Number', '{SheetNumber}'),
            (BuiltInParameter.SHEET_NAME, 'Sheet Name', '{SheetName}'),
            (BuiltInParameter.SHEET_CURRENT_REVISION, 'Current Revision', '{Revision}'),
            (BuiltInParameter.SHEET_CURRENT_REVISION_DATE, 'Revision Date', '{RevisionDate}'),
            (BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION, 'Revision Description', '{RevisionDescription}'),
            (BuiltInParameter.SHEET_DRAWN_BY, 'Drawn By', '{DrawnBy}'),
            (BuiltInParameter.SHEET_CHECKED_BY, 'Checked By', '{CheckedBy}'),
            (BuiltInParameter.SHEET_ISSUE_DATE, 'Issue Date', '{IssueDate}'),
            (BuiltInParameter.SHEET_APPROVED_BY, 'Approved By', '{ApprovedBy}'),
        ]

        for builtin_param, name, display_name in sheet_params:
            all_params[name] = ParameterItem(name, display_name, True, builtin_param)

    def _add_view_builtin_params(self, all_params):
        """Add view-specific built-in parameters."""
        view_params = [
            (BuiltInParameter.VIEW_NAME, 'View Name', '{ViewName}'),
            (BuiltInParameter.VIEW_TYPE, 'View Type', '{ViewType}'),
            (BuiltInParameter.VIEW_SCALE, 'View Scale', '{ViewScale}'),
            (BuiltInParameter.VIEW_PHASE, 'Phase', '{Phase}'),
            (BuiltInParameter.VIEW_LEVEL, 'Associated Level', '{Level}'),
            (BuiltInParameter.VIEW_DISCIPLINE, 'Discipline', '{Discipline}'),
        ]

        for builtin_param, name, display_name in view_params:
            all_params[name] = ParameterItem(name, display_name, True, builtin_param)

    def _add_project_info_params(self, all_params):
        """Add project information parameters."""
        project_params = [
            ('ProjectNumber', '{ProjectNumber} - Project Number', False),
            ('ProjectName', '{ProjectName} - Project Name', False),
            ('ProjectAddress', '{ProjectAddress} - Project Address', False),
            ('ClientName', '{ClientName} - Client Name', False),
            ('ProjectStatus', '{ProjectStatus} - Project Status', False),
        ]

        for name, display_name, _ in project_params:
            all_params[name] = ParameterItem(name, display_name, False)

    def toggle_project_params(self, sender, e):
        """Toggle inclusion of project information parameters."""
        self.load_parameters()

    def button_add_parameter(self, sender, e):
        """Add selected parameters from available to selected list."""
        selected_items = list(self.list_available.SelectedItems)
        for item in selected_items:
            self.selected_params.Add(item)
            self.available_params.Remove(item)

    def button_remove_parameter(self, sender, e):
        """Remove selected parameters from selected list back to available."""
        selected_items = list(self.list_selected.SelectedItems)
        for item in selected_items:
            self.available_params.Add(item)
            self.selected_params.Remove(item)

    def button_move_up(self, sender, e):
        """Move selected parameter up in the list."""
        if self.list_selected.SelectedIndex > 0:
            index = self.list_selected.SelectedIndex
            item = self.selected_params[index]
            self.selected_params.RemoveAt(index)
            self.selected_params.Insert(index - 1, item)
            self.list_selected.SelectedIndex = index - 1

    def button_move_down(self, sender, e):
        """Move selected parameter down in the list."""
        if self.list_selected.SelectedIndex < len(self.selected_params) - 1:
            index = self.list_selected.SelectedIndex
            item = self.selected_params[index]
            self.selected_params.RemoveAt(index)
            self.selected_params.Insert(index + 1, item)
            self.list_selected.SelectedIndex = index + 1

    def button_refresh(self, sender, e):
        """Refresh the parameter list."""
        # Save currently selected parameters
        selected = list(self.selected_params)

        # Reload available parameters
        self.load_parameters()

        # Restore selected parameters
        for param in selected:
            # Try to find the parameter in available list
            for available_param in list(self.available_params):
                if available_param.Name == param.Name:
                    self.available_params.Remove(available_param)
                    self.selected_params.Add(available_param)
                    break

    def button_add_custom_field(self, sender, e):
        """Add a custom field to the selected parameters."""
        custom_field = self.txt_custom_field.Text.strip()
        if custom_field:
            custom_param = ParameterItem(
                'Custom_{}'.format(custom_field),
                '{{{}}}'.format(custom_field),
                False
            )
            self.selected_params.Add(custom_param)
            self.txt_custom_field.Text = ''

    def button_add_custom_separator(self, sender, e):
        """Add a custom separator to the selected parameters."""
        custom_sep = self.txt_custom_separator.Text.strip()
        if custom_sep:
            sep_param = ParameterItem(
                'Separator_{}'.format(custom_sep),
                custom_sep,
                False
            )
            self.selected_params.Add(sep_param)
            self.txt_custom_separator.Text = ''

    def button_preview(self, sender, e):
        """Preview the filename pattern with current parameters."""
        pattern = self.build_pattern()
        forms.alert(
            "Filename Pattern:\n\n{}".format(pattern),
            title="Preview Filename Pattern"
        )

    def button_ok(self, sender, e):
        """OK button - build pattern and close dialog."""
        self.field_separator = self.txt_field_separator.Text if self.chk_field_separator.IsChecked else ''
        self.selected_result = self.build_pattern()
        self.DialogResult = True
        self.Close()

    def button_cancel(self, sender, e):
        """Cancel button - close dialog without saving."""
        self.selected_result = None
        self.DialogResult = False
        self.Close()

    def button_close(self, sender, e):
        """Close button - same as cancel."""
        self.button_cancel(sender, e)

    def header_drag(self, sender, e):
        """Allow dragging the window by its header."""
        try:
            if e.ChangedButton == System.Windows.Input.MouseButton.Left:
                self.DragMove()
        except:
            pass

    def build_pattern(self):
        """Build filename pattern from selected parameters.

        Returns:
            String pattern with placeholders like {SheetNumber}_{SheetName}
        """
        if len(self.selected_params) == 0:
            return ''

        pattern_parts = []
        separator = self.txt_field_separator.Text if self.chk_field_separator.IsChecked else ''

        for param in self.selected_params:
            # If it's a separator item, add it directly
            if param.Name.startswith('Separator_'):
                pattern_parts.append(param.DisplayName)
            # If it's a custom field, add it as-is
            elif param.Name.startswith('Custom_'):
                pattern_parts.append(param.DisplayName)
            # For standard parameters, extract the placeholder name
            else:
                # Extract placeholder from display name (e.g., "{SheetNumber} - Sheet Number" -> "{SheetNumber}")
                display = param.DisplayName
                if display.startswith('{') and '}' in display:
                    placeholder = display.split('}')[0] + '}'
                    pattern_parts.append(placeholder)
                else:
                    # If no placeholder format, create one from the name
                    pattern_parts.append('{{{}}}'.format(param.Name))

        # Join with separator
        if separator and self.chk_field_separator.IsChecked:
            return separator.join(pattern_parts)
        else:
            return '_'.join(pattern_parts)

    @staticmethod
    def show_dialog(doc, element_type='sheet'):
        """Show the parameter selector dialog.

        Args:
            doc: Revit document
            element_type: 'sheet' or 'view'

        Returns:
            String pattern if OK was clicked, None if cancelled
        """
        try:
            dialog = ParameterSelectorDialog(doc, element_type)
            result = dialog.ShowDialog()
            if result:
                return dialog.selected_result
            return None
        except Exception as e:
            forms.alert("Error showing dialog: {}".format(str(e)))
            return None
