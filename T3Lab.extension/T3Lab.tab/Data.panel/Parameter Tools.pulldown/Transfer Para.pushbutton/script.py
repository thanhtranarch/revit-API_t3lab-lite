# -*- coding: utf-8 -*-
"""
Transfer Parameter Value v1.0 - DQT
Transfers parameter values from one parameter to another within the same elements.
Select a category, pick Source Parameter and Target Parameter, preview and apply.

Supports all StorageType: String, Double, Integer, ElementId.
Handles type conversion between different StorageTypes where possible.

Copyright (c) 2026 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Transfer\nParam Value"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Transfer parameter values between parameters within the same elements."

# =============================================================================
# IMPORTS
# =============================================================================
import clr
import sys
import codecs
import os
import traceback

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xml")

import System
from System import IO, Text, Windows
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxResult, MessageBoxImage
from System.Windows import Visibility, HorizontalAlignment, VerticalAlignment, TextAlignment
from System.Windows import Thickness, GridLength, GridUnitType, RoutedEventArgs
from System.Windows.Controls import (
    StackPanel, DockPanel, Border, TextBlock, TextBox,
    Button, ComboBox, ComboBoxItem, ScrollViewer,
    DataGrid, DataGridTextColumn, DataGridLength,
    CheckBox, Dock, Orientation, SelectionMode
)
from System.Windows.Media import BrushConverter
from System.Windows.Markup import XamlReader

from pyrevit import revit, DB, script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    StorageType, Transaction, ElementId,
    UnitUtils
)

doc = revit.doc
output = script.get_output()
logger = script.get_logger()


# =============================================================================
# REVIT API COMPATIBILITY (2024 / 2025 / 2026 / 2027)
# =============================================================================
def _eid_int(eid):
    """Get integer value from ElementId.
    Revit 2024/2025: .IntegerValue
    Revit 2026+:     .Value (IntegerValue removed)
    """
    if eid is None:
        return -1
    try:
        return eid.Value
    except AttributeError:
        try:
            return eid.IntegerValue
        except:
            return -1


def _make_eid(int_val):
    """Create ElementId from integer.
    Revit 2024/2025: ElementId(int)
    Revit 2026+:     ElementId(long)
    """
    try:
        return ElementId(int(int_val))
    except:
        try:
            return ElementId(System.Int64(int(int_val)))
        except:
            return ElementId(-1)


# =============================================================================
# CONSTANTS
# =============================================================================
BC = BrushConverter()

# DQT Brand Colors - synced with Contains Manager
CLR_HEADER_BG = "#0F172A"           # Gold header background
CLR_HEADER_FG = "#0F172A"           # Dark text on gold header
CLR_PRIMARY = "#0F172A"             # DQT Gold
CLR_ACCENT = "#3B82F6"              # Darker gold for hover/accent
CLR_BG = "#FFFFFF"                  # White main background
CLR_CARD_BORDER = "#E0E0E0"         # Light gray borders
CLR_FOOTER_BG = "#0F172A"           # Gold footer
CLR_FOOTER_FG = "#5D4E37"           # Dark brown text on footer
CLR_TEXT = "#0F172A"                 # Primary text
CLR_SUBTEXT = "#475569"             # Secondary text
CLR_WHITE = "#FFFFFF"
CLR_SUCCESS = "#10B981"
CLR_WARNING = "#F59E0B"
CLR_DANGER = "#EF4444"
CLR_BTN_PRIMARY_BG = "#0F172A"      # Gold button
CLR_BTN_PRIMARY_FG = "#0F172A"      # Dark text on gold button
CLR_BTN_SECONDARY_BG = "#FFFFFF"    # White button
CLR_BTN_SECONDARY_FG = "#0F172A"    # Dark text on white button
CLR_BTN_SECONDARY_BORDER = "#E0E0E0"

CATEGORY_MAP = [
    ("Walls", BuiltInCategory.OST_Walls),
    ("Floors", BuiltInCategory.OST_Floors),
    ("Ceilings", BuiltInCategory.OST_Ceilings),
    ("Roofs", BuiltInCategory.OST_Roofs),
    ("Doors", BuiltInCategory.OST_Doors),
    ("Windows", BuiltInCategory.OST_Windows),
    ("Rooms", BuiltInCategory.OST_Rooms),
    ("Areas", BuiltInCategory.OST_Areas),
    ("Spaces", BuiltInCategory.OST_MEPSpaces),
    ("Columns", BuiltInCategory.OST_Columns),
    ("Structural Columns", BuiltInCategory.OST_StructuralColumns),
    ("Structural Framing", BuiltInCategory.OST_StructuralFraming),
    ("Structural Foundations", BuiltInCategory.OST_StructuralFoundation),
    ("Furniture", BuiltInCategory.OST_Furniture),
    ("Casework", BuiltInCategory.OST_Casework),
    ("Generic Models", BuiltInCategory.OST_GenericModel),
    ("Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment),
    ("Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures),
    ("Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment),
    ("Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures),
    ("Lighting Fixtures", BuiltInCategory.OST_LightingFixtures),
    ("Pipe Fittings", BuiltInCategory.OST_PipeFitting),
    ("Pipe Accessories", BuiltInCategory.OST_PipeAccessory),
    ("Duct Fittings", BuiltInCategory.OST_DuctFitting),
    ("Duct Accessories", BuiltInCategory.OST_DuctAccessory),
    ("Conduit Fittings", BuiltInCategory.OST_ConduitFitting),
    ("Cable Trays", BuiltInCategory.OST_CableTray),
    ("Curtain Panels", BuiltInCategory.OST_CurtainWallPanels),
    ("Curtain Wall Mullions", BuiltInCategory.OST_CurtainWallMullions),
    ("Parking", BuiltInCategory.OST_Parking),
    ("Planting", BuiltInCategory.OST_Planting),
    ("Site", BuiltInCategory.OST_Site),
    ("Topography", BuiltInCategory.OST_Topography),
    ("Sheets", BuiltInCategory.OST_Sheets),
]


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================
def get_elements_by_category(bic):
    """Collect all elements of given BuiltInCategory"""
    try:
        collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        return list(collector)
    except Exception:
        return []


def get_instance_parameters(elements):
    """Scan elements and return sorted list of instance parameter info dicts.
    Returns list of dicts: {name, storage_type, is_read_only, definition}
    """
    param_dict = {}
    for elem in elements:
        try:
            for param in elem.Parameters:
                if param.Definition is None:
                    continue
                name = param.Definition.Name
                if name not in param_dict:
                    param_dict[name] = {
                        'name': name,
                        'storage_type': param.StorageType,
                        'is_read_only': param.IsReadOnly,
                    }
        except Exception:
            continue
    
    result = sorted(param_dict.values(), key=lambda x: x['name'].lower())
    return result


def get_param_value_as_string(elem, param_name):
    """Get parameter value as display string"""
    try:
        param = elem.LookupParameter(param_name)
        if param is None:
            return "<not found>"
        if not param.HasValue:
            return "<empty>"
        
        st = param.StorageType
        if st == StorageType.String:
            v = param.AsString()
            return v if v else "<empty>"
        elif st == StorageType.Integer:
            return str(param.AsInteger())
        elif st == StorageType.Double:
            # Try to display with unit formatting
            try:
                return param.AsValueString() or str(param.AsDouble())
            except Exception:
                return str(param.AsDouble())
        elif st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid and _eid_int(eid) != -1:
                ref_elem = doc.GetElement(eid)
                if ref_elem and hasattr(ref_elem, 'Name'):
                    return ref_elem.Name
                return str(_eid_int(eid))
            return "<None>"
        return "<unknown>"
    except Exception:
        return "<error>"


def get_element_display_name(elem):
    """Get a readable name for element"""
    try:
        # Try Family + Type name
        type_id = elem.GetTypeId()
        if type_id and _eid_int(type_id) != -1:
            elem_type = doc.GetElement(type_id)
            if elem_type:
                fam_name = ""
                try:
                    fam_name = elem_type.FamilyName
                except Exception:
                    pass
                type_name = getattr(elem_type, 'Name', '')
                if fam_name and type_name:
                    return "{}: {}".format(fam_name, type_name)
                elif type_name:
                    return type_name
        
        # Fallback
        name_param = elem.LookupParameter("Name")
        if name_param and name_param.HasValue:
            return name_param.AsString()
        
        return "ID: {}".format(_eid_int(elem.Id))
    except Exception:
        return "ID: {}".format(_eid_int(elem.Id))


def transfer_value(elem, source_param_name, target_param_name):
    """Transfer value from source parameter to target parameter on same element.
    Handles StorageType conversion where possible.
    Returns (success, message)
    """
    try:
        src_param = elem.LookupParameter(source_param_name)
        tgt_param = elem.LookupParameter(target_param_name)
        
        if src_param is None:
            return (False, "Source param not found")
        if tgt_param is None:
            return (False, "Target param not found")
        if tgt_param.IsReadOnly:
            return (False, "Target is read-only")
        if not src_param.HasValue:
            return (False, "Source has no value")
        
        src_st = src_param.StorageType
        tgt_st = tgt_param.StorageType
        
        # Same StorageType - direct copy
        if src_st == tgt_st:
            if src_st == StorageType.String:
                tgt_param.Set(src_param.AsString() or "")
            elif src_st == StorageType.Integer:
                tgt_param.Set(src_param.AsInteger())
            elif src_st == StorageType.Double:
                tgt_param.Set(src_param.AsDouble())
            elif src_st == StorageType.ElementId:
                tgt_param.Set(src_param.AsElementId())
            return (True, "OK")
        
        # Different StorageType - convert
        # Source -> get as appropriate type -> set to target
        
        if tgt_st == StorageType.String:
            # Anything -> String
            val = get_param_value_as_string(elem, source_param_name)
            if val.startswith("<"):
                val = ""
            tgt_param.Set(val)
            return (True, "Converted to String")
        
        # Get source as string first
        src_str = ""
        if src_st == StorageType.String:
            src_str = src_param.AsString() or ""
        elif src_st == StorageType.Integer:
            src_str = str(src_param.AsInteger())
        elif src_st == StorageType.Double:
            src_str = str(src_param.AsDouble())
        elif src_st == StorageType.ElementId:
            src_str = str(_eid_int(src_param.AsElementId()))
        
        if tgt_st == StorageType.Integer:
            try:
                tgt_param.Set(int(float(src_str)))
                return (True, "Converted to Integer")
            except Exception:
                return (False, "Cannot convert '{}' to Integer".format(src_str))
        
        elif tgt_st == StorageType.Double:
            try:
                tgt_param.Set(float(src_str))
                return (True, "Converted to Double")
            except Exception:
                return (False, "Cannot convert '{}' to Double".format(src_str))
        
        elif tgt_st == StorageType.ElementId:
            try:
                eid = _make_eid(int(float(src_str)))
                tgt_param.Set(eid)
                return (True, "Converted to ElementId")
            except Exception:
                return (False, "Cannot convert to ElementId")
        
        return (False, "Unsupported conversion")
    
    except Exception as e:
        return (False, str(e))


def storage_type_label(st):
    """Return short label for StorageType"""
    if st == StorageType.String:
        return "Text"
    elif st == StorageType.Integer:
        return "Int"
    elif st == StorageType.Double:
        return "Number"
    elif st == StorageType.ElementId:
        return "Element"
    return "?"


# =============================================================================
# XAML WINDOW
# =============================================================================
def _style_button_primary(btn):
    """Apply DQT primary button style - gold bg, dark text (matches Contains Manager)"""
    btn.Background = BC.ConvertFromString(CLR_BTN_PRIMARY_BG)
    btn.Foreground = BC.ConvertFromString(CLR_BTN_PRIMARY_FG)
    btn.FontWeight = System.Windows.FontWeights.SemiBold
    btn.Padding = Thickness(16, 8, 16, 8)
    btn.BorderBrush = BC.ConvertFromString(CLR_ACCENT)
    btn.BorderThickness = Thickness(1)
    btn.Cursor = System.Windows.Input.Cursors.Hand
    btn.FontSize = 12


def _style_button_secondary(btn):
    """Apply DQT secondary button style - white bg, dark text (matches Contains Manager)"""
    btn.Background = BC.ConvertFromString(CLR_BTN_SECONDARY_BG)
    btn.Foreground = BC.ConvertFromString(CLR_BTN_SECONDARY_FG)
    btn.Padding = Thickness(12, 6, 12, 6)
    btn.BorderBrush = BC.ConvertFromString(CLR_BTN_SECONDARY_BORDER)
    btn.BorderThickness = Thickness(1)
    btn.Cursor = System.Windows.Input.Cursors.Hand
    btn.FontSize = 11


# =============================================================================
# DATA ROW CLASS
# =============================================================================
class PreviewRow(object):
    """Data row for DataGrid preview"""
    def __init__(self, element_name, elem_id, source_val, target_val, status=""):
        self.Element = element_name
        self.ID = str(elem_id)
        self.SourceValue = source_val
        self.TargetValue = target_val
        self.Status = status


# =============================================================================
# PARAM ITEM (Radio-like selection in StackPanel)
# =============================================================================
class ParamItem(object):
    """A selectable parameter item displayed as a row in the list"""
    def __init__(self, param_info, on_click_callback, panel_type):
        self.param_info = param_info
        self.name = param_info['name']
        self.storage_type = param_info['storage_type']
        self.is_read_only = param_info['is_read_only']
        self.panel_type = panel_type  # 'source' or 'target'
        self.is_selected = False
        self._callback = on_click_callback
        
        # Build UI
        self.border = Border()
        self.border.Padding = Thickness(8, 5, 8, 5)
        self.border.Margin = Thickness(1)
        self.border.Background = BC.ConvertFromString(CLR_WHITE)
        self.border.Cursor = System.Windows.Input.Cursors.Hand
        self.border.BorderThickness = Thickness(1)
        self.border.BorderBrush = BC.ConvertFromString("Transparent")
        
        grid = System.Windows.Controls.Grid()
        col0 = System.Windows.Controls.ColumnDefinition()
        col0.Width = GridLength(1, GridUnitType.Star)
        col1 = System.Windows.Controls.ColumnDefinition()
        col1.Width = GridLength(55, GridUnitType.Pixel)
        col2 = System.Windows.Controls.ColumnDefinition()
        col2.Width = GridLength(65, GridUnitType.Pixel)
        grid.ColumnDefinitions.Add(col0)
        grid.ColumnDefinitions.Add(col1)
        grid.ColumnDefinitions.Add(col2)
        
        # Name
        self.txt_name = TextBlock()
        self.txt_name.Text = self.name
        self.txt_name.FontSize = 11.5
        self.txt_name.VerticalAlignment = VerticalAlignment.Center
        self.txt_name.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        self.txt_name.Foreground = BC.ConvertFromString(CLR_TEXT)
        System.Windows.Controls.Grid.SetColumn(self.txt_name, 0)
        grid.Children.Add(self.txt_name)
        
        # StorageType badge
        type_lbl = storage_type_label(self.storage_type)
        self.txt_type = TextBlock()
        self.txt_type.Text = type_lbl
        self.txt_type.FontSize = 9.5
        self.txt_type.Foreground = BC.ConvertFromString(CLR_SUBTEXT)
        self.txt_type.VerticalAlignment = VerticalAlignment.Center
        self.txt_type.HorizontalAlignment = HorizontalAlignment.Center
        System.Windows.Controls.Grid.SetColumn(self.txt_type, 1)
        grid.Children.Add(self.txt_type)
        
        # Read-only badge (for target panel only)
        if self.panel_type == 'target' and self.is_read_only:
            ro_txt = TextBlock()
            ro_txt.Text = "RO"
            ro_txt.FontSize = 9
            ro_txt.Foreground = BC.ConvertFromString(CLR_DANGER)
            ro_txt.FontWeight = System.Windows.FontWeights.SemiBold
            ro_txt.VerticalAlignment = VerticalAlignment.Center
            ro_txt.HorizontalAlignment = HorizontalAlignment.Center
            System.Windows.Controls.Grid.SetColumn(ro_txt, 2)
            grid.Children.Add(ro_txt)
        
        self.border.Child = grid
        self.border.MouseLeftButtonUp += self._on_click
    
    def _on_click(self, sender, args):
        if self.panel_type == 'target' and self.is_read_only:
            return  # Cannot select read-only as target
        self._callback(self)
    
    def set_selected(self, selected):
        self.is_selected = selected
        if selected:
            self.border.Background = BC.ConvertFromString(CLR_PRIMARY)
            self.border.BorderBrush = BC.ConvertFromString(CLR_ACCENT)
            self.txt_name.FontWeight = System.Windows.FontWeights.SemiBold
        else:
            self.border.Background = BC.ConvertFromString(CLR_WHITE)
            self.border.BorderBrush = BC.ConvertFromString("Transparent")
            self.txt_name.FontWeight = System.Windows.FontWeights.Normal
    
    def matches_search(self, keyword):
        if not keyword:
            return True
        return keyword.lower() in self.name.lower()


# =============================================================================
# MAIN WINDOW CONTROLLER
# =============================================================================
class TransferParamWindow(object):
    def __init__(self):
        # Load XAML - use safe %% delimiters
        # Find T3Lab.extension parent folder dynamically
        current_dir = os.path.dirname(__file__)
        while current_dir and not current_dir.endswith('T3Lab.extension'):
            parent = os.path.dirname(current_dir)
            if parent == current_dir:
                break
            current_dir = parent
        xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "TransferPara.xaml")
        
        with codecs.open(xaml_path, "r", "utf-8") as f:
            xaml_content = f.read()
        stream = IO.MemoryStream(Text.Encoding.UTF8.GetBytes(xaml_content))
        self.window = XamlReader.Load(stream)
        
        # Get controls
        self.cmb_category = self.window.FindName("cmbCategory")
        self.txt_count = self.window.FindName("txtCount")
        
        self.txt_search_source = self.window.FindName("txtSearchSource")
        self.txt_search_target = self.window.FindName("txtSearchTarget")
        self.pnl_source = self.window.FindName("pnlSource")
        self.pnl_target = self.window.FindName("pnlTarget")
        self.txt_source_info = self.window.FindName("txtSourceInfo")
        self.txt_target_info = self.window.FindName("txtTargetInfo")
        
        self.dg_preview = self.window.FindName("dgPreview")
        self.btn_preview = self.window.FindName("btnPreview")
        self.txt_preview_status = self.window.FindName("txtPreviewStatus")
        
        self.btn_transfer = self.window.FindName("btnTransfer")
        self.btn_close = self.window.FindName("btnClose")
        self.txt_summary = self.window.FindName("txtSummary")
        
        
        
        # Build DataGrid columns via code-behind (avoids XAML Binding curly-brace issues)
        self._setup_datagrid_columns()
        
        # State
        self.elements = []
        self.all_params = []
        self.source_items = []
        self.target_items = []
        self.selected_source = None
        self.selected_target = None
        
        # Populate categories
        for name, bic in CATEGORY_MAP:
            item = ComboBoxItem()
            item.Content = name
            item.Tag = bic
            self.cmb_category.Items.Add(item)
        
        if self.cmb_category.Items.Count > 0:
            self.cmb_category.SelectedIndex = 0
        
        # Events
        self.cmb_category.SelectionChanged += self._on_category_changed
        self.txt_search_source.TextChanged += self._on_search_source
        self.txt_search_target.TextChanged += self._on_search_target
        self.btn_preview.Click += self._on_preview
        self.btn_transfer.Click += self._on_transfer
        self.btn_close.Click += self._on_close
        
        # Initial load
        self._on_category_changed(None, None)
    
    def show(self):
        self.window.ShowDialog()
    
    def _setup_datagrid_columns(self):
        """Add DataGrid columns via code-behind to avoid XAML Binding issues"""
        from System.Windows.Data import Binding as WPFBinding
        
        self.dg_preview.Columns.Clear()
        
        col_defs = [
            ("Element", "Element", 220),
            ("ID", "ID", 80),
            ("Source Value", "SourceValue", 200),
            ("->", None, 35),
            ("Current Target Value", "TargetValue", 200),
            ("Status", "Status", 100),
        ]
        
        for header, binding_path, width in col_defs:
            col = DataGridTextColumn()
            col.Header = header
            col.Width = DataGridLength(width)
            col.IsReadOnly = True
            if binding_path:
                col.Binding = WPFBinding(binding_path)
            self.dg_preview.Columns.Add(col)
    
    # -------------------------------------------------------------------------
    # CATEGORY CHANGED
    # -------------------------------------------------------------------------
    def _on_category_changed(self, sender, args):
        selected = self.cmb_category.SelectedItem
        if selected is None:
            return
        
        bic = selected.Tag
        self.elements = get_elements_by_category(bic)
        count = len(self.elements)
        self.txt_count.Text = "{} element(s)".format(count)
        
        # Clear selections
        self.selected_source = None
        self.selected_target = None
        self.txt_source_info.Text = "Select a source parameter"
        self.txt_target_info.Text = "Select a target parameter"
        self.dg_preview.ItemsSource = None
        self.txt_preview_status.Text = ""
        self.txt_summary.Text = ""
        
        # Get parameters
        if count > 0:
            self.all_params = get_instance_parameters(self.elements)
        else:
            self.all_params = []
        
        # Build param lists
        self._build_param_lists()
    
    def _build_param_lists(self):
        """Build source and target param item lists"""
        self.source_items = []
        self.target_items = []
        self.pnl_source.Children.Clear()
        self.pnl_target.Children.Clear()
        
        for p in self.all_params:
            # Source item
            si = ParamItem(p, self._on_source_selected, 'source')
            self.source_items.append(si)
            self.pnl_source.Children.Add(si.border)
            
            # Target item
            ti = ParamItem(p, self._on_target_selected, 'target')
            self.target_items.append(ti)
            self.pnl_target.Children.Add(ti.border)
    
    # -------------------------------------------------------------------------
    # PARAM SELECTION
    # -------------------------------------------------------------------------
    def _on_source_selected(self, item):
        # Deselect previous
        if self.selected_source:
            self.selected_source.set_selected(False)
        item.set_selected(True)
        self.selected_source = item
        self.txt_source_info.Text = "Source: {} ({})".format(
            item.name, storage_type_label(item.storage_type))
        self._clear_preview()
    
    def _on_target_selected(self, item):
        if item.is_read_only:
            return
        if self.selected_target:
            self.selected_target.set_selected(False)
        item.set_selected(True)
        self.selected_target = item
        self.txt_target_info.Text = "Target: {} ({}{})".format(
            item.name,
            storage_type_label(item.storage_type),
            " - Read Only!" if item.is_read_only else "")
        self._clear_preview()
    
    # -------------------------------------------------------------------------
    # SEARCH / FILTER
    # -------------------------------------------------------------------------
    def _on_search_source(self, sender, args):
        keyword = self.txt_search_source.Text.strip()
        for item in self.source_items:
            if item.matches_search(keyword):
                item.border.Visibility = Visibility.Visible
            else:
                item.border.Visibility = Visibility.Collapsed
    
    def _on_search_target(self, sender, args):
        keyword = self.txt_search_target.Text.strip()
        for item in self.target_items:
            if item.matches_search(keyword):
                item.border.Visibility = Visibility.Visible
            else:
                item.border.Visibility = Visibility.Collapsed
    
    # -------------------------------------------------------------------------
    # PREVIEW
    # -------------------------------------------------------------------------
    def _clear_preview(self):
        self.dg_preview.ItemsSource = None
        self.txt_preview_status.Text = ""
        self.txt_summary.Text = ""
    
    def _on_preview(self, sender, args):
        if not self.selected_source or not self.selected_target:
            MessageBox.Show(
                "Please select both Source and Target parameters.",
                "Missing Selection",
                MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if self.selected_source.name == self.selected_target.name:
            MessageBox.Show(
                "Source and Target must be different parameters.",
                "Same Parameter",
                MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        src_name = self.selected_source.name
        tgt_name = self.selected_target.name
        
        rows = []
        transferable = 0
        skipped = 0
        
        for elem in self.elements:
            elem_name = get_element_display_name(elem)
            elem_id = _eid_int(elem.Id)
            src_val = get_param_value_as_string(elem, src_name)
            tgt_val = get_param_value_as_string(elem, tgt_name)
            
            # Determine status
            src_param = elem.LookupParameter(src_name)
            tgt_param = elem.LookupParameter(tgt_name)
            
            status = ""
            if src_param is None:
                status = "No source"
                skipped += 1
            elif tgt_param is None:
                status = "No target"
                skipped += 1
            elif tgt_param.IsReadOnly:
                status = "Read-only"
                skipped += 1
            elif not src_param.HasValue:
                status = "Empty source"
                skipped += 1
            elif src_val == tgt_val:
                status = "Same value"
                transferable += 1  # Still count, will just overwrite same
            else:
                status = "Ready"
                transferable += 1
            
            rows.append(PreviewRow(elem_name, elem_id, src_val, tgt_val, status))
        
        self.dg_preview.ItemsSource = rows
        self.txt_preview_status.Text = "{} element(s) previewed".format(len(rows))
        self.txt_summary.Text = "Ready: {} | Skipped: {} | Total: {}".format(
            transferable, skipped, len(rows))
    
    # -------------------------------------------------------------------------
    # TRANSFER
    # -------------------------------------------------------------------------
    def _on_transfer(self, sender, args):
        if not self.selected_source or not self.selected_target:
            MessageBox.Show(
                "Please select both Source and Target parameters.",
                "Missing Selection",
                MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        if self.selected_source.name == self.selected_target.name:
            MessageBox.Show(
                "Source and Target must be different parameters.",
                "Same Parameter",
                MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        src_name = self.selected_source.name
        tgt_name = self.selected_target.name
        
        # Confirm
        result = MessageBox.Show(
            "Transfer values:\n\n"
            "  Source: {}\n"
            "  Target: {}\n"
            "  Elements: {}\n\n"
            "Proceed?".format(src_name, tgt_name, len(self.elements)),
            "Confirm Transfer",
            MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        # Execute
        success = 0
        fail = 0
        skip = 0
        errors = []
        
        t = Transaction(doc, "DQT - Transfer Parameter Value")
        t.Start()
        
        try:
            for elem in self.elements:
                src_param = elem.LookupParameter(src_name)
                tgt_param = elem.LookupParameter(tgt_name)
                
                if src_param is None or tgt_param is None:
                    skip += 1
                    continue
                if tgt_param.IsReadOnly:
                    skip += 1
                    continue
                if not src_param.HasValue:
                    skip += 1
                    continue
                
                ok, msg = transfer_value(elem, src_name, tgt_name)
                if ok:
                    success += 1
                else:
                    fail += 1
                    if len(errors) < 10:
                        errors.append("ID {}: {}".format(
                            _eid_int(elem.Id), msg))
            
            t.Commit()
        except Exception as e:
            t.RollBack()
            MessageBox.Show(
                "Transaction failed:\n{}".format(str(e)),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            return
        
        # Result message
        msg_lines = [
            "Transfer Complete!",
            "",
            "Success: {}".format(success),
            "Failed: {}".format(fail),
            "Skipped: {}".format(skip),
        ]
        if errors:
            msg_lines.append("")
            msg_lines.append("Errors (first 10):")
            for err in errors:
                msg_lines.append("  - {}".format(err))
        
        MessageBox.Show(
            "\n".join(msg_lines),
            "Results",
            MessageBoxButton.OK, MessageBoxImage.Information)
        
        # Log to output
        output.print_md("# Transfer Parameter Value - Results")
        output.print_md("**Source:** {}".format(src_name))
        output.print_md("**Target:** {}".format(tgt_name))
        output.print_md("**Success:** {} | **Failed:** {} | **Skipped:** {}".format(
            success, fail, skip))
        
        # Refresh preview
        self._on_preview(None, None)
    
    # -------------------------------------------------------------------------
    # CLOSE
    # -------------------------------------------------------------------------
    def _on_close(self, sender, args):
        self.window.Close()


# =============================================================================
# ENTRY POINT
# =============================================================================
def main():
    try:
        win = TransferParamWindow()
        win.show()
    except Exception as e:
        logger.error("Transfer Parameter Value error: {}".format(str(e)))
        traceback.print_exc()
        forms.alert("Error:\n{}".format(str(e)), title="Error")


if __name__ == "__main__":
    main()