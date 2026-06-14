# -*- coding: utf-8 -*-
"""
IFC-SG Parameter Loader v1.0 - DQT
Automatically adds required IFC+SG parameters to Revit model categories.
Reads parameter requirements from Autodesk Model Checker XML or Excel mapping.
Creates project parameters (Instance) bound to correct categories.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "IFC-SG\nLoader"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Add required IFC+SG parameters to model. Import from XML or Excel."

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Xml')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import System
from System.Windows import *, WindowState
from System.Windows.Controls import *
from System.Windows.Media import *
from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader
import os
import sys
import json
import codecs
import traceback
import datetime

# Find T3Lab.extension parent folder dynamically
current_dir = os.path.dirname(__file__)
while current_dir and not current_dir.endswith('T3Lab.extension'):
    parent = os.path.dirname(current_dir)
    if parent == current_dir:
        break
    current_dir = parent

MAPPER_XAML_PATH = os.path.join(current_dir, "lib", "GUI", "Tools", "ParameterLoaderMapper.xaml")
CATPICKER_XAML_PATH = os.path.join(current_dir, "lib", "GUI", "Tools", "ParameterLoaderCatPicker.xaml")
MAIN_XAML_PATH = os.path.join(current_dir, "lib", "GUI", "Tools", "ParameterLoaderMain.xaml")

# =====================================================================
# REVIT API COMPATIBILITY (2024/2025/2026+)
# =====================================================================
def _eid_int(eid):
    """Get integer value from ElementId - compatible with Revit 2024-2026+"""
    try:
        return eid.Value  # Revit 2026+
    except:
        return eid.IntegerValue  # Revit 2024/2025

def _get_group_type_id(pg_key):
    """Get ForgeTypeId for parameter group - compatible with Revit 2024-2026+
    pg_key: e.g. 'PG_IFC', 'PG_GEOMETRY', 'PG_FIRE_PROTECTION'
    """
    # Revit 2026+ uses GroupTypeId
    group_map = {
        "PG_IFC": "Ifc",
        "PG_GEOMETRY": "Geometry",
        "PG_FIRE_PROTECTION": "FireProtection",
        "PG_MATERIALS": "Materials",
        "PG_IDENTITY_DATA": "IdentityData",
        "PG_STRUCTURAL": "Structural",
        "PG_MECHANICAL": "Mechanical",
        "PG_CONSTRUCTION": "Construction",
        "PG_PLUMBING": "Plumbing",
        "PG_ELECTRICAL": "Electrical",
        "PG_PHASING": "Phasing",
        "PG_GENERAL": "General",
        "PG_DATA": "Data",
    }
    
    # Try GroupTypeId first (Revit 2022+, required in 2026)
    try:
        from Autodesk.Revit.DB import GroupTypeId
        attr_name = group_map.get(pg_key, "Ifc")
        return getattr(GroupTypeId, attr_name)
    except:
        pass
    
    # Fallback to BuiltInParameterGroup (Revit 2024/2025)
    try:
        return getattr(BuiltInParameterGroup, pg_key, BuiltInParameterGroup.PG_IFC)
    except:
        pass
    
    return None

def _create_ext_def_options(param_name):
    """Create ExternalDefinitionCreationOptions - compatible with Revit 2024-2026+"""
    # Try SpecTypeId first (Revit 2022+)
    try:
        opt = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
        opt.Visible = True
        return opt
    except:
        pass
    
    # Fallback to ParameterType (Revit 2021 and below - removed in 2026)
    try:
        opt = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
        opt.Visible = True
        return opt
    except:
        pass
    
    return None

def _bind_param_insert(document, defn, binding, pg_key="PG_IFC"):
    """Insert parameter binding - compatible with Revit 2024-2026+"""
    group_id = _get_group_type_id(pg_key)
    
    if group_id is not None:
        try:
            return document.ParameterBindings.Insert(defn, binding, group_id)
        except:
            pass
    
    # Fallback: try without group
    try:
        return document.ParameterBindings.Insert(defn, binding)
    except:
        pass
    
    return False


# =====================================================================
# REVIT CONTEXT
# =====================================================================
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

SCRIPT_DIR = os.path.dirname(__file__)

# =====================================================================
# CATEGORY MAPPING
# =====================================================================
CATEGORY_MAP = {
    "Areas": BuiltInCategory.OST_Areas,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    "Plumbing Fixtures": BuiltInCategory.OST_PlumbingFixtures,
    "Project Information": BuiltInCategory.OST_ProjectInformation,
    "Ceilings": BuiltInCategory.OST_Ceilings,
    "Doors": BuiltInCategory.OST_Doors,
    "Toposolid": BuiltInCategory.OST_Topography,
    "Floors": BuiltInCategory.OST_Floors,
    "Shaft Openings": BuiltInCategory.OST_ShaftOpening,
    "Windows": BuiltInCategory.OST_Windows,
    "Planting": BuiltInCategory.OST_Planting,
    "Specialty Equipment": BuiltInCategory.OST_SpecialityEquipment,
    "Parking": BuiltInCategory.OST_Parking,
    "Rooms": BuiltInCategory.OST_Rooms,
    "Walls": BuiltInCategory.OST_Walls,
    "Railings": BuiltInCategory.OST_StairsRailing,
    "Ramps": BuiltInCategory.OST_Ramps,
    "Model Groups": BuiltInCategory.OST_IOSModelGroups,
    "Roofs": BuiltInCategory.OST_Roofs,
    "Furniture": BuiltInCategory.OST_Furniture,
    "Stairs": BuiltInCategory.OST_Stairs,
    "Structural Framing": BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": BuiltInCategory.OST_StructuralColumns,
    "Columns": BuiltInCategory.OST_Columns,
    "Structural Foundations": BuiltInCategory.OST_StructuralFoundation,
    "Electrical Equipment": BuiltInCategory.OST_ElectricalEquipment,
    "Duct Accessories": BuiltInCategory.OST_DuctAccessory,
    "Mechanical Equipment": BuiltInCategory.OST_MechanicalEquipment,
    "Pipes": BuiltInCategory.OST_PipeCurves,
    "Pipe Fittings": BuiltInCategory.OST_PipeFitting,
    "Ducts": BuiltInCategory.OST_DuctCurves,
    "Duct Fittings": BuiltInCategory.OST_DuctFitting,
    "Pipe Accessories": BuiltInCategory.OST_PipeAccessory,
}


# =====================================================================
# PARAMETER REQUIREMENT PARSER
# =====================================================================
class ParamRequirement:
    """A single parameter requirement: param name -> list of categories"""
    
    # Auto-mapping rules: keyword -> group name
    GROUP_MAP_RULES = [
        # (keywords_in_name, group_name, BuiltInParameterGroup_name)
        (["Ifc", "IfcObject", "IfcExport"], "IFC Parameters", "PG_IFC"),
        (["AGF_", "AST_", "ACN_", "ALS_", "AVF_"], "IFC Parameters", "PG_IFC"),
        (["AI_", "TED_"], "IFC Parameters", "PG_IFC"),
        (["Fire", "Smoke", "Shelter", "Sprinkler"], "Fire Protection", "PG_FIRE_PROTECTION"),
        (["Width", "Height", "Length", "Depth", "Thickness", "Diameter",
          "Area", "Volume", "Gradient", "Girth", "InnerDiameter", "OuterDiameter",
          "InternalLength", "InternalWidth", "ClearWidth", "ClearHeight", "ClearDepth",
          "OverallWidth", "Breadth", "InnerLength", "InnerWidth", "OuterLength", "OuterWidth",
          "RiserHeight", "TreadLength", "BasePlateThickness", "NominalDiameter",
          "StructuralWidth", "StructuralHeight", "MountingHeight", "SafetyBarrierHeight",
          "Hose_NominalDiameter"], "Dimensions", "PG_GEOMETRY"),
        (["Material", "MaterialGrade", "ReinforcementSteelGrade", "SectionFabricationMethod"], 
         "Materials and Finishes", "PG_MATERIALS"),
        (["Phase", "Status", "Retrofit"], "Phasing", "PG_PHASING"),
        (["Mark", "SpaceName", "UnitNumber", "LotNumber", "FamilyLot",
          "BoreholeRef", "TreeNumber", "HedgeNumber"], "Identity Data", "PG_IDENTITY_DATA"),
        (["System", "SystemType", "SystemName"], "Mechanical", "PG_MECHANICAL"),
        (["Ventilation", "VentilationType", "VentilationMode", "CValue", 
          "SoundPower", "SoundPressure"], "Mechanical", "PG_MECHANICAL"),
        (["Capacity", "NominalCapacity", "EffectiveCapacity", "LoadingCapacity",
          "OccupancyLoad", "WorkingLoad", "PumpHead", "Duty", "Standby",
          "CompactionRatio"], "Structural", "PG_STRUCTURAL"),
        (["Rebar", "Stirrups", "StirrupsType", "MainRebar", "TopMain", "BottomMain",
          "TopDistribution", "BottomDistribution", "SideBar", "WeldedMesh",
          "TopLeft", "TopMiddle", "TopRight", "BottomLeft", "BottomMiddle", "BottomRight",
          "LatticeGirder", "ColumnCage", "SpliceConnection", "SpliceDetail",
          "PrefabricationReinforcement", "ReinforcementLength"], "Structural", "PG_STRUCTURAL"),
        (["Connection", "ConnectionType", "ConnectionDetail", "MechanicalConnectionType"],
         "Structural", "PG_STRUCTURAL"),
        (["Construction", "ConstructionMethod"], "Construction", "PG_CONSTRUCTION"),
        (["Plumbing", "WELS", "TradeEffluent", "IsPotable", "WaterSupply",
          "Perforated", "PreInsulated", "DemountableStructure"], "Plumbing", "PG_PLUMBING"),
        (["Electrical", "PWCS_Flushing", "Purpose"], "Electrical", "PG_ELECTRICAL"),
    ]
    
    def __init__(self, name):
        self.name = name
        self.categories = []
        self.disciplines = set()
        self.selected = True
        self.already_exists = False
        self.status = ""
        self.group_under = ""  # Display name: "IFC Parameters", "Dimensions", etc.
        self.group_pg = ""     # BuiltInParameterGroup key: "PG_IFC", "PG_GEOMETRY", etc.
        self._auto_map_group()
    
    def _auto_map_group(self):
        """Auto-detect parameter group based on name patterns"""
        for keywords, group_name, pg_key in self.GROUP_MAP_RULES:
            for kw in keywords:
                if self.name.startswith(kw) or self.name == kw:
                    self.group_under = group_name
                    self.group_pg = pg_key
                    return
        # Default
        self.group_under = "IFC Parameters"
        self.group_pg = "PG_IFC"
    
    @property
    def group_display(self):
        return self.group_under or "IFC Parameters"


class RequirementParser:
    """Parse XML or Excel into list of ParamRequirements"""
    
    @staticmethod
    def from_xml(filepath):
        """Parse Autodesk Model Checker XML"""
        import xml.etree.ElementTree as ET
        tree = ET.parse(filepath)
        root = tree.getroot()
        
        # Build: param_name -> set of (discipline, category)
        param_map = {}
        
        for h in root.findall("Heading"):
            disc = h.get("HeadingText", "")
            for s in h.findall("Section"):
                cat = s.get("SectionName", "")
                for c in s.findall("Check"):
                    param = c.get("CheckName", "")
                    if not param:
                        continue
                    if param not in param_map:
                        param_map[param] = {"cats": set(), "discs": set()}
                    param_map[param]["cats"].add(cat)
                    param_map[param]["discs"].add(disc)
        
        # Convert to ParamRequirement list
        reqs = []
        for name, data in sorted(param_map.items()):
            req = ParamRequirement(name)
            req.categories = sorted(data["cats"])
            req.disciplines = data["discs"]
            reqs.append(req)
        
        return reqs
    
    @staticmethod
    def read_excel_headers(filepath):
        """Read sheet names and headers from Excel for mapping dialog"""
        try:
            clr.AddReference('Microsoft.Office.Interop.Excel')
            from Microsoft.Office.Interop import Excel as ExcelInterop
            
            excel_app = ExcelInterop.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            wb = excel_app.Workbooks.Open(filepath)
            
            sheets_data = {}
            for si in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets[si]
                sheet_name = ws.Name
                
                # Read headers from row 1
                headers = []
                cols = ws.UsedRange.Columns.Count
                total_rows = ws.UsedRange.Rows.Count
                for ci in range(1, min(cols + 1, 30)):
                    val = ws.Cells[1, ci].Value2
                    headers.append(str(val) if val else "Column {}".format(ci))
                
                # Read preview rows (2-6)
                preview = []
                for ri in range(2, min(total_rows + 1, 7)):
                    row_data = []
                    for ci in range(1, min(cols + 1, 30)):
                        val = ws.Cells[ri, ci].Value2
                        row_data.append(str(val) if val else "")
                    preview.append(row_data)
                
                sheets_data[sheet_name] = {
                    "headers": headers,
                    "preview": preview,
                    "total_rows": total_rows,
                    "total_cols": cols
                }
            
            wb.Close(False)
            excel_app.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app)
            return sheets_data
            
        except Exception as e:
            raise Exception("Error reading Excel: {}".format(str(e)))
    
    @staticmethod
    def from_excel(filepath, sheet_name=None, col_param=1, col_category=2, 
                   col_discipline=None, header_row=1):
        """Parse Excel with user-specified column mapping"""
        try:
            clr.AddReference('Microsoft.Office.Interop.Excel')
            from Microsoft.Office.Interop import Excel as ExcelInterop
            
            excel_app = ExcelInterop.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            wb = excel_app.Workbooks.Open(filepath)
            
            if sheet_name:
                ws = wb.Sheets[sheet_name]
            else:
                ws = wb.Sheets[1]
            
            rows = ws.UsedRange.Rows.Count
            
            param_map = {}
            data_start = header_row + 1
            
            for r in range(data_start, rows + 1):
                raw_param = ws.Cells[r, col_param].Value2
                if raw_param is None:
                    continue
                # Fix: convert float numbers to int strings (1.0 -> "1")
                if isinstance(raw_param, float) and raw_param == int(raw_param):
                    param = str(int(raw_param))
                else:
                    param = str(raw_param).strip()
                
                cat = ""
                if col_category:
                    raw_cat = ws.Cells[r, col_category].Value2
                    if raw_cat is not None:
                        cat = str(raw_cat).strip()
                
                disc = ""
                if col_discipline:
                    raw_disc = ws.Cells[r, col_discipline].Value2
                    if raw_disc is not None:
                        disc = str(raw_disc).strip()
                
                if not param:
                    continue
                
                if param not in param_map:
                    param_map[param] = {"cats": set(), "discs": set()}
                if cat:
                    param_map[param]["cats"].add(cat)
                if disc:
                    param_map[param]["discs"].add(disc)
            
            wb.Close(False)
            excel_app.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app)
            
            reqs = []
            for name, data in sorted(param_map.items()):
                req = ParamRequirement(name)
                req.categories = sorted(data["cats"])
                req.disciplines = data["discs"]
                reqs.append(req)
            return reqs
            
        except Exception as e:
            raise Exception("Excel parse error: {}".format(str(e)))


# =====================================================================
# EXCEL COLUMN MAPPER DIALOG
# =====================================================================



class ExcelColumnMapper:
    """Dialog to let user map Excel columns to Parameter/Category/Discipline"""
    
    def __init__(self, filepath):
        self.filepath = filepath
        self.result = None  # Will hold mapping dict if OK
        self.sheets_data = RequirementParser.read_excel_headers(filepath)
        
        with codecs.open(MAPPER_XAML_PATH, "r", "utf-8") as f:
            xaml_content = f.read()
        xr = XmlReader.Create(StringReader(xaml_content))
        self.window = XamlReader.Load(xr)
        
        # Get controls
        self.cmbSheet = self.window.FindName("cmbSheet")
        self.txtHeaderRow = self.window.FindName("txtHeaderRow")
        self.cmbColParam = self.window.FindName("cmbColParam")
        self.cmbColCategory = self.window.FindName("cmbColCategory")
        self.cmbColDiscipline = self.window.FindName("cmbColDiscipline")
        self.txtPreview = self.window.FindName("txtPreview")
        self.txtMapInfo = self.window.FindName("txtMapInfo")
        self.btnMapOK = self.window.FindName("btnMapOK")
        self.btnMapCancel = self.window.FindName("btnMapCancel")
        
        # Events
        self.cmbSheet.SelectionChanged += self._on_sheet_changed
        self.btnMapOK.Click += self._on_ok
        self.btnMapCancel.Click += self._on_cancel
        
        # Populate sheets
        for name in self.sheets_data.keys():
            self.cmbSheet.Items.Add(name)
        if self.cmbSheet.Items.Count > 0:
            self.cmbSheet.SelectedIndex = 0
    
    def _on_sheet_changed(self, sender, args):
        sheet_name = str(self.cmbSheet.SelectedItem)
        if sheet_name not in self.sheets_data:
            return
        
        data = self.sheets_data[sheet_name]
        headers = data["headers"]
        preview = data["preview"]
        
        # Populate column dropdowns
        none_option = "(None)"
        
        for cmb in [self.cmbColParam, self.cmbColCategory, self.cmbColDiscipline]:
            cmb.Items.Clear()
        
        self.cmbColDiscipline.Items.Add(none_option)
        self.cmbColCategory.Items.Add(none_option)
        
        for i, h in enumerate(headers):
            display = "Col {} - {}".format(i + 1, h)
            self.cmbColParam.Items.Add(display)
            self.cmbColCategory.Items.Add(display)
            self.cmbColDiscipline.Items.Add(display)
        
        # Auto-detect columns by header keywords
        param_idx = 0
        cat_idx = 0  # (None)
        disc_idx = 0  # (None)
        
        for i, h in enumerate(headers):
            hl = h.lower()
            if any(k in hl for k in ["parameter", "param", "property", "field"]):
                param_idx = i
            if any(k in hl for k in ["category", "revit category", "element"]):
                cat_idx = i + 1  # +1 because "(None)" is at index 0
            if any(k in hl for k in ["discipline", "disc", "group", "heading"]):
                disc_idx = i + 1
        
        if self.cmbColParam.Items.Count > param_idx:
            self.cmbColParam.SelectedIndex = param_idx
        if self.cmbColCategory.Items.Count > cat_idx:
            self.cmbColCategory.SelectedIndex = cat_idx
        if self.cmbColDiscipline.Items.Count > disc_idx:
            self.cmbColDiscipline.SelectedIndex = disc_idx
        
        # Preview
        lines = []
        # Header line
        header_line = " | ".join("{:15s}".format(h[:15]) for h in headers[:8])
        lines.append(header_line)
        lines.append("-" * len(header_line))
        for row in preview:
            line = " | ".join("{:15s}".format(str(v)[:15]) for v in row[:8])
            lines.append(line)
        
        self.txtPreview.Text = "\n".join(lines)
        self.txtMapInfo.Text = "{} rows, {} columns".format(
            data["total_rows"], data["total_cols"])
    
    def _on_ok(self, sender, args):
        # Validate - Parameter column is required
        if self.cmbColParam.SelectedIndex < 0:
            System.Windows.MessageBox.Show(
                "Please select the Parameter Name column.",
                "Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        sheet_name = str(self.cmbSheet.SelectedItem)
        
        # Get column numbers (1-based for Excel)
        col_param = self.cmbColParam.SelectedIndex + 1
        
        col_category = None
        if self.cmbColCategory.SelectedIndex > 0:  # 0 = "(None)"
            col_category = self.cmbColCategory.SelectedIndex  # already offset by (None)
        
        col_discipline = None
        if self.cmbColDiscipline.SelectedIndex > 0:
            col_discipline = self.cmbColDiscipline.SelectedIndex
        
        try:
            header_row = int(self.txtHeaderRow.Text.strip())
        except:
            header_row = 1
        
        self.result = {
            "sheet_name": sheet_name,
            "col_param": col_param,
            "col_category": col_category,
            "col_discipline": col_discipline,
            "header_row": header_row
        }
        self.window.DialogResult = System.Nullable[System.Boolean](True)
        self.window.Close()
    
    def _on_cancel(self, sender, args):
        self.result = None
        self.window.Close()
    
    def show(self):
        self.window.ShowDialog()
        return self.result


# =====================================================================
# CATEGORY PICKER DIALOG
# =====================================================================



class CategoryPickerDialog:
    """Dialog for user to select which Revit categories to assign to parameters"""
    
    def __init__(self, params_without_cats):
        """
        params_without_cats: list of ParamRequirement with empty categories
        """
        self.params = params_without_cats
        self.selected_categories = []
        self._checkboxes = []
        
        with codecs.open(CATPICKER_XAML_PATH, "r", "utf-8") as f:
            xaml_content = f.read()
        xr = XmlReader.Create(StringReader(xaml_content))
        self.window = XamlReader.Load(xr)
        
        self.txtPickerInfo = self.window.FindName("txtPickerInfo")
        self.spPickCategories = self.window.FindName("spPickCategories")
        self.txtPickSearch = self.window.FindName("txtPickSearch")
        self.txtPickCount = self.window.FindName("txtPickCount")
        self.btnPickAll = self.window.FindName("btnPickAll")
        self.btnPickNone = self.window.FindName("btnPickNone")
        self.btnPickOK = self.window.FindName("btnPickOK")
        self.btnPickCancel = self.window.FindName("btnPickCancel")
        
        self.txtPickerInfo.Text = "{} parameters need category assignment".format(len(params_without_cats))
        
        self.btnPickAll.Click += lambda s, e: self._toggle_all(True)
        self.btnPickNone.Click += lambda s, e: self._toggle_all(False)
        self.btnPickOK.Click += self._on_ok
        self.btnPickCancel.Click += lambda s, e: self.window.Close()
        self.txtPickSearch.TextChanged += lambda s, e: self._render_categories()
        
        self._render_categories()
    
    def _render_categories(self):
        self.spPickCategories.Children.Clear()
        self._checkboxes = []
        converter = BrushConverter()
        
        search = (self.txtPickSearch.Text or "").strip().lower()
        
        # Sort categories alphabetically
        sorted_cats = sorted(CATEGORY_MAP.keys())
        
        for cat_name in sorted_cats:
            if search and search not in cat_name.lower():
                continue
            
            chk = CheckBox()
            chk.Content = cat_name
            chk.FontSize = 11
            chk.Margin = System.Windows.Thickness(4, 2, 4, 2)
            chk.IsChecked = System.Nullable[System.Boolean](
                cat_name in self.selected_categories)
            chk.Tag = cat_name
            chk.Checked += self._on_cat_toggle
            chk.Unchecked += self._on_cat_toggle
            
            self._checkboxes.append(chk)
            self.spPickCategories.Children.Add(chk)
    
    def _on_cat_toggle(self, sender, args):
        cat_name = str(sender.Tag)
        if bool(sender.IsChecked):
            if cat_name not in self.selected_categories:
                self.selected_categories.append(cat_name)
        else:
            if cat_name in self.selected_categories:
                self.selected_categories.remove(cat_name)
        self.txtPickCount.Text = "{} selected".format(len(self.selected_categories))
    
    def _toggle_all(self, state):
        self.selected_categories = []
        if state:
            self.selected_categories = sorted(CATEGORY_MAP.keys())
        self._render_categories()
        self.txtPickCount.Text = "{} selected".format(len(self.selected_categories))
    
    def _on_ok(self, sender, args):
        if not self.selected_categories:
            System.Windows.MessageBox.Show(
                "Please select at least 1 category.",
                "Required", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        self.window.DialogResult = System.Nullable[System.Boolean](True)
        self.window.Close()
    
    def show(self):
        self.window.ShowDialog()
        if self.selected_categories:
            return self.selected_categories
        return None


# =====================================================================
# PARAMETER ADDER ENGINE
# =====================================================================
class ParameterAdder:
    """Add shared parameters to Revit model"""
    
    def __init__(self, document, application):
        self.doc = document
        self.app = application
        self.log = []
        self._original_sp_path = None
        self._temp_sp_path = None
    
    def _get_existing_params(self):
        """Get set of parameter names already in model"""
        existing = set()
        bm = self.doc.ParameterBindings
        it = bm.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            try:
                existing.add(it.Key.Name)
            except:
                pass
        return existing
    
    def _get_existing_param_categories(self, param_name):
        """Get categories already bound to a parameter"""
        bound_cats = set()
        bm = self.doc.ParameterBindings
        it = bm.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            try:
                if it.Key.Name == param_name:
                    binding = it.Current
                    if hasattr(binding, 'Categories'):
                        for cat in binding.Categories:
                            bound_cats.add(cat.Name)
            except:
                pass
        return bound_cats
    
    def _setup_temp_shared_param_file(self):
        """
        Create a dedicated temp shared parameter file for adding new params.
        This avoids issues with read-only or network shared param files.
        """
        # Save original
        self._original_sp_path = self.app.SharedParametersFilename
        
        # Create temp file in script directory (guaranteed writable)
        self._temp_sp_path = os.path.join(SCRIPT_DIR, "DQT_IFC_SG_SharedParams.txt")
        
        if not os.path.exists(self._temp_sp_path):
            with open(self._temp_sp_path, 'w') as f:
                f.write("# IFC+SG Shared Parameters - Auto-generated by DQT\n")
                f.write("*META\tVERSION\tMINVERSION\n")
                f.write("META\t2\t1\n")
        
        self.app.SharedParametersFilename = self._temp_sp_path
        sp_file = self.app.OpenSharedParameterFile()
        return sp_file
    
    def _restore_shared_param_file(self):
        """Restore original shared parameter file"""
        if self._original_sp_path:
            try:
                self.app.SharedParametersFilename = self._original_sp_path
            except:
                pass
    
    def _find_definition_in_file(self, sp_file, param_name):
        """Search all groups in shared param file for a definition"""
        for group in sp_file.Groups:
            for defn in group.Definitions:
                if defn.Name == param_name:
                    return defn
        return None
    
    def _create_definition(self, sp_file, group_name, param_name):
        """Create shared parameter definition with Revit version compatibility"""
        # Get or create group
        group = None
        for g in sp_file.Groups:
            if g.Name == group_name:
                group = g
                break
        if not group:
            group = sp_file.Groups.Create(group_name)
        
        # Check if already exists in this group
        for d in group.Definitions:
            if d.Name == param_name:
                return d
        
        # Try Revit 2024+ API first (ForgeTypeId / SpecTypeId)
        try:
            opt = _create_ext_def_options(param_name)
            if opt:
                return group.Definitions.Create(opt)
        except:
            pass
        
        try:
            opt = _create_ext_def_options(param_name)
            if opt:
                return group.Definitions.Create(opt)
        except:
            pass
        
        return None
    
    def check_existing(self, requirements):
        """Check which parameters already exist in model"""
        existing = self._get_existing_params()
        for req in requirements:
            if req.name in existing:
                req.already_exists = True
                req.status = "exists"
                req.bound_categories = self._get_existing_param_categories(req.name)
            else:
                req.already_exists = False
                req.status = ""
                req.bound_categories = set()
        return requirements
    
    def add_parameters(self, requirements, progress_callback=None):
        """Add selected parameters to model"""
        self.log = []
        
        # Setup dedicated temp shared param file
        sp_file = self._setup_temp_shared_param_file()
        if not sp_file:
            self.log.append("ERROR: Could not create shared parameter file at {}".format(
                self._temp_sp_path))
            self._restore_shared_param_file()
            return self.log
        
        selected = [r for r in requirements if r.selected]
        total = len(selected)
        added = 0
        skipped = 0
        failed = 0
        
        for idx, req in enumerate(selected):
            try:
                # Determine group name
                group_name = "IFC+SG Parameters"
                if req.disciplines:
                    disc_list = sorted(req.disciplines)
                    if len(disc_list) == 1:
                        group_name = "IFC+SG_{}".format(disc_list[0])
                
                # First check if definition exists in temp file already
                defn = self._find_definition_in_file(sp_file, req.name)
                
                if not defn:
                    # Create new definition
                    defn = self._create_definition(sp_file, group_name, req.name)
                
                if not defn:
                    req.status = "failed"
                    self.log.append("ERROR: {} - Could not create shared param definition".format(req.name))
                    failed += 1
                    continue
                
                # Build category set
                cat_set = CategorySet()
                for cat_name in req.categories:
                    bic = CATEGORY_MAP.get(cat_name)
                    if bic is not None:
                        try:
                            cat = Category.GetCategory(self.doc, bic)
                            if cat and cat.AllowsBoundParameters:
                                cat_set.Insert(cat)
                        except:
                            pass
                
                if cat_set.Size == 0:
                    req.status = "skipped"
                    self.log.append("SKIP: {} - No valid categories found".format(req.name))
                    skipped += 1
                    continue
                
                # Try to get existing binding from model
                existing_binding = None
                bm = self.doc.ParameterBindings
                it = bm.ForwardIterator()
                it.Reset()
                while it.MoveNext():
                    try:
                        if it.Key.Name == req.name:
                            existing_binding = it.Current
                            defn = it.Key  # Use the existing definition key
                            break
                    except:
                        pass
                
                if existing_binding and hasattr(existing_binding, 'Categories'):
                    # Update existing binding with additional categories
                    changed = False
                    for cat in cat_set:
                        try:
                            if not existing_binding.Categories.Contains(cat):
                                existing_binding.Categories.Insert(cat)
                                changed = True
                        except:
                            pass
                    
                    if changed:
                        self.doc.ParameterBindings.ReInsert(defn, existing_binding)
                        req.status = "updated"
                        self.log.append("UPDATE: {} - Added new category bindings".format(req.name))
                    else:
                        req.status = "exists"
                        self.log.append("SKIP: {} - Already bound to all categories".format(req.name))
                    added += 1
                else:
                    # Create new instance binding
                    binding = InstanceBinding(cat_set)
                    
                    # Use the mapped parameter group
                    success = False
                    pg_key = req.group_pg if hasattr(req, 'group_pg') and req.group_pg else "PG_IFC"
                    
                    try:
                        success = _bind_param_insert(self.doc, defn, binding, pg_key)
                    except:
                        pass
                    
                    # Fallback to PG_IFC
                    if not success:
                        try:
                            success = _bind_param_insert(self.doc, defn, binding, "PG_IFC")
                        except:
                            pass
                    
                    # Fallback to no group
                    if not success:
                        try:
                            success = self.doc.ParameterBindings.Insert(defn, binding)
                        except:
                            pass
                    
                    if success:
                        req.status = "added"
                        self.log.append("ADD: {} -> {} cats [{}]".format(
                            req.name, cat_set.Size, req.group_display))
                        added += 1
                    else:
                        req.status = "failed"
                        self.log.append("FAIL: {} - ParameterBindings.Insert returned false".format(req.name))
                        failed += 1
                
            except Exception as e:
                req.status = "failed"
                self.log.append("ERROR: {} - {}".format(req.name, str(e)))
                failed += 1
            
            if progress_callback:
                progress_callback(idx + 1, total)
        
        # Restore original shared parameter file
        self._restore_shared_param_file()
        
        self.log.insert(0, "SUMMARY: {} added/updated, {} skipped, {} failed out of {}".format(
            added, skipped, failed, total))
        
        return self.log


# =====================================================================
# WPF UI
# =====================================================================



# =====================================================================
# MAIN WINDOW
# =====================================================================
class ParamLoaderWindow:
    
    def __init__(self):
        self.requirements = []
        self.adder = ParameterAdder(doc, app)
        self._show_filter = "all"
        self._checkboxes = []  # track checkbox -> requirement index
        
        with codecs.open(MAIN_XAML_PATH, "r", "utf-8") as f:
            xaml_content = f.read()
        xr = XmlReader.Create(StringReader(xaml_content))
        self.window = XamlReader.Load(xr)
        self._get_controls()
        self._bind_events()
    
    def _get_controls(self):
        names = [
            "btnImportXML", "btnImportExcel", "txtSourceInfo",
            "txtTotalParams", "txtExisting", "txtToAdd", "txtCategories",
            "btnSelectNew", "btnSelectAll", "btnSelectNone",
            "txtSearch", "btnShowAll", "btnShowNew", "btnShowExist",
            "btnSortName", "btnSortGroup", "btnSortCat", "btnSortStatus", "chkHeaderAll",
            "spParams", "txtLog",
            "txtStatus", "btnAddParams", "btnClose"
        ]
        for n in names:
            setattr(self, n, self.window.FindName(n))
    
    def _bind_events(self):
        self.btnImportXML.Click += self._on_import_xml
        self.btnImportExcel.Click += self._on_import_excel
        self.btnSelectNew.Click += self._on_select_new
        self.btnSelectAll.Click += self._on_select_all
        self.btnSelectNone.Click += self._on_select_none
        self.btnShowAll.Click += lambda s, e: self._set_filter("all")
        self.btnShowNew.Click += lambda s, e: self._set_filter("new")
        self.btnShowExist.Click += lambda s, e: self._set_filter("exist")
        self.btnSortName.Click += lambda s, e: self._set_sort("name")
        self.btnSortGroup.Click += lambda s, e: self._set_sort("group")
        self.btnSortCat.Click += lambda s, e: self._set_sort("category")
        self.btnSortStatus.Click += lambda s, e: self._set_sort("status")
        self.chkHeaderAll.Checked += self._on_header_check
        self.chkHeaderAll.Unchecked += self._on_header_check
        self.txtSearch.TextChanged += lambda s, e: self._render_params()
        self.btnAddParams.Click += self._on_add_params
        self.btnClose.Click += lambda s, e: self.window.Close()
        self._current_filter = "all"
        self._sort_key = "name"
        self._sort_asc = True
        self._last_clicked_idx = -1  # for shift-select
    
    def _log(self, text):
        current = self.txtLog.Text or ""
        self.txtLog.Text = text + "\n" + current
    
    # =================================================================
    # IMPORT
    # =================================================================
    def _on_import_xml(self, sender, args):
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dlg = OpenFileDialog()
        dlg.Filter = "XML Files (*.xml)|*.xml"
        dlg.Title = "Import Autodesk Model Checker XML"
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                self.requirements = RequirementParser.from_xml(dlg.FileName)
                self._post_import(os.path.basename(dlg.FileName))
            except Exception as e:
                System.Windows.MessageBox.Show(str(e), "Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_import_excel(self, sender, args):
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dlg = OpenFileDialog()
        dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV Files (*.csv)|*.csv"
        dlg.Title = "Import Excel Parameter Mapping"
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                # Show column mapper dialog
                mapper = ExcelColumnMapper(dlg.FileName)
                mapping = mapper.show()
                
                if mapping is None:
                    return  # User cancelled
                
                self.requirements = RequirementParser.from_excel(
                    dlg.FileName,
                    sheet_name=mapping.get("sheet_name"),
                    col_param=mapping.get("col_param", 1),
                    col_category=mapping.get("col_category"),
                    col_discipline=mapping.get("col_discipline"),
                    header_row=mapping.get("header_row", 1)
                )
                self._post_import(os.path.basename(dlg.FileName))
            except Exception as e:
                System.Windows.MessageBox.Show(
                    "Error:\n{}".format(str(e)),
                    "Import Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _post_import(self, source_name):
        """After import: check existing, update UI"""
        self.window.Cursor = System.Windows.Input.Cursors.Wait
        
        # Check which already exist
        t = Transaction(doc, "Check Parameters")
        t.Start()
        try:
            self.adder.check_existing(self.requirements)
            t.RollBack()
        except:
            t.RollBack()
        
        # Update stats
        total = len(self.requirements)
        existing = len([r for r in self.requirements if r.already_exists])
        to_add = total - existing
        all_cats = set()
        for r in self.requirements:
            all_cats.update(r.categories)
        
        self.txtTotalParams.Text = str(total)
        self.txtExisting.Text = str(existing)
        self.txtToAdd.Text = str(to_add)
        self.txtCategories.Text = str(len(all_cats))
        
        self.txtSourceInfo.Text = "{} ({} params)".format(source_name, total)
        
        # Auto-select new params only
        for r in self.requirements:
            r.selected = not r.already_exists
        
        self._show_filter = "all"
        self._render_params()
        
        self.btnAddParams.IsEnabled = True
        self.txtStatus.Text = "{} params loaded. {} new, {} already exist.".format(total, to_add, existing)
        self._log("Imported: {} ({} params, {} existing)".format(source_name, total, existing))
        
        self.window.Cursor = System.Windows.Input.Cursors.Arrow
    
    # =================================================================
    # SELECTION
    # =================================================================
    def _on_select_new(self, sender, args):
        for r in self.requirements:
            r.selected = not r.already_exists
        self._render_params()
    
    def _on_select_all(self, sender, args):
        for r in self.requirements:
            r.selected = True
        self._render_params()
    
    def _on_select_none(self, sender, args):
        for r in self.requirements:
            r.selected = False
        self._render_params()
    
    def _set_filter(self, f):
        self._show_filter = f
        self._render_params()
    
    def _set_sort(self, key):
        """Toggle sort by key"""
        if self._sort_key == key:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_key = key
            self._sort_asc = True
        
        # Update button labels
        arrow_up = u" \u25B2"
        arrow_down = u" \u25BC"
        arrow = arrow_up if self._sort_asc else arrow_down
        
        self.btnSortName.Content = "Name" + (arrow if key == "name" else "")
        self.btnSortGroup.Content = "Group" + (arrow if key == "group" else "")
        self.btnSortCat.Content = "Category" + (arrow if key == "category" else "")
        self.btnSortStatus.Content = "Status" + (arrow if key == "status" else "")
        
        self._render_params()
    
    def _on_header_check(self, sender, args):
        """Toggle all visible params"""
        checked = bool(sender.IsChecked)
        search = (self.txtSearch.Text or "").strip().lower()
        for req in self.requirements:
            if self._show_filter == "new" and req.already_exists:
                continue
            if self._show_filter == "exist" and not req.already_exists:
                continue
            if search and search not in req.name.lower():
                continue
            req.selected = checked
        self._render_params()
    
    # =================================================================
    # RENDER
    # =================================================================
    def _render_params(self):
        self.spParams.Children.Clear()
        self._checkboxes = []
        self._visible_indices = []  # track which requirement indices are visible
        converter = BrushConverter()
        
        search = (self.txtSearch.Text or "").strip().lower()
        
        # Build filtered list with original indices
        filtered = []
        for idx, req in enumerate(self.requirements):
            if self._show_filter == "new" and req.already_exists:
                continue
            if self._show_filter == "exist" and not req.already_exists:
                continue
            if search and search not in req.name.lower() and \
               not any(search in c.lower() for c in req.categories):
                continue
            filtered.append((idx, req))
        
        # Sort
        if self._sort_key == "name":
            filtered.sort(key=lambda x: x[1].name.lower(), reverse=not self._sort_asc)
        elif self._sort_key == "group":
            filtered.sort(key=lambda x: (x[1].group_display.lower(), x[1].name.lower()),
                         reverse=not self._sort_asc)
        elif self._sort_key == "category":
            filtered.sort(key=lambda x: (x[1].categories[0].lower() if x[1].categories else "", x[1].name.lower()), 
                         reverse=not self._sort_asc)
        elif self._sort_key == "status":
            status_order = {"": 0, "failed": 1, "added": 2, "updated": 3, "exists": 4, "skipped": 5}
            filtered.sort(key=lambda x: (
                0 if not x[1].already_exists else 1,
                status_order.get(x[1].status, 0),
                x[1].name.lower()
            ), reverse=not self._sort_asc)
        
        self._visible_indices = [idx for idx, _ in filtered]
        
        # Group headers when sorted by category or group
        current_group_header = ""
        
        # Group color map for parameter groups
        group_colors = {
            "IFC Parameters": "#1565C0",
            "Dimensions": "#6A1B9A",
            "Fire Protection": "#C62828",
            "Materials and Finishes": "#E65100",
            "Structural": "#4E342E",
            "Identity Data": "#00695C",
            "Mechanical": "#2E7D32",
            "Construction": "#F57F17",
            "Plumbing": "#0277BD",
            "Electrical": "#AD1457",
            "Phasing": "#546E7A",
        }
        group_bg = {
            "IFC Parameters": "#E3F2FD",
            "Dimensions": "#F3E5F5",
            "Fire Protection": "#FFEBEE",
            "Materials and Finishes": "#FFF3E0",
            "Structural": "#EFEBE9",
            "Identity Data": "#E0F2F1",
            "Mechanical": "#E8F5E9",
            "Construction": "#FFFDE7",
            "Plumbing": "#E1F5FE",
            "Electrical": "#FCE4EC",
            "Phasing": "#ECEFF1",
        }
        
        for pos, (orig_idx, req) in enumerate(filtered):
            # Group/Category header when sorted by group or category
            if self._sort_key in ("category", "group"):
                if self._sort_key == "category":
                    header_val = req.categories[0] if req.categories else ""
                else:
                    header_val = req.group_display
                
                if header_val != current_group_header:
                    current_group_header = header_val
                    
                    grp_header = System.Windows.Controls.Border()
                    grp_header.Margin = System.Windows.Thickness(0, 4, 0, 2)
                    grp_header.Padding = System.Windows.Thickness(6, 2, 6, 2)
                    try:
                        if self._sort_key == "group":
                            # Use discipline-specific color
                            bg = group_bg.get(header_val, "#0F172A")
                            grp_header.Background = converter.ConvertFromString(bg)
                        else:
                            grp_header.Background = converter.ConvertFromString("#0F172A")
                    except:
                        pass
                    grp_header.CornerRadius = System.Windows.CornerRadius(2)
                    
                    # Count params in this group
                    if self._sort_key == "group":
                        grp_count = len([r for _, r in filtered if r.group_display == header_val])
                    else:
                        grp_count = len([r for _, r in filtered 
                                        if r.categories and r.categories[0] == header_val])
                    
                    grp_txt = TextBlock()
                    grp_txt.Text = u"{} ({})".format(header_val, grp_count)
                    grp_txt.FontWeight = System.Windows.FontWeights.Bold
                    grp_txt.FontSize = 11
                    try:
                        if self._sort_key == "group":
                            fg = group_colors.get(header_val, "#5D4E37")
                            grp_txt.Foreground = converter.ConvertFromString(fg)
                        else:
                            grp_txt.Foreground = converter.ConvertFromString("#5D4E37")
                    except:
                        pass
                    grp_header.Child = grp_txt
                    self.spParams.Children.Add(grp_header)
            
            # Row using Grid layout matching header columns (5 columns now)
            row_border = System.Windows.Controls.Border()
            row_border.Padding = System.Windows.Thickness(6, 3, 6, 3)
            row_border.CornerRadius = System.Windows.CornerRadius(2)
            row_border.Margin = System.Windows.Thickness(0, 0, 0, 1)
            
            if req.status == "added" or req.status == "updated":
                bg_color = "#E8F5E9"
            elif req.already_exists:
                bg_color = "#F9F9F9"
            else:
                bg_color = "#FFF8E1"
            
            try:
                row_border.Background = converter.ConvertFromString(bg_color)
            except:
                pass
            
            row_grid = Grid()
            gc1 = ColumnDefinition()
            gc1.Width = System.Windows.GridLength(28)
            gc2 = ColumnDefinition()
            gc2.Width = System.Windows.GridLength(65)
            gc3 = ColumnDefinition()
            gc3.Width = System.Windows.GridLength(200)
            gc4 = ColumnDefinition()
            gc4.Width = System.Windows.GridLength(70)
            gc5 = ColumnDefinition()
            gc5.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            row_grid.ColumnDefinitions.Add(gc1)
            row_grid.ColumnDefinitions.Add(gc2)
            row_grid.ColumnDefinitions.Add(gc3)
            row_grid.ColumnDefinitions.Add(gc4)
            row_grid.ColumnDefinitions.Add(gc5)
            
            # Col 0: Checkbox
            chk = CheckBox()
            chk.IsChecked = System.Nullable[System.Boolean](bool(req.selected))
            chk.VerticalAlignment = System.Windows.VerticalAlignment.Center
            chk.Tag = orig_idx
            chk.Checked += self._on_param_toggled
            chk.Unchecked += self._on_param_toggled
            self._checkboxes.append((chk, orig_idx, pos))
            Grid.SetColumn(chk, 0)
            row_grid.Children.Add(chk)
            
            # Col 1: Status
            status_txt = TextBlock()
            status_txt.FontSize = 10
            status_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
            
            if req.status == "added":
                status_txt.Text = u"\u2714 ADDED"
                try: status_txt.Foreground = converter.ConvertFromString("#2E7D32")
                except: pass
            elif req.status == "updated":
                status_txt.Text = u"\u2714 UPDATED"
                try: status_txt.Foreground = converter.ConvertFromString("#2E7D32")
                except: pass
            elif req.already_exists:
                status_txt.Text = u"\u2714 EXISTS"
                try: status_txt.Foreground = converter.ConvertFromString("#888")
                except: pass
            elif req.status == "failed":
                status_txt.Text = u"\u2718 FAILED"
                try: status_txt.Foreground = converter.ConvertFromString("#C62828")
                except: pass
            else:
                status_txt.Text = u"\u25CB NEW"
                try: status_txt.Foreground = converter.ConvertFromString("#F57F17")
                except: pass
            
            Grid.SetColumn(status_txt, 1)
            row_grid.Children.Add(status_txt)
            
            # Col 2: Param name
            name_txt = TextBlock()
            name_txt.Text = req.name
            name_txt.FontSize = 11
            name_txt.FontWeight = System.Windows.FontWeights.SemiBold
            name_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
            Grid.SetColumn(name_txt, 2)
            row_grid.Children.Add(name_txt)
            
            # Col 3: Group badge
            grp_badge = TextBlock()
            grp_short = req.group_display
            # Shorten for display
            short_map = {
                "IFC Parameters": "IFC",
                "Fire Protection": "Fire",
                "Materials and Finishes": "Material",
                "Identity Data": "Identity",
                "Dimensions": "Dims",
                "Mechanical": "Mech",
                "Construction": "Constr",
                "Structural": "Struct",
                "Plumbing": "Plumb",
                "Electrical": "Elec",
                "Phasing": "Phase",
            }
            grp_badge.Text = short_map.get(grp_short, grp_short[:8])
            grp_badge.FontSize = 9
            grp_badge.FontWeight = System.Windows.FontWeights.SemiBold
            grp_badge.VerticalAlignment = System.Windows.VerticalAlignment.Center
            try:
                grp_badge.Foreground = converter.ConvertFromString(
                    group_colors.get(req.group_display, "#888"))
            except:
                pass
            grp_badge.ToolTip = req.group_display
            Grid.SetColumn(grp_badge, 3)
            row_grid.Children.Add(grp_badge)
            
            # Col 4: Categories
            cats_txt = TextBlock()
            cat_str = ", ".join(req.categories[:3])
            if len(req.categories) > 3:
                cat_str += " +{}".format(len(req.categories) - 3)
            cats_txt.Text = cat_str
            cats_txt.FontSize = 9
            cats_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
            cats_txt.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            try:
                cats_txt.Foreground = converter.ConvertFromString("#888")
            except:
                pass
            cats_txt.ToolTip = "\n".join(req.categories)
            Grid.SetColumn(cats_txt, 4)
            row_grid.Children.Add(cats_txt)
            
            # Make row clickable for shift-select
            row_border.Tag = pos
            row_border.MouseLeftButtonDown += self._on_row_click
            row_border.Cursor = System.Windows.Input.Cursors.Hand
            
            row_border.Child = row_grid
            self.spParams.Children.Add(row_border)
    
    def _on_param_toggled(self, sender, args):
        orig_idx = sender.Tag
        if orig_idx is not None and 0 <= orig_idx < len(self.requirements):
            self.requirements[orig_idx].selected = bool(sender.IsChecked)
            # Track last clicked position for shift-select
            for chk, oi, pos in self._checkboxes:
                if oi == orig_idx:
                    self._last_clicked_idx = pos
                    break
    
    def _on_row_click(self, sender, args):
        """Handle row click - supports Shift+Click for range selection"""
        pos = sender.Tag
        if pos is None:
            return
        
        # Check if Shift is held
        shift_held = (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) != 0
        
        if shift_held and self._last_clicked_idx >= 0 and self._last_clicked_idx != pos:
            # Range select between last clicked and current
            start = min(self._last_clicked_idx, pos)
            end = max(self._last_clicked_idx, pos)
            
            # Determine select state from current click target
            target_orig_idx = None
            for chk, oi, p in self._checkboxes:
                if p == pos:
                    target_orig_idx = oi
                    break
            
            if target_orig_idx is not None:
                new_state = not self.requirements[target_orig_idx].selected
                
                # Apply to range
                count = 0
                for chk, oi, p in self._checkboxes:
                    if start <= p <= end:
                        self.requirements[oi].selected = new_state
                        count += 1
                
                self.txtStatus.Text = "{} params {} (Shift+Click range)".format(
                    count, "selected" if new_state else "deselected")
                self._render_params()
        else:
            # Single click - toggle the checkbox
            for chk, oi, p in self._checkboxes:
                if p == pos:
                    new_state = not self.requirements[oi].selected
                    self.requirements[oi].selected = new_state
                    self._last_clicked_idx = pos
                    self._render_params()
                    break
    
    # =================================================================
    # ADD PARAMETERS
    # =================================================================
    def _on_add_params(self, sender, args):
        selected = [r for r in self.requirements if r.selected]
        if not selected:
            System.Windows.MessageBox.Show(
                "No parameters selected.", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        # Check for params without categories
        no_cat_params = [r for r in selected if not r.categories]
        
        if no_cat_params:
            # Show category picker dialog
            result = System.Windows.MessageBox.Show(
                "{} of {} selected parameters have no category assigned.\n\n".format(
                    len(no_cat_params), len(selected)) +
                "Would you like to choose target categories for these parameters?\n\n" +
                "Yes = Pick categories\n" +
                "No = Skip params without categories",
                "Parameters Without Categories",
                MessageBoxButton.YesNo, MessageBoxImage.Question)
            
            if result == MessageBoxResult.Yes:
                picker = CategoryPickerDialog(no_cat_params)
                chosen_cats = picker.show()
                
                if chosen_cats is None:
                    return  # User cancelled picker
                
                # Apply chosen categories to params without categories
                for req in no_cat_params:
                    req.categories = sorted(chosen_cats)
                
                self._render_params()
                self._log("Applied {} categories to {} params".format(
                    len(chosen_cats), len(no_cat_params)))
            else:
                # Deselect params without categories
                for req in no_cat_params:
                    req.selected = False
                selected = [r for r in self.requirements if r.selected]
                if not selected:
                    self.txtStatus.Text = "No parameters with categories to add."
                    return
        
        # Confirm
        result = System.Windows.MessageBox.Show(
            "Add {} parameters to the model?\n\n".format(len(selected)) +
            "This will:\n" +
            "- Create a shared parameter file if needed\n" +
            "- Add parameters as Instance parameters\n" +
            "- Bind to required categories\n\n" +
            "This operation can be undone (Ctrl+Z).",
            "Confirm Add Parameters",
            MessageBoxButton.YesNo, MessageBoxImage.Question)
        
        if result != MessageBoxResult.Yes:
            return
        
        self.window.Cursor = System.Windows.Input.Cursors.Wait
        self.txtStatus.Text = "Adding parameters..."
        
        t = Transaction(doc, "IFC-SG Add Parameters")
        t.Start()
        
        try:
            log = self.adder.add_parameters(self.requirements)
            t.Commit()
            
            # Update log
            for entry in log:
                self._log(entry)
            
            # Re-check existing
            self.adder.check_existing(self.requirements)
            
            # Update stats
            existing = len([r for r in self.requirements if r.already_exists])
            to_add = len(self.requirements) - existing
            self.txtExisting.Text = str(existing)
            self.txtToAdd.Text = str(to_add)
            
            self._render_params()
            
            added_count = len([r for r in self.requirements if r.status in ("added", "updated")])
            failed_count = len([r for r in self.requirements if r.status == "failed"])
            self.txtStatus.Text = "Done! {} added/updated, {} failed".format(added_count, failed_count)
            
        except Exception as e:
            t.RollBack()
            self._log("ERROR: {}".format(str(e)))
            self.txtStatus.Text = "Error: {}".format(str(e))
            System.Windows.MessageBox.Show(
                "Error:\n{}".format(traceback.format_exc()),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self.window.Cursor = System.Windows.Input.Cursors.Arrow
    
    def show(self):
        self.window.ShowDialog()


# =====================================================================
# ENTRY POINT
# =====================================================================
try:
    window = ParamLoaderWindow()
    window.show()
except Exception as e:
    print("Error: {}".format(str(e)))
    print(traceback.format_exc())