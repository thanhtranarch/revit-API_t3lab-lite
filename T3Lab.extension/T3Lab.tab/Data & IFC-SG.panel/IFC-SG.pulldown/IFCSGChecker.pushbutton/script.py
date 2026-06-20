# -*- coding: utf-8 -*-
"""
IFC-SG Parameter Checker v1.0 - DQT
Checks that required IFC+SG parameters exist and have values in Revit model elements.
Supports import from:
  - Autodesk Model Checker XML configuration files
  - Excel parameter mapping files (LTA/BCA format)

Based on CORENET X Code of Practice 3rd Edition September 2025.

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "IFC-SG\nChecker"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Check IFC+SG required parameters. Import rules from Autodesk XML or Excel."

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

# =====================================================================
# PATHS
# =====================================================================
SCRIPT_DIR = os.path.dirname(__file__)
CONFIG_DIR = os.path.join(SCRIPT_DIR, "configs")
REPORTS_DIR = os.path.join(SCRIPT_DIR, "reports")

for d in [CONFIG_DIR, REPORTS_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

# =====================================================================
# CATEGORY NAME MAPPING: Revit Category Name <-> BuiltInCategory
# =====================================================================
CATEGORY_MAP = {
    "Areas": BuiltInCategory.OST_Areas,
    "Generic Models": BuiltInCategory.OST_GenericModel,
    "Plumbing Fixtures": BuiltInCategory.OST_PlumbingFixtures,
    "Project Information": None,  # Special handling
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
# CONFIG PARSER - Parse XML / Excel / JSON into internal format
# =====================================================================
class ParamCheckConfig:
    """
    Internal config format:
    {
        "name": "IFC+SG COP3",
        "source": "XML",
        "disciplines": {
            "ARC": {
                "enabled": True,
                "categories": {
                    "Doors": {
                        "enabled": True,
                        "params": ["ClearHeight", "ClearWidth", "FireRating", ...]
                    }
                }
            }
        }
    }
    """
    
    def __init__(self):
        self.name = ""
        self.source = ""
        self.description = ""
        self.disciplines = {}
    
    @staticmethod
    def from_xml(filepath):
        """Parse Autodesk Model Checker XML configuration"""
        import xml.etree.ElementTree as ET
        
        config = ParamCheckConfig()
        config.source = "XML"
        
        tree = ET.parse(filepath)
        root = tree.getroot()
        config.name = root.get("Name", "Imported XML Config")
        config.description = root.get("Description", "")
        
        for heading in root.findall("Heading"):
            disc_name = heading.get("HeadingText", "")
            disc_enabled = heading.get("IsChecked", "True") == "True"
            
            categories = {}
            for section in heading.findall("Section"):
                cat_name = section.get("SectionName", "")
                cat_enabled = section.get("IsChecked", "True") == "True"
                
                params = []
                for check in section.findall("Check"):
                    param_name = check.get("CheckName", "")
                    if param_name:
                        params.append(param_name)
                
                if params:
                    categories[cat_name] = {
                        "enabled": cat_enabled,
                        "params": sorted(set(params))
                    }
            
            if categories:
                config.disciplines[disc_name] = {
                    "enabled": disc_enabled,
                    "categories": categories
                }
        
        return config
    
    @staticmethod
    def from_excel(filepath):
        """
        Parse Excel parameter mapping file.
        Expected format:
        Column A: Discipline (ARC/STR/MEP)
        Column B: Revit Category
        Column C: Parameter Name
        Column D: Required (Yes/No) [optional]
        """
        config = ParamCheckConfig()
        config.source = "Excel"
        config.name = os.path.splitext(os.path.basename(filepath))[0]
        
        try:
            clr.AddReference('Microsoft.Office.Interop.Excel')
            from Microsoft.Office.Interop import Excel as ExcelInterop
            
            excel_app = ExcelInterop.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            wb = excel_app.Workbooks.Open(filepath)
            ws = wb.Sheets[1]
            
            # Find data range
            used = ws.UsedRange
            rows = used.Rows.Count
            
            for r in range(2, rows + 1):  # Skip header
                disc = str(ws.Cells[r, 1].Value2 or "").strip()
                cat = str(ws.Cells[r, 2].Value2 or "").strip()
                param = str(ws.Cells[r, 3].Value2 or "").strip()
                required = str(ws.Cells[r, 4].Value2 or "Yes").strip().lower()
                
                if not disc or not cat or not param:
                    continue
                if required in ("no", "false", "0"):
                    continue
                
                if disc not in config.disciplines:
                    config.disciplines[disc] = {"enabled": True, "categories": {}}
                if cat not in config.disciplines[disc]["categories"]:
                    config.disciplines[disc]["categories"][cat] = {"enabled": True, "params": []}
                
                if param not in config.disciplines[disc]["categories"][cat]["params"]:
                    config.disciplines[disc]["categories"][cat]["params"].append(param)
            
            wb.Close(False)
            excel_app.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app)
            
        except Exception as e:
            raise Exception("Excel parse error: {}".format(str(e)))
        
        return config
    
    @staticmethod
    def from_json(filepath):
        """Load from saved JSON config"""
        config = ParamCheckConfig()
        with codecs.open(filepath, 'r', 'utf-8') as f:
            data = json.load(f)
        config.name = data.get("name", "")
        config.source = data.get("source", "JSON")
        config.description = data.get("description", "")
        config.disciplines = data.get("disciplines", {})
        return config
    
    def to_json(self, filepath):
        """Save config as JSON"""
        data = {
            "name": self.name,
            "source": self.source,
            "description": self.description,
            "disciplines": self.disciplines
        }
        with codecs.open(filepath, 'w', 'utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    def get_total_stats(self):
        """Return total disciplines, categories, parameters"""
        total_disc = len(self.disciplines)
        total_cat = 0
        total_param = 0
        for d in self.disciplines.values():
            cats = d.get("categories", {})
            total_cat += len(cats)
            for c in cats.values():
                total_param += len(c.get("params", []))
        return total_disc, total_cat, total_param


# =====================================================================
# CHECKER ENGINE - Run parameter checks against model
# =====================================================================
class CheckResult:
    """Result for one category check"""
    def __init__(self, discipline, category, param_name, status, 
                 total_elements=0, missing_count=0, element_ids=None):
        self.discipline = discipline
        self.category = category
        self.param_name = param_name
        self.status = status  # "pass", "fail", "warning", "skip", "no_elements"
        self.total_elements = total_elements
        self.missing_count = missing_count
        self.element_ids = element_ids or []


class ParamChecker:
    """Check IFC+SG parameters in Revit model"""
    
    def __init__(self, document):
        self.doc = document
        self._element_cache = {}  # category -> list of elements
    
    def _get_elements(self, category_name):
        """Get all instance elements for a Revit category"""
        if category_name in self._element_cache:
            return self._element_cache[category_name]
        
        bic = CATEGORY_MAP.get(category_name)
        elements = []
        
        if category_name == "Project Information":
            # Special: only 1 element
            elements = [self.doc.ProjectInformation]
        elif bic is not None:
            try:
                collector = FilteredElementCollector(self.doc)\
                    .OfCategory(bic)\
                    .WhereElementIsNotElementType()
                elements = list(collector)
            except:
                elements = []
        
        self._element_cache[category_name] = elements
        return elements
    
    def _check_param_has_value(self, element, param_name):
        """Check if an element has a parameter with a non-empty value"""
        # Try by name
        for p in element.Parameters:
            if p.Definition.Name == param_name:
                if not p.HasValue:
                    return False
                if p.StorageType == StorageType.String:
                    val = p.AsString()
                    return val is not None and val.strip() != ""
                elif p.StorageType == StorageType.Integer:
                    return True  # Has value
                elif p.StorageType == StorageType.Double:
                    return True
                elif p.StorageType == StorageType.ElementId:
                    return p.AsElementId() != ElementId.InvalidElementId
                return True
        
        # Parameter not found on element
        return False
    
    def run_check(self, config, progress_callback=None):
        """
        Run all parameter checks.
        Returns list of CheckResult per discipline/category/param.
        """
        results = []
        self._element_cache = {}
        
        total_checks = 0
        for d_data in config.disciplines.values():
            if not d_data.get("enabled", True):
                continue
            for c_data in d_data.get("categories", {}).values():
                if not c_data.get("enabled", True):
                    continue
                total_checks += len(c_data.get("params", []))
        
        current = 0
        
        for disc_name, disc_data in config.disciplines.items():
            if not disc_data.get("enabled", True):
                continue
            
            for cat_name, cat_data in disc_data.get("categories", {}).items():
                if not cat_data.get("enabled", True):
                    continue
                
                elements = self._get_elements(cat_name)
                
                if not elements:
                    for param_name in cat_data.get("params", []):
                        results.append(CheckResult(
                            disc_name, cat_name, param_name,
                            "no_elements", 0, 0))
                        current += 1
                        if progress_callback:
                            progress_callback(current, total_checks)
                    continue
                
                for param_name in cat_data.get("params", []):
                    missing_ids = []
                    total = len(elements)
                    
                    for el in elements:
                        try:
                            if not self._check_param_has_value(el, param_name):
                                missing_ids.append(_eid_int(el.Id))
                        except:
                            pass
                    
                    missing = len(missing_ids)
                    
                    if missing == 0:
                        status = "pass"
                    elif missing == total:
                        # All missing - likely parameter doesn't exist
                        status = "fail"
                    else:
                        status = "warning"  # Partial
                    
                    results.append(CheckResult(
                        disc_name, cat_name, param_name,
                        status, total, missing, missing_ids[:100]))
                    
                    current += 1
                    if progress_callback:
                        progress_callback(current, total_checks)
        
        return results


# =====================================================================
# EXCEL REPORT
# =====================================================================
class ExcelReporter:
    """Generate Excel report for IFC-SG parameter check"""
    
    def __init__(self, doc):
        self.doc = doc
    
    def _rgb(self, r, g, b):
        return r + (g * 256) + (b * 256 * 256)
    
    def generate(self, config, results, filepath):
        clr.AddReference('Microsoft.Office.Interop.Excel')
        from Microsoft.Office.Interop import Excel as ExcelInterop
        
        excel_app = ExcelInterop.ApplicationClass()
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        
        try:
            wb = excel_app.Workbooks.Add()
            
            # --- Sheet 1: Summary ---
            ws = wb.Sheets[1]
            ws.Name = "Summary"
            
            ws.Cells[1, 1].Value2 = "IFC-SG PARAMETER CHECK REPORT"
            ws.Cells[1, 1].Font.Size = 16
            ws.Cells[1, 1].Font.Bold = True
            ws.Range["A1:E1"].Merge()
            ws.Range["A1:E1"].Interior.Color = self._rgb(240, 204, 136)
            
            row = 3
            info = [
                ("Project", self.doc.ProjectInformation.Name or "N/A"),
                ("Config", config.name),
                ("Source", config.source),
                ("Date", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ]
            for label, val in info:
                ws.Cells[row, 1].Value2 = label
                ws.Cells[row, 1].Font.Bold = True
                ws.Cells[row, 2].Value2 = val
                row += 1
            
            row += 1
            # Stats
            total = len(results)
            passed = len([r for r in results if r.status == "pass"])
            failed = len([r for r in results if r.status == "fail"])
            warning = len([r for r in results if r.status == "warning"])
            no_elem = len([r for r in results if r.status == "no_elements"])
            
            stats = [("Total Checks", total), ("Passed", passed),
                     ("Failed (all missing)", failed), ("Warning (partial)", warning),
                     ("No Elements", no_elem)]
            for label, val in stats:
                ws.Cells[row, 1].Value2 = label
                ws.Cells[row, 1].Font.Bold = True
                ws.Cells[row, 2].Value2 = val
                row += 1
            
            ws.Columns["A:E"].AutoFit()
            
            # --- Sheet 2: Detailed Results ---
            ws2 = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
            ws2.Name = "Detailed Results"
            
            headers = ["Discipline", "Category", "Parameter", "Status",
                       "Total Elements", "Missing Count", "Element IDs (sample)"]
            for i, h in enumerate(headers, 1):
                ws2.Cells[1, i].Value2 = h
                ws2.Cells[1, i].Font.Bold = True
                ws2.Cells[1, i].Interior.Color = self._rgb(240, 204, 136)
            
            row = 2
            status_colors = {
                "pass": self._rgb(200, 230, 201),
                "fail": self._rgb(255, 205, 210),
                "warning": self._rgb(255, 236, 179),
                "no_elements": self._rgb(224, 224, 224),
            }
            
            for r in results:
                ws2.Cells[row, 1].Value2 = r.discipline
                ws2.Cells[row, 2].Value2 = r.category
                ws2.Cells[row, 3].Value2 = r.param_name
                ws2.Cells[row, 4].Value2 = r.status.upper()
                ws2.Cells[row, 5].Value2 = r.total_elements
                ws2.Cells[row, 6].Value2 = r.missing_count
                ws2.Cells[row, 7].Value2 = ", ".join(str(eid) for eid in r.element_ids[:20])
                
                color = status_colors.get(r.status)
                if color:
                    ws2.Cells[row, 4].Interior.Color = color
                row += 1
            
            ws2.Columns["A:G"].AutoFit()
            
            # --- Sheet 3: Failed Only ---
            ws3 = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
            ws3.Name = "Failed Parameters"
            
            fail_headers = ["Discipline", "Category", "Parameter", "Missing Count", "Total Elements"]
            for i, h in enumerate(fail_headers, 1):
                ws3.Cells[1, i].Value2 = h
                ws3.Cells[1, i].Font.Bold = True
                ws3.Cells[1, i].Interior.Color = self._rgb(255, 205, 210)
            
            row = 2
            for r in results:
                if r.status in ("fail", "warning"):
                    ws3.Cells[row, 1].Value2 = r.discipline
                    ws3.Cells[row, 2].Value2 = r.category
                    ws3.Cells[row, 3].Value2 = r.param_name
                    ws3.Cells[row, 4].Value2 = r.missing_count
                    ws3.Cells[row, 5].Value2 = r.total_elements
                    row += 1
            
            if row == 2:
                ws3.Cells[2, 1].Value2 = "All parameters passed!"
                ws3.Range["A2:E2"].Merge()
            
            ws3.Columns["A:E"].AutoFit()
            
            wb.SaveAs(filepath)
            wb.Close()
            excel_app.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app)
            return True
            
        except Exception as e:
            try:
                wb.Close(False)
                excel_app.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app)
            except:
                pass
            raise e


# # =====================================================================
# XAML PATH DEFINITION
# =====================================================================
import os
import io

current_dir = os.path.dirname(__file__)
while current_dir and not current_dir.endswith('T3Lab.extension'):
    parent = os.path.dirname(current_dir)
    if parent == current_dir:
        break
    current_dir = parent
xaml_path = os.path.join(current_dir, "lib", "GUI", "Tools", "IFCSGChecker.xaml")


# =====================================================================
# MAIN WINDOW
# =====================================================================
class IFCSGCheckerWindow:
    """Main IFC-SG Parameter Checker window"""
    
    def __init__(self):
        self.config = None
        self.results = None
        self.all_results = []
        self.checker = ParamChecker(doc)
        self.reporter = ExcelReporter(doc)
        
        # Parse XAML
        with io.open(xaml_path, 'r', encoding='utf-8') as f:
            xaml_content = f.read()
        xr = XmlReader.Create(StringReader(xaml_content))
        self.window = XamlReader.Load(xr)
        
        self._get_controls()
        self._bind_events()
        self._load_saved_configs()
        
        self.txtFooter.Text = "{} | {}".format(
            doc.ProjectInformation.Name or "Untitled",
            datetime.datetime.now().strftime("%Y-%m-%d"))
    
    def _get_controls(self):
        names = [
            "cmbConfig", "btnImportXML", "btnImportExcel", "btnSaveConfig", "btnDeleteConfig",
            "txtTotalParams", "txtCategories", "txtPassed", "txtFailed", "txtWarning", "txtNoElem",
            "tvCategories", "btnExpandAll", "btnCollapseAll",
            "txtResultHeader", "spResults",
            "btnFilterAll", "btnFilterFail", "btnFilterWarn", "btnFilterPass", 
            "btnSelectAllFailed", "txtSearch",
            "txtStatus", "btnRunCheck", "btnExportExcel", "btnClose", "txtFooter"
        ]
        for name in names:
            setattr(self, name, self.window.FindName(name))
    
    def _bind_events(self):
        self.btnImportXML.Click += self._on_import_xml
        self.btnImportExcel.Click += self._on_import_excel
        self.btnSaveConfig.Click += self._on_save_config
        self.btnDeleteConfig.Click += self._on_delete_config
        self.cmbConfig.SelectionChanged += self._on_config_changed
        self.btnExpandAll.Click += self._on_expand_all
        self.btnCollapseAll.Click += self._on_collapse_all
        self.btnFilterAll.Click += lambda s, e: self._apply_filter("all")
        self.btnFilterFail.Click += lambda s, e: self._apply_filter("fail")
        self.btnFilterWarn.Click += lambda s, e: self._apply_filter("warning")
        self.btnFilterPass.Click += lambda s, e: self._apply_filter("pass")
        self.txtSearch.TextChanged += lambda s, e: self._apply_filter(self._current_filter)
        self.btnSelectAllFailed.Click += self._on_select_all_failed
        self.btnRunCheck.Click += self._on_run_check
        self.btnExportExcel.Click += self._on_export_excel
        self.btnClose.Click += lambda s, e: self.window.Close()
        self._current_filter = "all"
    
    # =================================================================
    # CONFIG MANAGEMENT
    # =================================================================
    def _load_saved_configs(self):
        self.cmbConfig.Items.Clear()
        if os.path.exists(CONFIG_DIR):
            for f in sorted(os.listdir(CONFIG_DIR)):
                if f.endswith('.json'):
                    self.cmbConfig.Items.Add(os.path.splitext(f)[0])
        if self.cmbConfig.Items.Count > 0:
            self.cmbConfig.SelectedIndex = 0
    
    def _on_config_changed(self, sender, args):
        sel = self.cmbConfig.SelectedItem
        if sel:
            path = os.path.join(CONFIG_DIR, str(sel) + ".json")
            try:
                self.config = ParamCheckConfig.from_json(path)
                self._refresh_tree()
                self._update_config_stats()
                self.btnRunCheck.IsEnabled = True
                self.txtStatus.Text = "Config loaded: {} ({})".format(
                    self.config.name, self.config.source)
            except Exception as e:
                self.txtStatus.Text = "Error loading config: {}".format(str(e))
    
    def _on_import_xml(self, sender, args):
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dlg = OpenFileDialog()
        dlg.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        dlg.Title = "Import Autodesk Model Checker XML"
        
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                self.config = ParamCheckConfig.from_xml(dlg.FileName)
                # Auto-save as JSON
                name = os.path.splitext(os.path.basename(dlg.FileName))[0]
                save_path = os.path.join(CONFIG_DIR, name + ".json")
                self.config.to_json(save_path)
                
                self._load_saved_configs()
                # Select the new one
                for i in range(self.cmbConfig.Items.Count):
                    if str(self.cmbConfig.Items[i]) == name:
                        self.cmbConfig.SelectedIndex = i
                        break
                
                d, c, p = self.config.get_total_stats()
                self.txtStatus.Text = "Imported XML: {} disciplines, {} categories, {} params".format(d, c, p)
            except Exception as e:
                System.Windows.MessageBox.Show(
                    "Error importing XML:\n{}".format(str(e)),
                    "Import Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_import_excel(self, sender, args):
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dlg = OpenFileDialog()
        dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
        dlg.Title = "Import Excel Parameter Mapping"
        
        if dlg.ShowDialog() == DialogResult.OK:
            try:
                self.config = ParamCheckConfig.from_excel(dlg.FileName)
                name = os.path.splitext(os.path.basename(dlg.FileName))[0]
                save_path = os.path.join(CONFIG_DIR, name + ".json")
                self.config.to_json(save_path)
                
                self._load_saved_configs()
                for i in range(self.cmbConfig.Items.Count):
                    if str(self.cmbConfig.Items[i]) == name:
                        self.cmbConfig.SelectedIndex = i
                        break
                
                d, c, p = self.config.get_total_stats()
                self.txtStatus.Text = "Imported Excel: {} disciplines, {} categories, {} params".format(d, c, p)
            except Exception as e:
                System.Windows.MessageBox.Show(
                    "Error importing Excel:\n{}".format(str(e)),
                    "Import Error", MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_save_config(self, sender, args):
        if not self.config:
            return
        from System.Windows.Forms import SaveFileDialog, DialogResult
        dlg = SaveFileDialog()
        dlg.Filter = "JSON Files (*.json)|*.json"
        dlg.Title = "Save Config"
        dlg.InitialDirectory = CONFIG_DIR
        if dlg.ShowDialog() == DialogResult.OK:
            self.config.to_json(dlg.FileName)
            self.txtStatus.Text = "Config saved: {}".format(dlg.FileName)
    
    def _on_delete_config(self, sender, args):
        sel = self.cmbConfig.SelectedItem
        if not sel:
            return
        result = System.Windows.MessageBox.Show(
            "Delete config '{}'?".format(sel),
            "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        if result == MessageBoxResult.Yes:
            path = os.path.join(CONFIG_DIR, str(sel) + ".json")
            if os.path.exists(path):
                os.remove(path)
            self._load_saved_configs()
    
    # =================================================================
    # TREE VIEW
    # =================================================================
    def _refresh_tree(self):
        self.tvCategories.Items.Clear()
        if not self.config:
            return
        
        converter = BrushConverter()
        
        for disc_name, disc_data in self.config.disciplines.items():
            # Discipline node
            disc_item = TreeViewItem()
            disc_item.IsExpanded = True
            
            disc_sp = StackPanel()
            disc_sp.Orientation = System.Windows.Controls.Orientation.Horizontal
            
            chk_disc = CheckBox()
            chk_disc.IsChecked = System.Nullable[System.Boolean](
                bool(disc_data.get("enabled", True)))
            chk_disc.Margin = System.Windows.Thickness(0, 0, 6, 0)
            chk_disc.Tag = disc_name
            chk_disc.Checked += self._on_disc_toggled
            chk_disc.Unchecked += self._on_disc_toggled
            
            lbl_disc = TextBlock()
            lbl_disc.Text = u"{} ({} categories)".format(
                disc_name, len(disc_data.get("categories", {})))
            lbl_disc.FontWeight = System.Windows.FontWeights.Bold
            lbl_disc.FontSize = 12
            try:
                lbl_disc.Foreground = converter.ConvertFromString("#5D4E37")
            except:
                pass
            
            disc_sp.Children.Add(chk_disc)
            disc_sp.Children.Add(lbl_disc)
            disc_item.Header = disc_sp
            
            # Category nodes
            for cat_name, cat_data in disc_data.get("categories", {}).items():
                cat_item = TreeViewItem()
                
                cat_sp = StackPanel()
                cat_sp.Orientation = System.Windows.Controls.Orientation.Horizontal
                
                chk_cat = CheckBox()
                chk_cat.IsChecked = System.Nullable[System.Boolean](
                    bool(cat_data.get("enabled", True)))
                chk_cat.Margin = System.Windows.Thickness(0, 0, 6, 0)
                chk_cat.Tag = "{}|{}".format(disc_name, cat_name)
                chk_cat.Checked += self._on_cat_toggled
                chk_cat.Unchecked += self._on_cat_toggled
                
                param_count = len(cat_data.get("params", []))
                lbl_cat = TextBlock()
                lbl_cat.Text = u"{} ({} params)".format(cat_name, param_count)
                lbl_cat.FontSize = 11
                
                cat_sp.Children.Add(chk_cat)
                cat_sp.Children.Add(lbl_cat)
                cat_item.Header = cat_sp
                
                disc_item.Items.Add(cat_item)
            
            self.tvCategories.Items.Add(disc_item)
    
    def _on_disc_toggled(self, sender, args):
        disc_name = str(sender.Tag)
        if disc_name in self.config.disciplines:
            self.config.disciplines[disc_name]["enabled"] = bool(sender.IsChecked)
    
    def _on_cat_toggled(self, sender, args):
        tag = str(sender.Tag)
        parts = tag.split("|")
        if len(parts) == 2:
            disc, cat = parts
            if disc in self.config.disciplines:
                cats = self.config.disciplines[disc].get("categories", {})
                if cat in cats:
                    cats[cat]["enabled"] = bool(sender.IsChecked)
    
    def _on_expand_all(self, sender, args):
        for item in self.tvCategories.Items:
            item.IsExpanded = True
    
    def _on_collapse_all(self, sender, args):
        for item in self.tvCategories.Items:
            item.IsExpanded = False
    
    def _update_config_stats(self):
        if self.config:
            d, c, p = self.config.get_total_stats()
            self.txtTotalParams.Text = str(p)
            self.txtCategories.Text = str(c)
    
    # =================================================================
    # RUN CHECK
    # =================================================================
    def _on_run_check(self, sender, args):
        if not self.config:
            return
        
        self.txtStatus.Text = "Running IFC-SG parameter checks..."
        self.window.Cursor = System.Windows.Input.Cursors.Wait
        
        try:
            self.results = self.checker.run_check(self.config)
            self.all_results = list(self.results)
            
            # Update cards
            passed = len([r for r in self.results if r.status == "pass"])
            failed = len([r for r in self.results if r.status == "fail"])
            warning = len([r for r in self.results if r.status == "warning"])
            no_elem = len([r for r in self.results if r.status == "no_elements"])
            
            self.txtPassed.Text = str(passed)
            self.txtFailed.Text = str(failed)
            self.txtWarning.Text = str(warning)
            self.txtNoElem.Text = str(no_elem)
            
            self._current_filter = "all"
            self._render_results(self.results)
            
            self.btnExportExcel.IsEnabled = True
            self.btnSelectAllFailed.IsEnabled = True
            self.txtResultHeader.Text = "Check Results ({} checks)".format(len(self.results))
            self.txtStatus.Text = "Done: {} passed, {} failed, {} partial, {} no elements".format(
                passed, failed, warning, no_elem)
            
        except Exception as e:
            self.txtStatus.Text = "Error: {}".format(str(e))
            System.Windows.MessageBox.Show(
                "Error:\n{}".format(traceback.format_exc()),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self.window.Cursor = System.Windows.Input.Cursors.Arrow
    
    def _apply_filter(self, filter_type):
        self._current_filter = filter_type
        if not self.all_results:
            return
        
        search_text = self.txtSearch.Text.strip().lower() if self.txtSearch.Text else ""
        
        filtered = []
        for r in self.all_results:
            # Status filter
            if filter_type == "fail" and r.status not in ("fail",):
                continue
            if filter_type == "warning" and r.status not in ("warning",):
                continue
            if filter_type == "pass" and r.status not in ("pass",):
                continue
            
            # Search filter
            if search_text:
                searchable = "{}{}{}".format(
                    r.discipline, r.category, r.param_name).lower()
                if search_text not in searchable:
                    continue
            
            filtered.append(r)
        
        self._render_results(filtered)
    
    def _on_select_all_failed(self, sender, args):
        """Select ALL failed elements across all categories in Revit"""
        if not self.all_results:
            return
        all_ids = []
        for r in self.all_results:
            if r.status in ("fail", "warning"):
                all_ids.extend(r.element_ids)
        unique_ids = list(set(all_ids))
        if unique_ids:
            self._select_elements_in_revit(unique_ids[:2000])
            self.txtStatus.Text = "Selected {} failed elements in Revit".format(len(unique_ids))
        else:
            self.txtStatus.Text = "No failed elements to select"
    
    def _select_elements_in_revit(self, element_ids):
        """Select elements in Revit and zoom to them"""
        try:
            ids = System.Collections.Generic.List[ElementId]()
            for eid in element_ids:
                try:
                    ids.Add(ElementId(int(eid)))  # Revit 2026: accepts Int64
                except:
                    pass
            if ids.Count > 0:
                uidoc.Selection.SetElementIds(ids)
                self.txtStatus.Text = "Selected {} elements in Revit".format(ids.Count)
        except Exception as e:
            self.txtStatus.Text = "Select error: {}".format(str(e))
    
    def _compute_category_stats(self, results):
        """Compute % completion per discipline > category"""
        stats = {}  # "disc|cat" -> {total, passed, failed, warning, no_elem, pct}
        for r in results:
            key = "{}|{}".format(r.discipline, r.category)
            if key not in stats:
                stats[key] = {"total": 0, "pass": 0, "fail": 0, "warning": 0, "no_elements": 0}
            stats[key]["total"] += 1
            stats[key][r.status] = stats[key].get(r.status, 0) + 1
        
        for key, s in stats.items():
            checkable = s["total"] - s.get("no_elements", 0)
            if checkable > 0:
                s["pct"] = int(round(s["pass"] / float(checkable) * 100))
            else:
                s["pct"] = -1  # No elements to check
        return stats
    
    def _render_results(self, results):
        self.spResults.Children.Clear()
        converter = BrushConverter()
        
        status_bg = {
            "pass": "#E8F5E9", "fail": "#FFEBEE",
            "warning": "#FFF8E1", "no_elements": "#ECEFF1"
        }
        status_fg = {
            "pass": "#2E7D32", "fail": "#C62828",
            "warning": "#F57F17", "no_elements": "#78909C"
        }
        status_icon = {
            "pass": u"\u2714", "fail": u"\u2718",
            "warning": u"\u26A0", "no_elements": u"\u23F8"
        }
        
        # Compute category stats for progress bars
        cat_stats = self._compute_category_stats(self.all_results)
        
        current_disc = ""
        current_cat = ""
        
        for r in results:
            # Discipline header
            if r.discipline != current_disc:
                current_disc = r.discipline
                current_cat = ""
                
                disc_border = System.Windows.Controls.Border()
                disc_border.Margin = System.Windows.Thickness(0, 8, 0, 2)
                disc_border.Padding = System.Windows.Thickness(8, 4, 8, 4)
                try:
                    disc_border.Background = converter.ConvertFromString("#0F172A")
                except:
                    pass
                disc_border.CornerRadius = System.Windows.CornerRadius(3)
                
                disc_txt = TextBlock()
                disc_txt.Text = r.discipline
                disc_txt.FontWeight = System.Windows.FontWeights.Bold
                disc_txt.FontSize = 13
                try:
                    disc_txt.Foreground = converter.ConvertFromString("#5D4E37")
                except:
                    pass
                disc_border.Child = disc_txt
                self.spResults.Children.Add(disc_border)
            
            # Category header with progress bar
            if r.category != current_cat:
                current_cat = r.category
                stat_key = "{}|{}".format(r.discipline, r.category)
                stat = cat_stats.get(stat_key, {})
                pct = stat.get("pct", 0)
                cat_pass = stat.get("pass", 0)
                cat_total = stat.get("total", 0)
                cat_no_elem = stat.get("no_elements", 0)
                cat_fail = stat.get("fail", 0)
                cat_warn = stat.get("warning", 0)
                
                # Category container
                cat_border = System.Windows.Controls.Border()
                cat_border.Margin = System.Windows.Thickness(0, 4, 0, 2)
                cat_border.Padding = System.Windows.Thickness(4, 3, 4, 3)
                cat_border.CornerRadius = System.Windows.CornerRadius(3)
                try:
                    cat_border.Background = converter.ConvertFromString("#F9F6EE")
                    cat_border.BorderBrush = converter.ConvertFromString("#E8E0D0")
                except:
                    pass
                cat_border.BorderThickness = System.Windows.Thickness(1)
                
                cat_grid = Grid()
                cg1 = ColumnDefinition()
                cg1.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
                cg2 = ColumnDefinition()
                cg2.Width = System.Windows.GridLength(200)
                cg3 = ColumnDefinition()
                cg3.Width = System.Windows.GridLength(80)
                cat_grid.ColumnDefinitions.Add(cg1)
                cat_grid.ColumnDefinitions.Add(cg2)
                cat_grid.ColumnDefinitions.Add(cg3)
                
                # Category name + counts
                cat_info = StackPanel()
                cat_name_txt = TextBlock()
                cat_name_txt.Text = u"\u25B8 {}".format(r.category)
                cat_name_txt.FontWeight = System.Windows.FontWeights.SemiBold
                cat_name_txt.FontSize = 11
                try:
                    cat_name_txt.Foreground = converter.ConvertFromString("#5D4E37")
                except:
                    pass
                cat_info.Children.Add(cat_name_txt)
                
                # Sub stats text
                sub_parts = []
                if cat_pass > 0:
                    sub_parts.append("{} pass".format(cat_pass))
                if cat_fail > 0:
                    sub_parts.append("{} fail".format(cat_fail))
                if cat_warn > 0:
                    sub_parts.append("{} partial".format(cat_warn))
                if cat_no_elem > 0:
                    sub_parts.append("{} N/A".format(cat_no_elem))
                
                sub_txt = TextBlock()
                sub_txt.Text = " | ".join(sub_parts)
                sub_txt.FontSize = 9
                try:
                    sub_txt.Foreground = converter.ConvertFromString("#999")
                except:
                    pass
                cat_info.Children.Add(sub_txt)
                Grid.SetColumn(cat_info, 0)
                cat_grid.Children.Add(cat_info)
                
                # Progress bar
                if pct >= 0:
                    prog_sp = StackPanel()
                    prog_sp.VerticalAlignment = System.Windows.VerticalAlignment.Center
                    prog_sp.Margin = System.Windows.Thickness(4, 0, 4, 0)
                    
                    # Bar background
                    bar_border = System.Windows.Controls.Border()
                    bar_border.Height = 10
                    bar_border.CornerRadius = System.Windows.CornerRadius(5)
                    try:
                        bar_border.Background = converter.ConvertFromString("#E0E0E0")
                    except:
                        pass
                    
                    # Bar fill
                    bar_grid = Grid()
                    bar_bg = System.Windows.Controls.Border()
                    bar_bg.Height = 10
                    bar_bg.CornerRadius = System.Windows.CornerRadius(5)
                    try:
                        bar_bg.Background = converter.ConvertFromString("#E0E0E0")
                    except:
                        pass
                    bar_grid.Children.Add(bar_bg)
                    
                    bar_fill = System.Windows.Controls.Border()
                    bar_fill.Height = 10
                    bar_fill.CornerRadius = System.Windows.CornerRadius(5)
                    bar_fill.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
                    # Width as percentage
                    bar_fill.Width = max(1, pct * 1.8)  # 180px max width
                    
                    if pct >= 80:
                        fill_color = "#66BB6A"
                    elif pct >= 50:
                        fill_color = "#FFA726"
                    else:
                        fill_color = "#EF5350"
                    try:
                        bar_fill.Background = converter.ConvertFromString(fill_color)
                    except:
                        pass
                    bar_grid.Children.Add(bar_fill)
                    
                    prog_sp.Children.Add(bar_grid)
                    
                    Grid.SetColumn(prog_sp, 1)
                    cat_grid.Children.Add(prog_sp)
                
                # Percentage text + Select All Failed button
                pct_sp = StackPanel()
                pct_sp.VerticalAlignment = System.Windows.VerticalAlignment.Center
                pct_sp.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
                
                if pct >= 0:
                    pct_txt = TextBlock()
                    pct_txt.Text = "{}%".format(pct)
                    pct_txt.FontSize = 12
                    pct_txt.FontWeight = System.Windows.FontWeights.Bold
                    pct_txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
                    try:
                        if pct >= 80:
                            pct_txt.Foreground = converter.ConvertFromString("#2E7D32")
                        elif pct >= 50:
                            pct_txt.Foreground = converter.ConvertFromString("#F57F17")
                        else:
                            pct_txt.Foreground = converter.ConvertFromString("#C62828")
                    except:
                        pass
                    pct_sp.Children.Add(pct_txt)
                else:
                    na_txt = TextBlock()
                    na_txt.Text = "N/A"
                    na_txt.FontSize = 11
                    na_txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
                    try:
                        na_txt.Foreground = converter.ConvertFromString("#999")
                    except:
                        pass
                    pct_sp.Children.Add(na_txt)
                
                # Select all failed in this category
                if cat_fail > 0 or cat_warn > 0:
                    all_fail_ids = []
                    for ar in self.all_results:
                        if ar.discipline == r.discipline and ar.category == r.category:
                            if ar.status in ("fail", "warning"):
                                all_fail_ids.extend(ar.element_ids)
                    
                    if all_fail_ids:
                        sel_all_btn = Button()
                        sel_all_btn.Content = "Select"
                        sel_all_btn.FontSize = 9
                        sel_all_btn.Padding = System.Windows.Thickness(4, 1, 4, 1)
                        sel_all_btn.Margin = System.Windows.Thickness(0, 2, 0, 0)
                        sel_all_btn.Cursor = System.Windows.Input.Cursors.Hand
                        sel_all_btn.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
                        try:
                            sel_all_btn.Background = converter.ConvertFromString("#FFCDD2")
                            sel_all_btn.Foreground = converter.ConvertFromString("#C62828")
                            sel_all_btn.BorderBrush = converter.ConvertFromString("#EF9A9A")
                        except:
                            pass
                        sel_all_btn.BorderThickness = System.Windows.Thickness(1)
                        # Store IDs - deduplicate
                        unique_ids = list(set(all_fail_ids))[:500]
                        sel_all_btn.Tag = unique_ids
                        sel_all_btn.Click += self._on_select_btn_click
                        pct_sp.Children.Add(sel_all_btn)
                
                Grid.SetColumn(pct_sp, 2)
                cat_grid.Children.Add(pct_sp)
                
                cat_border.Child = cat_grid
                self.spResults.Children.Add(cat_border)
            
            # Parameter row
            row_border = System.Windows.Controls.Border()
            row_border.Margin = System.Windows.Thickness(16, 1, 0, 1)
            row_border.Padding = System.Windows.Thickness(8, 3, 8, 3)
            row_border.CornerRadius = System.Windows.CornerRadius(2)
            try:
                row_border.Background = converter.ConvertFromString(
                    status_bg.get(r.status, "#FAFAFA"))
            except:
                pass
            
            row_grid = Grid()
            c1 = ColumnDefinition()
            c1.Width = System.Windows.GridLength(28)
            c2 = ColumnDefinition()
            c2.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            c3 = ColumnDefinition()
            c3.Width = System.Windows.GridLength(120)
            c4 = ColumnDefinition()
            c4.Width = System.Windows.GridLength(55)
            row_grid.ColumnDefinitions.Add(c1)
            row_grid.ColumnDefinitions.Add(c2)
            row_grid.ColumnDefinitions.Add(c3)
            row_grid.ColumnDefinitions.Add(c4)
            
            # Icon
            icon = TextBlock()
            icon.Text = status_icon.get(r.status, "?")
            icon.FontSize = 12
            icon.VerticalAlignment = System.Windows.VerticalAlignment.Center
            try:
                icon.Foreground = converter.ConvertFromString(
                    status_fg.get(r.status, "#666"))
            except:
                pass
            Grid.SetColumn(icon, 0)
            row_grid.Children.Add(icon)
            
            # Param name
            name_txt = TextBlock()
            name_txt.Text = r.param_name
            name_txt.FontSize = 11
            name_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
            Grid.SetColumn(name_txt, 1)
            row_grid.Children.Add(name_txt)
            
            # Count info
            if r.status == "no_elements":
                count_text = "No elements"
            elif r.status == "pass":
                count_text = "{} OK".format(r.total_elements)
            else:
                count_text = "{}/{} missing".format(r.missing_count, r.total_elements)
            
            count_txt = TextBlock()
            count_txt.Text = count_text
            count_txt.FontSize = 10
            count_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
            count_txt.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            try:
                count_txt.Foreground = converter.ConvertFromString(
                    status_fg.get(r.status, "#888"))
            except:
                pass
            Grid.SetColumn(count_txt, 2)
            row_grid.Children.Add(count_txt)
            
            # Select button for failed/warning params
            if r.status in ("fail", "warning") and r.element_ids:
                sel_btn = Button()
                sel_btn.Content = u"\u25BA Select"
                sel_btn.FontSize = 9
                sel_btn.Padding = System.Windows.Thickness(3, 1, 3, 1)
                sel_btn.VerticalAlignment = System.Windows.VerticalAlignment.Center
                sel_btn.Cursor = System.Windows.Input.Cursors.Hand
                try:
                    sel_btn.Background = converter.ConvertFromString("#FFF3E0")
                    sel_btn.Foreground = converter.ConvertFromString("#E65100")
                    sel_btn.BorderBrush = converter.ConvertFromString("#FFCC80")
                except:
                    pass
                sel_btn.BorderThickness = System.Windows.Thickness(1)
                sel_btn.Tag = list(r.element_ids)[:200]
                sel_btn.Click += self._on_select_btn_click
                Grid.SetColumn(sel_btn, 3)
                row_grid.Children.Add(sel_btn)
            
            row_border.Child = row_grid
            self.spResults.Children.Add(row_border)
    
    def _on_select_btn_click(self, sender, args):
        """Handle select button click - select elements in Revit"""
        ids = sender.Tag
        if ids:
            self._select_elements_in_revit(ids)
    
    # =================================================================
    # EXPORT
    # =================================================================
    def _on_export_excel(self, sender, args):
        if not self.results or not self.config:
            return
        from System.Windows.Forms import SaveFileDialog, DialogResult
        
        dlg = SaveFileDialog()
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FileName = "IFC-SG_Check_{}_{}".format(
            doc.ProjectInformation.Name or "Project",
            datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        dlg.InitialDirectory = REPORTS_DIR
        
        if dlg.ShowDialog() == DialogResult.OK:
            self.txtStatus.Text = "Exporting..."
            self.window.Cursor = System.Windows.Input.Cursors.Wait
            try:
                self.reporter.generate(self.config, self.all_results, dlg.FileName)
                self.txtStatus.Text = "Exported: {}".format(os.path.basename(dlg.FileName))
                result = System.Windows.MessageBox.Show(
                    "Report exported!\nOpen now?", "Done",
                    MessageBoxButton.YesNo, MessageBoxImage.Information)
                if result == MessageBoxResult.Yes:
                    os.startfile(dlg.FileName)
            except Exception as e:
                System.Windows.MessageBox.Show(
                    "Export error:\n{}".format(str(e)),
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            finally:
                self.window.Cursor = System.Windows.Input.Cursors.Arrow
    
    def show(self):
        self.window.ShowDialog()


# =====================================================================
# ENTRY POINT
# =====================================================================
try:
    window = IFCSGCheckerWindow()
    window.show()
except Exception as e:
    print("Error: {}".format(str(e)))
    print(traceback.format_exc())