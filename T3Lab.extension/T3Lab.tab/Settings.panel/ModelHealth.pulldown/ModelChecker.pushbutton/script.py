# -*- coding: utf-8 -*-
"""
Model Checker v2.0 - DQT
Rule-based BIM model compliance checker with customizable JSON checksets.
Supports project-specific BEP rules, configurable parameters, and Excel reporting.

Phase 1: Project Settings rules
- Revit Version, Project Info, Survey/Base Point coordinates
- Design Options, Worksets, Starting View, True North

Phase 2: Naming Convention + Model Performance rules
- View/Sheet/Level/Grid/Family naming patterns
- File size, Warnings, CAD imports, In-Place families
- RVT links, Groups, Line patterns, Rooms, Duplicate marks

Copyright (c) 2025 Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Model\nChecker"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Rule-based model compliance checker. Customizable JSON checksets for different BEP requirements."

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
from System.Windows import *
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
CHECKSETS_DIR = os.path.join(SCRIPT_DIR, "checksets")
REPORTS_DIR = os.path.join(SCRIPT_DIR, "reports")

# Create directories if not exist
for d in [CHECKSETS_DIR, REPORTS_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

# =====================================================================
# DEFAULT CHECKSET - PROJECT SETTINGS
# =====================================================================
DEFAULT_CHECKSET = {
    "name": "Default Project Settings",
    "version": "1.0",
    "author": "DQT",
    "description": "Standard project settings compliance checks",
    "created": "",
    "modified": "",
    "rules": [
        {
            "id": "PS_001",
            "name": "Revit Version",
            "category": "Project Settings",
            "description": "Check if model is using the expected Revit version",
            "type": "value_match",
            "severity": "warning",
            "enabled": True,
            "params": {
                "expected_version": "2024"
            }
        },
        {
            "id": "PS_002",
            "name": "Project Information - Required Fields",
            "category": "Project Settings",
            "description": "Check that required Project Information fields are filled in",
            "type": "not_empty",
            "severity": "error",
            "enabled": True,
            "params": {
                "fields": [
                    "Project Name",
                    "Project Number",
                    "Client Name",
                    "Project Address"
                ]
            }
        },
        {
            "id": "PS_003",
            "name": "Survey Point - N/S",
            "category": "Project Settings",
            "description": "Check if Survey Point N/S coordinate matches expected value",
            "type": "coordinate_match",
            "severity": "error",
            "enabled": True,
            "params": {
                "point_type": "survey",
                "axis": "NS",
                "expected_value": 0.0,
                "tolerance": 0.001
            }
        },
        {
            "id": "PS_004",
            "name": "Survey Point - E/W",
            "category": "Project Settings",
            "description": "Check if Survey Point E/W coordinate matches expected value",
            "type": "coordinate_match",
            "severity": "error",
            "enabled": True,
            "params": {
                "point_type": "survey",
                "axis": "EW",
                "expected_value": 0.0,
                "tolerance": 0.001
            }
        },
        {
            "id": "PS_005",
            "name": "Survey Point - Elevation",
            "category": "Project Settings",
            "description": "Check if Survey Point elevation matches expected value",
            "type": "coordinate_match",
            "severity": "warning",
            "enabled": True,
            "params": {
                "point_type": "survey",
                "axis": "Elev",
                "expected_value": 0.0,
                "tolerance": 0.001
            }
        },
        {
            "id": "PS_006",
            "name": "Project Base Point - N/S",
            "category": "Project Settings",
            "description": "Check if Project Base Point N/S coordinate matches expected value",
            "type": "coordinate_match",
            "severity": "warning",
            "enabled": True,
            "params": {
                "point_type": "base",
                "axis": "NS",
                "expected_value": 0.0,
                "tolerance": 0.001
            }
        },
        {
            "id": "PS_007",
            "name": "Project Base Point - E/W",
            "category": "Project Settings",
            "description": "Check if Project Base Point E/W coordinate matches expected value",
            "type": "coordinate_match",
            "severity": "warning",
            "enabled": True,
            "params": {
                "point_type": "base",
                "axis": "EW",
                "expected_value": 0.0,
                "tolerance": 0.001
            }
        },
        {
            "id": "PS_008",
            "name": "Project Base Point - Elevation",
            "category": "Project Settings",
            "description": "Check if Project Base Point elevation matches expected value",
            "type": "coordinate_match",
            "severity": "warning",
            "enabled": True,
            "params": {
                "point_type": "base",
                "axis": "Elev",
                "expected_value": 0.0,
                "tolerance": 0.001
            }
        },
        {
            "id": "PS_009",
            "name": "Design Options",
            "category": "Project Settings",
            "description": "Check if model contains Design Options (should be cleaned before submission)",
            "type": "count_check",
            "severity": "warning",
            "enabled": True,
            "params": {
                "target": "design_options",
                "max_count": 0,
                "message_over": "Design Options should be resolved before submission"
            }
        },
        {
            "id": "PS_010",
            "name": "Workset Naming Convention",
            "category": "Project Settings",
            "description": "Check if workset names follow naming convention pattern",
            "type": "naming_pattern",
            "severity": "warning",
            "enabled": False,
            "params": {
                "target": "worksets",
                "pattern": "^[A-Z][a-zA-Z0-9_ -]+$",
                "description": "Workset names should start with uppercase letter"
            }
        },
        {
            "id": "PS_011",
            "name": "Starting View",
            "category": "Project Settings",
            "description": "Check if a Starting View is set for the project",
            "type": "exists_check",
            "severity": "warning",
            "enabled": True,
            "params": {
                "target": "starting_view"
            }
        },
        {
            "id": "PS_012",
            "name": "Project Coordinates - True North",
            "category": "Project Settings",
            "description": "Check True North angle value",
            "type": "value_match_numeric",
            "severity": "info",
            "enabled": True,
            "params": {
                "target": "true_north",
                "expected_value": 0.0,
                "tolerance": 0.01,
                "report_only": True
            }
        }
    ]
}


# =====================================================================
# CHECKSET MANAGER - Load / Save / List JSON checksets
# =====================================================================
class ChecksetManager:
    """Manages JSON checkset files"""
    
    def __init__(self):
        self.checksets_dir = CHECKSETS_DIR
        self._ensure_default()
    
    def _ensure_default(self):
        """Create default checkset if none exist"""
        default_path = os.path.join(self.checksets_dir, "default.json")
        if not os.path.exists(default_path):
            self.save_checkset("default", DEFAULT_CHECKSET)
    
    def list_checksets(self):
        """List all available checkset files"""
        result = []
        if os.path.exists(self.checksets_dir):
            for f in os.listdir(self.checksets_dir):
                if f.endswith('.json'):
                    name = os.path.splitext(f)[0]
                    result.append(name)
        return sorted(result)
    
    def load_checkset(self, name):
        """Load a checkset from JSON file"""
        path = os.path.join(self.checksets_dir, name + ".json")
        if os.path.exists(path):
            with codecs.open(path, 'r', 'utf-8') as f:
                return json.load(f)
        return None
    
    def save_checkset(self, name, data):
        """Save a checkset to JSON file"""
        data["modified"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if not data.get("created"):
            data["created"] = data["modified"]
        path = os.path.join(self.checksets_dir, name + ".json")
        with codecs.open(path, 'w', 'utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return path
    
    def delete_checkset(self, name):
        """Delete a checkset file"""
        path = os.path.join(self.checksets_dir, name + ".json")
        if os.path.exists(path) and name != "default":
            os.remove(path)
            return True
        return False
    
    def duplicate_checkset(self, source_name, new_name):
        """Duplicate a checkset with a new name"""
        data = self.load_checkset(source_name)
        if data:
            data["name"] = new_name
            data["created"] = ""
            self.save_checkset(new_name, data)
            return True
        return False
    
    def import_checkset(self, filepath):
        """Import a checkset from external JSON file"""
        try:
            with codecs.open(filepath, 'r', 'utf-8') as f:
                data = json.load(f)
            if "rules" in data:
                name = os.path.splitext(os.path.basename(filepath))[0]
                self.save_checkset(name, data)
                return name
        except:
            pass
        return None
    
    def export_checkset(self, name, filepath):
        """Export a checkset to external location"""
        data = self.load_checkset(name)
        if data:
            with codecs.open(filepath, 'w', 'utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            return True
        return False


# =====================================================================
# RULE ENGINE - Execute rules against Revit model
# =====================================================================
class RuleResult:
    """Stores the result of a single rule check"""
    def __init__(self, rule, status, message, details=None, elements=None):
        self.rule = rule
        self.rule_id = rule.get("id", "")
        self.rule_name = rule.get("name", "")
        self.category = rule.get("category", "")
        self.severity = rule.get("severity", "info")
        self.status = status  # "pass", "fail", "warning", "info", "error", "skipped"
        self.message = message
        self.details = details or []  # list of detail strings
        self.elements = elements or []  # list of element ids


class RuleEngine:
    """Executes rules against the Revit model"""
    
    def __init__(self, document):
        self.doc = document
        self.app = document.Application
    
    def run_checkset(self, checkset):
        """Run all enabled rules in a checkset"""
        results = []
        rules = checkset.get("rules", [])
        for rule in rules:
            if not rule.get("enabled", True):
                results.append(RuleResult(rule, "skipped", "Rule is disabled"))
                continue
            try:
                result = self._execute_rule(rule)
                results.append(result)
            except Exception as e:
                results.append(RuleResult(rule, "error", "Error: {}".format(str(e))))
        return results
    
    def _execute_rule(self, rule):
        """Execute a single rule"""
        rule_type = rule.get("type", "")
        
        if rule_type == "value_match":
            return self._check_value_match(rule)
        elif rule_type == "not_empty":
            return self._check_not_empty(rule)
        elif rule_type == "coordinate_match":
            return self._check_coordinate(rule)
        elif rule_type == "count_check":
            return self._check_count(rule)
        elif rule_type == "naming_pattern":
            return self._check_naming(rule)
        elif rule_type == "exists_check":
            return self._check_exists(rule)
        elif rule_type == "value_match_numeric":
            return self._check_value_match_numeric(rule)
        elif rule_type == "element_naming":
            return self._check_element_naming(rule)
        elif rule_type == "element_count":
            return self._check_element_count(rule)
        elif rule_type == "category_count":
            return self._check_category_count(rule)
        elif rule_type == "warning_count":
            return self._check_warning_count(rule)
        elif rule_type == "file_size":
            return self._check_file_size(rule)
        elif rule_type == "cad_import_check":
            return self._check_cad_imports(rule)
        elif rule_type == "inplace_family_check":
            return self._check_inplace_families(rule)
        elif rule_type == "rvt_link_check":
            return self._check_rvt_links(rule)
        elif rule_type == "group_check":
            return self._check_groups(rule)
        elif rule_type == "line_pattern_check":
            return self._check_line_patterns(rule)
        elif rule_type == "unplaced_rooms":
            return self._check_unplaced_rooms(rule)
        elif rule_type == "room_area_check":
            return self._check_room_area(rule)
        elif rule_type == "duplicate_mark":
            return self._check_duplicate_marks(rule)
        else:
            return RuleResult(rule, "error", "Unknown rule type: {}".format(rule_type))
    
    # -----------------------------------------------------------------
    # RULE TYPE: value_match
    # -----------------------------------------------------------------
    def _check_value_match(self, rule):
        params = rule.get("params", {})
        expected = params.get("expected_version", "")
        
        # Get Revit version
        actual_version = str(self.app.VersionNumber)
        actual_build = self.app.VersionBuild
        
        if expected and expected in actual_version:
            return RuleResult(rule, "pass", 
                "Revit version {} (Build: {})".format(actual_version, actual_build))
        else:
            return RuleResult(rule, "fail",
                "Expected Revit {}, found {} (Build: {})".format(expected, actual_version, actual_build))
    
    # -----------------------------------------------------------------
    # RULE TYPE: not_empty - Check required Project Info fields
    # -----------------------------------------------------------------
    def _check_not_empty(self, rule):
        params = rule.get("params", {})
        fields = params.get("fields", [])
        
        project_info = self.doc.ProjectInformation
        empty_fields = []
        filled_fields = []
        details = []
        
        for field_name in fields:
            value = self._get_project_info_value(project_info, field_name)
            if value and str(value).strip():
                filled_fields.append(field_name)
                details.append(u"  \u2714 {}: {}".format(field_name, value))
            else:
                empty_fields.append(field_name)
                details.append(u"  \u2718 {}: <EMPTY>".format(field_name))
        
        if empty_fields:
            return RuleResult(rule, "fail",
                "{}/{} fields empty: {}".format(len(empty_fields), len(fields), ", ".join(empty_fields)),
                details=details)
        else:
            return RuleResult(rule, "pass",
                "All {}/{} required fields are filled".format(len(filled_fields), len(fields)),
                details=details)
    
    def _get_project_info_value(self, project_info, field_name):
        """Get value from Project Information by field name"""
        # Try built-in parameters first
        builtin_map = {
            "Project Name": BuiltInParameter.PROJECT_NAME,
            "Project Number": BuiltInParameter.PROJECT_NUMBER,
            "Client Name": BuiltInParameter.CLIENT_NAME,
            "Project Address": BuiltInParameter.PROJECT_ADDRESS,
            "Author": BuiltInParameter.PROJECT_AUTHOR,
            "Building Name": BuiltInParameter.PROJECT_BUILDING_NAME,
            "Organization Name": BuiltInParameter.PROJECT_ORGANIZATION_NAME,
            "Organization Description": BuiltInParameter.PROJECT_ORGANIZATION_DESCRIPTION,
            "Project Issue Date": BuiltInParameter.PROJECT_ISSUE_DATE,
            "Project Status": BuiltInParameter.PROJECT_STATUS,
        }
        
        if field_name in builtin_map:
            param = project_info.get_Parameter(builtin_map[field_name])
            if param and param.HasValue:
                return param.AsString()
        
        # Try by parameter name (for shared/project parameters)
        for param in project_info.Parameters:
            if param.Definition.Name == field_name:
                if param.HasValue:
                    if param.StorageType == StorageType.String:
                        return param.AsString()
                    elif param.StorageType == StorageType.Integer:
                        return str(param.AsInteger())
                    elif param.StorageType == StorageType.Double:
                        return str(param.AsDouble())
                    elif param.StorageType == StorageType.ElementId:
                        return str(_eid_int(param.AsElementId()))
        return None
    
    # -----------------------------------------------------------------
    # RULE TYPE: coordinate_match - Check Survey/Base point coords
    # -----------------------------------------------------------------
    def _check_coordinate(self, rule):
        params = rule.get("params", {})
        point_type = params.get("point_type", "survey")  # "survey" or "base"
        axis = params.get("axis", "NS")  # "NS", "EW", "Elev"
        expected = params.get("expected_value", 0.0)
        tolerance = params.get("tolerance", 0.001)
        
        # Get base/survey points
        collector = FilteredElementCollector(self.doc).OfClass(BasePoint)
        
        actual_value = None
        point_name = "Survey Point" if point_type == "survey" else "Project Base Point"
        
        for bp in collector:
            is_survey = bp.IsShared
            if (point_type == "survey" and is_survey) or (point_type == "base" and not is_survey):
                pos = bp.Position
                if axis == "NS":
                    actual_value = pos.Y  # North/South = Y in Revit internal
                elif axis == "EW":
                    actual_value = pos.X  # East/West = X in Revit internal
                elif axis == "Elev":
                    actual_value = pos.Z
                break
        
        if actual_value is None:
            return RuleResult(rule, "error", "Could not find {}".format(point_name))
        
        # Convert from internal units (feet) to display units for reporting
        # Note: comparison uses internal units, display converts for user
        actual_ft = actual_value
        expected_ft = expected  # User provides in project units - might need conversion
        
        diff = abs(actual_ft - expected_ft)
        
        if diff <= tolerance:
            return RuleResult(rule, "pass",
                "{} {} = {:.6f} (expected: {})".format(point_name, axis, actual_ft, expected))
        else:
            return RuleResult(rule, "fail",
                "{} {} = {:.6f} (expected: {}, diff: {:.6f})".format(
                    point_name, axis, actual_ft, expected, diff))
    
    # -----------------------------------------------------------------
    # RULE TYPE: count_check - Check element counts
    # -----------------------------------------------------------------
    def _check_count(self, rule):
        params = rule.get("params", {})
        target = params.get("target", "")
        max_count = params.get("max_count", 0)
        msg_over = params.get("message_over", "Count exceeds maximum")
        
        count = 0
        details = []
        elements = []
        
        if target == "design_options":
            collector = FilteredElementCollector(self.doc).OfClass(DesignOption)
            for opt in collector:
                count += 1
                details.append("  - {} (ID: {})".format(opt.Name, _eid_int(opt.Id)))
                elements.append(_eid_int(opt.Id))
        
        elif target == "phases":
            phases = self.doc.Phases
            count = phases.Size
            for i in range(count):
                p = phases[i]
                details.append("  - {} (ID: {})".format(p.Name, _eid_int(p.Id)))
        
        elif target == "warnings":
            warnings = self.doc.GetWarnings()
            count = len(warnings) if warnings else 0
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "Found {} (max allowed: {}). {}".format(count, max_count, msg_over),
                details=details, elements=elements)
        else:
            return RuleResult(rule, "pass",
                "Count: {} (max allowed: {})".format(count, max_count),
                details=details)
    
    # -----------------------------------------------------------------
    # RULE TYPE: naming_pattern - Check naming conventions
    # -----------------------------------------------------------------
    def _check_naming(self, rule):
        import re
        params = rule.get("params", {})
        target = params.get("target", "")
        pattern = params.get("pattern", "")
        desc = params.get("description", "")
        
        violations = []
        all_names = []
        details = []
        
        if target == "worksets":
            if self.doc.IsWorkshared:
                worksets = FilteredWorksetCollector(self.doc).OfKind(WorksetKind.UserWorkset)
                for ws in worksets:
                    name = ws.Name
                    all_names.append(name)
                    if pattern and not re.match(pattern, name):
                        violations.append(name)
                        details.append(u"  \u2718 {}".format(name))
                    else:
                        details.append(u"  \u2714 {}".format(name))
            else:
                return RuleResult(rule, "info", "Model is not workshared", details=["Workset check skipped"])
        
        if violations:
            return RuleResult(rule, "fail",
                "{}/{} worksets violate pattern. {}".format(len(violations), len(all_names), desc),
                details=details)
        else:
            return RuleResult(rule, "pass",
                "All {} names match pattern".format(len(all_names)),
                details=details)
    
    # -----------------------------------------------------------------
    # RULE TYPE: exists_check - Check if something exists
    # -----------------------------------------------------------------
    def _check_exists(self, rule):
        params = rule.get("params", {})
        target = params.get("target", "")
        
        if target == "starting_view":
            starting_view_id = self.doc.Application.Create \
                if hasattr(self.doc, 'GetStartingViewSettings') \
                else None
            
            # Try to get starting view settings
            try:
                sv_settings = StartingViewSettings.GetStartingViewSettings(self.doc)
                if sv_settings:
                    view_id = sv_settings.ViewId
                    if view_id and view_id != ElementId.InvalidElementId:
                        view = self.doc.GetElement(view_id)
                        if view:
                            return RuleResult(rule, "pass",
                                "Starting View is set: {}".format(view.Name))
                return RuleResult(rule, "fail", "No Starting View is set")
            except:
                return RuleResult(rule, "info", "Could not determine Starting View setting")
        
        return RuleResult(rule, "error", "Unknown target: {}".format(target))
    
    # -----------------------------------------------------------------
    # RULE TYPE: value_match_numeric - Check numeric values
    # -----------------------------------------------------------------
    def _check_value_match_numeric(self, rule):
        params = rule.get("params", {})
        target = params.get("target", "")
        expected = params.get("expected_value", 0.0)
        tolerance = params.get("tolerance", 0.01)
        report_only = params.get("report_only", False)
        
        if target == "true_north":
            try:
                # Get active project location
                project_location = self.doc.ActiveProjectLocation
                project_position = project_location.GetProjectPosition(XYZ.Zero)
                angle_rad = project_position.Angle
                angle_deg = angle_rad * 180.0 / 3.14159265358979
                
                diff = abs(angle_deg - expected)
                
                if report_only:
                    return RuleResult(rule, "info",
                        "True North angle: {:.2f} degrees".format(angle_deg))
                
                if diff <= tolerance:
                    return RuleResult(rule, "pass",
                        "True North angle: {:.2f} deg (expected: {})".format(angle_deg, expected))
                else:
                    return RuleResult(rule, "fail",
                        "True North angle: {:.2f} deg (expected: {}, diff: {:.2f})".format(
                            angle_deg, expected, diff))
            except Exception as e:
                return RuleResult(rule, "error", "Error reading True North: {}".format(str(e)))
        
        return RuleResult(rule, "error", "Unknown target: {}".format(target))

    # =================================================================
    # PHASE 2: NAMING CONVENTION RULES
    # =================================================================
    
    def _check_element_naming(self, rule):
        """Check naming of views/sheets/levels/grids/families"""
        import re
        params = rule.get("params", {})
        target = params.get("target", "")
        pattern = params.get("pattern", "")
        desc = params.get("description", "")
        exclude_templates = params.get("exclude_templates", True)
        
        violations = []
        all_items = []
        details = []
        elements = []
        
        if target == "views":
            collector = FilteredElementCollector(self.doc).OfClass(View)
            for v in collector:
                if v.IsTemplate and exclude_templates:
                    continue
                try:
                    vtype = v.ViewType
                    if vtype in (ViewType.SystemBrowser, ViewType.ProjectBrowser,
                                 ViewType.Undefined, ViewType.Internal, ViewType.DrawingSheet):
                        continue
                except:
                    continue
                name = v.Name or ""
                all_items.append(name)
                if pattern and not re.match(pattern, name):
                    violations.append(name)
                    details.append(u"  \u2718 {} (ID: {})".format(name, _eid_int(v.Id)))
                    elements.append(_eid_int(v.Id))
        
        elif target == "sheets":
            collector = FilteredElementCollector(self.doc).OfClass(ViewSheet)
            for s in collector:
                sheet_num = s.SheetNumber or ""
                sheet_name = s.Name or ""
                full_name = "{} - {}".format(sheet_num, sheet_name)
                all_items.append(full_name)
                if pattern and not re.match(pattern, sheet_num):
                    violations.append(full_name)
                    details.append(u"  \u2718 {} (ID: {})".format(full_name, _eid_int(s.Id)))
                    elements.append(_eid_int(s.Id))
        
        elif target == "levels":
            collector = FilteredElementCollector(self.doc).OfClass(Level)
            for lv in collector:
                name = lv.Name or ""
                all_items.append(name)
                if pattern and not re.match(pattern, name):
                    violations.append(name)
                    details.append(u"  \u2718 {} (ID: {})".format(name, _eid_int(lv.Id)))
                    elements.append(_eid_int(lv.Id))
        
        elif target == "grids":
            collector = FilteredElementCollector(self.doc).OfClass(Grid)
            for g in collector:
                name = g.Name or ""
                all_items.append(name)
                if pattern and not re.match(pattern, name):
                    violations.append(name)
                    details.append(u"  \u2718 {} (ID: {})".format(name, _eid_int(g.Id)))
                    elements.append(_eid_int(g.Id))
        
        elif target == "families":
            collector = FilteredElementCollector(self.doc).OfClass(Family)
            for fam in collector:
                if fam.IsInPlace:
                    continue
                name = fam.Name or ""
                all_items.append(name)
                if pattern and not re.match(pattern, name):
                    violations.append(name)
                    details.append(u"  \u2718 {} (ID: {})".format(name, _eid_int(fam.Id)))
                    elements.append(_eid_int(fam.Id))
        
        total = len(all_items)
        fail_count = len(violations)
        
        if not pattern:
            return RuleResult(rule, "info",
                "Found {} {} (no pattern defined)".format(total, target))
        
        if violations:
            if len(details) > 20:
                details = details[:20]
                details.append("  ... and {} more violations".format(fail_count - 20))
            return RuleResult(rule, "fail",
                "{}/{} {} violate naming. {}".format(fail_count, total, target, desc),
                details=details, elements=elements[:50])
        else:
            return RuleResult(rule, "pass",
                "All {} {} match naming pattern".format(total, target))

    # =================================================================
    # PHASE 2: MODEL PERFORMANCE RULES
    # =================================================================
    
    def _check_element_count(self, rule):
        """Count total elements in model"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 500000)
        
        collector = FilteredElementCollector(self.doc).WhereElementIsNotElementType()
        count = collector.GetElementCount()
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "{} elements (max: {})".format(count, max_count))
        else:
            return RuleResult(rule, "pass",
                "{} elements (max: {})".format(count, max_count))
    
    def _check_category_count(self, rule):
        """Count elements in a specific category"""
        params = rule.get("params", {})
        category_label = params.get("category", "")
        max_count = params.get("max_count", 0)
        bic_name = params.get("built_in_category", "")
        
        count = 0
        elements = []
        
        if bic_name:
            try:
                bic = getattr(BuiltInCategory, bic_name)
                collector = FilteredElementCollector(self.doc).OfCategory(bic).WhereElementIsNotElementType()
                for el in collector:
                    count += 1
                    if count <= 50:
                        elements.append(_eid_int(el.Id))
            except:
                return RuleResult(rule, "error",
                    "Invalid BuiltInCategory: {}".format(bic_name))
        
        if max_count > 0 and count > max_count:
            return RuleResult(rule, "fail",
                "{}: {} elements (max: {})".format(category_label or bic_name, count, max_count),
                elements=elements)
        else:
            return RuleResult(rule, "pass",
                "{}: {} elements{}".format(
                    category_label or bic_name, count,
                    " (max: {})".format(max_count) if max_count > 0 else ""))
    
    def _check_warning_count(self, rule):
        """Check Revit warnings count"""
        params = rule.get("params", {})
        max_warnings = params.get("max_count", 100)
        
        try:
            warnings = self.doc.GetWarnings()
            count = len(warnings) if warnings else 0
        except:
            count = 0
        
        details = []
        if count > 0:
            try:
                warning_types = {}
                for w in warnings:
                    desc = w.GetDescriptionText() or "Unknown"
                    warning_types[desc] = warning_types.get(desc, 0) + 1
                sorted_types = sorted(warning_types.items(), key=lambda x: x[1], reverse=True)
                for desc, cnt in sorted_types[:15]:
                    short_desc = desc[:80] + "..." if len(desc) > 80 else desc
                    details.append("  {}x {}".format(cnt, short_desc))
                if len(sorted_types) > 15:
                    details.append("  ... and {} more types".format(len(sorted_types) - 15))
            except:
                pass
        
        if count > max_warnings:
            return RuleResult(rule, "fail",
                "{} warnings (max: {})".format(count, max_warnings), details=details)
        else:
            return RuleResult(rule, "pass",
                "{} warnings (max: {})".format(count, max_warnings), details=details)
    
    def _check_file_size(self, rule):
        """Check model file size"""
        params = rule.get("params", {})
        max_mb = params.get("max_size_mb", 300)
        
        filepath = self.doc.PathName
        if not filepath or not os.path.exists(filepath):
            return RuleResult(rule, "info", "File not saved or path not accessible")
        
        size_bytes = os.path.getsize(filepath)
        size_mb = size_bytes / (1024.0 * 1024.0)
        
        if size_mb > max_mb:
            return RuleResult(rule, "fail",
                "File size: {:.1f} MB (max: {} MB)".format(size_mb, max_mb))
        else:
            return RuleResult(rule, "pass",
                "File size: {:.1f} MB (max: {} MB)".format(size_mb, max_mb))
    
    def _check_cad_imports(self, rule):
        """Detect imported/linked CAD files"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 0)
        check_type = params.get("check_type", "imported")
        
        count = 0
        details = []
        elements = []
        
        collector = FilteredElementCollector(self.doc).OfClass(ImportInstance)
        for imp in collector:
            is_linked = imp.IsLinked
            if check_type == "imported" and is_linked:
                continue
            if check_type == "linked" and not is_linked:
                continue
            count += 1
            try:
                name_p = imp.LookupParameter("Name")
                imp_name = name_p.AsString() if name_p and name_p.HasValue else "Unknown"
            except:
                imp_name = "CAD Instance"
            link_type = "Linked" if is_linked else "Imported"
            details.append("  {} - {} (ID: {})".format(link_type, imp_name, _eid_int(imp.Id)))
            elements.append(_eid_int(imp.Id))
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "{} CAD {} (max: {})".format(count, check_type, max_count),
                details=details[:20], elements=elements[:50])
        else:
            return RuleResult(rule, "pass",
                "{} CAD {} (max: {})".format(count, check_type, max_count),
                details=details[:10])
    
    def _check_inplace_families(self, rule):
        """Detect In-Place families"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 0)
        
        count = 0
        details = []
        elements = []
        
        collector = FilteredElementCollector(self.doc).OfClass(Family)
        for fam in collector:
            if fam.IsInPlace:
                count += 1
                cat_name = fam.FamilyCategory.Name if fam.FamilyCategory else "N/A"
                details.append("  {} [{}] (ID: {})".format(fam.Name, cat_name, _eid_int(fam.Id)))
                elements.append(_eid_int(fam.Id))
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "{} In-Place families (max: {}). Convert to loadable families.".format(count, max_count),
                details=details[:20], elements=elements[:50])
        else:
            return RuleResult(rule, "pass",
                "{} In-Place families (max: {})".format(count, max_count),
                details=details[:10])
    
    def _check_rvt_links(self, rule):
        """Check Revit links status"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 20)
        check_loaded = params.get("check_loaded", True)
        
        count = 0
        unloaded = 0
        details = []
        elements = []
        
        collector = FilteredElementCollector(self.doc).OfClass(RevitLinkInstance)
        for link in collector:
            count += 1
            try:
                link_type = self.doc.GetElement(link.GetTypeId())
                link_name = link_type.Name if link_type else "Unknown"
                is_loaded = True
                try:
                    link_doc = link.GetLinkDocument()
                    if link_doc is None:
                        is_loaded = False
                        unloaded += 1
                except:
                    is_loaded = False
                    unloaded += 1
                status = "Loaded" if is_loaded else "NOT LOADED"
                details.append("  {} [{}] (ID: {})".format(link_name, status, _eid_int(link.Id)))
                elements.append(_eid_int(link.Id))
            except:
                details.append("  Unknown link (ID: {})".format(_eid_int(link.Id)))
        
        msg = "{} RVT links".format(count)
        if unloaded > 0:
            msg += ", {} NOT loaded".format(unloaded)
        
        failed = count > max_count or (check_loaded and unloaded > 0)
        if failed:
            return RuleResult(rule, "fail", msg, details=details[:20], elements=elements[:50])
        else:
            return RuleResult(rule, "pass", msg, details=details[:10])
    
    def _check_groups(self, rule):
        """Check model/detail groups"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 20)
        
        count = 0
        details = []
        elements = []
        
        collector = FilteredElementCollector(self.doc).OfClass(GroupType)
        for gt in collector:
            count += 1
            details.append("  {} (ID: {})".format(gt.Name, _eid_int(gt.Id)))
            elements.append(_eid_int(gt.Id))
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "{} group types (max: {}). Consider ungrouping.".format(count, max_count),
                details=details[:20], elements=elements[:50])
        else:
            return RuleResult(rule, "pass",
                "{} group types (max: {})".format(count, max_count),
                details=details[:10])
    
    def _check_line_patterns(self, rule):
        """Check line patterns count"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 50)
        
        count = 0
        details = []
        
        collector = FilteredElementCollector(self.doc).OfClass(LinePatternElement)
        for lp in collector:
            count += 1
            details.append("  {} (ID: {})".format(lp.Name, _eid_int(lp.Id)))
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "{} line patterns (max: {}). Purge unused.".format(count, max_count),
                details=details[:20])
        else:
            return RuleResult(rule, "pass",
                "{} line patterns (max: {})".format(count, max_count))
    
    def _check_unplaced_rooms(self, rule):
        """Check for unplaced/unbounded rooms"""
        params = rule.get("params", {})
        max_count = params.get("max_count", 0)
        
        count = 0
        details = []
        elements = []
        
        collector = FilteredElementCollector(self.doc).OfCategory(
            BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
        for room in collector:
            try:
                if room.Area == 0 or room.Location is None:
                    count += 1
                    rn = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                    rnum = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                    name = rn.AsString() if rn and rn.HasValue else "No Name"
                    number = rnum.AsString() if rnum and rnum.HasValue else "-"
                    details.append("  {} - {} (ID: {})".format(number, name, _eid_int(room.Id)))
                    elements.append(_eid_int(room.Id))
            except:
                pass
        
        if count > max_count:
            return RuleResult(rule, "fail",
                "{} unplaced/unbounded rooms".format(count),
                details=details[:20], elements=elements[:50])
        else:
            return RuleResult(rule, "pass", "{} unplaced rooms".format(count))
    
    def _check_room_area(self, rule):
        """Check rooms with very small area"""
        params = rule.get("params", {})
        min_area_sqm = params.get("min_area_sqm", 0.5)
        
        count = 0
        total_rooms = 0
        details = []
        elements = []
        
        collector = FilteredElementCollector(self.doc).OfCategory(
            BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
        for room in collector:
            try:
                if room.Location is None:
                    continue
                total_rooms += 1
                area_sqft = room.Area
                area_sqm = area_sqft * 0.092903
                if 0 < area_sqm < min_area_sqm:
                    count += 1
                    rn = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                    name = rn.AsString() if rn and rn.HasValue else "No Name"
                    details.append("  {} - {:.2f} sqm (ID: {})".format(name, area_sqm, _eid_int(room.Id)))
                    elements.append(_eid_int(room.Id))
            except:
                pass
        
        if count > 0:
            return RuleResult(rule, "fail",
                "{}/{} rooms < {} sqm".format(count, total_rooms, min_area_sqm),
                details=details[:20], elements=elements[:50])
        else:
            return RuleResult(rule, "pass",
                "All {} rooms >= {} sqm".format(total_rooms, min_area_sqm))
    
    def _check_duplicate_marks(self, rule):
        """Check duplicate Mark parameter values"""
        params = rule.get("params", {})
        bic_name = params.get("built_in_category", "")
        category_name = params.get("category", "All")
        ignore_empty = params.get("ignore_empty", True)
        
        mark_map = {}
        if bic_name:
            try:
                bic = getattr(BuiltInCategory, bic_name)
                collector = FilteredElementCollector(self.doc).OfCategory(bic).WhereElementIsNotElementType()
            except:
                return RuleResult(rule, "error", "Invalid category: {}".format(bic_name))
        else:
            collector = FilteredElementCollector(self.doc).WhereElementIsNotElementType()
        
        for el in collector:
            try:
                mark_p = el.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                if mark_p and mark_p.HasValue:
                    val = mark_p.AsString()
                    if ignore_empty and (not val or not val.strip()):
                        continue
                    if val not in mark_map:
                        mark_map[val] = []
                    mark_map[val].append(_eid_int(el.Id))
            except:
                pass
        
        duplicates = {k: v for k, v in mark_map.items() if len(v) > 1}
        dup_count = len(duplicates)
        
        details = []
        elements = []
        for mark_val, ids in sorted(duplicates.items(), key=lambda x: len(x[1]), reverse=True)[:15]:
            details.append("  '{}' used by {} elements".format(mark_val, len(ids)))
            elements.extend(ids[:5])
        
        if dup_count > 0:
            total_affected = sum(len(v) for v in duplicates.values())
            return RuleResult(rule, "fail",
                "{} duplicate Marks ({} elements) in {}".format(dup_count, total_affected, category_name),
                details=details, elements=elements[:50])
        else:
            return RuleResult(rule, "pass",
                "No duplicate Marks in {}".format(category_name))
class ExcelReporter:
    """Generate Excel compliance report"""
    
    def __init__(self, doc):
        self.doc = doc
    
    def _rgb_to_ole(self, r, g, b):
        """Convert RGB to OLE color (avoids System.Drawing dependency)"""
        return r + (g * 256) + (b * 256 * 256)
    
    def generate_report(self, checkset, results, filepath):
        """Generate a full Excel report"""
        try:
            clr.AddReference('Microsoft.Office.Interop.Excel')
            from Microsoft.Office.Interop import Excel as ExcelInterop
            excel_app = ExcelInterop.ApplicationClass()
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            wb = excel_app.Workbooks.Add()
            
            # --- Sheet 1: Summary ---
            ws_summary = wb.Sheets[1]
            ws_summary.Name = "Summary"
            self._write_summary(ws_summary, checkset, results)
            
            # --- Sheet 2: Detailed Results ---
            ws_details = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
            ws_details.Name = "Detailed Results"
            self._write_details(ws_details, results)
            
            # --- Sheet 3: Failed Items ---
            ws_failed = wb.Sheets.Add(After=wb.Sheets[wb.Sheets.Count])
            ws_failed.Name = "Failed Items"
            self._write_failed(ws_failed, results)
            
            wb.SaveAs(filepath)
            wb.Close()
            excel_app.Quit()
            
            # Release COM objects
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
    
    def _write_summary(self, ws, checkset, results):
        """Write summary sheet"""
        # Title
        ws.Cells[1, 1].Value2 = "MODEL CHECKER REPORT"
        ws.Cells[1, 1].Font.Size = 16
        ws.Cells[1, 1].Font.Bold = True
        ws.Range["A1:D1"].Merge()
        
        # DQT branding
        header_range = ws.Range["A1:D1"]
        header_range.Interior.Color = self._rgb_to_ole(240, 204, 136)
        
        # Project info
        row = 3
        info = [
            ("Project Name", self.doc.ProjectInformation.Name or "N/A"),
            ("File Path", self.doc.PathName or "Not saved"),
            ("Checkset", checkset.get("name", "N/A")),
            ("Description", checkset.get("description", "")),
            ("Date", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Revit Version", "{} ({})".format(
                self.doc.Application.VersionNumber,
                self.doc.Application.VersionBuild)),
        ]
        
        for label, value in info:
            ws.Cells[row, 1].Value2 = label
            ws.Cells[row, 1].Font.Bold = True
            ws.Cells[row, 2].Value2 = value
            row += 1
        
        # Statistics
        row += 1
        total = len(results)
        passed = len([r for r in results if r.status == "pass"])
        failed = len([r for r in results if r.status == "fail"])
        warnings = len([r for r in results if r.status == "warning"])
        info_count = len([r for r in results if r.status == "info"])
        errors = len([r for r in results if r.status == "error"])
        skipped = len([r for r in results if r.status == "skipped"])
        
        ws.Cells[row, 1].Value2 = "RESULTS SUMMARY"
        ws.Cells[row, 1].Font.Size = 12
        ws.Cells[row, 1].Font.Bold = True
        ws.Range[ws.Cells[row, 1], ws.Cells[row, 4]].Interior.Color = \
            self._rgb_to_ole(254, 248, 231)
        row += 1
        
        stats = [
            ("Total Rules", total),
            ("Passed", passed),
            ("Failed", failed),
            ("Info", info_count),
            ("Errors", errors),
            ("Skipped", skipped),
        ]
        
        for label, value in stats:
            ws.Cells[row, 1].Value2 = label
            ws.Cells[row, 1].Font.Bold = True
            ws.Cells[row, 2].Value2 = value
            
            # Color code
            if label == "Passed":
                ws.Cells[row, 2].Font.Color = self._rgb_to_ole(46, 125, 50)
            elif label == "Failed":
                ws.Cells[row, 2].Font.Color = self._rgb_to_ole(198, 40, 40)
            row += 1
        
        # Auto-fit
        ws.Columns["A:D"].AutoFit()
    
    def _write_details(self, ws, results):
        """Write detailed results sheet"""
        headers = ["Rule ID", "Rule Name", "Category", "Severity", "Status", "Message"]
        
        # Header row
        for i, h in enumerate(headers, 1):
            ws.Cells[1, i].Value2 = h
            ws.Cells[1, i].Font.Bold = True
            ws.Cells[1, i].Interior.Color = self._rgb_to_ole(240, 204, 136)
        
        # Data rows
        row = 2
        for r in results:
            ws.Cells[row, 1].Value2 = r.rule_id
            ws.Cells[row, 2].Value2 = r.rule_name
            ws.Cells[row, 3].Value2 = r.category
            ws.Cells[row, 4].Value2 = r.severity.upper()
            ws.Cells[row, 5].Value2 = r.status.upper()
            ws.Cells[row, 6].Value2 = r.message
            
            # Color code status
            status_colors = {
                "pass": self._rgb_to_ole(200, 230, 201),
                "fail": self._rgb_to_ole(255, 205, 210),
                "warning": self._rgb_to_ole(255, 236, 179),
                "info": self._rgb_to_ole(187, 222, 251),
                "error": self._rgb_to_ole(255, 171, 145),
                "skipped": self._rgb_to_ole(224, 224, 224),
            }
            
            color = status_colors.get(r.status, self._rgb_to_ole(255, 255, 255))
            ws.Cells[row, 5].Interior.Color = color
            
            # Write details as sub-rows
            if r.details:
                for detail in r.details:
                    row += 1
                    ws.Cells[row, 6].Value2 = detail
                    ws.Cells[row, 6].Font.Color = self._rgb_to_ole(128, 128, 128)
                    ws.Cells[row, 6].Font.Size = 9
            
            row += 1
        
        ws.Columns["A:F"].AutoFit()
    
    def _write_failed(self, ws, results):
        """Write failed items sheet"""
        headers = ["Rule ID", "Rule Name", "Severity", "Message", "Element IDs"]
        
        for i, h in enumerate(headers, 1):
            ws.Cells[1, i].Value2 = h
            ws.Cells[1, i].Font.Bold = True
            ws.Cells[1, i].Interior.Color = self._rgb_to_ole(255, 205, 210)
        
        row = 2
        for r in results:
            if r.status == "fail":
                ws.Cells[row, 1].Value2 = r.rule_id
                ws.Cells[row, 2].Value2 = r.rule_name
                ws.Cells[row, 3].Value2 = r.severity.upper()
                ws.Cells[row, 4].Value2 = r.message
                ws.Cells[row, 5].Value2 = ", ".join(str(e) for e in r.elements) if r.elements else "N/A"
                row += 1
        
        if row == 2:
            ws.Cells[2, 1].Value2 = "No failed items - All checks passed!"
            ws.Range["A2:E2"].Merge()
            ws.Cells[2, 1].Font.Color = self._rgb_to_ole(46, 125, 50)
        
        ws.Columns["A:E"].AutoFit()


# =====================================================================
# WPF MAIN WINDOW
# =====================================================================
XAML_STR = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Model Checker v1.0 - DQT"
        Height="800" Width="1100"
        MinHeight="650" MinWidth="900"
        WindowStartupLocation="CenterScreen"
        Background="#F8FAFC">
    
    <Window.Resources>
        <Style x:Key="CardBorder" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="4"/>
            <Setter Property="Padding" Value="12,8"/>
        </Style>
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background" Value="#0F172A"/>
            <Setter Property="Foreground" Value="#5D4E37"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style x:Key="BtnSecondary" TargetType="Button">
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#5D4E37"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style x:Key="BtnDanger" TargetType="Button">
            <Setter Property="Background" Value="#FFCDD2"/>
            <Setter Property="Foreground" Value="#C62828"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="BorderBrush" Value="#EF9A9A"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="BtnSuccess" TargetType="Button">
            <Setter Property="Background" Value="#C8E6C9"/>
            <Setter Property="Foreground" Value="#2E7D32"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="BorderBrush" Value="#81C784"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Row 0: Header -->
        <Border Grid.Row="0" Background="#0F172A" CornerRadius="6" Padding="14,10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="&#x2714; Model Checker" FontSize="20" FontWeight="Bold" Foreground="#333"/>
                    <TextBlock Text="Rule-based BIM model compliance checker" FontSize="11" Foreground="#666" Margin="0,3,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" VerticalAlignment="Center" HorizontalAlignment="Right">
                    <TextBlock Text="DQT" FontSize="14" FontWeight="Bold" Foreground="#C89650"/>
                    <TextBlock Text="v1.0" FontSize="9" Foreground="#999" HorizontalAlignment="Right"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Row 1: Checkset Selection Bar -->
        <Border Grid.Row="1" Style="{StaticResource CardBorder}" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="Checkset:" FontWeight="SemiBold" 
                           FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#FFFFFF"/>
                <ComboBox x:Name="cmbCheckset" Grid.Column="1" Padding="8,5" FontSize="11"/>
                
                <Button x:Name="btnNewCheckset" Grid.Column="2" Content="New" 
                        Style="{StaticResource BtnSecondary}" Margin="6,0,0,0"/>
                <Button x:Name="btnDuplicate" Grid.Column="3" Content="Duplicate" 
                        Style="{StaticResource BtnSecondary}" Margin="4,0,0,0"/>
                <Button x:Name="btnImport" Grid.Column="4" Content="Import" 
                        Style="{StaticResource BtnSecondary}" Margin="4,0,0,0"/>
                <Button x:Name="btnExport" Grid.Column="5" Content="Export" 
                        Style="{StaticResource BtnSecondary}" Margin="4,0,0,0"/>
                <Button x:Name="btnDeleteCheckset" Grid.Column="6" Content="Delete" 
                        Style="{StaticResource BtnDanger}" Margin="4,0,0,0"/>
            </Grid>
        </Border>
        
        <!-- Row 2: Summary Cards -->
        <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <Border Grid.Column="0" Style="{StaticResource CardBorder}" Margin="0,0,4,0">
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock x:Name="txtTotalRules" Text="0" FontSize="22" FontWeight="Bold" Foreground="#FFFFFF" HorizontalAlignment="Center"/>
                    <TextBlock Text="Total Rules" FontSize="10" Foreground="#999" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
            <Border Grid.Column="1" Style="{StaticResource CardBorder}" Margin="2,0,2,0" Background="#E8F5E9">
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock x:Name="txtPassed" Text="0" FontSize="22" FontWeight="Bold" Foreground="#2E7D32" HorizontalAlignment="Center"/>
                    <TextBlock Text="Passed" FontSize="10" Foreground="#388E3C" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
            <Border Grid.Column="2" Style="{StaticResource CardBorder}" Margin="2,0,2,0" Background="#FFEBEE">
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock x:Name="txtFailed" Text="0" FontSize="22" FontWeight="Bold" Foreground="#C62828" HorizontalAlignment="Center"/>
                    <TextBlock Text="Failed" FontSize="10" Foreground="#D32F2F" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
            <Border Grid.Column="3" Style="{StaticResource CardBorder}" Margin="2,0,2,0" Background="#E3F2FD">
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock x:Name="txtInfo" Text="0" FontSize="22" FontWeight="Bold" Foreground="#1565C0" HorizontalAlignment="Center"/>
                    <TextBlock Text="Info" FontSize="10" Foreground="#1976D2" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
            <Border Grid.Column="4" Style="{StaticResource CardBorder}" Margin="4,0,0,0" Background="#ECEFF1">
                <StackPanel HorizontalAlignment="Center">
                    <TextBlock x:Name="txtSkipped" Text="0" FontSize="22" FontWeight="Bold" Foreground="#546E7A" HorizontalAlignment="Center"/>
                    <TextBlock Text="Skipped" FontSize="10" Foreground="#78909C" HorizontalAlignment="Center"/>
                </StackPanel>
            </Border>
        </Grid>
        
        <!-- Row 3: Main Content - Rule List + Detail Panel -->
        <Grid Grid.Row="3" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="3*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Left: Rule Tree with checkboxes -->
            <Border Grid.Column="0" Style="{StaticResource CardBorder}" Margin="0,0,4,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" Text="Rules" FontWeight="Bold" FontSize="13" 
                               Foreground="#FFFFFF" Margin="0,0,0,6"/>
                    
                    <!-- Select All / None -->
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,6">
                        <Button x:Name="btnSelectAll" Content="Select All" 
                                Style="{StaticResource BtnSecondary}" Padding="6,3" FontSize="10" Margin="0,0,4,0"/>
                        <Button x:Name="btnSelectNone" Content="Select None" 
                                Style="{StaticResource BtnSecondary}" Padding="6,3" FontSize="10"/>
                    </StackPanel>
                    
                    <!-- Rule ListBox with checkboxes -->
                    <ListBox x:Name="lstRules" Grid.Row="2" 
                             BorderBrush="#E0E0E0" BorderThickness="1"
                             Background="White" ScrollViewer.VerticalScrollBarVisibility="Auto">
                    </ListBox>
                </Grid>
            </Border>
            
            <!-- Right: Rule Detail / Results Panel -->
            <Border Grid.Column="1" Style="{StaticResource CardBorder}" Margin="4,0,0,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <TextBlock Grid.Row="0" x:Name="txtDetailHeader" Text="Select a rule to view details" 
                               FontWeight="Bold" FontSize="13" Foreground="#FFFFFF" Margin="0,0,0,6"/>
                    
                    <!-- Results ScrollViewer (shown after running) -->
                    <ScrollViewer x:Name="dgResults" Grid.Row="1"
                              VerticalScrollBarVisibility="Auto"
                              Visibility="Collapsed">
                        <StackPanel x:Name="spResults"/>
                    </ScrollViewer>
                    
                    <!-- Rule Detail Panel (shown before running) -->
                    <ScrollViewer x:Name="pnlRuleDetail" Grid.Row="1" 
                                  VerticalScrollBarVisibility="Auto">
                        <StackPanel x:Name="spRuleDetail" Margin="4">
                            <TextBlock x:Name="txtRuleId" Text="" FontSize="10" Foreground="#999" Margin="0,0,0,4"/>
                            <TextBlock x:Name="txtRuleDesc" Text="" FontSize="11" Foreground="#666" 
                                       TextWrapping="Wrap" Margin="0,0,0,8"/>
                            <TextBlock x:Name="txtRuleType" Text="" FontSize="10" Foreground="#888" Margin="0,0,0,4"/>
                            <TextBlock x:Name="txtRuleSeverity" Text="" FontSize="10" Margin="0,0,0,8"/>
                            
                            <!-- Parameters editing area -->
                            <TextBlock Text="Parameters:" FontWeight="SemiBold" FontSize="11" 
                                       Foreground="#FFFFFF" Margin="0,4,0,4"/>
                            <Border x:Name="pnlParams" BorderBrush="#E0E0E0" BorderThickness="1" 
                                    CornerRadius="6" Padding="8" Background="#FAFAFA">
                                <StackPanel x:Name="spParams"/>
                            </Border>
                            
                            <Button x:Name="btnSaveParams" Content="Save Parameters" 
                                    Style="{StaticResource BtnPrimary}" Margin="0,8,0,0"
                                    HorizontalAlignment="Left" Padding="10,5"/>
                        </StackPanel>
                    </ScrollViewer>
                </Grid>
            </Border>
        </Grid>
        
        <!-- Row 4: Action Buttons -->
        <Grid Grid.Row="4" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <!-- Status Text -->
            <TextBlock x:Name="txtStatus" Grid.Column="0" Text="Ready. Select a checkset and click Run Check." 
                       FontSize="11" Foreground="#888" VerticalAlignment="Center"/>
            
            <Button x:Name="btnRunCheck" Grid.Column="1" Content="&#x25B6; Run Check" 
                    Style="{StaticResource BtnSuccess}" Margin="0,0,6,0"/>
            <Button x:Name="btnExportExcel" Grid.Column="2" Content="&#x1F4CA; Export Excel" 
                    Style="{StaticResource BtnPrimary}" Margin="0,0,6,0" IsEnabled="False"/>
            <Button x:Name="btnClose" Grid.Column="3" Content="Close" 
                    Style="{StaticResource BtnSecondary}"/>
        </Grid>
        
        <!-- Row 5: Footer -->
        <Border Grid.Row="5" Background="#F5F0E0" CornerRadius="6" Padding="8,4">
            <Grid>
                <TextBlock Text="Model Checker v1.0 | Dang Quoc Truong (DQT)" 
                           FontSize="9" Foreground="#999" HorizontalAlignment="Left"/>
                <TextBlock x:Name="txtFooterInfo" Text="" 
                           FontSize="9" Foreground="#999" HorizontalAlignment="Right"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'''


# =====================================================================
# DATA ITEMS for WPF Binding
# =====================================================================
class RuleDisplayItem:
    """Item for rule listbox display"""
    def __init__(self, rule, index):
        self.rule = rule
        self.index = index
        self.Enabled = rule.get("enabled", True)
        self.DisplayName = "[{}] {}".format(rule.get("id", ""), rule.get("name", ""))
        severity = rule.get("severity", "info")
        color_map = {
            "error": "#C62828",
            "warning": "#F57F17",
            "info": "#1565C0"
        }
        self.SeverityColor = color_map.get(severity, "#999")


class ResultDisplayItem:
    """Item for results DataGrid display"""
    def __init__(self, result):
        self.result = result
        status_icons = {
            "pass": u"\u2714",
            "fail": u"\u2718",
            "warning": u"\u26A0",
            "info": u"\u2139",
            "error": u"\u2716",
            "skipped": u"\u23F8"
        }
        self.StatusIcon = status_icons.get(result.status, "?")
        self.RuleId = result.rule_id
        self.RuleName = result.rule_name
        self.Severity = result.severity.upper()
        self.Message = result.message
        self.Status = result.status


# =====================================================================
# MAIN WINDOW CLASS
# =====================================================================
class ModelCheckerWindow:
    """Main Model Checker window"""
    
    def __init__(self):
        self.checkset_mgr = ChecksetManager()
        self.rule_engine = RuleEngine(doc)
        self.reporter = ExcelReporter(doc)
        self.current_checkset = None
        self.current_results = None
        self.rule_items = []
        
        # Parse XAML
        xr = XmlReader.Create(StringReader(XAML_STR))
        self.window = XamlReader.Load(xr)
        
        # Get controls
        self._get_controls()
        self._bind_events()
        self._load_checksets()
        
        # Footer info
        self.txtFooterInfo.Text = "{} | {}".format(
            doc.ProjectInformation.Name or "Untitled",
            datetime.datetime.now().strftime("%Y-%m-%d"))
    
    def _get_controls(self):
        """Get all named controls from XAML"""
        names = [
            "cmbCheckset", "btnNewCheckset", "btnDuplicate", "btnImport",
            "btnExport", "btnDeleteCheckset",
            "txtTotalRules", "txtPassed", "txtFailed", "txtInfo", "txtSkipped",
            "lstRules", "btnSelectAll", "btnSelectNone",
            "txtDetailHeader", "dgResults", "spResults", "pnlRuleDetail", "spRuleDetail",
            "txtRuleId", "txtRuleDesc", "txtRuleType", "txtRuleSeverity",
            "spParams", "btnSaveParams",
            "txtStatus", "btnRunCheck", "btnExportExcel", "btnClose",
            "txtFooterInfo"
        ]
        for name in names:
            ctrl = self.window.FindName(name)
            setattr(self, name, ctrl)
    
    def _bind_events(self):
        """Bind event handlers"""
        self.cmbCheckset.SelectionChanged += self._on_checkset_changed
        self.btnNewCheckset.Click += self._on_new_checkset
        self.btnDuplicate.Click += self._on_duplicate
        self.btnImport.Click += self._on_import
        self.btnExport.Click += self._on_export
        self.btnDeleteCheckset.Click += self._on_delete_checkset
        self.btnSelectAll.Click += self._on_select_all
        self.btnSelectNone.Click += self._on_select_none
        self.lstRules.SelectionChanged += self._on_rule_selected
        self.btnSaveParams.Click += self._on_save_params
        self.btnRunCheck.Click += self._on_run_check
        self.btnExportExcel.Click += self._on_export_excel
        self.btnClose.Click += self._on_close
    
    # =================================================================
    # CHECKSET MANAGEMENT
    # =================================================================
    def _load_checksets(self):
        """Load checkset list into combo"""
        self.cmbCheckset.Items.Clear()
        for name in self.checkset_mgr.list_checksets():
            self.cmbCheckset.Items.Add(name)
        if self.cmbCheckset.Items.Count > 0:
            self.cmbCheckset.SelectedIndex = 0
    
    def _on_checkset_changed(self, sender, args):
        """When checkset selection changes"""
        sel = self.cmbCheckset.SelectedItem
        if sel:
            self.current_checkset = self.checkset_mgr.load_checkset(str(sel))
            self._refresh_rules()
            self._reset_results()
    
    def _refresh_rules(self):
        """Refresh rule list from current checkset"""
        self.lstRules.Items.Clear()
        self.rule_items = []
        
        if not self.current_checkset:
            return
        
        rules = self.current_checkset.get("rules", [])
        converter = BrushConverter()
        
        for i, rule in enumerate(rules):
            item = RuleDisplayItem(rule, i)
            self.rule_items.append(item)
            
            # Create CheckBox + label manually (IronPython DataTemplate binding workaround)
            sp = StackPanel()
            sp.Orientation = System.Windows.Controls.Orientation.Horizontal
            sp.Margin = System.Windows.Thickness(2, 3, 2, 3)
            sp.Tag = i
            
            chk = CheckBox()
            is_enabled = bool(rule.get("enabled", True))
            chk.IsChecked = System.Nullable[System.Boolean](is_enabled)
            chk.VerticalAlignment = System.Windows.VerticalAlignment.Center
            chk.Margin = System.Windows.Thickness(0, 0, 6, 0)
            chk.Tag = i
            chk.Checked += self._on_rule_toggled
            chk.Unchecked += self._on_rule_toggled
            
            # Severity indicator using unicode circle
            severity = rule.get("severity", "info")
            sev_symbols = {"error": u"\u2B24", "warning": u"\u2B24", "info": u"\u2B24"}
            sev_colors = {"error": "#C62828", "warning": "#F57F17", "info": "#1565C0"}
            
            dot = TextBlock()
            dot.Text = sev_symbols.get(severity, u"\u2B24")
            dot.FontSize = 8
            dot.VerticalAlignment = System.Windows.VerticalAlignment.Center
            dot.Margin = System.Windows.Thickness(0, 0, 6, 0)
            try:
                dot.Foreground = converter.ConvertFromString(sev_colors.get(severity, "#999"))
            except:
                dot.Foreground = converter.ConvertFromString("#999")
            
            lbl = TextBlock()
            lbl.Text = str(item.DisplayName)
            lbl.FontSize = 11
            lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            try:
                lbl.Foreground = converter.ConvertFromString("#0F172A")
            except:
                pass
            
            sp.Children.Add(chk)
            sp.Children.Add(dot)
            sp.Children.Add(lbl)
            
            self.lstRules.Items.Add(sp)
        
        # Update total
        enabled_count = len([r for r in rules if r.get("enabled", True)])
        self.txtTotalRules.Text = str(len(rules))
        self.txtStatus.Text = "{} rules loaded, {} enabled".format(len(rules), enabled_count)
    
    def _on_rule_toggled(self, sender, args):
        """When a rule checkbox is toggled"""
        idx = sender.Tag
        if idx is not None and self.current_checkset:
            rules = self.current_checkset.get("rules", [])
            if 0 <= idx < len(rules):
                rules[idx]["enabled"] = bool(sender.IsChecked)
                # Auto-save
                sel = self.cmbCheckset.SelectedItem
                if sel:
                    self.checkset_mgr.save_checkset(str(sel), self.current_checkset)
    
    def _on_select_all(self, sender, args):
        if self.current_checkset:
            for rule in self.current_checkset.get("rules", []):
                rule["enabled"] = True
            self._refresh_rules()
            sel = self.cmbCheckset.SelectedItem
            if sel:
                self.checkset_mgr.save_checkset(str(sel), self.current_checkset)
    
    def _on_select_none(self, sender, args):
        if self.current_checkset:
            for rule in self.current_checkset.get("rules", []):
                rule["enabled"] = False
            self._refresh_rules()
            sel = self.cmbCheckset.SelectedItem
            if sel:
                self.checkset_mgr.save_checkset(str(sel), self.current_checkset)
    
    def _on_new_checkset(self, sender, args):
        """Create new checkset"""
        from System.Windows.Forms import Form, TextBox, Button as WFButton, Label, DialogResult
        
        form = Form()
        form.Text = "New Checkset"
        form.Width = 350
        form.Height = 150
        form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        
        lbl = Label()
        lbl.Text = "Checkset Name:"
        lbl.Top = 15
        lbl.Left = 10
        lbl.Width = 100
        form.Controls.Add(lbl)
        
        txt = TextBox()
        txt.Top = 12
        txt.Left = 115
        txt.Width = 200
        form.Controls.Add(txt)
        
        btn_ok = WFButton()
        btn_ok.Text = "Create"
        btn_ok.Top = 55
        btn_ok.Left = 115
        btn_ok.Width = 95
        btn_ok.DialogResult = DialogResult.OK
        form.Controls.Add(btn_ok)
        
        btn_cancel = WFButton()
        btn_cancel.Text = "Cancel"
        btn_cancel.Top = 55
        btn_cancel.Left = 220
        btn_cancel.Width = 95
        btn_cancel.DialogResult = DialogResult.Cancel
        form.Controls.Add(btn_cancel)
        
        form.AcceptButton = btn_ok
        form.CancelButton = btn_cancel
        
        if form.ShowDialog() == DialogResult.OK and txt.Text.strip():
            name = txt.Text.strip().replace(" ", "_")
            new_data = dict(DEFAULT_CHECKSET)
            new_data["name"] = txt.Text.strip()
            new_data["created"] = ""
            self.checkset_mgr.save_checkset(name, new_data)
            self._load_checksets()
            # Select the new one
            for i in range(self.cmbCheckset.Items.Count):
                if str(self.cmbCheckset.Items[i]) == name:
                    self.cmbCheckset.SelectedIndex = i
                    break
        form.Dispose()
    
    def _on_duplicate(self, sender, args):
        """Duplicate current checkset"""
        sel = self.cmbCheckset.SelectedItem
        if not sel:
            return
        
        from System.Windows.Forms import Form, TextBox, Button as WFButton, Label, DialogResult
        
        form = Form()
        form.Text = "Duplicate Checkset"
        form.Width = 350
        form.Height = 150
        form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        
        lbl = Label()
        lbl.Text = "New Name:"
        lbl.Top = 15
        lbl.Left = 10
        lbl.Width = 100
        form.Controls.Add(lbl)
        
        txt = TextBox()
        txt.Top = 12
        txt.Left = 115
        txt.Width = 200
        txt.Text = str(sel) + "_copy"
        form.Controls.Add(txt)
        
        btn_ok = WFButton()
        btn_ok.Text = "Duplicate"
        btn_ok.Top = 55
        btn_ok.Left = 115
        btn_ok.Width = 95
        btn_ok.DialogResult = DialogResult.OK
        form.Controls.Add(btn_ok)
        
        btn_cancel = WFButton()
        btn_cancel.Text = "Cancel"
        btn_cancel.Top = 55
        btn_cancel.Left = 220
        btn_cancel.Width = 95
        btn_cancel.DialogResult = DialogResult.Cancel
        form.Controls.Add(btn_cancel)
        
        form.AcceptButton = btn_ok
        form.CancelButton = btn_cancel
        
        if form.ShowDialog() == DialogResult.OK and txt.Text.strip():
            new_name = txt.Text.strip().replace(" ", "_")
            self.checkset_mgr.duplicate_checkset(str(sel), new_name)
            self._load_checksets()
            for i in range(self.cmbCheckset.Items.Count):
                if str(self.cmbCheckset.Items[i]) == new_name:
                    self.cmbCheckset.SelectedIndex = i
                    break
        form.Dispose()
    
    def _on_import(self, sender, args):
        """Import checkset from file"""
        from System.Windows.Forms import OpenFileDialog, DialogResult
        
        dlg = OpenFileDialog()
        dlg.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        dlg.Title = "Import Checkset"
        
        if dlg.ShowDialog() == DialogResult.OK:
            name = self.checkset_mgr.import_checkset(dlg.FileName)
            if name:
                self._load_checksets()
                for i in range(self.cmbCheckset.Items.Count):
                    if str(self.cmbCheckset.Items[i]) == name:
                        self.cmbCheckset.SelectedIndex = i
                        break
                self.txtStatus.Text = "Imported checkset: {}".format(name)
            else:
                System.Windows.MessageBox.Show(
                    "Failed to import checkset. Invalid JSON format.",
                    "Import Error",
                    MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_export(self, sender, args):
        """Export checkset to file"""
        sel = self.cmbCheckset.SelectedItem
        if not sel:
            return
        
        from System.Windows.Forms import SaveFileDialog, DialogResult
        
        dlg = SaveFileDialog()
        dlg.Filter = "JSON Files (*.json)|*.json"
        dlg.FileName = str(sel) + ".json"
        dlg.Title = "Export Checkset"
        
        if dlg.ShowDialog() == DialogResult.OK:
            if self.checkset_mgr.export_checkset(str(sel), dlg.FileName):
                self.txtStatus.Text = "Exported to: {}".format(dlg.FileName)
            else:
                System.Windows.MessageBox.Show(
                    "Failed to export checkset.",
                    "Export Error",
                    MessageBoxButton.OK, MessageBoxImage.Error)
    
    def _on_delete_checkset(self, sender, args):
        """Delete current checkset"""
        sel = self.cmbCheckset.SelectedItem
        if not sel or str(sel) == "default":
            System.Windows.MessageBox.Show(
                "Cannot delete the default checkset.",
                "Delete", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        result = System.Windows.MessageBox.Show(
            "Delete checkset '{}'?\nThis cannot be undone.".format(sel),
            "Confirm Delete",
            MessageBoxButton.YesNo, MessageBoxImage.Warning)
        
        if result == MessageBoxResult.Yes:
            self.checkset_mgr.delete_checkset(str(sel))
            self._load_checksets()
    
    # =================================================================
    # RULE DETAIL PANEL
    # =================================================================
    def _on_rule_selected(self, sender, args):
        """When a rule is selected in the list"""
        idx = self.lstRules.SelectedIndex
        if idx < 0 or not self.current_checkset:
            return
        
        rules = self.current_checkset.get("rules", [])
        if idx >= len(rules):
            return
        
        rule = rules[idx]
        
        # Show detail panel, hide results
        self.pnlRuleDetail.Visibility = System.Windows.Visibility.Visible
        self.dgResults.Visibility = System.Windows.Visibility.Collapsed
        
        self.txtDetailHeader.Text = rule.get("name", "")
        self.txtRuleId.Text = "ID: {}  |  Category: {}".format(
            rule.get("id", ""), rule.get("category", ""))
        self.txtRuleDesc.Text = rule.get("description", "")
        self.txtRuleType.Text = "Type: {}".format(rule.get("type", ""))
        
        severity = rule.get("severity", "info")
        sev_colors = {"error": "#C62828", "warning": "#F57F17", "info": "#1565C0"}
        self.txtRuleSeverity.Text = "Severity: {}".format(severity.upper())
        converter = BrushConverter()
        self.txtRuleSeverity.Foreground = converter.ConvertFromString(
            sev_colors.get(severity, "#666"))
        
        # Build parameter editing UI
        self.spParams.Children.Clear()
        params = rule.get("params", {})
        
        for key, value in params.items():
            param_sp = StackPanel()
            param_sp.Orientation = System.Windows.Controls.Orientation.Horizontal
            param_sp.Margin = System.Windows.Thickness(0, 2, 0, 2)
            
            lbl = TextBlock()
            lbl.Text = "{}:".format(key)
            lbl.Width = 150
            lbl.FontSize = 11
            lbl.Foreground = converter.ConvertFromString("#5D4E37")
            lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            
            if isinstance(value, bool):
                ctrl = CheckBox()
                ctrl.IsChecked = value
                ctrl.Tag = key
            elif isinstance(value, list):
                ctrl = TextBox()
                ctrl.Text = ", ".join(str(v) for v in value)
                ctrl.Width = 280
                ctrl.Padding = System.Windows.Thickness(4, 2, 4, 2)
                ctrl.Tag = key
                ctrl.ToolTip = "Comma-separated values"
            else:
                ctrl = TextBox()
                ctrl.Text = str(value)
                ctrl.Width = 280
                ctrl.Padding = System.Windows.Thickness(4, 2, 4, 2)
                ctrl.Tag = key
            
            param_sp.Children.Add(lbl)
            param_sp.Children.Add(ctrl)
            self.spParams.Children.Add(param_sp)
    
    def _on_save_params(self, sender, args):
        """Save edited parameters back to checkset"""
        idx = self.lstRules.SelectedIndex
        if idx < 0 or not self.current_checkset:
            return
        
        rules = self.current_checkset.get("rules", [])
        if idx >= len(rules):
            return
        
        rule = rules[idx]
        params = rule.get("params", {})
        
        # Read values from UI controls
        for i in range(self.spParams.Children.Count):
            sp = self.spParams.Children[i]
            if sp.Children.Count >= 2:
                ctrl = sp.Children[1]
                key = str(ctrl.Tag) if ctrl.Tag else None
                if not key or key not in params:
                    continue
                
                original_value = params[key]
                
                if isinstance(ctrl, CheckBox):
                    params[key] = bool(ctrl.IsChecked)
                elif isinstance(ctrl, TextBox):
                    text = ctrl.Text.strip()
                    if isinstance(original_value, list):
                        # Parse comma-separated back to list
                        params[key] = [v.strip() for v in text.split(",") if v.strip()]
                    elif isinstance(original_value, float):
                        try:
                            params[key] = float(text)
                        except:
                            pass
                    elif isinstance(original_value, int):
                        try:
                            params[key] = int(text)
                        except:
                            pass
                    else:
                        params[key] = text
        
        # Save checkset
        sel = self.cmbCheckset.SelectedItem
        if sel:
            self.checkset_mgr.save_checkset(str(sel), self.current_checkset)
            self.txtStatus.Text = "Parameters saved for rule: {}".format(rule.get("name", ""))
    
    # =================================================================
    # RUN CHECK
    # =================================================================
    def _on_run_check(self, sender, args):
        """Execute all enabled rules"""
        if not self.current_checkset:
            System.Windows.MessageBox.Show(
                "Please select a checkset first.",
                "No Checkset", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        
        self.txtStatus.Text = "Running checks..."
        self.window.Cursor = System.Windows.Input.Cursors.Wait
        
        try:
            # Run rules
            self.current_results = self.rule_engine.run_checkset(self.current_checkset)
            
            # Update summary cards
            total = len(self.current_results)
            passed = len([r for r in self.current_results if r.status == "pass"])
            failed = len([r for r in self.current_results if r.status == "fail"])
            info_count = len([r for r in self.current_results if r.status in ("info", "warning")])
            skipped = len([r for r in self.current_results if r.status in ("skipped", "error")])
            
            self.txtTotalRules.Text = str(total)
            self.txtPassed.Text = str(passed)
            self.txtFailed.Text = str(failed)
            self.txtInfo.Text = str(info_count)
            self.txtSkipped.Text = str(skipped)
            
            # Show results panel
            self.spResults.Children.Clear()
            self.dgResults.Visibility = System.Windows.Visibility.Visible
            self.pnlRuleDetail.Visibility = System.Windows.Visibility.Collapsed
            self.txtDetailHeader.Text = "Check Results"
            
            converter = BrushConverter()
            
            # Status styling
            status_icons = {
                "pass": u"\u2714", "fail": u"\u2718", "warning": u"\u26A0",
                "info": u"\u2139", "error": u"\u2716", "skipped": u"\u23F8"
            }
            status_bg = {
                "pass": "#E8F5E9", "fail": "#FFEBEE", "warning": "#FFF8E1",
                "info": "#E3F2FD", "error": "#FBE9E7", "skipped": "#ECEFF1"
            }
            status_fg = {
                "pass": "#2E7D32", "fail": "#C62828", "warning": "#F57F17",
                "info": "#1565C0", "error": "#D84315", "skipped": "#546E7A"
            }
            
            for result in self.current_results:
                # Main result row
                row_border = System.Windows.Controls.Border()
                row_border.Margin = System.Windows.Thickness(0, 0, 0, 2)
                row_border.Padding = System.Windows.Thickness(8, 6, 8, 6)
                row_border.CornerRadius = System.Windows.CornerRadius(3)
                try:
                    row_border.Background = converter.ConvertFromString(
                        status_bg.get(result.status, "#FAFAFA"))
                except:
                    row_border.Background = converter.ConvertFromString("#FAFAFA")
                
                row_grid = Grid()
                col1 = ColumnDefinition()
                col1.Width = System.Windows.GridLength(45)
                col2 = ColumnDefinition()
                col2.Width = System.Windows.GridLength(65)
                col3 = ColumnDefinition()
                col3.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
                row_grid.ColumnDefinitions.Add(col1)
                row_grid.ColumnDefinitions.Add(col2)
                row_grid.ColumnDefinitions.Add(col3)
                
                # Status icon
                icon_txt = TextBlock()
                icon_txt.Text = status_icons.get(result.status, "?")
                icon_txt.FontSize = 14
                icon_txt.FontWeight = System.Windows.FontWeights.Bold
                icon_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
                try:
                    icon_txt.Foreground = converter.ConvertFromString(
                        status_fg.get(result.status, "#666"))
                except:
                    pass
                Grid.SetColumn(icon_txt, 0)
                row_grid.Children.Add(icon_txt)
                
                # Severity badge
                sev_txt = TextBlock()
                sev_txt.Text = result.severity.upper()
                sev_txt.FontSize = 9
                sev_txt.FontWeight = System.Windows.FontWeights.SemiBold
                sev_txt.VerticalAlignment = System.Windows.VerticalAlignment.Center
                try:
                    sev_txt.Foreground = converter.ConvertFromString("#888")
                except:
                    pass
                Grid.SetColumn(sev_txt, 1)
                row_grid.Children.Add(sev_txt)
                
                # Rule name + message
                info_sp = StackPanel()
                info_sp.VerticalAlignment = System.Windows.VerticalAlignment.Center
                
                name_txt = TextBlock()
                name_txt.Text = u"[{}] {}".format(result.rule_id, result.rule_name)
                name_txt.FontSize = 11
                name_txt.FontWeight = System.Windows.FontWeights.SemiBold
                try:
                    name_txt.Foreground = converter.ConvertFromString("#333")
                except:
                    pass
                info_sp.Children.Add(name_txt)
                
                msg_txt = TextBlock()
                msg_txt.Text = result.message
                msg_txt.FontSize = 10
                msg_txt.TextWrapping = System.Windows.TextWrapping.Wrap
                try:
                    msg_txt.Foreground = converter.ConvertFromString(
                        status_fg.get(result.status, "#666"))
                except:
                    pass
                msg_txt.Margin = System.Windows.Thickness(0, 2, 0, 0)
                info_sp.Children.Add(msg_txt)
                
                # Details (sub-items)
                if result.details:
                    for detail in result.details:
                        det_txt = TextBlock()
                        det_txt.Text = str(detail)
                        det_txt.FontSize = 9
                        try:
                            det_txt.Foreground = converter.ConvertFromString("#888")
                        except:
                            pass
                        det_txt.Margin = System.Windows.Thickness(8, 1, 0, 0)
                        info_sp.Children.Add(det_txt)
                
                Grid.SetColumn(info_sp, 2)
                row_grid.Children.Add(info_sp)
                
                row_border.Child = row_grid
                self.spResults.Children.Add(row_border)
            
            # Enable export
            self.btnExportExcel.IsEnabled = True
            
            self.txtStatus.Text = "Check complete: {} passed, {} failed, {} info, {} skipped".format(
                passed, failed, info_count, skipped)
            
        except Exception as e:
            self.txtStatus.Text = "Error: {}".format(str(e))
            System.Windows.MessageBox.Show(
                "Error running checks:\n{}".format(traceback.format_exc()),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self.window.Cursor = System.Windows.Input.Cursors.Arrow
    
    def _reset_results(self):
        """Reset results display"""
        self.current_results = None
        self.spResults.Children.Clear()
        self.dgResults.Visibility = System.Windows.Visibility.Collapsed
        self.pnlRuleDetail.Visibility = System.Windows.Visibility.Visible
        self.btnExportExcel.IsEnabled = False
        self.txtPassed.Text = "0"
        self.txtFailed.Text = "0"
        self.txtInfo.Text = "0"
        self.txtSkipped.Text = "0"
        self.txtDetailHeader.Text = "Select a rule to view details"
    
    # =================================================================
    # EXPORT EXCEL
    # =================================================================
    def _on_export_excel(self, sender, args):
        """Export results to Excel"""
        if not self.current_results or not self.current_checkset:
            return
        
        from System.Windows.Forms import SaveFileDialog, DialogResult
        
        dlg = SaveFileDialog()
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        project_name = doc.ProjectInformation.Name or "Untitled"
        checkset_name = str(self.cmbCheckset.SelectedItem or "default")
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        dlg.FileName = "ModelCheck_{}_{}_{}".format(project_name, checkset_name, timestamp)
        dlg.Title = "Export Check Report"
        
        # Default to reports directory
        dlg.InitialDirectory = REPORTS_DIR
        
        if dlg.ShowDialog() == DialogResult.OK:
            self.txtStatus.Text = "Exporting to Excel..."
            self.window.Cursor = System.Windows.Input.Cursors.Wait
            
            try:
                self.reporter.generate_report(
                    self.current_checkset,
                    self.current_results,
                    dlg.FileName)
                
                self.txtStatus.Text = "Report exported: {}".format(os.path.basename(dlg.FileName))
                
                # Ask to open
                result = System.Windows.MessageBox.Show(
                    "Report exported successfully!\n\nOpen the file now?",
                    "Export Complete",
                    MessageBoxButton.YesNo, MessageBoxImage.Information)
                
                if result == MessageBoxResult.Yes:
                    os.startfile(dlg.FileName)
                    
            except Exception as e:
                self.txtStatus.Text = "Export failed"
                System.Windows.MessageBox.Show(
                    "Error exporting:\n{}".format(str(e)),
                    "Export Error",
                    MessageBoxButton.OK, MessageBoxImage.Error)
            finally:
                self.window.Cursor = System.Windows.Input.Cursors.Arrow
    
    def _on_close(self, sender, args):
        self.window.Close()
    
    def show(self):
        self.window.ShowDialog()


# =====================================================================
# MAIN ENTRY POINT
# =====================================================================
try:
    window = ModelCheckerWindow()
    window.show()
except Exception as e:
    print("Error: {}".format(str(e)))
    print(traceback.format_exc())