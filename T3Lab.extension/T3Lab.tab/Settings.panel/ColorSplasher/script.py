# -*- coding: utf-8 -*-
"""
Color Splasher v4.0 - DQT
Auto-color elements in active view based on parameter values.
Features: Gradient/Random colors, Create Legend, Create View Filters, Reset.

Copyright (c) 2026 by Dang Quoc Truong (DQT)
All rights reserved.
"""

__title__ = "Color\nSplasher"
__author__ = "Dang Quoc Truong (DQT)"
__doc__ = "Auto-color elements by parameter value. Visualize model data with color overrides."

from pyrevit import revit, DB, forms, script
from collections import OrderedDict
import random
import math

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')

import System
from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    GridLength, GridUnitType, WindowStartupLocation, TextTrimming
)
from System.Windows.Controls import (
    Grid as WPFGrid, StackPanel, TextBlock, TextBox, Button,
    ComboBox, ListBox, Border, ScrollViewer, ColumnDefinition,
    RowDefinition, ScrollBarVisibility, Orientation, SelectionMode
)
from System.Windows.Media import SolidColorBrush, Color

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# =====================================================
# DQT BRAND COLORS (synced with IFC-SG Checker)
# =====================================================
CLR_HEADER = Color.FromRgb(240, 204, 136)       # #0F172A Gold header
CLR_HEADER_TEXT = Color.FromRgb(51, 51, 51)      # #0F172A Dark text on gold
CLR_HEADER_SUB = Color.FromRgb(102, 102, 102)    # #475569 Subtitle
CLR_ACCENT = Color.FromRgb(93, 78, 55)           # #5D4E37 DQT brown accent
CLR_BG = Color.FromRgb(254, 248, 231)            # #F8FAFC Cream background
CLR_CARD = Color.FromRgb(255, 255, 255)
CLR_BORDER = Color.FromRgb(212, 184, 122)        # #CBD5E1 Gold border
CLR_FOOTER = Color.FromRgb(245, 240, 224)        # #F5F0E0 Footer bg
CLR_TEXT = Color.FromRgb(51, 51, 51)              # #0F172A
CLR_MUTED = Color.FromRgb(153, 153, 153)         # #94A3B8
CLR_ALT = Color.FromRgb(255, 248, 238)           # #FFF8EE

# =====================================================
# COLOR PALETTE - Vivid distinct colors for element fill
# =====================================================
PALETTE = [
    (231,76,60),(46,134,193),(39,174,96),(243,156,18),(142,68,173),
    (26,188,156),(211,84,0),(52,73,94),(22,160,133),(192,57,43),
    (41,128,185),(46,204,113),(230,126,34),(155,89,182),(52,152,219),
    (241,196,15),(127,140,141),(44,62,80),(220,118,51),(86,101,115),
    (175,122,197),(69,179,157),(205,97,85),(93,109,126),(162,155,254),
    (72,201,176),(189,195,199),(149,165,166),(255,127,80),(100,149,237),
]


# =====================================================
# ElementId helper
# =====================================================
def eid_int(eid):
    if eid is None:
        return -1
    for attr in ['IntegerValue', 'Value']:
        try:
            return int(getattr(eid, attr))
        except:
            continue
    try:
        import re
        m = re.search(r'(\d+)', str(eid))
        if m:
            return int(m.group(1))
    except:
        pass
    return -1


# =====================================================
# SPECIAL PARAMETER NAMES (computed, not from LookupParameter)
# =====================================================
SPECIAL_PARAMS = [
    "Family and Type",
    "Type Name (System)",
    "Family Name (System)",
    "Category Name",
    "Level Name",
]


def get_special_value(elem, param_name):
    """Get special computed parameter values"""
    try:
        if param_name == "Family and Type":
            # Use BuiltInParameter
            p = elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
            if p and p.HasValue:
                return p.AsValueString() or "<Empty>"
            # Fallback: build from type
            tid = elem.GetTypeId()
            if tid:
                etype = doc.GetElement(tid)
                if etype:
                    fname = ""
                    tname = etype.Name or ""
                    fp = etype.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                    if fp and fp.HasValue:
                        fname = fp.AsString() or ""
                    if not fname:
                        fp = etype.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                        if fp and fp.HasValue:
                            fname = fp.AsString() or ""
                    if fname:
                        return "{} : {}".format(fname, tname)
                    return tname
            return "<No Type>"
        
        elif param_name == "Type Name (System)":
            tid = elem.GetTypeId()
            if tid:
                etype = doc.GetElement(tid)
                if etype:
                    return etype.Name or "<Unnamed>"
            return "<No Type>"
        
        elif param_name == "Family Name (System)":
            tid = elem.GetTypeId()
            if tid:
                etype = doc.GetElement(tid)
                if etype:
                    fp = etype.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                    if fp and fp.HasValue:
                        return fp.AsString() or "<Unknown>"
                    fp = etype.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                    if fp and fp.HasValue:
                        return fp.AsString() or "<Unknown>"
            return "<No Family>"
        
        elif param_name == "Category Name":
            if elem.Category:
                return elem.Category.Name
            return "<No Category>"
        
        elif param_name == "Level Name":
            lp = elem.get_Parameter(DB.BuiltInParameter.WALL_BASE_CONSTRAINT)
            if lp and lp.HasValue:
                lid = lp.AsElementId()
                lev = doc.GetElement(lid)
                if lev:
                    return lev.Name
            # Generic level param
            lp = elem.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
            if lp and lp.HasValue:
                lid = lp.AsElementId()
                lev = doc.GetElement(lid)
                if lev:
                    return lev.Name
            lp = elem.get_Parameter(DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM)
            if lp and lp.HasValue:
                lid = lp.AsElementId()
                lev = doc.GetElement(lid)
                if lev:
                    return lev.Name
            return "<No Level>"
    except:
        pass
    return "<Error>"


# =====================================================
# LINKED MODEL SUPPORT
# =====================================================

def get_link_instances():
    """Get all RevitLinkInstance in the document"""
    links = []
    try:
        for link in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance):
            try:
                link_doc = link.GetLinkDocument()
                if link_doc:
                    links.append((link, link_doc))
            except:
                continue
    except:
        pass
    return links


def collect_categories_from_doc(source_doc, view_id=None):
    """Collect model categories from a given document"""
    cats = {}
    SKIP_NAMES = {
        'Views', 'Sheets', 'Schedules', 'Schedule Graphics', 'Viewports',
        'Scope Boxes', 'Matchline', 'Reference Planes', 'Grids', 'Levels',
        'Section Boxes', 'Cameras', 'Title Blocks', 'Revision Clouds',
        'Detail Items', 'Lines', 'Text Notes', 'Dimensions', 'Tags',
        'Keynote Tags', 'Multi-Category Tags', 'Generic Annotations',
        'Spot Elevations', 'Spot Coordinates', 'Spot Slopes',
        'Curtain Grid Lines', 'Curtain Grid Mullions',
        'Area Tags', 'Room Tags', 'Space Tags',
        'Project Information', 'Project Base Point', 'Survey Point',
        'Mass', 'Mass Floor', 'Analytical Links', 'Analytical Nodes',
    }
    try:
        col = DB.FilteredElementCollector(source_doc).WhereElementIsNotElementType()
        for elem in col:
            try:
                cat = elem.Category
                if cat is None:
                    continue
                cname = cat.Name
                if not cname or cname.startswith('<') or cname.startswith('_'):
                    continue
                if cname in SKIP_NAMES:
                    continue
                try:
                    if cat.CategoryType != DB.CategoryType.Model:
                        continue
                except:
                    pass
                if cname not in cats:
                    cats[cname] = cat
            except:
                continue
    except:
        pass
    return sorted(cats.items(), key=lambda x: x[0])


def collect_elements_from_doc(source_doc, category_name):
    """Collect elements matching category name from a document"""
    results = []
    try:
        for elem in DB.FilteredElementCollector(source_doc).WhereElementIsNotElementType():
            try:
                if elem.Category and elem.Category.Name == category_name:
                    results.append(elem)
            except:
                continue
    except:
        pass
    return results


def get_param_value_linked(elem, param_name, source_doc):
    """Get parameter value from a linked element"""
    if param_name in SPECIAL_PARAMS:
        return get_special_value_linked(elem, param_name, source_doc)
    
    def _read(param):
        if param is None or not param.HasValue:
            return None
        st = param.StorageType
        if st == DB.StorageType.String:
            v = param.AsString()
            return v if v else "<Empty>"
        elif st == DB.StorageType.Double:
            return param.AsValueString() or str(round(param.AsDouble(), 4))
        elif st == DB.StorageType.Integer:
            return param.AsValueString() or str(param.AsInteger())
        elif st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            ev = eid_int(eid)
            if ev > 0:
                el = source_doc.GetElement(eid)
                if el:
                    try:
                        return el.Name
                    except:
                        return str(ev)
            return param.AsValueString() or "<None>"
        return None
    
    try:
        val = _read(elem.LookupParameter(param_name))
        if val is not None:
            return val
    except:
        pass
    try:
        tid = elem.GetTypeId()
        if tid:
            etype = source_doc.GetElement(tid)
            if etype:
                val = _read(etype.LookupParameter(param_name))
                if val is not None:
                    return val
    except:
        pass
    return "<No Value>"


def get_special_value_linked(elem, param_name, source_doc):
    """Get special computed param value from linked element"""
    try:
        if param_name == "Family and Type":
            p = elem.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
            if p and p.HasValue:
                return p.AsValueString() or "<Empty>"
            tid = elem.GetTypeId()
            if tid:
                etype = source_doc.GetElement(tid)
                if etype:
                    fname = ""
                    tname = etype.Name or ""
                    fp = etype.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                    if fp and fp.HasValue:
                        fname = fp.AsString() or ""
                    if fname:
                        return "{} : {}".format(fname, tname)
                    return tname
            return "<No Type>"
        elif param_name == "Type Name (System)":
            tid = elem.GetTypeId()
            if tid:
                etype = source_doc.GetElement(tid)
                if etype:
                    return etype.Name or "<Unnamed>"
            return "<No Type>"
        elif param_name == "Family Name (System)":
            tid = elem.GetTypeId()
            if tid:
                etype = source_doc.GetElement(tid)
                if etype:
                    fp = etype.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                    if fp and fp.HasValue:
                        return fp.AsString() or "<Unknown>"
            return "<No Family>"
        elif param_name == "Category Name":
            if elem.Category:
                return elem.Category.Name
            return "<No Category>"
        elif param_name == "Level Name":
            for bip in [DB.BuiltInParameter.WALL_BASE_CONSTRAINT,
                       DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                       DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM]:
                lp = elem.get_Parameter(bip)
                if lp and lp.HasValue:
                    lid = lp.AsElementId()
                    lev = source_doc.GetElement(lid)
                    if lev:
                        return lev.Name
            return "<No Level>"
    except:
        pass
    return "<Error>"


def apply_overrides_linked(elements_by_value, color_map, link_instance):
    """Apply color to linked elements via View Filters (only way for links).
    SetElementOverrides does NOT work for individual linked elements.
    View Filters are the only API method that works across links."""
    active_view = doc.ActiveView
    link_doc = link_instance.GetLinkDocument()
    link_name = link_doc.Title if link_doc else "Link"
    
    t = DB.Transaction(doc, "DQT - Color Splasher (Link Filters)")
    t.Start()
    count = 0
    solid = get_solid_fill()
    
    try:
        # Find the category from elements
        cat_name = None
        sample_elem = None
        for val, elems in elements_by_value.items():
            if elems:
                sample_elem = elems[0]
                if sample_elem.Category:
                    cat_name = sample_elem.Category.Name
                break
        
        if not cat_name:
            t.RollBack()
            return 0
        
        # Get BuiltInCategory from host doc matching the category name
        host_cat = None
        try:
            for cat in doc.Settings.Categories:
                try:
                    if cat.Name == cat_name:
                        host_cat = cat
                        break
                except:
                    continue
        except:
            pass
        
        if not host_cat:
            t.RollBack()
            return 0
        
        cat_id_list = System.Collections.Generic.List[DB.ElementId]()
        cat_id_list.Add(host_cat.Id)
        
        for val, elems in elements_by_value.items():
            if val not in color_map:
                continue
            r, g, b = color_map[val]
            
            # Create filter name
            filter_name = "{} - {} - {}".format(link_name, cat_name, val)
            if len(filter_name) > 200:
                filter_name = filter_name[:200]
            
            # Check existing
            pfilter = None
            try:
                for pfe in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
                    try:
                        if pfe.Name == filter_name:
                            pfilter = pfe
                            break
                    except:
                        continue
            except:
                pass
            
            if not pfilter:
                try:
                    # Get param from linked element
                    sample = elems[0]
                    param = sample.LookupParameter(cat_name)
                    # Try common parameters
                    for pname in ["Type Name", "Family and Type", "Mark", "Comments"]:
                        param = sample.LookupParameter(pname)
                        if param and param.HasValue:
                            break
                    
                    # Try type parameter
                    if not param or not param.HasValue:
                        tid = sample.GetTypeId()
                        if tid:
                            etype = link_doc.GetElement(tid)
                            if etype:
                                for pname in ["Type Name", "Family Name"]:
                                    param = etype.LookupParameter(pname)
                                    if param and param.HasValue:
                                        break
                    
                    if param and param.HasValue:
                        param_id = param.Id
                        rule = None
                        st = param.StorageType
                        if st == DB.StorageType.String:
                            str_val = param.AsString() or ""
                            try:
                                rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, str_val, True)
                            except:
                                rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, str_val)
                        elif st == DB.StorageType.Integer:
                            rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, param.AsInteger())
                        elif st == DB.StorageType.Double:
                            rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, param.AsDouble(), 0.001)
                        
                        if rule:
                            elem_filter = DB.ElementParameterFilter(rule)
                            pfilter = DB.ParameterFilterElement.Create(doc, filter_name, cat_id_list, elem_filter)
                except:
                    continue
            
            if pfilter:
                try:
                    active_view.AddFilter(pfilter.Id)
                    active_view.SetFilterVisibility(pfilter.Id, True)
                    
                    color = DB.Color(r, g, b)
                    ogs = DB.OverrideGraphicSettings()
                    try:
                        ogs.SetSurfaceForegroundPatternColor(color)
                        ogs.SetSurfaceForegroundPatternVisible(True)
                        if solid:
                            ogs.SetSurfaceForegroundPatternId(solid.Id)
                    except:
                        pass
                    try:
                        ogs.SetCutForegroundPatternColor(color)
                        ogs.SetCutForegroundPatternVisible(True)
                        if solid:
                            ogs.SetCutForegroundPatternId(solid.Id)
                    except:
                        pass
                    active_view.SetFilterOverrides(pfilter.Id, ogs)
                    count += len(elems)
                except:
                    continue
        
        t.Commit()
    except:
        if t.HasStarted():
            t.RollBack()
    return count


def collect_categories():
    cats = {}
    active_view = doc.ActiveView
    
    # Categories to exclude (annotation, views, internal)
    SKIP_NAMES = {
        'Views', 'Sheets', 'Schedules', 'Schedule Graphics', 'Viewports',
        'Scope Boxes', 'Matchline', 'Reference Planes', 'Grids', 'Levels',
        'Section Boxes', 'Cameras', 'Title Blocks', 'Revision Clouds',
        'Detail Items', 'Lines', 'Text Notes', 'Dimensions', 'Tags',
        'Keynote Tags', 'Multi-Category Tags', 'Generic Annotations',
        'Spot Elevations', 'Spot Coordinates', 'Spot Slopes',
        'Curtain Grid Lines', 'Curtain Grid Mullions',
        'Area Tags', 'Room Tags', 'Space Tags',
        'Project Information', 'Project Base Point', 'Survey Point',
        'Mass', 'Mass Floor', 'Analytical Links', 'Analytical Nodes',
    }
    
    try:
        for elem in DB.FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType():
            try:
                cat = elem.Category
                if cat is None:
                    continue
                cname = cat.Name
                if not cname:
                    continue
                if cname.startswith('<') or cname.startswith('_'):
                    continue
                if cname in SKIP_NAMES:
                    continue
                # Only model categories (CategoryType.Model)
                try:
                    if cat.CategoryType != DB.CategoryType.Model:
                        continue
                except:
                    pass
                if cname not in cats:
                    cats[cname] = cat
            except:
                continue
    except:
        pass
    
    if not cats:
        try:
            for elem in DB.FilteredElementCollector(doc).WhereElementIsNotElementType():
                try:
                    cat = elem.Category
                    if cat is None:
                        continue
                    cname = cat.Name
                    if not cname or cname.startswith('<') or cname.startswith('_'):
                        continue
                    if cname in SKIP_NAMES:
                        continue
                    try:
                        if cat.CategoryType != DB.CategoryType.Model:
                            continue
                    except:
                        pass
                    if cname not in cats:
                        cats[cname] = cat
                except:
                    continue
        except:
            pass
    
    return sorted(cats.items(), key=lambda x: x[0])


def collect_elements(category):
    active_view = doc.ActiveView
    target = category.Name
    
    # BuiltInCategory approach
    try:
        cid = eid_int(category.Id)
        if cid > 0:
            bic = System.Enum.ToObject(DB.BuiltInCategory, cid)
            elems = list(DB.FilteredElementCollector(doc, active_view.Id)
                        .OfCategory(bic).WhereElementIsNotElementType().ToElements())
            if elems:
                return elems
    except:
        pass
    
    # Name match
    results = []
    try:
        for elem in DB.FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType():
            try:
                if elem.Category and elem.Category.Name == target:
                    results.append(elem)
            except:
                continue
    except:
        pass
    return results


def collect_param_names(category):
    elems = collect_elements(category)
    if not elems:
        return SPECIAL_PARAMS[:]
    names = set()
    for elem in elems[:20]:
        # Instance parameters
        try:
            for p in elem.Parameters:
                try:
                    if p and p.Definition and p.Definition.Name:
                        names.add(p.Definition.Name)
                except:
                    continue
        except:
            continue
        # Type parameters
        try:
            tid = elem.GetTypeId()
            if tid:
                etype = doc.GetElement(tid)
                if etype:
                    for p in etype.Parameters:
                        try:
                            if p and p.Definition and p.Definition.Name:
                                names.add(p.Definition.Name)
                        except:
                            continue
        except:
            continue
    # Prepend special params at top
    result = list(SPECIAL_PARAMS)
    for n in sorted(names):
        if n not in result:
            result.append(n)
    return result


def get_param_value(elem, param_name):
    # Check special params first
    if param_name in SPECIAL_PARAMS:
        return get_special_value(elem, param_name)
    
    # Helper to read param value
    def _read(param):
        if param is None or not param.HasValue:
            return None
        st = param.StorageType
        if st == DB.StorageType.String:
            v = param.AsString()
            return v if v else "<Empty>"
        elif st == DB.StorageType.Double:
            return param.AsValueString() or str(round(param.AsDouble(), 4))
        elif st == DB.StorageType.Integer:
            return param.AsValueString() or str(param.AsInteger())
        elif st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            ev = eid_int(eid)
            if ev > 0:
                el = doc.GetElement(eid)
                if el:
                    try:
                        return el.Name
                    except:
                        return str(ev)
            return param.AsValueString() or "<None>"
        return None
    
    # 1) Try instance parameter
    try:
        val = _read(elem.LookupParameter(param_name))
        if val is not None:
            return val
    except:
        pass
    
    # 2) Try type parameter
    try:
        tid = elem.GetTypeId()
        if tid:
            etype = doc.GetElement(tid)
            if etype:
                val = _read(etype.LookupParameter(param_name))
                if val is not None:
                    return val
    except:
        pass
    
    return "<No Value>"


def get_solid_fill():
    try:
        for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
            try:
                if fp.GetFillPattern().IsSolidFill:
                    return fp
            except:
                continue
    except:
        pass
    return None


def apply_overrides(elements_by_value, color_map):
    active_view = doc.ActiveView
    t = DB.Transaction(doc, "DQT - Color Splasher")
    t.Start()
    count = 0
    solid = get_solid_fill()
    try:
        for val, elems in elements_by_value.items():
            if val not in color_map:
                continue
            r, g, b = color_map[val]
            color = DB.Color(r, g, b)
            ogs = DB.OverrideGraphicSettings()
            try:
                ogs.SetSurfaceForegroundPatternColor(color)
                ogs.SetSurfaceForegroundPatternVisible(True)
                if solid:
                    ogs.SetSurfaceForegroundPatternId(solid.Id)
            except:
                try:
                    ogs.SetProjectionFillColor(color)
                    ogs.SetProjectionFillPatternVisible(True)
                    if solid:
                        ogs.SetProjectionFillPatternId(solid.Id)
                except:
                    pass
            try:
                ogs.SetCutForegroundPatternColor(color)
                ogs.SetCutForegroundPatternVisible(True)
                if solid:
                    ogs.SetCutForegroundPatternId(solid.Id)
            except:
                try:
                    ogs.SetCutFillColor(color)
                    ogs.SetCutFillPatternVisible(True)
                except:
                    pass
            
            for elem in elems:
                try:
                    active_view.SetElementOverrides(elem.Id, ogs)
                    count += 1
                except:
                    continue
        t.Commit()
    except:
        if t.HasStarted():
            t.RollBack()
    return count


def clear_overrides():
    active_view = doc.ActiveView
    t = DB.Transaction(doc, "DQT - Clear Color Overrides")
    t.Start()
    count = 0
    try:
        blank = DB.OverrideGraphicSettings()
        for elem in DB.FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType():
            try:
                active_view.SetElementOverrides(elem.Id, blank)
                count += 1
            except:
                continue
        t.Commit()
    except:
        if t.HasStarted():
            t.RollBack()
    return count


def generate_gradient(n):
    """Generate n gradient colors from blue to red through green"""
    colors = []
    for i in range(n):
        t = float(i) / max(n - 1, 1)
        if t < 0.5:
            r = int(0 + (0) * (t * 2))
            g = int(0 + (200) * (t * 2))
            b = int(200 + (-200) * (t * 2))
        else:
            t2 = (t - 0.5) * 2
            r = int(220 * t2)
            g = int(200 * (1 - t2))
            b = 0
        colors.append((max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b))))
    return colors


def generate_random(n):
    """Generate n random vivid distinct colors"""
    colors = []
    for i in range(n):
        h = (i * 137.508) % 360
        s = 0.6 + random.random() * 0.3
        v = 0.7 + random.random() * 0.25
        c = v * s
        x = c * (1 - abs((h / 60.0) % 2 - 1))
        m = v - c
        if h < 60: r1, g1, b1 = c, x, 0
        elif h < 120: r1, g1, b1 = x, c, 0
        elif h < 180: r1, g1, b1 = 0, c, x
        elif h < 240: r1, g1, b1 = 0, x, c
        elif h < 300: r1, g1, b1 = x, 0, c
        else: r1, g1, b1 = c, 0, x
        colors.append((int((r1+m)*255), int((g1+m)*255), int((b1+m)*255)))
    return colors


# =====================================================
# WPF WINDOW
# =====================================================

class ColorSplasherWindow(Window):
    
    def __init__(self):
        self.categories = []
        self.current_param_names = []
        self.elements_by_value = OrderedDict()
        self.color_map = {}
        self.sorted_values = []
        
        # Link support
        self.link_instances = get_link_instances()  # [(RevitLinkInstance, Document), ...]
        self.selected_link = None      # None = current model
        self.selected_link_doc = None  # linked document
        
        self.Title = "Color Splasher - By DQT"
        self.Width = 1000
        self.Height = 780
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(CLR_BG)
        
        self._build_ui()
        self._fill_sources()
        self._load_data_for_source()
        
        # Events
        self.cbSource.SelectionChanged += self._ev_source
        self.cbCat.SelectionChanged += self._ev_cat
        self.lbParams.SelectionChanged += self._ev_param
        self.txtSearch.TextChanged += self._ev_search
        self.btnApply.Click += self._ev_apply
        self.btnReset.Click += self._ev_reset
        self.btnGradient.Click += self._ev_gradient
        self.btnRandom.Click += self._ev_random
        self.btnLegend.Click += self._ev_legend
        self.btnFilters.Click += self._ev_filters
        self.btnClose.Click += self._ev_close
        
        if self.categories:
            self._ev_cat(None, None)
    
    def _make_row(self, grid, height):
        rd = RowDefinition(); rd.Height = height; grid.RowDefinitions.Add(rd)
    
    def _make_col(self, grid, width):
        cd = ColumnDefinition(); cd.Width = width; grid.ColumnDefinitions.Add(cd)
    
    def _build_ui(self):
        main = WPFGrid()
        self._make_row(main, GridLength.Auto)
        self._make_row(main, GridLength(1, GridUnitType.Star))
        self._make_row(main, GridLength.Auto)
        
        # === HEADER ===
        hdr = Border(); hdr.Background = SolidColorBrush(CLR_HEADER)
        hdr.Padding = Thickness(14, 10, 14, 10)
        hdr.CornerRadius = System.Windows.CornerRadius(5)
        hdr.Margin = Thickness(12, 12, 12, 0)
        WPFGrid.SetRow(hdr, 0)
        hg = WPFGrid()
        self._make_col(hg, GridLength(1, GridUnitType.Star))
        self._make_col(hg, GridLength.Auto)
        
        hs = StackPanel()
        t1 = TextBlock(); t1.Text = "COLOR SPLASHER"; t1.FontSize = 20
        t1.FontWeight = System.Windows.FontWeights.Bold; t1.Foreground = SolidColorBrush(CLR_HEADER_TEXT)
        hs.Children.Add(t1)
        t2 = TextBlock(); t2.Text = "Auto-color elements by parameter value"
        t2.FontSize = 11; t2.Foreground = SolidColorBrush(CLR_HEADER_SUB)
        t2.Margin = Thickness(0, 3, 0, 0); hs.Children.Add(t2)
        WPFGrid.SetColumn(hs, 0); hg.Children.Add(hs)
        
        # DQT badge (right side) like IFC-SG Checker
        badge = StackPanel(); badge.VerticalAlignment = VerticalAlignment.Center
        badge.HorizontalAlignment = HorizontalAlignment.Right
        b1 = TextBlock(); b1.Text = "DQT"; b1.FontSize = 14
        b1.FontWeight = System.Windows.FontWeights.Bold; b1.Foreground = SolidColorBrush(CLR_ACCENT)
        b1.HorizontalAlignment = HorizontalAlignment.Right
        badge.Children.Add(b1)
        b2 = TextBlock(); b2.Text = "v4.0"; b2.FontSize = 9
        b2.Foreground = SolidColorBrush(CLR_MUTED)
        b2.HorizontalAlignment = HorizontalAlignment.Right
        badge.Children.Add(b2)
        WPFGrid.SetColumn(badge, 1); hg.Children.Add(badge)
        hdr.Child = hg; main.Children.Add(hdr)
        
        # === CONTENT ===
        ct = WPFGrid(); ct.Margin = Thickness(15); WPFGrid.SetRow(ct, 1)
        self._make_col(ct, GridLength(260))
        self._make_col(ct, GridLength(10))
        self._make_col(ct, GridLength(1, GridUnitType.Star))
        
        # LEFT
        lb = Border(); lb.Background = SolidColorBrush(CLR_CARD)
        lb.BorderBrush = SolidColorBrush(CLR_BORDER); lb.BorderThickness = Thickness(1)
        lb.CornerRadius = System.Windows.CornerRadius(4); lb.Padding = Thickness(12)
        WPFGrid.SetColumn(lb, 0)
        
        lg = WPFGrid()
        for h in [GridLength.Auto, GridLength.Auto, GridLength.Auto, GridLength.Auto,
                   GridLength.Auto, GridLength.Auto,
                   GridLength(1, GridUnitType.Star), GridLength.Auto, GridLength.Auto]:
            self._make_row(lg, h)
        
        # Source (Current Model / Links)
        ls = TextBlock(); ls.Text = "SOURCE"; ls.FontWeight = System.Windows.FontWeights.Bold
        ls.FontSize = 11; ls.Foreground = SolidColorBrush(CLR_ACCENT); ls.Margin = Thickness(0,0,0,5)
        WPFGrid.SetRow(ls, 0); lg.Children.Add(ls)
        
        self.cbSource = ComboBox(); self.cbSource.Height = 28; self.cbSource.Margin = Thickness(0,0,0,10)
        self.cbSource.FontSize = 11; WPFGrid.SetRow(self.cbSource, 1); lg.Children.Add(self.cbSource)
        
        # Category
        l1 = TextBlock(); l1.Text = "CATEGORY"; l1.FontWeight = System.Windows.FontWeights.Bold
        l1.FontSize = 11; l1.Foreground = SolidColorBrush(CLR_ACCENT); l1.Margin = Thickness(0,0,0,5)
        WPFGrid.SetRow(l1, 2); lg.Children.Add(l1)
        
        self.cbCat = ComboBox(); self.cbCat.Height = 28; self.cbCat.Margin = Thickness(0,0,0,10)
        self.cbCat.FontSize = 12; WPFGrid.SetRow(self.cbCat, 3); lg.Children.Add(self.cbCat)
        
        # Search
        l3 = TextBlock(); l3.Text = "SEARCH PARAMETER"; l3.FontWeight = System.Windows.FontWeights.Bold
        l3.FontSize = 11; l3.Foreground = SolidColorBrush(CLR_ACCENT); l3.Margin = Thickness(0,0,0,5)
        WPFGrid.SetRow(l3, 4); lg.Children.Add(l3)
        
        self.txtSearch = TextBox(); self.txtSearch.Height = 26; self.txtSearch.Margin = Thickness(0,0,0,8)
        self.txtSearch.FontSize = 11; WPFGrid.SetRow(self.txtSearch, 5); lg.Children.Add(self.txtSearch)
        
        # Parameter list
        self.lbParams = ListBox(); self.lbParams.Margin = Thickness(0,0,0,8); self.lbParams.FontSize = 11
        self.lbParams.BorderBrush = SolidColorBrush(CLR_BORDER); self.lbParams.BorderThickness = Thickness(1)
        WPFGrid.SetRow(self.lbParams, 6); lg.Children.Add(self.lbParams)
        
        self.txtStatus = TextBlock(); self.txtStatus.FontSize = 10
        self.txtStatus.Foreground = SolidColorBrush(CLR_MUTED); self.txtStatus.Text = "Ready"
        WPFGrid.SetRow(self.txtStatus, 7); lg.Children.Add(self.txtStatus)
        
        lb.Child = lg; ct.Children.Add(lb)
        
        # RIGHT
        rb = Border(); rb.Background = SolidColorBrush(CLR_CARD)
        rb.BorderBrush = SolidColorBrush(CLR_BORDER); rb.BorderThickness = Thickness(1)
        rb.CornerRadius = System.Windows.CornerRadius(4); rb.Padding = Thickness(12)
        WPFGrid.SetColumn(rb, 2)
        
        rg = WPFGrid()
        self._make_row(rg, GridLength.Auto)
        self._make_row(rg, GridLength.Auto)
        self._make_row(rg, GridLength(1, GridUnitType.Star))
        self._make_row(rg, GridLength.Auto)
        
        # Legend header
        lhg = WPFGrid()
        self._make_col(lhg, GridLength(1, GridUnitType.Star))
        self._make_col(lhg, GridLength.Auto)
        ll = TextBlock(); ll.Text = "VALUES"; ll.FontWeight = System.Windows.FontWeights.Bold
        ll.FontSize = 11; ll.Foreground = SolidColorBrush(CLR_ACCENT)
        WPFGrid.SetColumn(ll, 0); lhg.Children.Add(ll)
        self.txtVC = TextBlock(); self.txtVC.Text = "0 unique values"; self.txtVC.FontSize = 10
        self.txtVC.Foreground = SolidColorBrush(CLR_MUTED)
        WPFGrid.SetColumn(self.txtVC, 1); lhg.Children.Add(self.txtVC)
        lhg.Margin = Thickness(0,0,0,8); WPFGrid.SetRow(lhg, 0); rg.Children.Add(lhg)
        
        # Summary
        sb = Border(); sb.Background = SolidColorBrush(CLR_BG)
        sb.CornerRadius = System.Windows.CornerRadius(3); sb.Padding = Thickness(8,6,8,6)
        sb.Margin = Thickness(0,0,0,8)
        self.txtSum = TextBlock(); self.txtSum.Text = "Select a category and parameter"
        self.txtSum.FontSize = 11; self.txtSum.TextWrapping = System.Windows.TextWrapping.Wrap
        self.txtSum.Foreground = SolidColorBrush(CLR_TEXT); sb.Child = self.txtSum
        WPFGrid.SetRow(sb, 1); rg.Children.Add(sb)
        
        # Legend scroll with colored rows
        sv = ScrollViewer(); sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        self.spLeg = StackPanel(); sv.Content = self.spLeg
        WPFGrid.SetRow(sv, 2); rg.Children.Add(sv)
        
        # Color mode buttons
        btn_panel = WPFGrid(); btn_panel.Margin = Thickness(0, 8, 0, 0)
        self._make_col(btn_panel, GridLength(1, GridUnitType.Star))
        self._make_col(btn_panel, GridLength(6))
        self._make_col(btn_panel, GridLength(1, GridUnitType.Star))
        self._make_col(btn_panel, GridLength(6))
        self._make_col(btn_panel, GridLength(1, GridUnitType.Star))
        self._make_col(btn_panel, GridLength(6))
        self._make_col(btn_panel, GridLength(1, GridUnitType.Star))
        
        self.btnGradient = self._make_btn("Gradient Colors", CLR_CARD, CLR_TEXT)
        WPFGrid.SetColumn(self.btnGradient, 0); btn_panel.Children.Add(self.btnGradient)
        
        self.btnRandom = self._make_btn("Random Colors", CLR_CARD, CLR_TEXT)
        WPFGrid.SetColumn(self.btnRandom, 2); btn_panel.Children.Add(self.btnRandom)
        
        self.btnLegend = self._make_btn("Create Legend", CLR_CARD, CLR_TEXT)
        WPFGrid.SetColumn(self.btnLegend, 4); btn_panel.Children.Add(self.btnLegend)
        
        self.btnFilters = self._make_btn("View Filters", CLR_CARD, CLR_TEXT)
        WPFGrid.SetColumn(self.btnFilters, 6); btn_panel.Children.Add(self.btnFilters)
        
        WPFGrid.SetRow(btn_panel, 3); rg.Children.Add(btn_panel)
        
        rb.Child = rg; ct.Children.Add(rb); main.Children.Add(ct)
        
        # === FOOTER (buttons row) ===
        btn_row = Border(); btn_row.Background = SolidColorBrush(CLR_BG)
        btn_row.Padding = Thickness(12, 8, 12, 8); WPFGrid.SetRow(btn_row, 2)
        brg = WPFGrid()
        self._make_col(brg, GridLength(1, GridUnitType.Star))
        self._make_col(brg, GridLength.Auto)
        self._make_col(brg, GridLength.Auto)
        self._make_col(brg, GridLength.Auto)
        
        # Status text on left
        self.txtFooterStatus = TextBlock()
        try:
            vn = doc.ActiveView.Name
            if len(vn) > 50: vn = vn[:47] + "..."
            self.txtFooterStatus.Text = "View: " + vn
        except:
            self.txtFooterStatus.Text = ""
        self.txtFooterStatus.FontSize = 11; self.txtFooterStatus.Foreground = SolidColorBrush(CLR_MUTED)
        self.txtFooterStatus.VerticalAlignment = VerticalAlignment.Center
        WPFGrid.SetColumn(self.txtFooterStatus, 0); brg.Children.Add(self.txtFooterStatus)
        
        self.btnReset = self._make_btn("Reset", CLR_CARD, CLR_TEXT)
        self.btnReset.Width = 100; self.btnReset.Margin = Thickness(0,0,8,0)
        WPFGrid.SetColumn(self.btnReset, 1); brg.Children.Add(self.btnReset)
        
        # Apply button - green success style like IFC-SG Checker
        self.btnApply = Button(); self.btnApply.Content = "Apply Colors"
        self.btnApply.Width = 130; self.btnApply.Height = 34; self.btnApply.FontSize = 13
        self.btnApply.FontWeight = System.Windows.FontWeights.Bold
        self.btnApply.Background = SolidColorBrush(Color.FromRgb(200, 230, 201))
        self.btnApply.Foreground = SolidColorBrush(Color.FromRgb(46, 125, 50))
        self.btnApply.BorderBrush = SolidColorBrush(Color.FromRgb(129, 199, 132))
        self.btnApply.BorderThickness = Thickness(1)
        self.btnApply.Padding = Thickness(14, 0, 14, 0)
        self.btnApply.Margin = Thickness(0,0,8,0)
        WPFGrid.SetColumn(self.btnApply, 2); brg.Children.Add(self.btnApply)
        
        self.btnClose = self._make_btn("Close", CLR_CARD, CLR_TEXT)
        self.btnClose.Width = 80
        WPFGrid.SetColumn(self.btnClose, 3); brg.Children.Add(self.btnClose)
        
        btn_row.Child = brg; main.Children.Add(btn_row)
        
        # === BOTTOM FOOTER BAR ===
        self._make_row(main, GridLength.Auto)
        fb = Border(); fb.Background = SolidColorBrush(CLR_FOOTER)
        fb.CornerRadius = System.Windows.CornerRadius(0, 0, 3, 3)
        fb.Padding = Thickness(8, 4, 8, 4); WPFGrid.SetRow(fb, 3)
        fbg = WPFGrid()
        self._make_col(fbg, GridLength(1, GridUnitType.Star))
        self._make_col(fbg, GridLength.Auto)
        
        fbl = TextBlock(); fbl.Text = "Color Splasher v4.0 | Dang Quoc Truong (DQT)"
        fbl.FontSize = 9; fbl.Foreground = SolidColorBrush(CLR_MUTED)
        WPFGrid.SetColumn(fbl, 0); fbg.Children.Add(fbl)
        
        import time
        fbr = TextBlock()
        try:
            fbr.Text = "{} | {}".format(doc.Title, time.strftime("%Y-%m-%d"))
        except:
            fbr.Text = time.strftime("%Y-%m-%d")
        fbr.FontSize = 9; fbr.Foreground = SolidColorBrush(CLR_MUTED)
        WPFGrid.SetColumn(fbr, 1); fbg.Children.Add(fbr)
        
        fb.Child = fbg; main.Children.Add(fb)
        
        self.Content = main
    
    def _make_btn(self, text, bg, fg):
        b = Button(); b.Content = text; b.Height = 30; b.FontSize = 11
        b.FontWeight = System.Windows.FontWeights.SemiBold
        b.Background = SolidColorBrush(bg); b.Foreground = SolidColorBrush(fg)
        b.BorderBrush = SolidColorBrush(CLR_BORDER); b.Padding = Thickness(10, 0, 10, 0)
        return b
    
    # ======== DATA ========
    def _fill_sources(self):
        self.cbSource.Items.Clear()
        self.cbSource.Items.Add("Current Model")
        for link_inst, link_doc in self.link_instances:
            try:
                name = link_doc.Title
            except:
                name = "Link"
            self.cbSource.Items.Add("Link: " + name)
        self.cbSource.SelectedIndex = 0
    
    def _load_data_for_source(self):
        """Load categories based on selected source"""
        idx = self.cbSource.SelectedIndex if self.cbSource.SelectedIndex >= 0 else 0
        if idx == 0:
            self.selected_link = None
            self.selected_link_doc = None
            self.categories = collect_categories()
        else:
            link_idx = idx - 1
            if link_idx < len(self.link_instances):
                self.selected_link, self.selected_link_doc = self.link_instances[link_idx]
                self.categories = collect_categories_from_doc(self.selected_link_doc)
            else:
                self.selected_link = None
                self.selected_link_doc = None
                self.categories = []
        self._fill_cats()
        if self.categories:
            self._ev_cat(None, None)
    
    def _get_source_doc(self):
        """Get the active source document (current or linked)"""
        return self.selected_link_doc if self.selected_link_doc else doc
    
    def _collect_elems(self, category):
        """Collect elements from current source"""
        if self.selected_link_doc:
            return collect_elements_from_doc(self.selected_link_doc, category.Name)
        else:
            return collect_elements(category)
    
    def _get_val(self, elem, pname):
        """Get param value from current source"""
        if self.selected_link_doc:
            return get_param_value_linked(elem, pname, self.selected_link_doc)
        else:
            return get_param_value(elem, pname)
    
    def _collect_params(self, category):
        """Collect param names from current source"""
        src_doc = self._get_source_doc()
        elems = self._collect_elems(category)
        if not elems:
            return SPECIAL_PARAMS[:]
        names = set()
        for elem in elems[:20]:
            try:
                for p in elem.Parameters:
                    try:
                        if p and p.Definition and p.Definition.Name:
                            names.add(p.Definition.Name)
                    except:
                        continue
            except:
                continue
            try:
                tid = elem.GetTypeId()
                if tid:
                    etype = src_doc.GetElement(tid)
                    if etype:
                        for p in etype.Parameters:
                            try:
                                if p and p.Definition and p.Definition.Name:
                                    names.add(p.Definition.Name)
                            except:
                                continue
            except:
                continue
        result = list(SPECIAL_PARAMS)
        for n in sorted(names):
            if n not in result:
                result.append(n)
        return result
    
    def _fill_cats(self):
        self.cbCat.Items.Clear()
        for name, _ in self.categories:
            self.cbCat.Items.Add(name)
        if self.cbCat.Items.Count > 0:
            self.cbCat.SelectedIndex = 0
        self.txtStatus.Text = "{} categories".format(len(self.categories))
    
    def _get_cat(self):
        i = self.cbCat.SelectedIndex
        if i < 0 or i >= len(self.categories): return None
        return self.categories[i][1]
    
    def _filter_params(self):
        self.lbParams.Items.Clear()
        s = (self.txtSearch.Text or "").lower().strip()
        for n in self.current_param_names:
            if not s or s in n.lower():
                self.lbParams.Items.Add(n)
    
    def _assign_palette_colors(self):
        """Assign colors from default palette"""
        self.color_map = {}
        for i, v in enumerate(self.sorted_values):
            self.color_map[v] = PALETTE[i % len(PALETTE)]
    
    def _rebuild_legend(self):
        """Rebuild legend UI from current color_map"""
        self.spLeg.Children.Clear()
        total = sum(len(self.elements_by_value[v]) for v in self.sorted_values)
        
        for val in self.sorted_values:
            cnt = len(self.elements_by_value[val])
            r, g, b = self.color_map.get(val, (128, 128, 128))
            
            # Full-width colored row (like pyRevit color splasher)
            row = Border(); row.Margin = Thickness(0, 0, 0, 1)
            row.Padding = Thickness(10, 6, 10, 6)
            row.Background = SolidColorBrush(Color.FromRgb(r, g, b))
            
            gr = WPFGrid()
            self._make_col(gr, GridLength(1, GridUnitType.Star))
            self._make_col(gr, GridLength.Auto)
            
            # Determine text color (white or black based on brightness)
            brightness = (r * 299 + g * 587 + b * 114) / 1000
            txt_color = Color.FromRgb(255, 255, 255) if brightness < 140 else Color.FromRgb(0, 0, 0)
            
            tv = TextBlock(); tv.Text = str(val); tv.FontSize = 12
            tv.FontWeight = System.Windows.FontWeights.SemiBold
            tv.VerticalAlignment = VerticalAlignment.Center
            tv.TextTrimming = TextTrimming.CharacterEllipsis
            tv.Foreground = SolidColorBrush(txt_color)
            WPFGrid.SetColumn(tv, 0); gr.Children.Add(tv)
            
            tc = TextBlock(); tc.Text = "({})".format(cnt); tc.FontSize = 10
            tc.Foreground = SolidColorBrush(txt_color); tc.Opacity = 0.8
            tc.VerticalAlignment = VerticalAlignment.Center; tc.Margin = Thickness(8, 0, 0, 0)
            WPFGrid.SetColumn(tc, 1); gr.Children.Add(tc)
            
            row.Child = gr; self.spLeg.Children.Add(row)
        
        u = len(self.sorted_values)
        self.txtVC.Text = "{} unique values".format(u)
        self.txtStatus.Text = "{} elements | {} values".format(total, u)
    
    def _analyze_param(self, pname):
        """Analyze parameter and build data"""
        cat = self._get_cat()
        if not cat: return
        
        elems = self._collect_elems(cat)
        self.elements_by_value = OrderedDict()
        for e in elems:
            v = self._get_val(e, pname)
            if v not in self.elements_by_value:
                self.elements_by_value[v] = []
            self.elements_by_value[v].append(e)
        
        self.sorted_values = sorted(self.elements_by_value.keys())
        self._assign_palette_colors()
        self._rebuild_legend()
        
        total = sum(len(x) for x in self.elements_by_value.values())
        src = "Link" if self.selected_link else "Model"
        self.txtSum.Text = "'{}' | {} values across {} elements [{}]".format(
            pname, len(self.sorted_values), total, src)
    
    # ======== EVENTS ========
    def _ev_source(self, s, a):
        """Handle source change (Current Model / Link)"""
        self._load_data_for_source()
    
    def _ev_cat(self, s, a):
        cat = self._get_cat()
        if not cat: return
        elems = self._collect_elems(cat)
        self.txtStatus.Text = "{} elements".format(len(elems))
        self.current_param_names = self._collect_params(cat)
        self._filter_params()
        self.spLeg.Children.Clear()
        self.elements_by_value = OrderedDict()
        self.color_map = {}
        self.sorted_values = []
        self.txtVC.Text = "0 unique values"
        self.txtSum.Text = "Select a parameter to preview colors"
    
    def _ev_param(self, s, a):
        sel = self.lbParams.SelectedItem
        if sel: self._analyze_param(str(sel))
    
    def _ev_search(self, s, a): self._filter_params()
    
    def _ev_gradient(self, s, a):
        if not self.sorted_values: return
        colors = generate_gradient(len(self.sorted_values))
        for i, v in enumerate(self.sorted_values):
            self.color_map[v] = colors[i]
        self._rebuild_legend()
    
    def _ev_random(self, s, a):
        if not self.sorted_values: return
        colors = generate_random(len(self.sorted_values))
        for i, v in enumerate(self.sorted_values):
            self.color_map[v] = colors[i]
        self._rebuild_legend()
    
    def _ev_legend(self, s, a):
        """Create Legend view with TextNotes + FilledRegions (pyRevit approach)"""
        if not self.sorted_values or not self.color_map:
            forms.alert("No data to create legend.", title="Color Splasher")
            return
        
        sel = self.lbParams.SelectedItem
        pname = str(sel) if sel else "Parameter"
        cat_name = str(self.cbCat.SelectedItem) if self.cbCat.SelectedItem else "Category"
        
        # Find existing legend to duplicate (prefer emptiest one)
        existing_legend = None
        min_elements = 999999
        for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
            try:
                if v.ViewType == DB.ViewType.Legend and not v.IsTemplate:
                    try:
                        elem_count = DB.FilteredElementCollector(doc, v.Id).GetElementCount()
                    except:
                        elem_count = 0
                    if existing_legend is None or elem_count < min_elements:
                        existing_legend = v
                        min_elements = elem_count
            except:
                continue
        
        if not existing_legend:
            forms.alert("No Legend view found in project.\nPlease create one manually first.",
                       title="Color Splasher")
            return
        
        t = DB.Transaction(doc, "DQT - Create Color Legend")
        t.Start()
        try:
            # 1) Duplicate legend
            new_id = existing_legend.Duplicate(DB.ViewDuplicateOption.Duplicate)
            new_legend = doc.GetElement(new_id)
            
            # 2) Rename (with fallback for duplicates)
            base_name = "Legend - {} - {}".format(cat_name, pname)
            renamed = False
            try:
                new_legend.Name = base_name
                renamed = True
            except:
                pass
            if not renamed:
                for i in range(1, 100):
                    try:
                        new_legend.Name = "{} - {}".format(base_name, i)
                        break
                    except:
                        continue
            
            # 3) Get TextNoteType
            txt_type_id = None
            # Try from existing legend elements first
            try:
                for ele in DB.FilteredElementCollector(doc, existing_legend.Id).ToElements():
                    if isinstance(ele, DB.TextNote):
                        txt_type_id = ele.GetTypeId()
                        break
            except:
                pass
            if not txt_type_id:
                for tnt in DB.FilteredElementCollector(doc).OfClass(DB.TextNoteType):
                    txt_type_id = tnt.Id
                    break
            
            if not txt_type_id:
                t.RollBack()
                forms.alert("No TextNote type found.", title="Color Splasher")
                return
            
            # 4) Get FilledRegionType with solid fill
            filled_type = None
            for frt in DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType):
                try:
                    pat = doc.GetElement(frt.ForegroundPatternId)
                    if pat and pat.GetFillPattern().IsSolidFill:
                        filled_type = frt
                        break
                except:
                    continue
            
            # Fallback: duplicate first FilledRegionType
            if not filled_type:
                all_frt = list(DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType))
                if all_frt:
                    for idx in range(100):
                        try:
                            filled_type = all_frt[0].Duplicate("DQT Color Swatch {}".format(idx))
                            break
                        except:
                            continue
                    if filled_type:
                        solid = get_solid_fill()
                        if solid:
                            filled_type.ForegroundPatternId = solid.Id
            
            if not filled_type:
                t.RollBack()
                forms.alert("No FilledRegion type available.", title="Color Splasher")
                return
            
            # 5) Create TextNotes and measure positions
            from pyrevit.framework import List as FWList
            
            y_pos = 0.0
            text_data = []  # (y_min, height, r, g, b)
            max_x_list = []
            
            for val in self.sorted_values:
                r, g, b = self.color_map.get(val, (128, 128, 128))
                cnt = len(self.elements_by_value[val])
                text_line = "{} / {} - {} ({})".format(cat_name, pname, val, cnt)
                
                pt = DB.XYZ(0, y_pos, 0)
                try:
                    tn = DB.TextNote.Create(doc, new_legend.Id, pt, text_line, txt_type_id)
                    doc.Regenerate()
                    
                    bbox = tn.get_BoundingBox(new_legend)
                    if bbox:
                        height = bbox.Max.Y - bbox.Min.Y
                        spacing = height * 0.25
                        max_x_list.append(bbox.Max.X)
                        text_data.append((bbox.Min.Y, height, r, g, b))
                        y_pos = bbox.Min.Y - (height + spacing)
                    else:
                        text_data.append((y_pos, 0.01, r, g, b))
                        y_pos -= 0.02
                except:
                    text_data.append((y_pos, 0.01, r, g, b))
                    y_pos -= 0.02
            
            # 7) Create FilledRegion color swatches
            ini_x = (max(max_x_list) if max_x_list else 0.3) + 0.005
            solid_fill = get_solid_fill()
            
            for td in text_data:
                y_min, height, r, g, b = td
                if height < 0.001:
                    height = 0.01
                rect_w = height * 2
                
                try:
                    p0 = DB.XYZ(ini_x, y_min, 0)
                    p1 = DB.XYZ(ini_x, y_min + height, 0)
                    p2 = DB.XYZ(ini_x + rect_w, y_min + height, 0)
                    p3 = DB.XYZ(ini_x + rect_w, y_min, 0)
                    
                    loop = DB.CurveLoop()
                    loop.Append(DB.Line.CreateBound(p0, p1))
                    loop.Append(DB.Line.CreateBound(p1, p2))
                    loop.Append(DB.Line.CreateBound(p2, p3))
                    loop.Append(DB.Line.CreateBound(p3, p0))
                    
                    loops = FWList[DB.CurveLoop]()
                    loops.Add(loop)
                    
                    region = DB.FilledRegion.Create(doc, filled_type.Id, new_legend.Id, loops)
                    
                    # Override color
                    color = DB.Color(r, g, b)
                    ogs = DB.OverrideGraphicSettings()
                    ogs.SetSurfaceForegroundPatternColor(color)
                    ogs.SetCutForegroundPatternColor(color)
                    if solid_fill:
                        ogs.SetSurfaceForegroundPatternId(solid_fill.Id)
                        ogs.SetCutForegroundPatternId(solid_fill.Id)
                    new_legend.SetElementOverrides(region.Id, ogs)
                except:
                    continue
            
            t.Commit()
            
            # Open legend view
            try:
                uidoc.ActiveView = new_legend
            except:
                try:
                    uidoc.RequestViewChange(new_legend)
                except:
                    pass
            
            forms.alert("Legend created:\n'{}'".format(new_legend.Name),
                       title="Color Splasher - DQT")
            
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            forms.alert("Error creating legend:\n{}".format(str(ex)),
                       title="Color Splasher")
    
    def _ev_filters(self, s, a):
        """Create View Filters for each value"""
        if not self.sorted_values or not self.color_map:
            forms.alert("No data to create filters.", title="Color Splasher")
            return
        
        sel = self.lbParams.SelectedItem
        pname = str(sel) if sel else ""
        cat = self._get_cat()
        if not cat or not pname:
            forms.alert("Select a category and parameter first.", title="Color Splasher")
            return
        
        # Check if parameter is a special (computed) param - can't create filter for those
        if pname in SPECIAL_PARAMS:
            forms.alert("Cannot create View Filters for computed parameter '{}'.\n"
                       "Please select a regular Revit parameter.".format(pname),
                       title="Color Splasher")
            return
        
        active_view = doc.ActiveView
        t = DB.Transaction(doc, "DQT - Create View Filters")
        t.Start()
        created = 0
        
        try:
            # Get category ids for filter
            cat_id = cat.Id
            cat_id_list = System.Collections.Generic.List[DB.ElementId]()
            cat_id_list.Add(cat_id)
            
            solid = get_solid_fill()
            
            for val in self.sorted_values:
                r, g, b = self.color_map.get(val, (128, 128, 128))
                
                # Create unique filter name
                filter_name = "{} - {}".format(pname, val)
                # Truncate if too long
                if len(filter_name) > 200:
                    filter_name = filter_name[:200]
                
                # Check if filter already exists
                existing = None
                try:
                    for pfe in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
                        try:
                            if pfe.Name == filter_name:
                                existing = pfe
                                break
                        except:
                            continue
                except:
                    pass
                
                pfilter = existing
                
                if not pfilter:
                    try:
                        # Find the ParameterId for this parameter name
                        # Sample an element to get the parameter
                        elems = self.elements_by_value.get(val, [])
                        if not elems:
                            continue
                        
                        sample = elems[0]
                        param = sample.LookupParameter(pname)
                        if param is None:
                            # Try type parameter
                            tid = sample.GetTypeId()
                            if tid:
                                etype = doc.GetElement(tid)
                                if etype:
                                    param = etype.LookupParameter(pname)
                        
                        if param is None:
                            continue
                        
                        param_id = param.Id
                        
                        # Create filter rule based on storage type
                        rule = None
                        st = param.StorageType
                        if st == DB.StorageType.String:
                            str_val = param.AsString() or ""
                            try:
                                rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, str_val, True)
                            except:
                                rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, str_val)
                        elif st == DB.StorageType.Integer:
                            rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, param.AsInteger())
                        elif st == DB.StorageType.Double:
                            rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, param.AsDouble(), 0.001)
                        elif st == DB.StorageType.ElementId:
                            rule = DB.ParameterFilterRuleFactory.CreateEqualsRule(param_id, param.AsElementId())
                        
                        if rule is None:
                            continue
                        
                        # Create ElementParameterFilter
                        elem_filter = DB.ElementParameterFilter(rule)
                        
                        pfilter = DB.ParameterFilterElement.Create(doc, filter_name, cat_id_list, elem_filter)
                    except Exception as ex:
                        output.print_md("Filter create error for '{}': {}".format(val, ex))
                        continue
                
                if pfilter:
                    try:
                        # Add filter to view and set override
                        active_view.AddFilter(pfilter.Id)
                        active_view.SetFilterVisibility(pfilter.Id, True)
                        
                        color = DB.Color(r, g, b)
                        ogs = DB.OverrideGraphicSettings()
                        try:
                            ogs.SetSurfaceForegroundPatternColor(color)
                            ogs.SetSurfaceForegroundPatternVisible(True)
                            if solid:
                                ogs.SetSurfaceForegroundPatternId(solid.Id)
                        except:
                            try:
                                ogs.SetProjectionFillColor(color)
                                ogs.SetProjectionFillPatternVisible(True)
                                if solid:
                                    ogs.SetProjectionFillPatternId(solid.Id)
                            except:
                                pass
                        try:
                            ogs.SetCutForegroundPatternColor(color)
                            ogs.SetCutForegroundPatternVisible(True)
                            if solid:
                                ogs.SetCutForegroundPatternId(solid.Id)
                        except:
                            pass
                        
                        active_view.SetFilterOverrides(pfilter.Id, ogs)
                        created += 1
                    except Exception as ex:
                        output.print_md("Filter apply error for '{}': {}".format(val, ex))
                        continue
            
            t.Commit()
            forms.alert("Created {} view filters in active view.".format(created),
                       title="Color Splasher - DQT")
        except Exception as ex:
            if t.HasStarted():
                t.RollBack()
            forms.alert("Error creating filters:\n{}".format(str(ex)),
                       title="Color Splasher")
    
    def _ev_apply(self, s, a):
        if not self.elements_by_value:
            forms.alert("Select a category and parameter first.", title="Color Splasher")
            return
        if self.selected_link:
            count = apply_overrides_linked(self.elements_by_value, self.color_map, self.selected_link)
            src = "linked"
            note = "\n(Applied via View Filters - works across links)"
        else:
            count = apply_overrides(self.elements_by_value, self.color_map)
            src = "model"
            note = ""
        sel = self.lbParams.SelectedItem
        pn = str(sel) if sel else ""
        msg = "Applied overrides to {} {} elements".format(count, src)
        if pn: msg += " by '{}'".format(pn)
        forms.alert(msg + "." + note, title="Color Splasher - DQT")
    
    def _ev_reset(self, s, a):
        if forms.alert("Clear ALL overrides in active view?",
                       title="Color Splasher", yes=True, no=True):
            count = clear_overrides()
            forms.alert("Cleared {} elements.".format(count), title="Color Splasher - DQT")
    
    def _ev_close(self, s, a): self.Close()


# =====================================================
# MAIN
# =====================================================
if __name__ == '__main__':
    active_view = doc.ActiveView
    vt = active_view.ViewType
    skip = [DB.ViewType.Schedule, DB.ViewType.DrawingSheet, DB.ViewType.Legend, DB.ViewType.Rendering]

    if vt in skip:
        forms.alert("Not supported in this view type.", title="Color Splasher")
    else:
        try:
            win = ColorSplasherWindow()
            win.ShowDialog()
        except Exception as ex:
            import traceback
            output.print_md("**ERROR:** {}".format(ex))
            output.print_md("```\n{}\n```".format(traceback.format_exc()))
            forms.alert("Error:\n{}".format(str(ex)), title="Color Splasher Error")