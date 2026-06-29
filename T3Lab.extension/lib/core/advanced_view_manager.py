# -*- coding: utf-8 -*-
"""
Advanced View Manager Core Execution Logic
Handles Revit API operations, element wrappers, and Excel import/export.
"""

import re
import os
import zipfile
import clr
clr.AddReference('System')
from System import Uri, Int64
from pyrevit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewPlan, ViewDrafting, View3D, ViewSection,
    Level, ViewFamilyType, ViewFamily, AreaScheme, ViewType, ElementId,
    BuiltInParameter, Viewport, ViewDetailLevel, StorageType, Transaction,
    BoundingBoxXYZ, Transform, XYZ
)

# =====================================================
# REVIT VERSION COMPATIBILITY (2024-2027)
# =====================================================

def _eid_int(eid):
    """Get integer value from ElementId - compatible with Revit 2024-2027.
    Revit 2024-2025: ElementId.IntegerValue (int)
    Revit 2026+: ElementId.Value (long, IntegerValue removed)
    """
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue

def _make_eid(int_val):
    """Create ElementId from integer - compatible with Revit 2024-2027.
    Revit 2024-2025: ElementId(int)
    Revit 2026+: ElementId(long)
    """
    try:
        return ElementId(int(int_val))
    except:
        try:
            return ElementId(Int64(int(int_val)))
        except:
            return ElementId(int_val)


# =====================================================
# ENHANCED VIEW ITEM
# =====================================================

class EnhancedViewItem(object):
    """Enhanced view item with all properties"""
    
    def __init__(self, view, doc):
        self.element = view
        self.doc = doc
        self.id = view.Id
        self.name = view.Name
        self.is_selected = False
        self.view_type = self._get_view_type_name(view)
        self.view_template = self._get_view_template(view)
        self.scale = self._get_scale(view)
        self.detail_level = self._get_detail_level(view)
        self.on_sheets = self._get_sheet_count(view)
        self.title_on_sheet = self._get_title_on_sheet(view)
        self.referencing_sheet = self._get_referencing_sheet(view)
        self.sheet_number = self._get_sheet_number(view)
        self.sheet_name = self._get_sheet_name(view)
        self.level_name = self._get_level_name(view)
        
        # Crop Box data
        crop_data = self._get_crop_data(view)
        self.crop_active = crop_data[0]
        self.crop_visible = crop_data[1]
        self.crop_min = crop_data[2]   # "x,y,z" string
        self.crop_max = crop_data[3]   # "x,y,z" string
    
    def _get_crop_data(self, view):
        """Get crop box data: (active, visible, min_str, max_str)"""
        try:
            crop_active = "Yes" if view.CropBoxActive else "No"
        except:
            crop_active = "No"
        
        try:
            crop_visible = "Yes" if view.CropBoxVisible else "No"
        except:
            crop_visible = "No"
        
        crop_min = ""
        crop_max = ""
        try:
            bb = view.CropBox
            if bb is not None:
                mn = bb.Min
                mx = bb.Max
                # Round to 6 decimal places (Revit internal units = feet)
                crop_min = "{},{},{}".format(
                    round(mn.X, 6), round(mn.Y, 6), round(mn.Z, 6))
                crop_max = "{},{},{}".format(
                    round(mx.X, 6), round(mx.Y, 6), round(mx.Z, 6))
        except:
            pass
        
        return (crop_active, crop_visible, crop_min, crop_max)
    
    def _get_level_name(self, view):
        """Get the associated level name for plan views"""
        try:
            if hasattr(view, 'GenLevel') and view.GenLevel is not None:
                return view.GenLevel.Name
        except:
            pass
        try:
            level_param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL)
            if level_param and level_param.AsString():
                return level_param.AsString()
        except:
            pass
        return ""
        
    def _get_view_type_name(self, view):
        view_type_dict = {
            ViewType.FloorPlan: "Floor Plan",
            ViewType.CeilingPlan: "Ceiling Plan",
            ViewType.Elevation: "Elevation",
            ViewType.Section: "Section",
            ViewType.ThreeD: "3D View",
            ViewType.DraftingView: "Drafting View",
            ViewType.EngineeringPlan: "Structural Plan",
            ViewType.AreaPlan: "Area Plan",
            ViewType.Detail: "Detail View",
            ViewType.Legend: "Legend",
            ViewType.Schedule: "Schedule",
            ViewType.DrawingSheet: "Sheet"
        }
        return view_type_dict.get(view.ViewType, str(view.ViewType))
    
    def _get_view_template(self, view):
        try:
            template_id = view.ViewTemplateId
            if template_id and template_id != ElementId.InvalidElementId:
                template = self.doc.GetElement(template_id)
                return template.Name if template else "None"
            return "None"
        except:
            return "None"
    
    def _get_scale(self, view):
        try:
            return view.Scale if hasattr(view, 'Scale') else 0
        except:
            return 0
    
    def _get_detail_level(self, view):
        try:
            detail_dict = {
                ViewDetailLevel.Coarse: "Coarse",
                ViewDetailLevel.Medium: "Medium",
                ViewDetailLevel.Fine: "Fine"
            }
            return detail_dict.get(view.DetailLevel, "N/A")
        except:
            return "N/A"
    
    def _get_sheet_count(self, view):
        try:
            count = 0
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    count += 1
            return count
        except:
            return 0
    
    def _get_title_on_sheet(self, view):
        try:
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    title_param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                    if title_param:
                        return title_param.AsString() or "N/A"
                    return "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _get_referencing_sheet(self, view):
        try:
            param = view.get_Parameter(BuiltInParameter.VIEW_REFERENCING_SHEET)
            if param:
                return param.AsString() or "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _get_sheet_number(self, view):
        try:
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    sheet = self.doc.GetElement(vp.SheetId)
                    if sheet:
                        return sheet.SheetNumber or "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _get_sheet_name(self, view):
        try:
            collector = FilteredElementCollector(self.doc)\
                .OfClass(Viewport)\
                .WhereElementIsNotElementType()
            
            for vp in collector:
                if vp.ViewId == view.Id:
                    sheet = self.doc.GetElement(vp.SheetId)
                    if sheet:
                        return sheet.Name or "N/A"
            return "N/A"
        except:
            return "N/A"


# =====================================================
# REVIT DB MODIFICATIONS
# =====================================================

def update_view_name(doc, item, new_name):
    """Update view name"""
    if not new_name or new_name.strip() == "":
        raise ValueError("View name cannot be empty")
    
    t = Transaction(doc, "Rename View")
    t.Start()
    try:
        view = item.element
        view.Name = new_name
        item.name = new_name
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        raise e

def update_view_template(doc, item, template_name):
    """Update template"""
    t = Transaction(doc, "Update Template")
    t.Start()
    try:
        view = item.element
        
        if template_name == "None":
            view.ViewTemplateId = ElementId.InvalidElementId
        else:
            collector = FilteredElementCollector(doc)\
                .OfClass(View)\
                .WhereElementIsNotElementType()
            
            for template in collector:
                if template.IsTemplate and template.Name == template_name:
                    view.ViewTemplateId = template.Id
                    break
        
        item.view_template = template_name
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        raise e

def update_scale(doc, item, scale_str):
    """Update scale"""
    try:
        scale_value = int(scale_str)
        if scale_value <= 0:
            raise ValueError("Scale must be positive integer")
    except:
        raise ValueError("Scale must be positive integer")

    t = Transaction(doc, "Update Scale")
    t.Start()
    try:
        view = item.element
        view.Scale = scale_value
        item.scale = scale_value
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        raise e

def update_detail_level(doc, item, detail_str):
    """Update detail level"""
    detail_map = {
        "Coarse": ViewDetailLevel.Coarse,
        "Medium": ViewDetailLevel.Medium,
        "Fine": ViewDetailLevel.Fine
    }
    
    if detail_str not in detail_map:
        raise ValueError("Invalid Detail Level")
        
    t = Transaction(doc, "Update Detail")
    t.Start()
    try:
        view = item.element
        view.DetailLevel = detail_map[detail_str]
        item.detail_level = detail_str
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        raise e

def update_title_on_sheet(doc, item, title_str):
    """Update title on sheet"""
    t = Transaction(doc, "Update Title")
    t.Start()
    try:
        view = item.element
        collector = FilteredElementCollector(doc)\
            .OfClass(Viewport)\
            .WhereElementIsNotElementType()
        
        updated = False
        for vp in collector:
            if vp.ViewId == view.Id:
                title_param = vp.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                if title_param:
                    title_param.Set(title_str)
                    item.title_on_sheet = title_str
                    updated = True
                    break
        
        if not updated:
            raise ValueError("View not on sheet")
            
        t.Commit()
        return True
    except Exception as e:
        t.RollBack()
        raise e

def duplicate_views(doc, views):
    """Duplicate list of views"""
    t = Transaction(doc, "Duplicate")
    t.Start()
    count = 0
    try:
        for item in views:
            try:
                view = item.element
                new_id = view.Duplicate(DB.ViewDuplicateOption.Duplicate)
                new_view = doc.GetElement(new_id)
                new_view.Name = view.Name + " - Copy"
                count += 1
            except:
                pass
        t.Commit()
        return count
    except Exception as e:
        t.RollBack()
        raise e

def delete_views(doc, views):
    """Delete list of views"""
    t = Transaction(doc, "Delete")
    t.Start()
    count = 0
    try:
        for item in views:
            try:
                doc.Delete(item.id)
                count += 1
            except:
                pass
        t.Commit()
        return count
    except Exception as e:
        t.RollBack()
        raise e

def match_level(view_name, levels):
    """Match a Level from view name. Exact match first, then longest substring."""
    view_name_lower = view_name.lower()
    
    for lvl_name, lvl in levels.items():
        if lvl_name.lower() == view_name_lower:
            return lvl
    
    best_match = None
    best_len = 0
    for lvl_name, lvl in levels.items():
        if lvl_name.lower() in view_name_lower:
            if len(lvl_name) > best_len:
                best_match = lvl
                best_len = len(lvl_name)
    return best_match

def _apply_crop_box(view, vd):
    """Apply crop box settings from view definition dict."""
    try:
        crop_min_str = vd.get('crop_min', '').strip()
        crop_max_str = vd.get('crop_max', '').strip()
        
        if crop_min_str and crop_max_str:
            min_parts = [float(v.strip()) for v in crop_min_str.split(',')]
            max_parts = [float(v.strip()) for v in crop_max_str.split(',')]
            
            if len(min_parts) == 3 and len(max_parts) == 3:
                new_bb = view.CropBox
                if new_bb is None:
                    new_bb = BoundingBoxXYZ()
                
                new_bb.Min = XYZ(min_parts[0], min_parts[1], min_parts[2])
                new_bb.Max = XYZ(max_parts[0], max_parts[1], max_parts[2])
                
                view.CropBox = new_bb
        
        crop_active = vd.get('crop_active', '').strip().lower()
        if crop_active == 'yes':
            view.CropBoxActive = True
        elif crop_active == 'no':
            view.CropBoxActive = False
        
        crop_visible = vd.get('crop_visible', '').strip().lower()
        if crop_visible == 'yes':
            view.CropBoxVisible = True
        elif crop_visible == 'no':
            view.CropBoxVisible = False
    except:
        pass


def create_views_from_defs(doc, view_defs):
    """Create views from list of definitions.
    Returns (created_count, skipped_count, failed_list)
    """
    existing_names = set()
    collector = FilteredElementCollector(doc)\
        .OfClass(View)\
        .WhereElementIsNotElementType()
    for v in collector:
        if not v.IsTemplate:
            existing_names.add(v.Name)
    
    levels = {}
    for lvl in FilteredElementCollector(doc).OfClass(Level):
        levels[lvl.Name] = lvl
    
    vf_types = {}
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        fam = vft.ViewFamily
        if fam not in vf_types:
            vf_types[fam] = vft
    
    templates = {}
    for v in FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType():
        if v.IsTemplate:
            templates[v.Name] = v.Id
    
    area_schemes = {}
    try:
        for scheme in FilteredElementCollector(doc).OfClass(AreaScheme):
            area_schemes[scheme.Name] = scheme.Id
    except:
        pass
    
    t = Transaction(doc, "DQT - Create Views from Excel")
    t.Start()
    
    created = 0
    failed = []
    dup_skipped = 0
    
    detail_map = {
        'Coarse': ViewDetailLevel.Coarse,
        'Medium': ViewDetailLevel.Medium,
        'Fine': ViewDetailLevel.Fine
    }
    
    type_to_family = {
        "Floor Plan": ViewFamily.FloorPlan,
        "Ceiling Plan": ViewFamily.CeilingPlan,
        "Structural Plan": ViewFamily.StructuralPlan,
        "Drafting View": ViewFamily.Drafting,
        "3D View": ViewFamily.ThreeDimensional,
        "Section": ViewFamily.Section,
        "Area Plan": ViewFamily.AreaPlan,
        "Legend": ViewFamily.Legend,
    }
    
    try:
        for vd in view_defs:
            view_name = vd['name']
            view_type = vd['type']
            
            if view_name in existing_names:
                dup_skipped += 1
                continue
            
            try:
                new_view = None
                vf = type_to_family.get(view_type)
                
                if vf is None:
                    failed.append((view_name, "Unsupported type: {}".format(view_type)))
                    continue
                
                vft = vf_types.get(vf)
                if vft is None:
                    failed.append((view_name, "No ViewFamilyType for {}".format(view_type)))
                    continue
                
                if view_type in ["Floor Plan", "Ceiling Plan", "Structural Plan"]:
                    matched_level = None
                    level_hint = vd.get('level', '').strip()
                    if level_hint:
                        matched_level = levels.get(level_hint)
                        if matched_level is None:
                            for lvl_name, lvl in levels.items():
                                if lvl_name.lower() == level_hint.lower():
                                    matched_level = lvl
                                    break
                    
                    if matched_level is None:
                        matched_level = match_level(view_name, levels)
                    
                    if matched_level is None:
                        available_levels = ", ".join(sorted(levels.keys()))
                        failed.append((view_name, 
                            "No matching Level found.\n"
                            "    Level hint: '{}'\n"
                            "    Available levels: {}".format(
                                level_hint or "(empty)", available_levels)))
                        continue
                    
                    try:
                        new_view = ViewPlan.Create(doc, vft.Id, matched_level.Id)
                    except:
                        # Fallback: duplicate existing view of same type+level
                        existing_plan = None
                        target_vt = {
                            "Floor Plan": ViewType.FloorPlan,
                            "Ceiling Plan": ViewType.CeilingPlan,
                            "Structural Plan": ViewType.EngineeringPlan,
                        }.get(view_type)
                        
                        for v in FilteredElementCollector(doc).OfClass(ViewPlan):
                            try:
                                if (v.ViewType == target_vt 
                                    and not v.IsTemplate 
                                    and hasattr(v, 'GenLevel') 
                                    and v.GenLevel is not None
                                    and v.GenLevel.Id == matched_level.Id):
                                    existing_plan = v
                                    break
                            except:
                                continue
                        
                        if existing_plan:
                            try:
                                new_id = existing_plan.Duplicate(DB.ViewDuplicateOption.Duplicate)
                                new_view = doc.GetElement(new_id)
                            except Exception as dup_ex:
                                failed.append((view_name, 
                                    "Cannot create (plan already exists for Level '{}') "
                                    "and Duplicate also failed: {}".format(
                                        matched_level.Name, str(dup_ex))))
                                continue
                        else:
                            failed.append((view_name, 
                                "Cannot create {} for Level '{}'. "
                                "A plan may already exist for this level+type combination.".format(
                                    view_type, matched_level.Name)))
                            continue
                
                elif view_type == "Drafting View":
                    new_view = ViewDrafting.Create(doc, vft.Id)
                
                elif view_type == "3D View":
                    new_view = View3D.CreateIsometric(doc, vft.Id)
                
                elif view_type == "Section":
                    bb = BoundingBoxXYZ()
                    bb.Min = XYZ(-10, -10, -10)
                    bb.Max = XYZ(10, 10, 10)
                    transform = Transform.Identity
                    transform.Origin = XYZ(0, 0, 0)
                    transform.BasisX = XYZ(1, 0, 0)
                    transform.BasisY = XYZ(0, 0, 1)
                    transform.BasisZ = XYZ(0, -1, 0)
                    bb.Transform = transform
                    new_view = ViewSection.CreateSection(doc, vft.Id, bb)
                
                elif view_type == "Area Plan":
                    matched_level = None
                    level_hint = vd.get('level', '').strip()
                    if level_hint:
                        matched_level = levels.get(level_hint)
                        if matched_level is None:
                            for lvl_name, lvl in levels.items():
                                if lvl_name.lower() == level_hint.lower():
                                    matched_level = lvl
                                    break
                    if matched_level is None:
                        matched_level = match_level(view_name, levels)
                    if matched_level is None:
                        failed.append((view_name, "No matching Level."))
                        continue
                    scheme_id = list(area_schemes.values())[0] if area_schemes else None
                    if scheme_id:
                        new_view = ViewPlan.CreateAreaPlan(doc, scheme_id, matched_level.Id)
                    else:
                        failed.append((view_name, "No Area Scheme in project."))
                        continue
                
                elif view_type == "Legend":
                    existing_legend = None
                    for v in FilteredElementCollector(doc).OfClass(View):
                        if v.ViewType == ViewType.Legend and not v.IsTemplate:
                            existing_legend = v
                            break
                    if existing_legend:
                        new_id = existing_legend.Duplicate(DB.ViewDuplicateOption.Duplicate)
                        new_view = doc.GetElement(new_id)
                    else:
                        failed.append((view_name, "No Legend to duplicate."))
                        continue
                
                if new_view:
                    try:
                        new_view.Name = view_name
                    except:
                        pass
                    
                    if vd.get('scale'):
                        try:
                            new_view.Scale = int(vd['scale'])
                        except:
                            pass
                    
                    if vd.get('detail_level') and vd['detail_level'] in detail_map:
                        try:
                            new_view.DetailLevel = detail_map[vd['detail_level']]
                        except:
                            pass
                    
                    if vd.get('template') and vd['template'] != "None" and vd['template'] in templates:
                        try:
                            new_view.ViewTemplateId = templates[vd['template']]
                        except:
                            pass
                    
                    # Apply Crop Box
                    _apply_crop_box(new_view, vd)
                    
                    created += 1
                    existing_names.add(view_name)
            
            except Exception as ex:
                failed.append((view_name, str(ex)))
        
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        raise e
        
    return created, dup_skipped, failed

def apply_excel_updates(doc, updates, all_view_items):
    """Apply updates from Excel to Revit views.
    Returns (updated_count, skipped_count, custom_param_updates, custom_param_errors)
    """
    t = Transaction(doc, "Import from Excel")
    t.Start()
    
    count = 0
    skipped = 0
    custom_param_updates = 0
    custom_param_errors = []
    
    try:
        for update in updates:
            view = None
            
            if update.get('element_id'):
                try:
                    view_elem = doc.GetElement(_make_eid(update['element_id']))
                    if view_elem and isinstance(view_elem, View):
                        view = view_elem
                except:
                    pass
            
            if not view and update.get('view_name'):
                for v in all_view_items:
                    if v.name == update['view_name']:
                        view = v.element
                        break
            
            if not view:
                skipped += 1
                continue
            
            # Update name
            if update.get('view_name') and update['view_name'] != view.Name:
                try:
                    view.Name = update['view_name']
                except:
                    pass
            
            # Update template
            if update.get('template') and update['template'] != "None":
                templates = FilteredElementCollector(doc)\
                    .OfClass(View)\
                    .WhereElementIsElementType()
                for tmpl in templates:
                    if tmpl.Name == update['template']:
                        try:
                            view.ViewTemplateId = tmpl.Id
                        except:
                            pass
                        break
            
            # Update scale
            if update.get('scale'):
                try:
                    view.Scale = int(update['scale'])
                except:
                    pass
            
            # Update detail level
            if update.get('detail_level'):
                detail_map = {
                    'Coarse': ViewDetailLevel.Coarse,
                    'Medium': ViewDetailLevel.Medium,
                    'Fine': ViewDetailLevel.Fine
                }
                if update['detail_level'] in detail_map:
                    try:
                        view.DetailLevel = detail_map[update['detail_level']]
                    except:
                        pass
            
            # Update custom parameters
            if update.get('custom_params'):
                for param_name, param_value in update['custom_params'].items():
                    try:
                        param = view.LookupParameter(param_name)
                        if param:
                            if param.IsReadOnly:
                                if param_name not in [e[0] for e in custom_param_errors]:
                                    custom_param_errors.append((param_name, "Read-only parameter"))
                                continue
                            
                            success = False
                            if param.StorageType == StorageType.String:
                                param.Set(str(param_value))
                                success = True
                            elif param.StorageType == StorageType.Integer:
                                try:
                                    param.Set(int(float(param_value)))
                                    success = True
                                except:
                                    pass
                            elif param.StorageType == StorageType.Double:
                                try:
                                    param.Set(float(param_value))
                                    success = True
                                except:
                                    pass
                            
                            if success:
                                custom_param_updates += 1
                        else:
                            if param_name not in [e[0] for e in custom_param_errors]:
                                custom_param_errors.append((param_name, "Parameter not found"))
                    except Exception as e:
                        if param_name not in [e[0] for e in custom_param_errors]:
                            custom_param_errors.append((param_name, str(e)))
            
            # Update crop box
            _apply_crop_box(view, update)
            
            count += 1
        
        t.Commit()
    except Exception as e:
        t.RollBack()
        raise e
        
    return count, skipped, custom_param_updates, custom_param_errors


# =====================================================
# XLSX WRITER & READER - Pure Python
# =====================================================

def write_xlsx(filepath, headers, rows, hidden_cols=None, header_colors=None):
    """Write data to .xlsx file using pure Python XML + zipfile."""
    if hidden_cols is None:
        hidden_cols = []
    if header_colors is None:
        header_colors = {}
    
    def _col_letter(idx):
        result = ""
        idx += 1
        while idx > 0:
            idx -= 1
            result = chr(65 + idx % 26) + result
            idx //= 26
        return result
    
    def _escape_xml(val):
        if val is None:
            return ""
        s = str(val)
        cleaned = []
        for ch in s:
            code = ord(ch)
            if code == 0x9 or code == 0xA or code == 0xD:
                cleaned.append(ch)
            elif code >= 0x20:
                cleaned.append(ch)
        s = "".join(cleaned)
        s = s.replace("&", "&amp;")
        s = s.replace("<", "&lt;")
        s = s.replace(">", "&gt;")
        s = s.replace('"', "&quot;")
        s = s.replace("'", "&apos;")
        return s
    
    default_header_color = "0F172A"  # DQT Gold
    fill_colors = [default_header_color]
    for ci, color in sorted(header_colors.items()):
        if color not in fill_colors:
            fill_colors.append(color)
    
    fills_xml = '<fills count="{}">\n'.format(len(fill_colors) + 2)
    fills_xml += '<fill><patternFill patternType="none"/></fill>\n'
    fills_xml += '<fill><patternFill patternType="gray125"/></fill>\n'
    for color in fill_colors:
        fills_xml += '<fill><patternFill patternType="solid"><fgColor rgb="FF{}"/><bgColor indexed="64"/></patternFill></fill>\n'.format(color)
    fills_xml += '</fills>\n'
    
    num_styles = 2 + len(fill_colors)
    xfs_xml = '<cellXfs count="{}">\n'.format(num_styles)
    xfs_xml += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />\n'
    xfs_xml += '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/>\n'
    for fi in range(len(fill_colors)):
        xfs_xml += '<xf numFmtId="0" fontId="1" fillId="{}" borderId="0" applyFont="1" applyFill="1"/>\n'.format(fi + 2)
    xfs_xml += '</cellXfs>\n'
    
    styles_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    styles_xml += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
    styles_xml += '<fonts count="2">\n'
    styles_xml += '<font><sz val="11"/><name val="Calibri"/></font>\n'
    styles_xml += '<font><b/><sz val="11"/><name val="Calibri"/></font>\n'
    styles_xml += '</fonts>\n'
    styles_xml += fills_xml
    styles_xml += '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
    styles_xml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
    styles_xml += xfs_xml
    styles_xml += '</styleSheet>'
    
    all_strings = []
    string_map = {}
    
    def _get_string_idx(val):
        s = str(val) if val is not None else ""
        if s not in string_map:
            string_map[s] = len(all_strings)
            all_strings.append(s)
        return string_map[s]
    
    sheet_rows = []
    
    # Header row
    header_cells = []
    for ci, h in enumerate(headers):
        col = _col_letter(ci)
        si = _get_string_idx(h)
        color = header_colors.get(ci, default_header_color)
        if color in fill_colors:
            style_id = fill_colors.index(color) + 2
        else:
            style_id = 1
        header_cells.append('<c r="{}1" t="s" s="{}"><v>{}</v></c>'.format(col, style_id, si))
    sheet_rows.append('<row r="1">{}</row>'.format("".join(header_cells)))
    
    # Data rows
    for ri, row_data in enumerate(rows):
        row_num = ri + 2
        cells = []
        for ci, val in enumerate(row_data):
            col = _col_letter(ci)
            ref = "{}{}".format(col, row_num)
            
            if val is None or val == "":
                cells.append('<c r="{}"><v></v></c>'.format(ref))
            elif isinstance(val, (int, float)):
                cells.append('<c r="{}"><v>{}</v></c>'.format(ref, val))
            else:
                si = _get_string_idx(val)
                cells.append('<c r="{}" t="s"><v>{}</v></c>'.format(ref, si))
        
        sheet_rows.append('<row r="{}">{}</row>'.format(row_num, "".join(cells)))
    
    cols_xml = '<cols>\n'
    for ci in range(len(headers)):
        width = 15
        hidden = ' hidden="1"' if ci in hidden_cols else ''
        cols_xml += '<col min="{}" max="{}" width="{}" bestFit="1" customWidth="1"{}/>'.format(ci+1, ci+1, width, hidden)
    cols_xml += '\n</cols>\n'
    
    sheet_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sheet_xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    sheet_xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
    sheet_xml += cols_xml
    sheet_xml += '<sheetData>\n'
    sheet_xml += '\n'.join(sheet_rows)
    sheet_xml += '\n</sheetData>\n'
    sheet_xml += '</worksheet>'
    
    sst_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    sst_xml += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">\n'.format(
        len(all_strings), len(all_strings))
    for s in all_strings:
        sst_xml += '<si><t>{}</t></si>\n'.format(_escape_xml(s))
    sst_xml += '</sst>'
    
    workbook_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    workbook_xml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    workbook_xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
    workbook_xml += '<sheets><sheet name="Views" sheetId="1" r:id="rId1"/></sheets>\n'
    workbook_xml += '</workbook>'
    
    workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    workbook_rels += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    workbook_rels += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n'
    workbook_rels += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
    workbook_rels += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
    workbook_rels += '</Relationships>'
    
    rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    rels_xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
    rels_xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
    rels_xml += '</Relationships>'
    
    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    content_types += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
    content_types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
    content_types += '<Default Extension="xml" ContentType="application/xml"/>\n'
    content_types += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
    content_types += '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
    content_types += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
    content_types += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n'
    content_types += '</Types>'
    
    if os.path.exists(filepath):
        os.remove(filepath)
    
    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels_xml)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        zf.writestr('xl/worksheets/sheet1.xml', sheet_xml)
        zf.writestr('xl/styles.xml', styles_xml)
        zf.writestr('xl/sharedStrings.xml', sst_xml)

def read_xlsx(filepath):
    """Read .xlsx file using pure Python zipfile + XML parsing."""
    try:
        from xml.etree import ElementTree as ET
    except:
        import xml.etree.ElementTree as ET
    
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    
    def _sanitize_xml(raw_bytes):
        text = None
        for enc in ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']:
            try:
                text = raw_bytes.decode(enc)
                break
            except:
                continue
        
        if text is None:
            text = raw_bytes.decode('utf-8', errors='replace')
        
        if text.startswith('\xef\xbb\xbf'):
            text = text[3:]
        if text.startswith('\ufeff'):
            text = text[1:]
        
        cleaned = []
        for ch in text:
            code = ord(ch)
            if code == 0x9 or code == 0xA or code == 0xD:
                cleaned.append(ch)
            elif code >= 0x20 and code <= 0xD7FF:
                cleaned.append(ch)
            elif code >= 0xE000 and code <= 0xFFFD:
                cleaned.append(ch)
        text = "".join(cleaned)
        
        text = re.sub(
            r'<\?xml[^?]*\?>',
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            text,
            count=1
        )
        return text
    
    def _parse_xml(raw_bytes):
        try:
            return ET.fromstring(raw_bytes)
        except:
            pass
        
        text = _sanitize_xml(raw_bytes)
        try:
            return ET.fromstring(text.encode('utf-8'))
        except:
            pass
        
        text = re.sub(r'<\?xml[^?]*\?>', '', text, count=1).strip()
        return ET.fromstring(text.encode('utf-8'))
    
    def _find_all(root, tag):
        results = root.findall('.//{%s}%s' % (ns, tag))
        if results:
            return results
        return root.findall('.//' + tag)
    
    def _find(elem, tag):
        result = elem.find('{%s}%s' % (ns, tag))
        if result is not None:
            return result
        return elem.find(tag)
    
    with zipfile.ZipFile(filepath, 'r') as zf:
        shared_strings = []
        if 'xl/sharedStrings.xml' in zf.namelist():
            sst_data = zf.read('xl/sharedStrings.xml')
            sst_root = _parse_xml(sst_data)
            for si in _find_all(sst_root, 'si'):
                texts = []
                t_elem = _find(si, 't')
                if t_elem is not None and t_elem.text:
                    texts.append(t_elem.text)
                else:
                    for t in _find_all(si, 't'):
                        if t.text:
                            texts.append(t.text)
                shared_strings.append("".join(texts))
        
        sheet_data = zf.read('xl/worksheets/sheet1.xml')
        sheet_root = _parse_xml(sheet_data)
    
    def _col_from_ref(ref):
        col = ""
        for ch in ref:
            if ch.isalpha():
                col += ch
            else:
                break
        return col
    
    def _col_to_index(col_str):
        idx = 0
        for ch in col_str.upper():
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1
    
    def _row_from_ref(ref):
        num = ""
        for ch in ref:
            if ch.isdigit():
                num += ch
        return int(num) if num else 0
    
    cells = {}
    max_row = 0
    max_col = 0
    
    sheet_data_elem = _find(sheet_root, 'sheetData')
    if sheet_data_elem is None:
        for sd in _find_all(sheet_root, 'sheetData'):
            sheet_data_elem = sd
            break
    
    if sheet_data_elem is not None:
        row_tag_ns = '{%s}row' % ns
        row_elems = list(sheet_data_elem)
        if not row_elems:
            row_elems = sheet_data_elem.findall(row_tag_ns)
        if not row_elems:
            row_elems = sheet_data_elem.findall('row')
        
        for row_elem in row_elems:
            cell_tag_ns = '{%s}c' % ns
            cell_elems = list(row_elem)
            if not cell_elems:
                cell_elems = row_elem.findall(cell_tag_ns)
            if not cell_elems:
                cell_elems = row_elem.findall('c')
            
            for cell in cell_elems:
                ref = cell.get('r', '')
                if not ref:
                    continue
                
                col_idx = _col_to_index(_col_from_ref(ref))
                row_idx = _row_from_ref(ref) - 1
                
                cell_type = cell.get('t', '')
                v_elem = _find(cell, 'v')
                
                value = None
                if v_elem is not None and v_elem.text is not None:
                    if cell_type == 's':
                        try:
                            si = int(v_elem.text)
                            value = shared_strings[si] if si < len(shared_strings) else ""
                        except:
                            value = v_elem.text
                    elif cell_type == 'b':
                        value = v_elem.text == '1'
                    else:
                        try:
                            fval = float(v_elem.text)
                            if fval == int(fval):
                                value = int(fval)
                            else:
                                value = fval
                        except:
                            value = v_elem.text
                else:
                    is_elem = None
                    is_parent = _find(cell, 'is')
                    if is_parent is not None:
                        is_elem = _find(is_parent, 't')
                    if is_elem is not None and is_elem.text:
                        value = is_elem.text
                
                if value is not None:
                    cells[(row_idx, col_idx)] = value
                    if row_idx > max_row:
                        max_row = row_idx
                    if col_idx > max_col:
                        max_col = col_idx
    
    if not cells:
        return [], []
    
    headers = []
    for ci in range(max_col + 1):
        val = cells.get((0, ci), "")
        headers.append(str(val) if val else "")
    
    rows = []
    for ri in range(1, max_row + 1):
        row = []
        has_data = False
        for ci in range(max_col + 1):
            val = cells.get((ri, ci))
            row.append(val)
            if val is not None and val != "":
                has_data = True
        if has_data:
            rows.append(row)
            
    return headers, rows
