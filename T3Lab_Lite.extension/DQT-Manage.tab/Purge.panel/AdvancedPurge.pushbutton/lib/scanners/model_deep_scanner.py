# -*- coding: utf-8 -*-
"""
Model Deep Scanner
Scans for deep model cleanup items

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

try:
    from base_scanner import BaseAdvancedScanner
except ImportError:
    import sys
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from base_scanner import BaseAdvancedScanner

from Autodesk.Revit.DB import (
    FilteredElementCollector, 
    BuiltInCategory,
    ElevationMarker,
    Element,
    ReferencePlane,
    ModelCurve,
    CurveElement
)


class ModelDeepScanner(BaseAdvancedScanner):
    """Scanner for model deep cleanup operations"""
    
    def scan(self, category):
        """
        Scan based on category ID
        
        Args:
            category: AdvancedPurgeCategory
            
        Returns:
            List of item dictionaries
        """
        category_id = category.id
        
        # Route to appropriate scan method
        if category_id == "orphaned_elevation_markers":
            return self.scan_orphaned_elevation_markers(category)
        elif category_id == "unused_scope_boxes":
            return self.scan_unused_scope_boxes(category)
        elif category_id == "room_separation_lines":
            return self.scan_room_separation_lines(category)
        elif category_id == "unnamed_reference_planes":
            return self.scan_unnamed_reference_planes(category)
        elif category_id == "all_reference_planes":
            return self.scan_all_reference_planes(category)
        elif category_id == "unused_constraints":
            return self.scan_unused_constraints(category)
        elif category_id == "all_constraints":
            return self.scan_all_constraints(category)
        elif category_id == "unused_external_links":
            return self.scan_unused_external_links(category)
        else:
            return []
    
    def scan_orphaned_elevation_markers(self, category):
        """Scan for elevation markers with no views"""
        items = []
        
        try:
            collector = FilteredElementCollector(self.doc)
            markers = collector.OfClass(ElevationMarker).ToElements()
            
            for marker in markers:
                # Check if marker has any views
                view_count = marker.CurrentViewCount
                
                if view_count == 0:
                    if self.can_delete(marker):
                        items.append(self.create_item_dict(marker, category.name))
                        
        except Exception as e:
            print("Error scanning orphaned elevation markers: {}".format(str(e)))
        
        return items
    
    def scan_unused_scope_boxes(self, category):
        """Scan for scope boxes not used in any view"""
        items = []
        
        try:
            # Get all scope boxes
            collector = FilteredElementCollector(self.doc)
            scope_boxes = collector.OfCategory(
                BuiltInCategory.OST_VolumeOfInterest
            ).WhereElementIsNotElementType().ToElements()
            
            # Get all views
            view_collector = FilteredElementCollector(self.doc)
            views = view_collector.OfClass(Autodesk.Revit.DB.View).ToElements()
            
            # Build set of used scope box ids
            used_scope_boxes = set()
            for view in views:
                if not view.IsTemplate:
                    scope_box_param = view.get_Parameter(
                        Autodesk.Revit.DB.BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP
                    )
                    if scope_box_param and scope_box_param.AsElementId().IntegerValue > 0:
                        used_scope_boxes.add(scope_box_param.AsElementId().IntegerValue)
            
            # Find unused scope boxes
            for scope_box in scope_boxes:
                if scope_box.Id.IntegerValue not in used_scope_boxes:
                    if self.can_delete(scope_box):
                        items.append(self.create_item_dict(scope_box, category.name))
                        
        except Exception as e:
            print("Error scanning unused scope boxes: {}".format(str(e)))
        
        return items
    
    def scan_room_separation_lines(self, category):
        """Scan for all room separation lines"""
        items = []
        
        try:
            collector = FilteredElementCollector(self.doc)
            room_sep_lines = collector.OfCategory(
                BuiltInCategory.OST_RoomSeparationLines
            ).WhereElementIsNotElementType().ToElements()
            
            for line in room_sep_lines:
                if self.can_delete(line):
                    items.append(self.create_item_dict(line, category.name))
                    
        except Exception as e:
            print("Error scanning room separation lines: {}".format(str(e)))
        
        return items
    
    def scan_unnamed_reference_planes(self, category):
        """Scan for reference planes without names"""
        items = []
        
        try:
            collector = FilteredElementCollector(self.doc)
            ref_planes = collector.OfClass(ReferencePlane).ToElements()
            
            for plane in ref_planes:
                # Check if name is empty or default
                name = plane.Name
                if not name or name.strip() == "":
                    if self.can_delete(plane):
                        items.append(self.create_item_dict(plane, category.name))
                        
        except Exception as e:
            print("Error scanning unnamed reference planes: {}".format(str(e)))
        
        return items
    
    def scan_all_reference_planes(self, category):
        """Scan for all reference planes"""
        items = []
        
        try:
            collector = FilteredElementCollector(self.doc)
            ref_planes = collector.OfClass(ReferencePlane).ToElements()
            
            for plane in ref_planes:
                if self.can_delete(plane):
                    items.append(self.create_item_dict(plane, category.name))
                    
        except Exception as e:
            print("Error scanning all reference planes: {}".format(str(e)))
        
        return items
    
    def scan_unused_constraints(self, category):
        """Scan for unused constraints"""
        items = []
        
        try:
            # This is complex - constraints are difficult to detect if unused
            # For now, return empty - needs more investigation
            pass
                    
        except Exception as e:
            print("Error scanning unused constraints: {}".format(str(e)))
        
        return items
    
    def scan_all_constraints(self, category):
        """Scan for all constraints"""
        items = []
        
        try:
            # Constraints don't have a direct category
            # They're properties of elements
            # For now, return empty - this needs special handling
            pass
                    
        except Exception as e:
            print("Error scanning all constraints: {}".format(str(e)))
        
        return items
    
    def scan_unused_external_links(self, category):
        """Scan for unused external links (RVT links, DWG, etc)"""
        items = []
        
        try:
            from Autodesk.Revit.DB import ExternalFileReference, ExternalFileReferenceType
            
            # Get all external file references
            ext_file_refs = self.doc.GetExternalFileReferences()
            
            for ref_type_id in ext_file_refs.Keys:
                ref = ext_file_refs[ref_type_id]
                
                # Check if reference is loaded
                if ref.GetLinkedFileStatus() == Autodesk.Revit.DB.LinkedFileStatus.Loaded:
                    # For RVT links, check if used in views
                    # For now, we'll just list all loaded links
                    # TODO: Add actual usage check
                    pass
                    
        except Exception as e:
            print("Error scanning unused external links: {}".format(str(e)))
        
        return items