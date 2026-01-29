# -*- coding: utf-8 -*-
"""
Unreferenced Views Scanner
Scans for views not placed on any sheet, categorized by view type

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, View, ViewType,
    ViewPlan, ViewSection, View3D
)
from base_scanner import BaseAdvancedScanner


class UnreferencedViewsScanner(BaseAdvancedScanner):
    """Scans for unreferenced views by type"""
    
    # View type mappings
    VIEW_TYPE_MAP = {
        'unreferenced_3d': ViewType.ThreeD,
        'unreferenced_area': ViewType.AreaPlan,
        'unreferenced_detail': ViewType.Detail,
        'unreferenced_drafting': ViewType.DraftingView,
        'unreferenced_elevation': ViewType.Elevation,
        'unreferenced_engineering': ViewType.EngineeringPlan,
        'unreferenced_floor': ViewType.FloorPlan,
        'unreferenced_rcp': ViewType.CeilingPlan,
        'unreferenced_section': ViewType.Section,
    }
    
    def scan(self, category, progress_callback=None):
        """
        Scan for unreferenced views of specific type
        
        Args:
            category: AdvancedPurgeCategory with id like 'unreferenced_3d'
            progress_callback: Optional callback(message)
            
        Returns:
            List of unreferenced view dictionaries
        """
        items = []
        
        try:
            if progress_callback:
                progress_callback("Scanning {}...".format(category.name))
            
            # Get view type for this category
            view_type = self.VIEW_TYPE_MAP.get(category.id)
            if not view_type:
                return items
            
            # Get all views of this type
            all_views = FilteredElementCollector(self.doc)\
                .OfClass(View)\
                .ToElements()
            
            # Filter by view type
            views_of_type = [v for v in all_views if v.ViewType == view_type]
            
            if progress_callback:
                progress_callback("Found {} {} views".format(
                    len(views_of_type), category.name
                ))
            
            # Get all sheets to check which views are placed
            all_sheets = FilteredElementCollector(self.doc)\
                .OfClass(ViewSheet)\
                .ToElements()
            
            # Collect all view IDs placed on sheets
            placed_view_ids = set()
            for sheet in all_sheets:
                try:
                    viewport_ids = sheet.GetAllViewports()
                    for vp_id in viewport_ids:
                        viewport = self.doc.GetElement(vp_id)
                        if viewport:
                            placed_view_ids.add(viewport.ViewId)
                except:
                    continue
            
            if progress_callback:
                progress_callback("Checking {} views for references...".format(
                    len(views_of_type)
                ))
            
            # Find unreferenced views
            for view in views_of_type:
                try:
                    # Skip if view is on a sheet
                    if view.Id in placed_view_ids:
                        continue
                    
                    # Skip view templates
                    if view.IsTemplate:
                        continue
                    
                    # Skip system views
                    view_name = view.Name if hasattr(view, 'Name') else ""
                    if view_name.startswith("{") or view_name.startswith("<"):
                        continue
                    
                    # Check if can be deleted
                    if not self.can_delete(view):
                        continue
                    
                    # Add to results
                    items.append(self.create_item_dict(view, category.name))
                    
                except Exception as e:
                    # Skip problematic views
                    continue
            
            if progress_callback:
                progress_callback("Found {} unreferenced views".format(len(items)))
            
        except Exception as e:
            if progress_callback:
                progress_callback("Error scanning {}: {}".format(
                    category.name, str(e)
                ))
        
        return items
    
    def can_delete(self, view):
        """
        Check if view can be deleted (enhanced for views)
        
        Args:
            view: View element
            
        Returns:
            True if view can be deleted
        """
        try:
            # Base checks
            if not super(UnreferencedViewsScanner, self).can_delete(view):
                return False
            
            # Don't delete active view
            if view.Id == self.doc.ActiveView.Id:
                return False
            
            # Check if view is template
            if view.IsTemplate:
                return False
            
            return True
            
        except:
            return False
