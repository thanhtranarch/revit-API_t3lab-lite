# -*- coding: utf-8 -*-
"""
Unshaded Views Scanner
Scans for views not placed on any sheet (unshaded), categorized by view type

Note: "Unshaded" is another term for unreferenced views in some contexts.
This provides alternate naming for user preference.

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, View, ViewType
)
from base_scanner import BaseAdvancedScanner


class UnshadedViewsScanner(BaseAdvancedScanner):
    """Scans for unshaded (unreferenced) views by type"""
    
    # View type mappings
    VIEW_TYPE_MAP = {
        'unshaded_3d': ViewType.ThreeD,
        'unshaded_area': ViewType.AreaPlan,
        'unshaded_detail': ViewType.Detail,
        'unshaded_drafting': ViewType.DraftingView,
        'unshaded_elevation': ViewType.Elevation,
        'unshaded_engineering': ViewType.EngineeringPlan,
        'unshaded_floor': ViewType.FloorPlan,
        'unshaded_rcp': ViewType.CeilingPlan,
        'unshaded_section': ViewType.Section,
    }
    
    def scan(self, category, progress_callback=None):
        """
        Scan for unshaded views of specific type
        
        Args:
            category: AdvancedPurgeCategory with id like 'unshaded_3d'
            progress_callback: Optional callback(message)
            
        Returns:
            List of unshaded view dictionaries
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
            views_of_type = [v for v in all_views 
                           if v.ViewType == view_type and not v.IsTemplate]
            
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
                progress_callback("Checking {} views...".format(len(views_of_type)))
            
            # Find unshaded (unreferenced) views
            for view in views_of_type:
                try:
                    # Skip if view is on a sheet
                    if view.Id in placed_view_ids:
                        continue
                    
                    # Skip system views (names starting with { or <)
                    view_name = view.Name if hasattr(view, 'Name') else ""
                    if view_name.startswith("{") or view_name.startswith("<"):
                        continue
                    
                    # Skip active view
                    if view.Id == self.doc.ActiveView.Id:
                        continue
                    
                    # Check if can be deleted
                    if not self.can_delete(view):
                        continue
                    
                    # Add to results
                    items.append(self.create_item_dict(view, category.name))
                    
                except Exception as e:
                    continue
            
            if progress_callback:
                progress_callback("Found {} unshaded views".format(len(items)))
            
        except Exception as e:
            if progress_callback:
                progress_callback("Error: {}".format(str(e)))
        
        return items
