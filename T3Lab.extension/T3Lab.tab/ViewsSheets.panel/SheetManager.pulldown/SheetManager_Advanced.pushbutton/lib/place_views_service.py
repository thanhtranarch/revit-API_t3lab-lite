# -*- coding: utf-8 -*-
"""
Sheet Manager - Place Views Service

Copyright © Dang Quoc Truong (DQT)
"""

from Autodesk.Revit.DB import FilteredElementCollector, View, Viewport, XYZ, UV


class PlaceViewsService(object):
    """Handle placing views on sheets"""
    
    def __init__(self, doc):
        self.doc = doc
    
    def get_placeable_views(self):
        """Get all views that can be placed on sheets"""
        try:
            # First, get all viewports to check which views are already placed
            viewports = FilteredElementCollector(self.doc).OfClass(Viewport)
            placed_view_ids = set()
            for vp in viewports:
                placed_view_ids.add(vp.ViewId)
            
            collector = FilteredElementCollector(self.doc).OfClass(View)
            
            placeable_views = []
            for view in collector:
                # Skip templates, schedules on sheets, legends on sheets
                if (not view.IsTemplate and 
                    view.CanBePrinted and
                    hasattr(view, 'ViewType')):
                    
                    # Check if view ID is in placed views
                    on_sheet = view.Id in placed_view_ids
                    
                    placeable_views.append({
                        'element': view,
                        'id': view.Id,
                        'name': view.Name,
                        'type': str(view.ViewType),
                        'on_sheet': on_sheet
                    })
                    
                    if on_sheet:
                        print("DEBUG: View '{}' is already on a sheet".format(view.Name))
            
            return placeable_views
        except Exception as e:
            print("Error getting placeable views: {}".format(str(e)))
            return []
    
    def place_view_on_sheet(self, sheet, view, location=None):
        """Place a view on a sheet"""
        try:
            # Check if view is already on a sheet
            from Autodesk.Revit.DB import FilteredElementCollector, Viewport
            viewports = FilteredElementCollector(self.doc).OfClass(Viewport)
            for vp in viewports:
                if vp.ViewId == view.Id:
                    print("WARNING: View '{}' is already placed on a sheet".format(view.Name))
                    return None
            
            # Default center location if not specified
            if location is None:
                # Get sheet center
                outline = sheet.Outline
                center_x = (outline.Min.U + outline.Max.U) / 2
                center_y = (outline.Min.V + outline.Max.V) / 2
                location = XYZ(center_x, center_y, 0)
            
            # Create viewport
            viewport = Viewport.Create(self.doc, sheet.Id, view.Id, location)
            print("SUCCESS: Placed view '{}' on sheet '{}'".format(view.Name, sheet.SheetNumber))
            return viewport
            
        except Exception as e:
            print("Error placing view on sheet: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return None
    
    def auto_arrange_views_on_sheet(self, sheet, views, rows=2, cols=2):
        """Auto-arrange multiple views on a sheet"""
        try:
            # Get sheet dimensions
            outline = sheet.Outline
            sheet_width = outline.Max.U - outline.Min.U
            sheet_height = outline.Max.V - outline.Min.V
            
            # Calculate spacing
            h_spacing = sheet_width / (cols + 1)
            v_spacing = sheet_height / (rows + 1)
            
            viewports = []
            view_index = 0
            
            for row in range(rows):
                for col in range(cols):
                    if view_index >= len(views):
                        break
                    
                    # Calculate position
                    x = outline.Min.U + (col + 1) * h_spacing
                    y = outline.Min.V + (row + 1) * v_spacing
                    location = XYZ(x, y, 0)
                    
                    # Place view
                    viewport = self.place_view_on_sheet(sheet, views[view_index], location)
                    if viewport:
                        viewports.append(viewport)
                    
                    view_index += 1
            
            return viewports
            
        except Exception as e:
            print("Error auto-arranging views: {}".format(str(e)))
            return []
    
    def batch_place_views(self, sheets, views, mode='one_per_sheet'):
        """Batch place views on multiple sheets
        
        Modes:
        - 'one_per_sheet': Place one view per sheet
        - 'all_on_each': Place all views on each sheet
        - 'distribute': Distribute views evenly across sheets
        """
        try:
            placements = []
            
            if mode == 'one_per_sheet':
                for i, sheet in enumerate(sheets):
                    if i < len(views):
                        viewport = self.place_view_on_sheet(sheet, views[i])
                        if viewport:
                            placements.append({
                                'sheet': sheet,
                                'view': views[i],
                                'viewport': viewport
                            })
            
            elif mode == 'all_on_each':
                for sheet in sheets:
                    sheet_placements = self.auto_arrange_views_on_sheet(
                        sheet, views, rows=2, cols=2)
                    for viewport in sheet_placements:
                        placements.append({
                            'sheet': sheet,
                            'viewport': viewport
                        })
            
            elif mode == 'distribute':
                views_per_sheet = max(1, len(views) / len(sheets))
                view_index = 0
                
                for sheet in sheets:
                    sheet_views = views[view_index:view_index + views_per_sheet]
                    sheet_placements = self.auto_arrange_views_on_sheet(
                        sheet, sheet_views, rows=2, cols=2)
                    
                    for viewport in sheet_placements:
                        placements.append({
                            'sheet': sheet,
                            'viewport': viewport
                        })
                    
                    view_index += views_per_sheet
            
            return placements
            
        except Exception as e:
            print("Error in batch place views: {}".format(str(e)))
            return []