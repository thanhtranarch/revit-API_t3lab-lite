# -*- coding: utf-8 -*-
"""
Sheet Manager - ViewSheet Sets Service

Copyright © Dang Quoc Truong (DQT)
"""

from Autodesk.Revit.DB import ViewSheetSet, FilteredElementCollector, ViewSheet


class ViewSheetSetsService(object):
    """Manage ViewSheet Sets"""
    
    def __init__(self, doc):
        self.doc = doc
    
    def get_all_sheet_sets(self):
        """Get all ViewSheet Sets from PrintManager"""
        try:
            print_mgr = self.doc.PrintManager
            view_sheet_setting = print_mgr.ViewSheetSetting
            
            sets = []
            
            # Get available view sheet sets
            available_sets = view_sheet_setting.AvailableViewSheetSets
            
            for set_name in available_sets.Keys:
                sheet_set = available_sets[set_name]
                sets.append({
                    'element': sheet_set,
                    'name': sheet_set.Name,
                    'id': None  # ViewSheetSet doesn't have Id
                })
            
            return sets
        except Exception as e:
            print("Error getting sheet sets: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []
    
    def create_sheet_set(self, name):
        """Create new ViewSheet Set"""
        try:
            # Use constructor, not static Create method
            new_set = ViewSheetSet()
            new_set.Name = name
            
            # Insert into document
            from Autodesk.Revit.DB import ViewSheetSetIterator
            # ViewSheetSet is stored in PrintManager
            print_mgr = self.doc.PrintManager
            print_mgr.ViewSheetSetting.CurrentViewSheetSet = new_set
            
            return new_set
        except Exception as e:
            print("Error creating sheet set: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return None
    
    def delete_sheet_set(self, set_name):
        """Delete a ViewSheet Set by name"""
        try:
            print_mgr = self.doc.PrintManager
            view_sheet_setting = print_mgr.ViewSheetSetting
            available_sets = view_sheet_setting.AvailableViewSheetSets
            
            # Remove from available sets
            if set_name in available_sets.Keys:
                available_sets.Remove(set_name)
                return True
            
            return False
        except Exception as e:
            print("Error deleting sheet set: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return False
    
    def rename_sheet_set(self, sheet_set, new_name):
        """Rename a ViewSheet Set"""
        try:
            sheet_set.Name = new_name
            return True
        except Exception as e:
            print("Error renaming sheet set: {}".format(str(e)))
            return False
    
    def add_sheets_to_set(self, sheet_set, sheet_ids):
        """Add sheets to a ViewSheet Set"""
        try:
            views = sheet_set.Views
            
            for sheet_id in sheet_ids:
                if sheet_id not in views:
                    views.Insert(sheet_id)
            
            return True
        except Exception as e:
            print("Error adding sheets to set: {}".format(str(e)))
            return False
    
    def remove_sheets_from_set(self, sheet_set, sheet_ids):
        """Remove sheets from a ViewSheet Set"""
        try:
            views = sheet_set.Views
            
            for sheet_id in sheet_ids:
                if sheet_id in views:
                    views.Erase(sheet_id)
            
            return True
        except Exception as e:
            print("Error removing sheets from set: {}".format(str(e)))
            return False
    
    def get_sheets_in_set(self, sheet_set):
        """Get all sheet IDs in a set"""
        try:
            views = sheet_set.Views
            sheet_ids = []
            
            for view_id in views:
                sheet_ids.append(view_id)
            
            return sheet_ids
        except Exception as e:
            print("Error getting sheets in set: {}".format(str(e)))
            return []