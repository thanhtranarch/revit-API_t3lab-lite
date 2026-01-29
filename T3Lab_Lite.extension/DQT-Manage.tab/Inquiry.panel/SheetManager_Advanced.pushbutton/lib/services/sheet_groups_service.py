# -*- coding: utf-8 -*-
"""
Sheet Manager - Custom Sheet Groups Service
Manage custom sheet groups with JSON persistence

Copyright © Dang Quoc Truong (DQT)
"""

import json
import os


class SheetGroupsService(object):
    """Manage custom sheet groups with JSON storage"""
    
    def __init__(self, doc):
        self.doc = doc
        self.groups = {}  # {group_name: [sheet_ids]}
        self.json_file = self._get_json_file_path()
        self.load_groups()
    
    def _get_json_file_path(self):
        """Get JSON file path for current document"""
        try:
            doc_path = self.doc.PathName
            if doc_path:
                # Save next to document
                base_name = os.path.splitext(doc_path)[0]
                json_path = base_name + "_SheetGroups.json"
            else:
                # Document not saved - use temp location
                import tempfile
                temp_dir = tempfile.gettempdir()
                json_path = os.path.join(temp_dir, "SheetManager_SheetGroups.json")
            
            return json_path
        except:
            # Fallback
            import tempfile
            temp_dir = tempfile.gettempdir()
            return os.path.join(temp_dir, "SheetManager_SheetGroups.json")
    
    def load_groups(self):
        """Load groups from JSON file"""
        try:
            if os.path.exists(self.json_file):
                with open(self.json_file, 'r') as f:
                    data = json.load(f)
                    # Convert string IDs back to ElementId
                    self.groups = {}
                    for group_name, sheet_id_strings in data.items():
                        from Autodesk.Revit.DB import ElementId
                        sheet_ids = [ElementId(int(id_str)) for id_str in sheet_id_strings]
                        self.groups[group_name] = sheet_ids
                    
                print("DEBUG: Loaded {} groups from {}".format(len(self.groups), self.json_file))
            else:
                self.groups = {}
                print("DEBUG: No existing groups file")
        except Exception as e:
            print("ERROR loading groups: {}".format(str(e)))
            self.groups = {}
    
    def save_groups(self):
        """Save groups to JSON file"""
        try:
            # Convert ElementId to string for JSON serialization
            data = {}
            for group_name, sheet_ids in self.groups.items():
                data[group_name] = [str(sheet_id.IntegerValue) for sheet_id in sheet_ids]
            
            # Ensure directory exists
            dir_path = os.path.dirname(self.json_file)
            if dir_path and not os.path.exists(dir_path):
                os.makedirs(dir_path)
            
            with open(self.json_file, 'w') as f:
                json.dump(data, f, indent=2)
            
            print("DEBUG: Saved {} groups to {}".format(len(self.groups), self.json_file))
            return True
        except Exception as e:
            print("ERROR saving groups: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return False
    
    def get_all_groups(self):
        """Get all group names"""
        return sorted(self.groups.keys())
    
    def create_group(self, group_name, sheet_ids=None):
        """Create new group"""
        if group_name in self.groups:
            return False  # Group already exists
        
        self.groups[group_name] = sheet_ids or []
        self.save_groups()
        return True
    
    def rename_group(self, old_name, new_name):
        """Rename a group"""
        if old_name not in self.groups or new_name in self.groups:
            return False
        
        self.groups[new_name] = self.groups.pop(old_name)
        self.save_groups()
        return True
    
    def delete_group(self, group_name):
        """Delete a group"""
        if group_name in self.groups:
            del self.groups[group_name]
            self.save_groups()
            return True
        return False
    
    def get_sheets_in_group(self, group_name):
        """Get sheet IDs in a group"""
        return self.groups.get(group_name, [])
    
    def set_sheets_in_group(self, group_name, sheet_ids):
        """Set sheets in a group"""
        if group_name not in self.groups:
            return False
        
        self.groups[group_name] = list(sheet_ids)
        self.save_groups()
        return True
    
    def add_sheets_to_group(self, group_name, sheet_ids):
        """Add sheets to a group"""
        if group_name not in self.groups:
            return False
        
        current = set(self.groups[group_name])
        current.update(sheet_ids)
        self.groups[group_name] = list(current)
        self.save_groups()
        return True
    
    def remove_sheets_from_group(self, group_name, sheet_ids):
        """Remove sheets from a group"""
        if group_name not in self.groups:
            return False
        
        current = set(self.groups[group_name])
        current.difference_update(sheet_ids)
        self.groups[group_name] = list(current)
        self.save_groups()
        return True
    
    def is_sheet_in_group(self, group_name, sheet_id):
        """Check if sheet is in group"""
        if group_name not in self.groups:
            return False
        return sheet_id in self.groups[group_name]
