# -*- coding: utf-8 -*-
"""
Dangerous Operations Scanner
Scans for extremely dangerous cleanup operations

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
    Group
)


class DangerousOpsScanner(BaseAdvancedScanner):
    """Scanner for dangerous operations"""
    
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
        if category_id == "area_separation_lines":
            return self.scan_area_separation_lines(category)
        elif category_id == "all_groups":
            return self.scan_all_groups(category)
        else:
            return []
    
    def scan_area_separation_lines(self, category):
        """Scan for all area separation lines"""
        items = []
        
        try:
            collector = FilteredElementCollector(self.doc)
            area_sep_lines = collector.OfCategory(
                BuiltInCategory.OST_AreaSchemeLines
            ).WhereElementIsNotElementType().ToElements()
            
            for line in area_sep_lines:
                if self.can_delete(line):
                    items.append(self.create_item_dict(line, category.name))
                    
        except Exception as e:
            print("Error scanning area separation lines: {}".format(str(e)))
        
        return items
    
    def scan_all_groups(self, category):
        """Scan for all groups in the model"""
        items = []
        
        try:
            collector = FilteredElementCollector(self.doc)
            groups = collector.OfClass(Group).ToElements()
            
            for group in groups:
                if self.can_delete(group):
                    items.append(self.create_item_dict(group, category.name))
                    
        except Exception as e:
            print("Error scanning groups: {}".format(str(e)))
        
        return items