# -*- coding: utf-8 -*-
"""
Collaboration Scanner
Scans for collaboration and BIM360 cleanup items

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

from Autodesk.Revit.DB import FilteredElementCollector, Category


class CollaborationScanner(BaseAdvancedScanner):
    """Scanner for collaboration cleanup operations"""
    
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
        if category_id == "bim360_cache":
            return self.scan_bim360_cache(category)
        elif category_id == "data_schema":
            return self.scan_data_schema(category)
        elif category_id == "unused_subcategories":
            return self.scan_unused_subcategories(category)
        elif category_id == "view_specific_constraints":
            return self.scan_view_specific_constraints(category)
        else:
            return []
    
    def scan_bim360_cache(self, category):
        """
        Scan for BIM360 collaboration cache
        
        NOTE: This is not a deletable element - it's a command to clear cache
        This would require special handling in executor
        """
        items = []
        
        # BIM360 cache clearing is not about deleting elements
        # It's about clearing cloud sync cache files
        # This needs special implementation in executor
        # For now, return empty
        
        print("INFO: BIM360 cache clearing requires special executor implementation")
        return items
    
    def scan_data_schema(self, category):
        """
        Scan for unused data schema
        
        NOTE: Data schemas are complex - they're stored in ExtensibleStorage
        This requires careful implementation to avoid breaking functionality
        """
        items = []
        
        try:
            from Autodesk.Revit.DB import ExtensibleStorage
            
            # Getting schema is complex and dangerous
            # ExtensibleStorage.Schema operations can break add-ins
            # For now, return empty - needs more research
            
            print("INFO: Data schema cleanup requires careful implementation")
                    
        except Exception as e:
            print("Error scanning data schema: {}".format(str(e)))
        
        return items
    
    def scan_unused_subcategories(self, category):
        """
        Scan for unused subcategories
        
        A subcategory is unused if no elements use it
        """
        items = []
        
        try:
            # Get all categories
            categories = self.doc.Settings.Categories
            
            # Collect all elements to build usage set
            collector = FilteredElementCollector(self.doc)
            all_elements = collector.WhereElementIsNotElementType().ToElements()
            
            # Build set of used subcategory ids
            used_subcategory_ids = set()
            for elem in all_elements:
                if elem.Category and elem.Category.Id:
                    used_subcategory_ids.add(elem.Category.Id.IntegerValue)
            
            # Check each category for unused subcategories
            for cat in categories:
                if cat and cat.SubCategories:
                    for subcat in cat.SubCategories:
                        if subcat.Id.IntegerValue not in used_subcategory_ids:
                            # This subcategory is unused
                            # Note: Subcategories are NOT elements, they're Category objects
                            # We can't delete them like regular elements
                            # This needs special handling in executor
                            
                            # Create a pseudo-item for display
                            item = {
                                'name': "{} > {}".format(cat.Name, subcat.Name),
                                'type': 'SubCategory',
                                'id': str(subcat.Id.IntegerValue),
                                'category': category.name,
                                'element': None  # No actual element
                            }
                            items.append(item)
                    
        except Exception as e:
            print("Error scanning unused subcategories: {}".format(str(e)))
        
        return items
    
    def scan_view_specific_constraints(self, category):
        """
        Scan for view-specific constraints
        
        NOTE: Similar to constraints in Model Deep - these are not standalone elements
        """
        items = []
        
        # View-specific constraints are properties of elements in specific views
        # They're not collectible as standalone elements
        # This requires iteration through all views and their elements
        # Very complex - TODO
        
        print("INFO: View-specific constraints scanning requires complex implementation")
        return items