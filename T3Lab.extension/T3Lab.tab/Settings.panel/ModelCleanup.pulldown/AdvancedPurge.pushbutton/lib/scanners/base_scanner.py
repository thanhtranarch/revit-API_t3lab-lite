# -*- coding: utf-8 -*-
"""
Base Scanner for Advanced Purge
Foundation class for all scanner implementations

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

from Autodesk.Revit.DB import FilteredElementCollector


class BaseAdvancedScanner(object):
    """Base class for all advanced purge scanners"""
    
    def __init__(self, doc):
        """
        Initialize scanner
        
        Args:
            doc: Revit document
        """
        self.doc = doc
    
    def scan(self, category, progress_callback=None):
        """
        Scan for items in this category
        
        Args:
            category: AdvancedPurgeCategory object
            progress_callback: Optional callback(message)
            
        Returns:
            List of dictionaries with item info:
            {
                'id': ElementId,
                'name': string,
                'type': string,
                'category': string
            }
        """
        raise NotImplementedError("Subclasses must implement scan()")
    
    def can_delete(self, element):
        """
        Check if element can be deleted
        
        Args:
            element: Revit element
            
        Returns:
            True if element can be deleted
        """
        try:
            if not element:
                return False
            
            # Check if pinned
            if hasattr(element, 'Pinned') and element.Pinned:
                return False
            
            # Check if can be deleted
            return True
            
        except:
            return False
    
    def get_element_name(self, element):
        """
        Get element name safely
        
        Args:
            element: Revit element
            
        Returns:
            Element name or "Unnamed"
        """
        try:
            # For views, use ViewName
            if hasattr(element, 'ViewName') and element.ViewName:
                return element.ViewName
            
            # For element types, use Name
            if hasattr(element, 'FamilyName'):
                # This is likely an ElementType
                if hasattr(element, 'Name') and element.Name:
                    return element.Name
            
            # For element instances, try to build descriptive name
            if hasattr(element, 'Name') and element.Name:
                name = element.Name
                
                # If name is generic like "Floor", "Wall", try to add more info
                if name and hasattr(element, 'Category') and element.Category:
                    cat_name = element.Category.Name
                    # If name equals category name, it's generic - add ID
                    if name == cat_name or name == "":
                        # Try to get type name
                        try:
                            elem_type = self.doc.GetElement(element.GetTypeId())
                            if elem_type and hasattr(elem_type, 'Name'):
                                return "{} : {}".format(elem_type.Name, element.Id.IntegerValue)
                        except:
                            pass
                        return "{} : {}".format(cat_name, element.Id.IntegerValue)
                
                return name
            
            # Try to get type name as fallback
            try:
                elem_type = self.doc.GetElement(element.GetTypeId())
                if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
                    return "{} : {}".format(elem_type.Name, element.Id.IntegerValue)
            except:
                pass
            
            # Last resort - use category + ID
            if hasattr(element, 'Category') and element.Category:
                return "{} : {}".format(element.Category.Name, element.Id.IntegerValue)
            
            return "Unnamed : {}".format(element.Id.IntegerValue)
            
        except Exception as e:
            try:
                return "Element : {}".format(element.Id.IntegerValue)
            except:
                return "Unnamed"
    
    def get_element_type_name(self, element):
        """
        Get element type name
        
        Args:
            element: Revit element
            
        Returns:
            Type name
        """
        try:
            element_type = self.doc.GetElement(element.GetTypeId())
            if element_type:
                return self.get_element_name(element_type)
            return "Unknown Type"
        except:
            return "Unknown Type"
    
    def create_item_dict(self, element, category_name):
        """
        Create item dictionary from element
        
        Args:
            element: Revit element
            category_name: Category name
            
        Returns:
            Dictionary with item info
        """
        return {
            'id': element.Id,
            'name': self.get_element_name(element),
            'type': self.get_element_type_name(element),
            'category': category_name
        }
    
    def filter_by_delete_system(self, items, delete_system):
        """
        Filter items based on delete_system setting
        
        Args:
            items: List of item dictionaries
            delete_system: Whether to include system types
            
        Returns:
            Filtered list
        """
        if delete_system:
            return items
        
        # Filter out system types (simplified check)
        filtered = []
        for item in items:
            element = self.doc.GetElement(item['id'])
            if element:
                # Check if element is system type
                is_system = False
                try:
                    if hasattr(element, 'Category') and element.Category:
                        # System families typically have negative category IDs
                        if element.Category.Id.IntegerValue < 0:
                            is_system = True
                except:
                    pass
                
                if not is_system:
                    filtered.append(item)
        
        return filtered