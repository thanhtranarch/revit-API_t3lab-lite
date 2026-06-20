# -*- coding: utf-8 -*-
"""
Revit Utilities - Helper functions for Revit API compatibility
Compatible with Revit 2024, 2025, 2026, 2027

Copyright (c) 2025 Dang Quoc Truong (DQT)

IMPORTANT COMPATIBILITY NOTES:
- Revit 2024/2025: ElementId.IntegerValue works
- Revit 2026+: ElementId.IntegerValue deprecated, use ElementId.Value instead
- This module provides helper functions that work across all versions
"""

__author__ = "Dang Quoc Truong (DQT)"


def get_element_id_value(element_id):
    """
    Get the integer value of an ElementId - works for Revit 2024-2027
    
    Revit 2024/2025: Uses IntegerValue
    Revit 2026+: Uses Value (IntegerValue deprecated)
    
    Args:
        element_id: An ElementId object
        
    Returns:
        int: The integer value of the ElementId
    """
    if element_id is None:
        return -1
    
    try:
        # Try Revit 2026+ method first (.Value)
        if hasattr(element_id, 'Value'):
            return element_id.Value
    except:
        pass
    
    try:
        # Fallback to Revit 2024/2025 method (.IntegerValue)
        if hasattr(element_id, 'IntegerValue'):
            return element_id.IntegerValue
    except:
        pass
    
    # Last resort - try to convert to int
    try:
        return int(str(element_id))
    except:
        return -1


def _eid_int(element_id):
    """
    Shorthand alias for get_element_id_value
    
    Usage:
        from revit_utils import _eid_int
        id_value = _eid_int(element.Id)
    """
    return get_element_id_value(element_id)


def get_revit_version():
    """
    Get the current Revit version number
    
    Returns:
        int: Revit version year (e.g., 2024, 2025, 2026, 2027)
    """
    try:
        from Autodesk.Revit.ApplicationServices import Application
        # This won't work directly, need to get from __revit__
        return 2024  # Default fallback
    except:
        return 2024


def get_revit_version_from_doc(doc):
    """
    Get Revit version from document
    
    Args:
        doc: Revit Document object
        
    Returns:
        int: Revit version year
    """
    try:
        app = doc.Application
        version_str = app.VersionNumber
        return int(version_str)
    except:
        return 2024


def is_revit_2026_or_newer(doc=None):
    """
    Check if running Revit 2026 or newer
    
    Args:
        doc: Optional Revit Document object
        
    Returns:
        bool: True if Revit 2026+
    """
    try:
        if doc:
            version = get_revit_version_from_doc(doc)
        else:
            version = 2024
        return version >= 2026
    except:
        return False


def safe_get_parameter_value(element, param_name):
    """
    Safely get parameter value from element
    
    Args:
        element: Revit Element
        param_name: Parameter name string
        
    Returns:
        Parameter value or None
    """
    try:
        param = element.LookupParameter(param_name)
        if param and param.HasValue:
            storage_type = param.StorageType
            
            # Import StorageType
            from Autodesk.Revit.DB import StorageType
            
            if storage_type == StorageType.String:
                return param.AsString()
            elif storage_type == StorageType.Integer:
                return param.AsInteger()
            elif storage_type == StorageType.Double:
                return param.AsDouble()
            elif storage_type == StorageType.ElementId:
                return param.AsElementId()
        return None
    except:
        return None


def safe_delete_elements(doc, element_ids, transaction_name="DQT - Delete Elements"):
    """
    Safely delete elements with proper error handling
    
    Args:
        doc: Revit Document
        element_ids: List of ElementId objects to delete
        transaction_name: Name for the transaction
        
    Returns:
        tuple: (deleted_count, failed_count, failed_ids)
    """
    from Autodesk.Revit.DB import Transaction
    
    deleted = 0
    failed = 0
    failed_ids = []
    
    if not element_ids:
        return (0, 0, [])
    
    t = Transaction(doc, transaction_name)
    t.Start()
    
    try:
        for eid in element_ids:
            try:
                doc.Delete(eid)
                deleted += 1
            except Exception as e:
                failed += 1
                failed_ids.append((eid, str(e)))
        
        t.Commit()
    except Exception as e:
        t.RollBack()
        return (0, len(element_ids), [(eid, "Transaction failed") for eid in element_ids])
    
    return (deleted, failed, failed_ids)


def format_element_id(element_id):
    """
    Format ElementId for display
    
    Args:
        element_id: ElementId object
        
    Returns:
        str: Formatted ID string
    """
    return str(get_element_id_value(element_id))


# Compatibility shims for deprecated APIs

def get_builtin_parameter_group(group_name):
    """
    Get BuiltInParameterGroup or GroupTypeId based on Revit version
    
    Revit 2024: BuiltInParameterGroup
    Revit 2025+: GroupTypeId
    
    Args:
        group_name: Name of the parameter group
        
    Returns:
        BuiltInParameterGroup or GroupTypeId
    """
    try:
        # Try new API first (Revit 2025+)
        from Autodesk.Revit.DB import GroupTypeId
        return getattr(GroupTypeId, group_name, None)
    except:
        pass
    
    try:
        # Fallback to old API (Revit 2024)
        from Autodesk.Revit.DB import BuiltInParameterGroup
        return getattr(BuiltInParameterGroup, group_name, None)
    except:
        return None


def get_spec_type_id(type_name):
    """
    Get SpecTypeId or ParameterType based on Revit version
    
    Revit 2024: ParameterType
    Revit 2025+: SpecTypeId
    
    Args:
        type_name: Name of the parameter type
        
    Returns:
        ParameterType or SpecTypeId
    """
    try:
        # Try new API first (Revit 2025+)
        from Autodesk.Revit.DB import SpecTypeId
        return getattr(SpecTypeId, type_name, None)
    except:
        pass
    
    try:
        # Fallback to old API (Revit 2024)
        from Autodesk.Revit.DB import ParameterType
        return getattr(ParameterType, type_name, None)
    except:
        return None
