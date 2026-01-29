# -*- coding: utf-8 -*-
"""
Sheet Manager - Custom Parameters Service

Copyright © Dang Quoc Truong (DQT)
"""

from Autodesk.Revit.DB import (FilteredElementCollector, ViewSheet)


class CustomParametersService(object):
    """Manage custom sheet parameters"""
    
    def __init__(self, doc, app):
        self.doc = doc
        self.app = app
    
    def get_all_sheet_parameters(self):
        """Get all parameters available on sheets"""
        try:
            # Get a sample sheet
            sheets = FilteredElementCollector(self.doc).OfClass(ViewSheet).WhereElementIsNotElementType()
            sample_sheet = None
            
            for sheet in sheets:
                if not sheet.IsPlaceholder:
                    sample_sheet = sheet
                    break
            
            if not sample_sheet:
                return []
            
            # Get all parameters
            parameters = []
            for param in sample_sheet.Parameters:
                param_info = {
                    'name': param.Definition.Name,
                    'type': str(param.StorageType),
                    'is_read_only': param.IsReadOnly,
                    'is_shared': param.IsShared,
                    'group': str(param.Definition.ParameterGroup)
                }
                parameters.append(param_info)
            
            # Sort alphabetically by name
            parameters.sort(key=lambda x: x['name'].lower())
            
            return parameters
            
        except Exception as e:
            print("Error getting sheet parameters: {}".format(str(e)))
            return []
    
    def get_parameter_values(self, sheet, param_name):
        """Get parameter value from a sheet"""
        try:
            param = sheet.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType.ToString() == "String":
                    return param.AsString()
                elif param.StorageType.ToString() == "Integer":
                    return param.AsInteger()
                elif param.StorageType.ToString() == "Double":
                    return param.AsDouble()
                else:
                    return param.AsValueString()
            return None
        except:
            return None
    
    def set_parameter_value(self, sheet, param_name, value):
        """Set parameter value on a sheet"""
        from Autodesk.Revit.DB import StorageType
        
        try:
            param = sheet.LookupParameter(param_name)
            if param and not param.IsReadOnly:
                # Use enum comparison, not string
                if param.StorageType == StorageType.String:
                    param.Set(str(value))
                    print("DEBUG: Set String param '{}' = '{}'".format(param_name, value))
                elif param.StorageType == StorageType.Integer:
                    param.Set(int(value))
                    print("DEBUG: Set Integer param '{}' = '{}'".format(param_name, value))
                elif param.StorageType == StorageType.Double:
                    param.Set(float(value))
                    print("DEBUG: Set Double param '{}' = '{}'".format(param_name, value))
                else:
                    print("DEBUG: Unsupported StorageType for '{}'".format(param_name))
                    return False
                return True
            else:
                if not param:
                    print("DEBUG: Parameter '{}' not found".format(param_name))
                elif param.IsReadOnly:
                    print("DEBUG: Parameter '{}' is read-only".format(param_name))
                return False
        except Exception as e:
            print("Error setting parameter '{}': {}".format(param_name, str(e)))
            import traceback
            traceback.print_exc()
            return False
    
    def bulk_update_parameter(self, sheets, param_name, value):
        """Update parameter on multiple sheets"""
        try:
            success_count = 0
            for sheet in sheets:
                if self.set_parameter_value(sheet, param_name, value):
                    success_count += 1
            return success_count
        except Exception as e:
            print("Error in bulk update: {}".format(str(e)))
            return 0
    
    def create_parameter_template(self, name, param_values):
        """Save a parameter template for reuse
        
        param_values: dict of {param_name: value}
        """
        try:
            # Store as simple dict (could be saved to file)
            template = {
                'name': name,
                'parameters': param_values
            }
            return template
        except Exception as e:
            print("Error creating template: {}".format(str(e)))
            return None
    
    def apply_parameter_template(self, sheets, template):
        """Apply a parameter template to sheets"""
        try:
            results = {}
            for param_name, value in template['parameters'].items():
                count = self.bulk_update_parameter(sheets, param_name, value)
                results[param_name] = count
            return results
        except Exception as e:
            print("Error applying template: {}".format(str(e)))
            return None