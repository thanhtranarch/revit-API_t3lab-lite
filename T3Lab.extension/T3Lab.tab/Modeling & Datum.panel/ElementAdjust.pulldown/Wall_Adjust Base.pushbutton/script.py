# -*- coding: utf-8 -*-
"""Auto Adjust Base Offset
Automatically adjusts Base Offset when changing Base Constraint to maintain element position.
Supports: Walls, Floors, Columns, Structural Columns, Beams, Structural Framing

Copyright (c) 2025 by Dang Quoc Truong (DQT)
"""

__title__ = "Auto Adjust\nBase Offset"
__author__ = "DQT"
__doc__ = """Automatically adjusts Base Offset when changing Base Constraint
to maintain element position and height.

Supports: Walls, Floors, Columns, Structural Columns, Beams, Structural Framing

Usage:
1. Select elements to adjust
2. Run this tool
3. Select new Base Constraint from levels list
4. Tool will automatically calculate and update Base Offset

Copyright (c) 2025 by Dang Quoc Truong (DQT)"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, forms, script
import sys

from Snippets._compat import eid_value

# Get current document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_all_levels():
    """Get all levels in the project sorted by elevation"""
    collector = FilteredElementCollector(doc).OfClass(Level)
    levels = list(collector)
    # Sort by elevation
    levels.sort(key=lambda x: x.Elevation)
    return levels

def get_level_elevation(level):
    """Get level elevation in project units"""
    return level.Elevation

def is_column(elem):
    """Check if element is a column (Architectural or Structural)"""
    if not isinstance(elem, FamilyInstance):
        return False
    
    try:
        category = elem.Category
        if not category:
            return False
        
        cat_id = eid_value(category.Id)
        # Check if it's Structural Columns or Columns category
        return (cat_id == BuiltInCategory.OST_StructuralColumns.value__ or 
                cat_id == BuiltInCategory.OST_Columns.value__)
    except:
        return False

def is_beam(elem):
    """Check if element is a beam or structural framing"""
    if not isinstance(elem, FamilyInstance):
        return False
    
    try:
        category = elem.Category
        if not category:
            return False
        
        cat_id = eid_value(category.Id)
        # Check if it's Structural Framing category
        return cat_id == BuiltInCategory.OST_StructuralFraming.value__
    except:
        return False

def get_element_type_name(elem):
    """Get element type name safely"""
    try:
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            return Element.Name.GetValue(elem_type)
    except:
        pass
    return "Unknown"

def get_element_category_name(elem):
    """Get element category name"""
    try:
        if elem.Category:
            return elem.Category.Name
    except:
        pass
    return "Unknown"

def get_element_base_info(elem):
    """Get element's current base constraint and offset information
    Works for Wall, Floor, Column, and Beam (Structural Framing)"""
    try:
        base_constraint_param = None
        base_offset_param = None
        
        # Wall
        if isinstance(elem, Wall):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
            base_offset_param = elem.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        
        # Floor
        elif isinstance(elem, Floor):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM)
            base_offset_param = elem.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
        
        # Column (Architectural or Structural)
        elif is_column(elem):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
            base_offset_param = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
        
        # Beam (Structural Framing)
        elif is_beam(elem):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
            base_offset_param = elem.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION)
        
        else:
            return None, None, None
        
        if not base_constraint_param or not base_offset_param:
            return None, None, None
        
        base_level_id = base_constraint_param.AsElementId()
        if base_level_id == ElementId.InvalidElementId:
            return None, None, None
            
        base_level = doc.GetElement(base_level_id)
        base_offset = base_offset_param.AsDouble() if base_offset_param else 0.0
        
        # Calculate actual elevation (absolute position)
        base_elevation = get_level_elevation(base_level)
        actual_elevation = base_elevation + base_offset
        
        return base_level, base_offset, actual_elevation
    except Exception as e:
        print("Error getting element base info: {}".format(str(e)))
        return None, None, None

def adjust_element_base_constraint(elem, new_level):
    """Adjust element's base constraint and offset to maintain position
    Works for Wall, Floor, Column, and Beam (Structural Framing)"""
    try:
        # Get current base information
        current_level, current_offset, actual_elevation = get_element_base_info(elem)
        
        if current_level is None or actual_elevation is None:
            print("Cannot get current element information: {}".format(elem.Id))
            return False
        
        # Calculate new offset needed
        new_level_elevation = get_level_elevation(new_level)
        new_offset = actual_elevation - new_level_elevation
        
        # Get parameters based on element type
        base_constraint_param = None
        base_offset_param = None
        
        # Wall
        if isinstance(elem, Wall):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
            base_offset_param = elem.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
        
        # Floor
        elif isinstance(elem, Floor):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM)
            base_offset_param = elem.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
        
        # Column (Architectural or Structural)
        elif is_column(elem):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
            base_offset_param = elem.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
        
        # Beam (Structural Framing)
        elif is_beam(elem):
            base_constraint_param = elem.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
            base_offset_param = elem.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION)
        
        if not base_constraint_param or not base_offset_param:
            print("Cannot access element parameters: {}".format(elem.Id))
            return False
        
        # Update Base Constraint
        base_constraint_param.Set(new_level.Id)
        
        # Update Base Offset
        base_offset_param.Set(new_offset)
        
        return True
    except Exception as e:
        print("Error adjusting element {}: {}".format(elem.Id, str(e)))
        return False

def main():
    """Main execution function"""
    
    try:
        # Get selected elements or prompt user to select
        selection = uidoc.Selection.GetElementIds()
        elements = []
        
        if selection.Count > 0:
            # Filter selected elements to get only walls, floors, columns, beams
            for elem_id in selection:
                elem = doc.GetElement(elem_id)
                if isinstance(elem, (Wall, Floor)) or is_column(elem) or is_beam(elem):
                    elements.append(elem)
        
        if not elements:
            # Prompt user to select elements
            try:
                from Autodesk.Revit.UI.Selection import ISelectionFilter
                
                class ElementSelectionFilter(ISelectionFilter):
                    def AllowElement(self, elem):
                        return isinstance(elem, (Wall, Floor)) or is_column(elem) or is_beam(elem)
                    
                    def AllowReference(self, reference, position):
                        return False
                
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    ElementSelectionFilter(),
                    "Select elements to adjust Base Constraint (Walls, Floors, Columns, Beams)"
                )
                
                elements = [doc.GetElement(ref.ElementId) for ref in refs]
                
            except Exception as e:
                forms.alert("No elements selected!", exitscript=True)
                return
        
        if not elements:
            forms.alert("No elements selected!", exitscript=True)
            return
        
        # Get all levels
        all_levels = get_all_levels()
        if not all_levels:
            forms.alert("No levels found in project!", exitscript=True)
            return
        
        # Create level selection options with display names
        class LevelOption:
            def __init__(self, level):
                self.level = level
                elevation_mm = level.Elevation * 304.8  # Convert to mm for display
                self.name = "{} (Elevation: {:.0f}mm)".format(level.Name, elevation_mm)
        
        level_options = [LevelOption(level) for level in all_levels]
        
        selected_option = forms.SelectFromList.show(
            level_options,
            title="Select New Base Constraint",
            button_name="Apply",
            name_attr='name',
            multiselect=False
        )
        
        if not selected_option:
            script.exit()
            return
        
        new_level = selected_option.level
        
        # Process elements in a transaction
        success_count = 0
        error_count = 0
        
        t = Transaction(doc, "Auto Adjust Base Constraint and Offset")
        t.Start()
        
        try:
            for elem in elements:
                try:
                    # Get current info
                    current_level, current_offset, actual_elevation = get_element_base_info(elem)
                    
                    if current_level is None:
                        error_count += 1
                        continue
                    
                    # Apply changes
                    if adjust_element_base_constraint(elem, new_level):
                        success_count += 1
                    else:
                        error_count += 1
                        
                except Exception as elem_error:
                    error_count += 1
            
            t.Commit()
            
        except Exception as e:
            t.RollBack()
            forms.alert("Error occurred while adjusting elements!\n\n{}".format(str(e)))
    
    except Exception as main_error:
        forms.alert("Error in main execution!\n\n{}".format(str(main_error)))

if __name__ == "__main__":
    main()