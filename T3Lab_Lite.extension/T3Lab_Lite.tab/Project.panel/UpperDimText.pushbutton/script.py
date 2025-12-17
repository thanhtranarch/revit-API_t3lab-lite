# -*- coding: utf-8 -*-
"""
UpperText

Adjust Dimension Text with Enhanced Control

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Upper Text"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# DEFINE VARIABLES
# ==================================================
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
active_view_id = uidoc.ActiveView.Id

# TEXT CONFIGURATION - Easy to modify
# ==================================================
# This script will convert ALL existing text in dimensions to UPPERCASE
# Only processes dimensions in CURRENT VIEW

# Optional: Set to True if you want to turn off leaders
TURN_OFF_LEADERS = False

# CLASS/FUNCTIONS
# ==================================================
def turn_off_leader(dim):
    """Turn off leader for dimension"""
    try:
        ord_para_list = dim.GetOrderedParameters()
        for para in ord_para_list:
            if para.Definition.Name == "Leader":
                para.Set(0)
                return True
    except Exception as e:
        print("Error turning off leader: {}".format(str(e)))
        return False

def convert_text_to_uppercase(text):
    """Convert text to uppercase, handle None values"""
    if text is None or text == "":
        return ""
    return str(text).upper()

def update_dimension_text_to_uppercase(dim):
    """
    Update all dimension text properties to uppercase
    
    Args:
        dim: Dimension element
    """
    try:
        if dim.HasOneSegment():
            # Single segment dimension - convert existing text to uppercase
            current_above = dim.Above if dim.Above else ""
            current_below = dim.Below if dim.Below else ""
            current_prefix = dim.Prefix if dim.Prefix else ""
            current_suffix = dim.Suffix if dim.Suffix else ""
            current_override = dim.ValueOverride if dim.ValueOverride else ""
            
            # Set uppercase versions
            dim.Above = convert_text_to_uppercase(current_above)
            dim.Below = convert_text_to_uppercase(current_below)
            dim.Prefix = convert_text_to_uppercase(current_prefix)
            dim.Suffix = convert_text_to_uppercase(current_suffix)
            if current_override:
                dim.ValueOverride = convert_text_to_uppercase(current_override)
        else:
            # Multi-segment dimension
            for segment in dim.Segments:
                current_above = segment.Above if segment.Above else ""
                current_below = segment.Below if segment.Below else ""
                current_prefix = segment.Prefix if segment.Prefix else ""
                current_suffix = segment.Suffix if segment.Suffix else ""
                current_override = segment.ValueOverride if segment.ValueOverride else ""
                
                # Set uppercase versions
                segment.Above = convert_text_to_uppercase(current_above)
                segment.Below = convert_text_to_uppercase(current_below)
                segment.Prefix = convert_text_to_uppercase(current_prefix)
                segment.Suffix = convert_text_to_uppercase(current_suffix)
                if current_override:
                    segment.ValueOverride = convert_text_to_uppercase(current_override)
        return True
    except Exception as e:
        print("Error updating dimension {}: {}".format(dim.Id, str(e)))
        return False

def get_all_dimensions_in_project():
    """Get all dimension elements from entire project"""
    # Create a FilteredElementCollector to get all dimensions in project
    collector = FilteredElementCollector(doc)
    dimensions = collector.OfClass(Dimension).ToElements()
    
    return list(dimensions)

def get_dimensions_in_current_view():
    """Get all dimension elements in current active view only"""
    # Create a FilteredElementCollector for active view
    collector = FilteredElementCollector(doc, active_view_id)
    dimensions = collector.OfClass(Dimension).ToElements()
    
    return list(dimensions)

def show_result_dialog(success_count, total_count, view_name=""):
    """Show result dialog with processing summary"""
    if success_count > 0:
        message = "Successfully converted {} out of {} dimensions to UPPERCASE in view '{}'!".format(success_count, total_count, view_name)
        if success_count < total_count:
            message += "\n{} dimensions had errors.".format(total_count - success_count)
        TaskDialog.Show("Convert to Uppercase", message)
    else:
        TaskDialog.Show("Convert to Uppercase", "No dimensions were converted in view '{}'.".format(view_name))

# MAIN SCRIPT
# ==================================================
def main():
    """Main execution function"""
    with Transaction(doc, "Convert Dimension Text to Uppercase") as transaction:
        transaction.Start()
        
        # Get dimensions in current view only
        current_view_dimensions = get_dimensions_in_current_view()
        
        if not current_view_dimensions:
            current_view_name = uidoc.ActiveView.Name
            TaskDialog.Show("Convert to Uppercase", "No dimensions found in current view '{}'!".format(current_view_name))
            transaction.RollBack()
            return
        
        success_count = 0
        total_count = len(current_view_dimensions)
        current_view_name = uidoc.ActiveView.Name
        
        # Process each dimension in current view
        for dim in current_view_dimensions:
            # Optional: Turn off leader
            if TURN_OFF_LEADERS:
                turn_off_leader(dim)
            
            # Convert all dimension text to uppercase
            if update_dimension_text_to_uppercase(dim):
                success_count += 1
        
        # Show results
        show_result_dialog(success_count, total_count, current_view_name)
        
        # Commit transaction
        transaction.Commit()

# Execute main function
if __name__ == "__main__":
    main()