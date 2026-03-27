# -*- coding: utf-8 -*-
"""
DWG2FAMILY - Convert DWG Import to Generic Model Family
Converts a preselected DWG import into a Generic Model family with:
- Detail lines from all DWG line instances
- Solid extrusion from DWG outline boundary
"""

import os
import datetime
import traceback
import time

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.Exceptions import OperationCanceledException

# Import helper functions
from Utils.DWGFamilyHelpers import (
    get_dwg_filename, extract_curves_from_dwg, get_outline_boundary,
    calculate_center_point, get_dwg_import_center_point,
    create_generic_model_family, save_and_load_family
)

# Initialize
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = __revit__.Application

# ==================== DEBUG LOGGING ====================

SCRIPT_NAME = os.path.splitext(os.path.basename(__file__))[0].replace('_script', '')
LOG_FILE = os.path.join(os.path.dirname(__file__), '{}_log.txt'.format(SCRIPT_NAME))

def init_log():
    """Initialize log file (clear previous run)"""
    with open(LOG_FILE, 'w') as f:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write("="*60 + "\n")
        f.write("{} - Script Started\n".format(timestamp))
        f.write("="*60 + "\n\n")

def write_log(message, level="INFO"):
    """Write message to log file with timestamp and level"""
    safe_message = message.encode('utf-8', errors='ignore').decode('utf-8')
    with open(LOG_FILE, 'a') as f:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write("[{}] [{}] {}\n".format(timestamp, level, safe_message))

def write_error(message, exception=None):
    """Write error with full traceback"""
    write_log(message, "ERROR")
    if exception:
        tb = traceback.format_exc()
        safe_tb = tb.encode('utf-8', errors='ignore').decode('utf-8')
        with open(LOG_FILE, 'a') as f:
            f.write(safe_tb + "\n")

# ==================== HELPER FUNCTIONS ====================

def delete_temp_file(file_path, max_retries=3, delay=0.5):
    """Delete temporary file with retry logic"""
    if not file_path or not os.path.exists(file_path):
        return True
    
    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            write_log("Deleted temp family file: {}".format(file_path))
            return True
        except Exception as del_err:
            if attempt < max_retries - 1:
                write_log("Retry {}/{}: Waiting {}s before retry...".format(
                    attempt + 1, max_retries, delay), "DEBUG")
                time.sleep(delay)
            else:
                write_log("Warning: Could not delete temp file after {} attempts: {}".format(
                    max_retries, str(del_err)), "WARN")
                return False
    
    return False

def get_preselected_dwg():
    """Get preselected DWG ImportInstance from current selection"""
    write_log("Getting preselected DWG import...")
    selection = uidoc.Selection
    selected_ids = selection.GetElementIds()

    if selected_ids.Count == 0:
        write_log("No elements selected", "WARN")
        return None

    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if isinstance(elem, ImportInstance):
            dwg_name = elem.Name if hasattr(elem, 'Name') else "Unknown"
            write_log("Found DWG import: {}".format(dwg_name))
            return elem

    write_log("No DWG import found in selection", "WARN")
    return None

# ==================== MAIN SCRIPT ====================

def main():
    """Main script execution"""
    init_log()

    try:
        write_log("="*60)
        write_log("Starting DWG to Family conversion")
        write_log("="*60)

        # Step 1: Get preselected DWG
        dwg_instance = get_preselected_dwg()
        if not dwg_instance:
            forms.alert('Please select a DWG import instance first!', exitscript=True)
            return

        # Get actual DWG filename
        dwg_filename = get_dwg_filename(dwg_instance, doc, write_log)
        if not dwg_filename:
            dwg_filename = dwg_instance.Name if hasattr(dwg_instance, 'Name') else "Unknown_DWG"
            write_log("Using instance name as fallback: {}".format(dwg_filename), "WARN")
        
        dwg_name = dwg_instance.Name if hasattr(dwg_instance, 'Name') else "Unknown_DWG"
        write_log("Processing DWG instance: {}".format(dwg_name))
        write_log("DWG source filename: {}".format(dwg_filename))

        # Get DWG location for family placement
        dwg_transform = dwg_instance.GetTotalTransform()
        dwg_location = dwg_transform.Origin
        write_log("DWG location: {}".format(dwg_location))

        # Step 2: Extract curves from DWG
        curves = extract_curves_from_dwg(dwg_instance, app, write_log)
        if not curves:
            forms.alert('No curves found in DWG import!\n\nCheck {} for details.'.format(LOG_FILE), exitscript=True)
            return

        write_log("Total curves extracted: {}".format(len(curves)))

        # Step 3: Get center point of DWG import instance
        dwg_center_point = get_dwg_import_center_point(dwg_instance, doc, write_log)
        if not dwg_center_point:
            write_log("Falling back to center point calculation from curves", "WARN")
            dwg_center_point = calculate_center_point(curves, write_log)
        
        write_log("DWG center point for instance placement: ({:.3f}, {:.3f}, {:.3f})".format(
            dwg_center_point.X, dwg_center_point.Y, dwg_center_point.Z))

        # Step 4: Calculate outline boundary
        outline_curves = get_outline_boundary(curves, write_log)
        if not outline_curves:
            forms.alert('Could not calculate outline boundary!\n\nCheck {} for details.'.format(LOG_FILE), exitscript=True)
            return

        # Step 5: Create Generic Model family
        fam_doc, has_extrusion = create_generic_model_family(curves, outline_curves, app, write_log, write_error)
        if not fam_doc:
            forms.alert('Failed to create family document!\n\nCheck {} for details.'.format(LOG_FILE), exitscript=True)
            return

        # Step 6: Save and load family into project
        save_path, family_symbol = save_and_load_family(fam_doc, dwg_filename, doc, write_log, write_error)
        if not save_path:
            fam_doc.Close(False)
            forms.alert('Family creation cancelled or failed!\n\nCheck {} for details.'.format(LOG_FILE), exitscript=True)
            return

        # Close family document (already saved)
        fam_doc.Close(False)
        
        # Delete temp file after closing family document
        time.sleep(0.1)
        delete_temp_file(save_path)

        # Step 7: Place family instance at DWG center point
        placed_instance = None
        if family_symbol:
            try:
                write_log("Placing family instance at DWG center point...")

                with revit.Transaction('Place Family Instance'):
                    if not family_symbol.IsActive:
                        family_symbol.Activate()
                        doc.Regenerate()
                        write_log("Activated family symbol")

                    placed_instance = doc.Create.NewFamilyInstance(
                        dwg_center_point,
                        family_symbol,
                        StructuralType.NonStructural
                    )
                    write_log("Family instance placed at DWG center point: {}".format(dwg_center_point))

            except Exception as e:
                write_error("Error placing family instance", e)

        # Success message
        write_log("="*60)
        write_log("Conversion completed successfully!")
        write_log("="*60)

        placement_msg = ""
        if placed_instance:
            placement_msg = "\n\nFamily instance placed at DWG center point in project!"
        elif family_symbol:
            placement_msg = "\n\nFamily loaded into project (placement failed - check log)"
        else:
            placement_msg = "\n\nFamily saved but not loaded into project (check log)"

        extrusion_msg = "Created" if has_extrusion else "SKIPPED (model lines only - check log)"
        display_family_name = os.path.splitext(os.path.basename(save_path))[0]

        forms.alert(
            'DWG converted to Generic Model family!\n\n'
            'Family name: {}\n'
            'Saved to: {}\n\n'
            'Model lines: {}\n'
            'Solid extrusion: {}{}'.format(
                display_family_name,
                save_path,
                len(curves),
                extrusion_msg,
                placement_msg
            )
        )

    except OperationCanceledException:
        write_log("Script cancelled by user", "WARN")
    except Exception as e:
        write_error("Unexpected error occurred", e)
        forms.alert(
            'An error occurred during conversion.\n\n'
            'Check {} for details.'.format(LOG_FILE),
            exitscript=True
        )

# Run main script
if __name__ == '__main__':
    main()
