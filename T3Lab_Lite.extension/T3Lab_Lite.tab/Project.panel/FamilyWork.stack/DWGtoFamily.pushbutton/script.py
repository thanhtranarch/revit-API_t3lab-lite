# -*- coding: utf-8 -*-
"""
DWG to Family

Convert a preselected DWG import into a Generic Model family with:
- Detail lines from all DWG line instances
- Solid extrusion from DWG outline boundary
- Family instance placed at DWG center point

Author: T3Lab
"""

from __future__ import unicode_literals

__author__  = "T3Lab"
__title__   = "DWG to Family"
__version__ = "1.0.0"

# IMPORTS
import os
import time
import traceback

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import ImportInstance
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.Exceptions import OperationCanceledException

from Utils.DWGFamilyHelpers import (
    get_dwg_filename, extract_curves_from_dwg, get_outline_boundary,
    calculate_center_point, get_dwg_import_center_point,
    create_generic_model_family, save_and_load_family
)

# ─── Revit context ────────────────────────────────────────────────────────────

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
app   = __revit__.Application

logger = script.get_logger()
output = script.get_output()


# ─── Logging adapters (keep DWGFamilyHelpers callback interface) ───────────────

def write_log(message, level="INFO"):
    """Route helper log calls to the pyrevit logger."""
    if level == "DEBUG":
        logger.debug(message)
    elif level == "WARN":
        logger.warning(message)
    elif level == "ERROR":
        logger.error(message)
    else:
        logger.info(message)


def write_error(message, exception=None):
    """Log an error and optional traceback via the pyrevit logger."""
    logger.error(message)
    if exception:
        logger.error(traceback.format_exc())


# ─── Helpers ──────────────────────────────────────────────────────────────────

def delete_temp_file(file_path, max_retries=3, delay=0.5):
    """Delete a temporary file with retry logic."""
    if not file_path or not os.path.exists(file_path):
        return True

    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            logger.info("Deleted temp family file: {}".format(file_path))
            return True
        except Exception as del_err:
            if attempt < max_retries - 1:
                logger.debug("Retry {}/{}: waiting {}s…".format(
                    attempt + 1, max_retries, delay))
                time.sleep(delay)
            else:
                logger.warning("Could not delete temp file after {} attempts: {}".format(
                    max_retries, str(del_err)))
                return False

    return False


def get_preselected_dwg():
    """Return the first ImportInstance in the current selection, or None."""
    logger.info("Getting preselected DWG import…")
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count == 0:
        logger.warning("No elements selected")
        return None

    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if isinstance(elem, ImportInstance):
            logger.info("Found DWG import: {}".format(
                elem.Name if hasattr(elem, 'Name') else "Unknown"))
            return elem

    logger.warning("No DWG ImportInstance found in selection")
    return None


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    logger.info("=" * 60)
    logger.info("Starting DWG to Family conversion")
    logger.info("=" * 60)

    try:
        # Step 1 – get preselected DWG
        dwg_instance = get_preselected_dwg()
        if not dwg_instance:
            forms.alert('Please select a DWG import instance first!', exitscript=True)
            return

        dwg_filename = get_dwg_filename(dwg_instance, doc, write_log)
        if not dwg_filename:
            dwg_filename = dwg_instance.Name if hasattr(dwg_instance, 'Name') else "Unknown_DWG"
            logger.warning("Using instance name as fallback: {}".format(dwg_filename))

        dwg_name = dwg_instance.Name if hasattr(dwg_instance, 'Name') else "Unknown_DWG"
        logger.info("Processing DWG instance: {}".format(dwg_name))
        logger.info("DWG source filename: {}".format(dwg_filename))

        dwg_location = dwg_instance.GetTotalTransform().Origin
        logger.info("DWG location: {}".format(dwg_location))

        # Step 2 – extract curves
        curves = extract_curves_from_dwg(dwg_instance, app, write_log)
        if not curves:
            forms.alert('No curves found in DWG import!', exitscript=True)
            return

        logger.info("Total curves extracted: {}".format(len(curves)))

        # Step 3 – center point
        dwg_center_point = get_dwg_import_center_point(dwg_instance, doc, write_log)
        if not dwg_center_point:
            logger.warning("Falling back to center point calculation from curves")
            dwg_center_point = calculate_center_point(curves, write_log)

        logger.info("DWG center point: ({:.3f}, {:.3f}, {:.3f})".format(
            dwg_center_point.X, dwg_center_point.Y, dwg_center_point.Z))

        # Step 4 – outline boundary
        outline_curves = get_outline_boundary(curves, write_log)
        if not outline_curves:
            forms.alert('Could not calculate outline boundary!', exitscript=True)
            return

        # Step 5 – create Generic Model family
        fam_doc, has_extrusion = create_generic_model_family(
            curves, outline_curves, app, write_log, write_error)
        if not fam_doc:
            forms.alert('Failed to create family document!', exitscript=True)
            return

        # Step 6 – save and load
        save_path, family_symbol = save_and_load_family(
            fam_doc, dwg_filename, doc, write_log, write_error)
        if not save_path:
            fam_doc.Close(False)
            forms.alert('Family creation cancelled or failed!', exitscript=True)
            return

        fam_doc.Close(False)
        time.sleep(0.1)
        delete_temp_file(save_path)

        # Step 7 – place instance
        placed_instance = None
        if family_symbol:
            try:
                logger.info("Placing family instance at DWG center point…")
                with revit.Transaction('Place Family Instance'):
                    if not family_symbol.IsActive:
                        family_symbol.Activate()
                        doc.Regenerate()
                        logger.info("Activated family symbol")

                    placed_instance = doc.Create.NewFamilyInstance(
                        dwg_center_point, family_symbol, StructuralType.NonStructural)
                    logger.info("Family instance placed at: {}".format(dwg_center_point))

            except Exception as e:
                write_error("Error placing family instance", e)

        # Result
        logger.info("=" * 60)
        logger.info("Conversion completed successfully!")
        logger.info("=" * 60)

        if placed_instance:
            placement_msg = "\n\nFamily instance placed at DWG center point in project!"
        elif family_symbol:
            placement_msg = "\n\nFamily loaded into project (placement failed – check log)"
        else:
            placement_msg = "\n\nFamily saved but not loaded into project (check log)"

        extrusion_msg  = "Created" if has_extrusion else "SKIPPED (model lines only)"
        display_name   = os.path.splitext(os.path.basename(save_path))[0]

        forms.alert(
            'DWG converted to Generic Model family!\n\n'
            'Family name: {}\n'
            'Saved to: {}\n\n'
            'Model lines: {}\n'
            'Solid extrusion: {}{}'.format(
                display_name, save_path,
                len(curves), extrusion_msg, placement_msg
            )
        )

    except OperationCanceledException:
        logger.warning("Script cancelled by user")
    except Exception as e:
        write_error("Unexpected error occurred", e)
        forms.alert('An error occurred during conversion.\n\nCheck the pyrevit log for details.',
                    exitscript=True)


if __name__ == '__main__':
    main()
