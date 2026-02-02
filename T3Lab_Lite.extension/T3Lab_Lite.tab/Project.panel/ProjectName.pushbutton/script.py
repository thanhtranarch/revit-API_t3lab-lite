# -*- coding: utf-8 -*-
"""
Project Name

Find and assign Project Name value
to Title Block on all sheets.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__ = "Project\nName"
__doc__ = "Find all Title Blocks and assign Project Name parameter value"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script

# ╔╦╗╔═╗╔╦╗╦ ╦╔═╗╔╦╗╔═╗
# ║║║║╣  ║ ╠═╣║ ║ ║║╚═╗
# ╩ ╩╚═╝ ╩ ╩ ╩╚═╝═╩╝╚═╝ VARIABLES
#====================================================================================================
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

# ╔═╗╦═╗╔═╗ ╦╔═╗╔═╗╔╦╗  ╔╗╔╔═╗╔╦╗╔═╗
# ╠═╝╠╦╝║ ║ ║║╣ ║   ║   ║║║╠═╣║║║║╣
# ╩  ╩╚═╚═╝╚╝╚═╝╚═╝ ╩   ╝╚╝╩ ╩╩ ╩╚═╝ PROJECT NAME
#====================================================================================================
PROJECT_NAME = (
    "PROPOSED NEW ERECTION OF MIXED DEVELOPMENT COMPRISING "
    "1 TOWER OF 45-STOREY SERVICE APARTMENTS (TOTAL: 241 UNITS)/"
    "RESIDENTIAL APARTMENTS (TOTAL: 246 UNITS) AND 9-STOREY PODIUM "
    "COMMERCIAL/OFFICE WITH BASEMENT CARPARKS AND COMMUNAL FACILITIES "
    "TS03 ON LOT 00595K, 00596N AND 00598L(PT) AT ANSON ROAD "
    "(DOWNTOWN CORE PLANNING AREA)."
)

# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝ FUNCTIONS
#====================================================================================================
def get_all_titleblocks(doc):
    """Get all Title Block instances in the document."""
    return FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsNotElementType()\
        .ToElements()


def find_project_name_param(titleblock):
    """Find the 'Project Name' parameter on a Title Block.
    Searches both built-in and custom/shared parameters."""
    # Try built-in PROJECT_NAME parameter first
    param = titleblock.get_Parameter(BuiltInParameter.SHEET_APPROVED_BY)
    # The built-in project name is on ProjectInfo, not on TitleBlock directly.
    # Title Blocks typically have a custom shared/family parameter called "Project Name".

    # Search all parameters for one named "Project Name"
    for p in titleblock.GetOrderedParameters():
        if p.Definition.Name == "Project Name":
            return p

    # Also try common variations
    name_variations = ["PROJECT NAME", "Project_Name", "PROJECT_NAME", "project name"]
    for p in titleblock.GetOrderedParameters():
        if p.Definition.Name in name_variations:
            return p

    return None


def set_project_name(titleblock, project_name):
    """Set the Project Name parameter value on a Title Block.
    Returns True if successful, False otherwise."""
    param = find_project_name_param(titleblock)
    if param is None:
        return False, "No 'Project Name' parameter found"

    if param.IsReadOnly:
        return False, "'Project Name' parameter is read-only"

    try:
        param.Set(project_name)
        return True, "Success"
    except Exception as e:
        return False, str(e)


# ╔═╗╦ ╦╔═╗╦═╗╔═╗  ╦  ╔═╗╦ ╦╦═╗
# ║ ║║ ║╠╣ ║╔╝╚═╗  ║  ║ ║║ ║╠╦╝
# ╚═╝╚═╝╚  ╩╩═╗╚═╝  ╩═╝╚═╝╚═╝╩╚═ MAIN
#====================================================================================================
if __name__ == '__main__':
    logger = script.get_logger()
    output = script.get_output()

    # Collect all Title Block instances
    titleblocks = get_all_titleblocks(doc)

    if not titleblocks:
        TaskDialog.Show("Project Name", "No Title Blocks found in the document.")
    else:
        print("Found {} Title Block(s) in the document.".format(len(titleblocks)))
        print("=" * 60)
        print("Assigning Project Name to all Title Blocks...")
        print("")

        success_count = 0
        fail_count = 0
        errors = []

        with Transaction(doc, "Set Project Name on Title Blocks") as t:
            t.Start()

            for tb in titleblocks:
                # Get the sheet number for reporting
                sheet_number = ""
                try:
                    owner_view = doc.GetElement(tb.OwnerViewId)
                    if owner_view:
                        sheet_number = owner_view.SheetNumber
                except:
                    sheet_number = "Unknown"

                ok, msg = set_project_name(tb, PROJECT_NAME)
                if ok:
                    success_count += 1
                    print("  [OK] Sheet {}: Project Name assigned".format(sheet_number))
                else:
                    fail_count += 1
                    errors.append("Sheet {}: {}".format(sheet_number, msg))
                    print("  [FAIL] Sheet {}: {}".format(sheet_number, msg))

            t.Commit()

        # Summary
        print("")
        print("=" * 60)
        print("SUMMARY")
        print("=" * 60)
        print("Total Title Blocks: {}".format(len(titleblocks)))
        print("Successfully updated: {}".format(success_count))
        print("Failed: {}".format(fail_count))

        if errors:
            print("")
            print("Errors:")
            for err in errors:
                print("  - {}".format(err))

        # Show result dialog
        if fail_count == 0:
            TaskDialog.Show(
                "Project Name",
                "Successfully assigned Project Name to all {} Title Block(s).".format(success_count)
            )
        else:
            TaskDialog.Show(
                "Project Name",
                "Updated {}/{} Title Block(s).\n\n{} failed - see output for details.".format(
                    success_count, len(titleblocks), fail_count
                )
            )
