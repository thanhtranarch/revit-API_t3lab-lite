# -*- coding: utf-8 -*-
"""
Project Name

Set Project Name in Project Information.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__ = "Project\nName"
__doc__ = "Set Project Name in Project Information"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
from Autodesk.Revit.DB import Transaction
from pyrevit import script, forms

# ╔╦╗╔═╗╔╦╗╦ ╦╔═╗╔╦╗╔═╗
# ║║║║╣  ║ ╠═╣║ ║ ║║╚═╗
# ╩ ╩╚═╝ ╩ ╩ ╩╚═╝═╩╝╚═╝ VARIABLES
#====================================================================================================
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

# ╔═╗╦ ╦╔═╗╦═╗╔═╗  ╦  ╔═╗╦ ╦╦═╗
# ║ ║║ ║╠╣ ║╔╝╚═╗  ║  ║ ║║ ║╠╦╝
# ╚═╝╚═╝╚  ╩╩═╗╚═╝  ╩═╝╚═╝╚═╝╩╚═ MAIN
#====================================================================================================
if __name__ == '__main__':
    logger = script.get_logger()
    output = script.get_output()

    # Get current Project Name from Project Information
    project_info = doc.ProjectInformation
    current_name = project_info.Name or ""

    # Ask user for new Project Name
    new_name = forms.ask_for_string(
        default=current_name,
        prompt="Enter Project Name:",
        title="Project Name"
    )

    if new_name is None:
        # User cancelled
        script.exit()

    if new_name == current_name:
        forms.alert("Project Name unchanged.", title="Project Name")
        script.exit()

    # Set Project Name in Project Information
    with Transaction(doc, "Set Project Name") as t:
        t.Start()
        project_info.Name = new_name
        t.Commit()

    print("Project Name updated successfully.")
    print("  Old: {}".format(current_name))
    print("  New: {}".format(new_name))

    forms.alert(
        "Project Name updated successfully.\n\n"
        "Old: {}\n\n"
        "New: {}".format(current_name, new_name),
        title="Project Name"
    )
