# -*- coding: utf-8 -*-
__title__ = 'Go To ACC'
__author__  = 'Tay Othman, AIA'
__doc__ = """This script will open the Autodesk Construction Cloud (ACC) website for the current project in the default web browser.
            Author: Tay Othman, AIA """

# _________________________________________________________________________________________.NET imports
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import Document
from Autodesk.Revit.UI import TaskDialog
import webbrowser
import os

# _________________________________________________________________________________________Get the current version of Revit
revit_version = __revit__.Application.VersionNumber
doc = __revit__.ActiveUIDocument.Document

if revit_version == "2020":
    # Show a task dialog with the message
    TaskDialog.Show("Revit Version", "Revit Version is 2020, this tool is compatible with Revit 2022 and Newer")
elif revit_version == "2022":
    # Continue running the script
    doc = __revit__.ActiveUIDocument.Document
    hub_id = Document.GetHubId(doc)
    proj_id = Document.GetProjectId(doc)
    hub_str = hub_id[2:]
    proj_str = proj_id[2:]
    accurl = "https://acc.autodesk.com/insight/accounts/" + hub_str + "/projects/" + proj_str + "/home"
    # TaskDialog.Show("GetHubId and GetProjectId", accurl)
    webbrowser.open_new_tab(accurl)
else:
    # Continue running the script
    doc = __revit__.ActiveUIDocument.Document
    hub_id = Document.GetHubId(doc)
    proj_id = Document.GetProjectId(doc)
    hub_str = hub_id[2:]
    proj_str = proj_id[2:]
    accurl = "https://acc.autodesk.com/insight/accounts/" + hub_str + "/projects/" + proj_str + "/home"
    # TaskDialog.Show("GetHubId and GetProjectId", accurl)
    webbrowser.open_new_tab(accurl)