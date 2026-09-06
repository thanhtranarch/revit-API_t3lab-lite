# -*- coding: utf-8 -*-
"""
Revisions Snippets

Code snippets for managing Revit revisions and revision clouds.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Revisions Snippets"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
import traceback

from Autodesk.Revit.DB import (RevisionNumberType,
                               Revision,
                               ViewSheet,
                               ElementId)
from Snippets._context_manager import try_except

# `__revit__` members are unavailable when no UIDocument is active, and at
# module scope that kills the import outright. Resolve defensively; the entry
# point reports the real problem (see Snippets._host.resolve_doc()).
try:
    from Snippets._host import resolve_doc, get_revit_version
    doc, _doc_err = resolve_doc()
    rvt_year = get_revit_version()
except Exception:
    doc = None
    rvt_year = 2024


def create_revision(description, date, revision_type=None):
    #type:(str,str,RevisionNumberType) -> Revision
    """Function to create new Revision.
    :param description:   string for Description
    :param date:          string for Date
    :param revision_type: RevisionNumberType; None = Revit's own "None"
                          numbering (resolved lazily, see below)
    :return:              new Revision"""
    # NOT `revision_type=RevisionNumberType.None` in the signature: `None` is a
    # Python keyword, so `X.None` is a SyntaxError and this whole module failed
    # to parse — it could never be imported by anything. Resolving it here also
    # means the lookup happens at call time rather than at import time, which
    # matters because NumberType is obsolete from RVT 2023 and the member may
    # not exist at all on newer hosts.
    if revision_type is None:
        revision_type = getattr(RevisionNumberType, 'None', None)

    with try_except(debug=True):
        target_doc = doc or (resolve_doc()[0] if 'resolve_doc' in globals() else None)
        if not target_doc:
            return None
        new_rev              = Revision.Create(target_doc)
        new_rev.Description  = description
        new_rev.RevisionDate = date

        try:
            new_rev.NumberType   = revision_type #OBSOLETE IN RVT 2023
        except:
            pass

        #FIXME LATER.Set NumberType to None
        # if rvt_year < 2023:
        #     pass
        # else:
        #     new_rev.RevisionNumberingSequenceId =
        #

        return new_rev

def add_revision_to_sheet(sheet, revision_id):
    #type:(ViewSheet, ElementId) -> None
    """ Function to add existing revision to the given sheet
    :param sheet:           ViewSheet
    :param revision_id:     Revision.Id that should be added to the ViewSheet."""
    with try_except(debug=True):
        # GET EXISTING ADDITIONAL REVISIONS
        revisions_on_sheet = sheet.GetAdditionalRevisionIds()

        # ADD NEW REVISION TO THE LIST
        revisions_on_sheet.Add(revision_id)

        # SET NEW LIST OF ADDITIONAL REVISIONS
        sheet.SetAdditionalRevisionIds(revisions_on_sheet)




# ╔╦╗╔═╗╔═╗╔╦╗╦╔╗╔╔═╗
#  ║ ║╣ ╚═╗ ║ ║║║║║ ╦
#  ╩ ╚═╝╚═╝ ╩ ╩╝╚╝╚═╝
#==================================================
def revision_data(revision):
    print("SequenceNumber: "            + str( revision.SequenceNumber))
    print("NumberType: "                + str( revision.NumberType))
    print("RevisionDate: "              + str( revision.RevisionDate))
    print("Description: "               + str( revision.Description))
    print("Issued: "                    + str( revision.Issued))
    print("IssuedTo: "                  + str( revision.IssuedTo))
    print("IssuedBy: "                  + str( revision.IssuedBy))
    print("Visibility: "                + str( revision.Visibility))
    print("ViewSpecific: "              + str( revision.ViewSpecific))
    print("OwnerViewId: "               + str( revision.OwnerViewId))
    print("GroupId: "                   + str( revision.GroupId))
    print("AssemblyInstanceId: "        + str( revision.AssemblyInstanceId))
    print("GetDependentElements(): "    + str( revision.GetDependentElements()))

def revision_cloud_data(reivsion_cloud):
    print(reivsion_cloud)
    print("OwnerViewId: " + str( reivsion_cloud.OwnerViewId) )
    owner_view = doc.GetElement(reivsion_cloud.OwnerViewId)
    print("OwnerView: " + owner_view.Name)
    print("Hidden: " + str( reivsion_cloud.IsHidden(owner_view)))

