# -*- coding: utf-8 -*-
"""
Groups Snippets

Code snippets for working with Revit groups.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Groups Snippets"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗
# ║║║║╠═╝║ ║╠╦╝ ║
# ╩╩ ╩╩  ╚═╝╩╚═ ╩  IMPORT
#==================================================

from Autodesk.Revit.DB import *
from pyrevit import forms

# CUSTOM IMPORTS
from GUI.forms           import select_from_dict


try:
    from Snippets._host import resolve_doc, host_uiapp
    default_doc, _ = resolve_doc()
    _uiapp = host_uiapp()
    default_uidoc = _uiapp.ActiveUIDocument if _uiapp else None
    default_app = _uiapp.Application if _uiapp else None
except Exception:
    default_doc = None
    default_uidoc = None
    default_app = None



def select_group_types(given_groups = None, uidoc = None, title='__title__', version = 'Version 0.1' ,exit_if_none = False):
    """Function to select group names from a list.
    :param given_groups: List of groups. If none then all groups in project will be used.
    :param uidoc:
    :param exit_if_none:
    :return: list of selected group types_names
    """
    if uidoc is None:
        uidoc = default_uidoc
        if uidoc is None:
            try:
                from Snippets._host import host_uiapp
                _u = host_uiapp()
                uidoc = _u.ActiveUIDocument if _u else None
            except Exception:
                uidoc = None
    if uidoc is None:
        return []

    #TODO if given_groups , verify that all elements are Groups
    if not given_groups:
        given_groups = FilteredElementCollector(uidoc.Document).OfCategory(BuiltInCategory.OST_IOSModelGroups).ToElements()


    dict_all_groups = {}
    for g in given_groups:
        group_name = g.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        if group_name and group_name not in dict_all_groups:
            dict_all_groups[group_name] = g

    selected_groups = select_from_dict(elements_dict=dict_all_groups,
                                       title=title,
                                       label='Select Groups:',
                                       version=version)

    #>>>>>>>>>> EXIT IF NONE SELECTED
    if not selected_groups and exit_if_none:
        forms.alert("No GroupTypes were selected. \nPlease try again.", exitscript=True)

    return selected_groups




def select_attached_groups(list_of_groups, uidoc = None, title="__title__", label = "Select Groups:", version = 'Version 0.1', exit_if_none = False):
    """Function to select attached groups from given list of groups.
    :param list_of_groups: List containing groups from which to take attached groups.
    :return: List of selected attached groups
    """
    if uidoc is None:
        uidoc = default_uidoc
        if uidoc is None:
            try:
                from Snippets._host import host_uiapp
                _u = host_uiapp()
                uidoc = _u.ActiveUIDocument if _u else None
            except Exception:
                uidoc = None
    if uidoc is None:
        return []

    dict_of_attached_group_names = {}

    for g in list_of_groups:
        attached_groups_ids = g.GetAvailableAttachedDetailGroupTypeIds()
        for a_group_id in attached_groups_ids:
            a_group = uidoc.Document.GetElement(a_group_id)
            a_group_name = a_group.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            if a_group_name and a_group_name not in dict_of_attached_group_names:
                dict_of_attached_group_names[a_group_name] = a_group


    selected_a_groups = select_from_dict(elements_dict = dict_of_attached_group_names,
                                         title=title,
                                         label=label,
                                         version=version)

    # >>>>>>>>>> EXIT IF NONE SELECTED
    if not selected_a_groups and exit_if_none:
        forms.alert("No AttachedGroups were selected. \nPlease try again.", exitscript=True)

    return selected_a_groups




def show_attached_group(view, group, list_a_group_names_to_show, uidoc = None):
    """Function to show attached groups that match list_a_groups_to_show in the selected view for selected groups.
    :param view:
    :param group:
    :param list_a_group_names_to_show:
    :return:
    """
    if uidoc is None:
        uidoc = default_uidoc
        if uidoc is None:
            try:
                from Snippets._host import host_uiapp
                _u = host_uiapp()
                uidoc = _u.ActiveUIDocument if _u else None
            except Exception:
                uidoc = None
    if uidoc is None:
        return

    all_attached_groups = group.GetAvailableAttachedDetailGroupTypeIds()
    attached_group_id = None
    # print("\n\nAtached Groups:")
    for a_group_id in all_attached_groups:
        a_group = uidoc.Document.GetElement(a_group_id)
        a_group_name = a_group.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        # print(a_group_name)
        if a_group_name in list_a_group_names_to_show:
            attached_group_id = a_group_id


    if attached_group_id:
        print("Showing attached group on the group [{}] in view - [{}]".format(group.Id, view.Name))
        group.ShowAttachedDetailGroups(view,attached_group_id )
