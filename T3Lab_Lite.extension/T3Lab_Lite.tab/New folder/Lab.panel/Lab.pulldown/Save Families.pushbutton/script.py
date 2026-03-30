"""
Make Plans
Create Plans from Room List
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Save Families"


import sys
import os
import clr
from collections import namedtuple

from Autodesk.Revit.DB import *

import rpw
from rpw import revit, doc, uidoc, DB, UI, db, ui
#variable
docUI=__revit__.ActiveUIDocument
doc=docUI.Document

Extend_Crop = '0.75'  # 9"
# Get Rooms
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
def list_room():
	#Iterate through the room elements and print their names and other properties
	for room in rooms:
		room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
		room_name = room.LookupParameter("Name").AsString()
		area = room.Area
		print("Room Number: {}\n Room Name: {}\n Area: {}\n *********".format(room_number,room_name,area))


# Get View Types and Prompt User
plan_types = db.Collector(of_class='ViewFamilyType', is_type=True).wrapped_elements
# Filter all view types that are FloorPlan or CeilingPlan
plan_types_options = {t.name: t for t in plan_types
                      if t.view_family.name in ('FloorPlan', 'CeilingPlan')}

plan_type = ui.forms.SelectFromList('MakeViews', plan_types_options,
                                    description='Select View Type')
view_type_id = plan_type.Id

crop_offset = ui.forms.TextInput('MakeViews', default=Extend_Crop,
                                 description='View Crop Offset [feet]'
                                 )

crop_offset = float(crop_offset)


def offset_bbox(bbox, offset):
    """
    Offset Bounding Box by given offset
    http://archi-lab.net/create-view-by-room-with-dynamo/
    """
    bboxMinX = bbox.Min.X - offset
    bboxMinY = bbox.Min.Y - offset
    bboxMinZ = bbox.Min.Z - offset
    bboxMaxX = bbox.Max.X + offset
    bboxMaxY = bbox.Max.Y + offset
    bboxMaxZ = bbox.Max.Z + offset
    newBbox = DB.BoundingBoxXYZ()
    newBbox.Min = DB.XYZ(bboxMinX, bboxMinY, bboxMinZ)
    newBbox.Max = DB.XYZ(bboxMaxX, bboxMaxY, bboxMaxZ)
    return newBbox


@rpw.db.Transaction.ensure('Create View')
def create_plan(new_view, view_type_id, cropbox_visible=False, remove_underlay=True):
    """Create a Drafting View"""
    view_type_id
    name = new_view.name
    bbox = new_view.bbox
    level_id = new_view.level_id

    viewplan = DB.ViewPlan.Create(doc, view_type_id, level_id)
    viewplan.CropBoxActive = True
    viewplan.CropBoxVisible = cropbox_visible
    if remove_underlay and revit.version.year == '2015':
        underlay_param = viewplan.get_Parameter(DB.BuiltInParameter.VIEW_UNDERLAY_ID)
        underlay_param.Set(DB.ElementId.InvalidElementId)
    viewplan.CropBox = bbox

    counter = 1
    while True:
        # Auto Increment Room Number
        try:
            viewplan.Name = name
        except Exception:
            try:
                viewplan.Name = '{} - Copy {}'.format(name, counter)
            except Exception as errmsg:
                counter += 1
                if counter > 100:
                    raise Exception('Exceeded Maximum Loop')
            else:
                break
        else:
            break
    return viewplan

# def set_template():
# template_name = "Your Template Name"
# view_type = ViewType.FloorPlan  # Change this to the desired view type
# view_templates = FilteredElementCollector(doc).OfClass(View).ToElements()
#
# template_exists = False
# template_id = None
#
# for template in view_templates:
#     if template.ViewType == view_type and template.Name == template_name:
#         template_exists = True
#         template_id = template.Id
#         break
#
# if not template_exists:
#     template_id = ViewTemplate.Create(doc, view_type)
#     view_template = doc.GetElement(template_id)
#     view_template.Name = template_name
# view = ViewPlan.Create(doc, template_id, ElementId.InvalidElementId)
# def wall_tag():
# def light_tag():

NewView = namedtuple('NewView', ['name', 'bbox', 'level_id'])
new_views = []

for room in rooms:
    room = db.Element(room)
    room_name = room.LookupParameter("Name").AsString()
    room_number = room.LookupParameter("Number").AsString()

    new_room_name = '{} {}'.format(room_name, room_number)
    room_bbox = room.get_BoundingBox(doc.ActiveView)
    new_bbox = offset_bbox(room_bbox, crop_offset)

    new_view = NewView(name=new_room_name, bbox=new_bbox)
    new_views.append(new_view)

for new_view in new_views:
    view = create_plan(new_view= new_view, view_type_id=view_type_id)
