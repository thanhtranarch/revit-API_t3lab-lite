# -*- coding: utf-8 -*-
"""
Make Plans

Create Plans from Room List

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Create Plan Views"

# IMPORT LIBRARIES
# ==================================================
from rpw import revit, DB, db

from Autodesk.Revit.DB import (
    Transaction,
    View,
    FilteredElementCollector,
    BuiltInCategory,
    ViewType,
    ViewFamilyType,
    BoundingBoxXYZ
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms
from collections import namedtuple
import re

# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
pattern=r'[^a-zA-Z0-9.]'

# CLASS/FUNCTIONS
# ==================================================
def offset_bbox(bbox, offset=1):
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

def create_plan(view, view_type_id, cropbox_visible=True):
    t = Transaction(doc, "Create Plan View")
    t.Start()
    try:
        viewplan = DB.ViewPlan.Create(doc, view_type_id, view.room_level_id)
        viewplan.CropBoxActive = True
        viewplan.CropBoxVisible = cropbox_visible
        viewplan.CropBox = view.bbox
        viewplan.Name = view.name
        t.Commit()
        return viewplan
    except Exception as e:
        print("{}".format(e))
        raise e
    
# MAIN SCRIPT
# ==================================================   
#Collect Room & Filter:
room_obj        = namedtuple("room_obj",["name","number","type","obj"])
room_collected  = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
room_list       = map(lambda r_obj: room_obj(r_obj.LookupParameter("Name").AsString(),r_obj.LookupParameter("Number").AsString(),r_obj.LookupParameter("Room Type").AsString() if r_obj.LookupParameter("Room Type") else None ,r_obj),room_collected)
room_obj_unit   = map(lambda r_obj: "UNIT TYPE: {} | #{} | {}".format(r_obj.type,r_obj.number,r_obj.name),filter(lambda r_obj: "UNIT" in r_obj.name.upper() and r_obj.type != None,room_list ))
room_obj_common = map(lambda r_obj: "{} | #{}".format(r_obj.name,r_obj.number),filter(lambda r_obj: "UNIT" not in r_obj.name.upper(),room_list))
room_general    = sorted(room_obj_common + room_obj_unit)

#Collect Viewtemplate:
view_temp       = namedtuple("view_temp",["name","obj"])
view_collected  = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType().ToElements()
view_temp_plan  = map(lambda v_obj: view_temp(v_obj.Name,v_obj), filter(lambda v_obj: v_obj.ViewType == ViewType.FloorPlan and v_obj.IsTemplate,view_collected))
view_temp_RCP   = map(lambda v_obj: view_temp(v_obj.Name,v_obj), filter(lambda v_obj: v_obj.ViewType == ViewType.CeilingPlan and v_obj.IsTemplate,view_collected))

#Collect Viewtype:
view_type       = FilteredElementCollector(doc).OfClass(ViewFamilyType).WhereElementIsElementType().ToElements()
view_type_plan  = {v_obj.FamilyName for v_obj in filter(lambda v_obj: v_obj.FamilyName in ('Floor Plan', 'Ceiling Plan'),view_type)}


in4_room        = namedtuple("in4_room",["name","bbox","room_level_id","obj"])
"""-----------------------------------"""
#Show list of room to select
selected_rooms = forms.SelectFromList.show( room_general,
                                            title="Select Room",
                                            width=500,
                                            button_name="Run",
                                            multiselect=True)
#Check selected_rooms, if True -> continues, else -> raise
if selected_rooms:
    room_set={}
    for room in selected_rooms:
        a=room.split("|")
        if "UNIT" in room:
            room_set[re.sub(pattern, '',a[1])] = a[-1].strip(" ")
        else:
            room_set[re.sub(pattern, '',a[1])] = a[0].strip(" ")

    #Get romm in4 to view
    view_in4_list=[]
    for room_number in room_set:
        room_name = ""
        for room in room_list:
            if room.number == room_number:
                if ("UNIT" in room.name) and (room.type is not None):
                    name_planview = "ENLARGED PLAN - TYPE {} ({})".format(room.type,room.name)
                else:
                    name_planview = "ENLARGED PLAN - {} - (#{})".format(room.name,room_number)

                room_level_id= room.obj.LevelId
                room_bbox   = room.obj.get_BoundingBox(doc.ActiveView)
                new_bbox    = offset_bbox(room_bbox)
                view_plan   = in4_room(name_planview,new_bbox,room_level_id,room.obj)
                view_in4_list.append(view_plan)
    #Chose Plan Type
    plan_types = forms.SelectFromList.show(
        view_type_plan,
        title="Type View",
        width=500,
        button_name="Run",
        multiselect=True)
    if plan_types:
        #chose Template
        template_plan   = None
        template_RCP    = None
        if "Floor Plan" in plan_types:
            template_plan = forms.SelectFromList.show(
                sorted([v_obj.name for v_obj in view_temp_plan]),
                title="View templates (Plan)",
                width=500,
                button_name="Run",
                multiselect=False)
        if "Ceiling Plan" in plan_types:
            template_RCP = forms.SelectFromList.show(
                sorted([v_obj.name for v_obj in view_temp_RCP]),
                title="View templates (RCP)",
                width=500,
                button_name="Run",
                multiselect=False)

        #Asign Template & Create View
        for plan_type in plan_types:
            for view in view_in4_list:
                plan_type_id=next((i.Id for i in view_type if i.FamilyName == plan_type), None)
                created_view = create_plan(view, plan_type_id)
                if template_RCP or template_plan:
                    template_id = None
                    try:
                        if created_view.ViewType == ViewType.FloorPlan and template_plan:
                            template_id     = next((v.obj.Id for v in view_temp_plan if v.obj.Name == template_plan), None)
                        elif created_view.ViewType == ViewType.CeilingPlan and template_RCP:
                            template_id     = next((v.obj.Id for v in view_temp_RCP if v.obj.Name == template_RCP), None)
                    except:
                        pass
                    if template_id is not None:
                        with Transaction(doc, "Assign Template") as t:
                            t.Start()
                            viewplan = doc.GetElement(created_view.Id)
                            viewplan.ViewTemplateId = template_id
                            t.Commit()
                    else:
                        TaskDialog.Show("Create Plan", "Plan {} was created without a template".format(view.name))
                else:
                    TaskDialog.Show("Create Plan", "Please proceed")
else:
    TaskDialog.Show("Create Plan", "Please proceed")
# view_template  = map(lambda v_obj: view_temp(v_obj.Name,v_obj), filter(lambda v_obj: v_obj.IsTemplate,view_collected))
# with Transaction(doc, "Remove KTA name") as t:
#     t.Start()
#     for view in view_template:
#         view[1].Name = view[0].strip("_")
            
#     t.Commit()