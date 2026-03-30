# -*- coding: utf-8 -*-
"""
Room to room

This script will loop through existing rooms to asign room in new phase views

Author: Tran Tien Thanh
Email: trantienthanh909@gmail.com
LinkedIn: linkedin.com/in/sunarch7899/
"""
__author__ ="Tran Tien Thanh"
__title__ = "Room to room"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from collections import namedtuple

# DEFINE VARIABLES
# ==================================================
uidoc   = __revit__.ActiveUIDocument
doc     = __revit__.ActiveUIDocument.Document
active_view_id = doc.ActiveView.Id
phase_views = ["Demo", "Existing"]
Roomtag = namedtuple('Roomtag', ['obj_tag', 'name'])
area_room = ["PATIO", "PORCH"]
sub_room = ["CLOSET", "STORAGE", "PANTRY" ]
# CLASS/FUNCTIONS
# ==================================================

def lcs(str1, str2):
    m, n = len(str1), len(str2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if str1[i - 1] == str2[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    lcs_str = ""
    i, j = m, n
    while i > 0 and j > 0:
        if str1[i - 1] == str2[j - 1]:
            lcs_str = str1[i - 1] + lcs_str
            i -= 1
            j -= 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1

    return lcs_str

def string_intersection(s1, s2):
    intersection = set(s1) & set(s2)
    return ''.join(sorted(intersection))

def compare_box(room1, room2):
    room1_box = room1.get_BoundingBox(doc.ActiveView)
    room2_box = room2.get_BoundingBox(doc.ActiveView)
    if room1_box is not None and room2_box is not None:
        compare = [
            abs(room1_box.Min.X - room2_box.Min.X),
            abs(room1_box.Min.Y - room2_box.Min.Y),
            abs(room1_box.Min.Z - room2_box.Min.Z),
            abs(room1_box.Max.X - room2_box.Max.X),
            abs(room1_box.Max.Y - room2_box.Max.Y),]
        result = all(i <= 1 for i in compare)
        return result

def assign_room_to_room(room1, room2):
    try: 
        roomname1 = room1.LookupParameter("Name").AsString()
        room2.LookupParameter("Name").Set(roomname1)
    except Exception as e:
        pass
    
    try:
        cei_height1 = room1.LookupParameter("Comments").AsValueString()
        room2.LookupParameter("Comments").Set(cei_height1)
    except Exception as e:
        pass
    
    try:
        flo_finish1 = room1.LookupParameter("Floor Finish").AsString()
        room2.LookupParameter("Floor Finish").Set(flo_finish1)
    except Exception as e:
        pass
  

def set_roomtype_to_room(room, roomtype_Id):
    room.LookupParameter("Room Type").Set(roomtype_Id)

def loopparameter(obj):
    count = 0
    for param in obj.Parameters:
        param_value = None
        storage_type = param.StorageType

        if storage_type == StorageType.String:
            param_value = param.AsString()
        elif storage_type == StorageType.Double:
            param_value = param.AsDouble()
        elif storage_type == StorageType.Integer:
            param_value = param.AsInteger()
        else:
            # Handle other storage types or unknown types
            print("Unsupported parameter type: {}".format(storage_type))
        print("Room Name: {}, Parameter Name: {}, Value: {}".format(obj, param, param_value))
        

# MAIN SCRIPT
# ==================================================
#get elements
rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
roomtagtypes_dict = {i: i.LookupParameter("Type Name").AsString() for i in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RoomTags).WhereElementIsElementType().ToElements() if i.LookupParameter("Type Name") is not None}
roomtag_appendage = "Room Tag (Appendage)"
roomtag_area = "Room Tag (Area)"
print(roomtagtypes_dict)

with Transaction(doc, "Assign Room to Room") as t:
    t.Start()
    count = 0
    
    for room1 in rooms:
        # print("Processing Room 1: {},{}".format(room1.Id, room1.LookupParameter("Name").AsString()))
        if (room1.LookupParameter("Phase").AsValueString()) in phase_views: 
            for room2 in rooms:
                
                    if compare_box(room1, room2):
                        # print("Room {} matches with Room {}".format(room1.Id, room2.Id))
                        # print(compare_box(room1, room2))
                        if room1.LookupParameter("Name") and len(room1.LookupParameter("Name").AsString()) >= 6:
                                    assign_room_to_room(room1, room2)
    
    for room in rooms:
        for sub in sub_room:
            if string_intersection(room.LookupParameter("Name").AsString(), sub) == sub:
                roomtag_type_id = next(
                    (roomtag_type.Id for roomtag_type in roomtagtypes_dict 
                    if roomtag_type.LookupParameter("Type Name").AsString() == "Room Tag (Appendage)"), 
                    None
                )

                if roomtag_type_id:
                    room.ChangeTypeId(roomtag_type_id)
                    count += 1
                    
        for area in area_room:
            if string_intersection(room.LookupParameter("Name").AsString(), area) == area:
                roomtag_type_id = next(
                    (roomtag_type.Id for roomtag_type in roomtagtypes_dict 
                    if roomtag_type.LookupParameter("Type Name").AsString() == "Room Tag (Area)"), 
                    None
                )
                print(roomtag_type_id)
                if roomtag_type_id:
                    room.ChangeTypeId(roomtag_type_id)
                    count += 1

    t.Commit()


if count > 0:
    TaskDialog.Show("Room to room", "Done!")
else:
    TaskDialog.Show("Room to room", "No matching rooms found.")
