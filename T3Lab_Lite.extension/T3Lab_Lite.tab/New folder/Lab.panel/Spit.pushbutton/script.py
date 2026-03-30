# -*- coding: utf-8 -*-
import clr 
import os
import sys

clr.AddReference("System")
from System import Windows
import System

clr.AddReference("RevitServices")
import RevitServices
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Collections.Generic import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Structure import *


# clr.AddReference('PresentationCore')
# clr.AddReference('PresentationFramework')
# clr.AddReference("System.Windows.Forms")

from System.Windows import MessageBox
from System.IO import FileStream, FileMode, FileAccess, File
from System.Windows.Markup import XamlReader


#Get UIdoc
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
levels_inModel = FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType().ToElements()
all_levels = sorted(levels_inModel, key=lambda lv: lv.Elevation)

# Get the directory path of the .py file
dir_path = os.path.dirname(os.path.realpath(__file__))
#xaml_file_path = os.path.join(dir_path, "Window.xaml")
# Get the file path of the .xaml file

class ColumnItem:
    def __init__(self, base_level, top_level, base_offset, top_offset):
        self.base_level = base_level
        self.top_level = top_level
        self.base_offset = base_offset
        self.top_offset = top_offset

    def __eq__(self, other):
        if not isinstance(other, ColumnItem): return False
        return (self.base_level == other.base_level and
                self.top_level == other.top_level and
                self.base_offset == other.base_offset and
                self.top_offset == other.top_offset)

    def __hash__(self):
        return hash((self.base_level, self.top_level, self.base_offset, self.top_offset))

class Utils:
    def __init__(self):
        pass

    def modify_wall_unconnected (self, wall):

        base_level_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        base_level = doc.GetElement(base_level_id)
        base_offset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble()
        bottom_elevation = base_level.Elevation + base_offset

        top_level_id = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId()
        top_level = all_levels[len(all_levels)-1]


        if top_level_id == ElementId.InvalidElementId:
            height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble()
            top_elevation = height + bottom_elevation

            for lv in all_levels:
                if round(lv.Elevation - top_elevation, 0) >= 0:
                    top_level = lv
                    break

            top_offset = top_elevation - top_level.Elevation
            wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(top_level.Id)
            wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(top_offset)
        
    def get_column_item(self, element):
        #get level Id
        top_level_id = None
        base_level_id = None
        base_offset = 0
        top_offset = 0

        if isinstance(element, FamilyInstance):
            top_level_id = element.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).AsElementId()
            base_level_id = element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsElementId()
            top_offset = element.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).AsDouble()
            base_offset = element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).AsDouble()
        else:
            top_level_id = element.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId()
            base_level_id = element.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
            top_offset = element.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble()
            base_offset = element.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble()

        top_level = doc.GetElement(top_level_id)
        base_level = doc.GetElement(base_level_id)

        return ColumnItem(base_level, top_level,base_offset, top_offset)

    def get_bottom_and_top_elevation (self, element):

        #get original information
        item = self.get_column_item(element)
        base_level = item.base_level
        top_level = item.top_level
        base_offset = item.base_offset
        top_offset = item.top_offset
        
        #get bottom an top elevation
        bottom_elevation = base_level.Elevation + base_offset
        top_elevation = top_level.Elevation + top_offset

        return (bottom_elevation, top_elevation)
        

    def levels_by_element(self, element):

        #get bottom an top elevation
        elevations = self.get_bottom_and_top_elevation(element)
        bottom_elevation = elevations[0]
        top_elevation = elevations[1]

        #get real top and botom level
        new_base_level = all_levels[0]
        new_top_level = all_levels[len(all_levels)-1]
        for lv in all_levels:
            if lv.Elevation >= top_elevation:
                new_top_level = lv
                break
            
        for lv in reversed(all_levels):
            if lv.Elevation <= bottom_elevation:
                new_base_level = lv
                break

        #get list level cutting element
        levels_by_element = []
        for lv in all_levels:
            if lv.Elevation > new_base_level.Elevation and lv.Elevation < new_top_level.Elevation:
                levels_by_element.append(lv)

        levels_by_element.insert(0, new_base_level)
        levels_by_element.append(new_top_level)


        return levels_by_element
    
    def reoder_column_levels (self, element):

        #get bottom an top elevation
        elevations = self.get_bottom_and_top_elevation(element)
        bottom_elevation = elevations[0]
        top_elevation = elevations[1]

        #get list level cutting element
        levels_by_element = self.levels_by_element(element)
        total = len(levels_by_element)
        new_base_level = levels_by_element[0]
        new_top_level = levels_by_element[total-1]
        
        #reset to real level
        if total > 0:
            new_base_offset = bottom_elevation - new_base_level.Elevation
            new_top_offset = top_elevation - new_top_level.Elevation
        
            #set real level and base/top offset
            if isinstance(element, FamilyInstance):
                element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).Set(new_base_level.Id)
                element.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).Set(new_top_level.Id)
                element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(new_base_offset)
                element.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(new_top_offset)
            else:
                element.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).Set(new_base_level.Id)
                element.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(new_top_level.Id)
                element.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(new_base_offset)
                element.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(new_top_offset)
    

    def get_list_items (self, element):

        item = self.get_column_item(element)
        top_offset = item.top_offset
        base_offset = item.base_offset

        levels_by_element = self.levels_by_element(element)
        total = len(levels_by_element)

        if total < 2: return []
        else:

            #for first item
            base_item = None
            if (base_offset < 0):
                first_level = None
                is_exist = False
                for lv in all_levels:
                    if lv.Elevation < levels_by_element[0].Elevation:
                        is_exist = True
                        break
                
                if is_exist: first_level = all_levels[all_levels.index(levels_by_element[0]) - 1]
                
                if first_level is not None:
                    levels_by_element.insert(0, first_level)
                    distance = levels_by_element[1].Elevation - levels_by_element[0].Elevation
                    offset = distance + base_offset
                    base_item = ColumnItem(first_level, None, offset, 0)
                else:
                    base_item = ColumnItem(levels_by_element[0], levels_by_element[0], base_offset, 0)
                    levels_by_element.insert(0, levels_by_element[0])

            else: base_item = ColumnItem(levels_by_element[0], None, base_offset, 0)
            

            #for last item
            top_item = None
            if top_offset <= 0:
                del levels_by_element[len(levels_by_element)-1]
                last_level = levels_by_element[len(levels_by_element)-1]
                top_item = ColumnItem(last_level, None, 0,top_offset)
            else:
                last_level = levels_by_element[len(levels_by_element)-1]
                top_item = ColumnItem(last_level, last_level, 0, top_offset)

            #return list items
            list_items = []   
            try:
                for i in range(1, len(levels_by_element)-1):
                    list_items.append(ColumnItem(levels_by_element[i], None, 0, 0))
            except:
                pass

            list_items.insert(0, base_item)
            list_items.append(top_item)
           
            return list_items
        

    def split_columns (self, list_columns):

        for column in list_columns:
            Utils().reoder_column_levels(column)
        doc.Regenerate()

        for column in list_columns:
            symbol = column.Symbol
            location = column.Location.Point
            rotation = column.Location.Rotation
            column_items = self.get_list_items(column)
            unique_items = list(set(column_items))
            try:
                if len(unique_items) > 1:
                    for item in unique_items:
                        if item is not None:
                            new_column = doc.Create.NewFamilyInstance(location, symbol, item.base_level, StructuralType.Column)
                            new_column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM).Set(item.base_offset)
                            new_column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).Set(item.top_offset)
                            if item.top_level is not None:
                                new_column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM).Set(item.top_level.Id)
                            if rotation != 0:
                                axis = Line.CreateBound(location, location + XYZ.BasisZ)
                                new_column.Location.Rotate(axis, rotation)

                    #delete old element
                    try:
                        doc.Delete(column.Id)
                    except:
                        pass

            except:
                pass


    def get_upper_level (self, level, base_offset):

        if all_levels[0].Name == level.Name and base_offset < 0:
            return level
        
        for lv in all_levels:
            if lv.Elevation > level.Elevation:
                return lv
            
        return level


    def split_walls (self, list_walls):

        for wall in list_walls:
            wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM).Set(0)
            Utils().modify_wall_unconnected(wall)
            Utils().reoder_column_levels(wall)
           
        doc.Regenerate()

        for wall in list_walls:
            type_id = wall.GetTypeId()
            curve = wall.Location.Curve
            is_structure = False
            if wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT).AsInteger() == 1:
                is_structure = True

            wall_items = self.get_list_items(wall)
            unique_items = list(set(wall_items))
            try:
                if len(unique_items) > 1:
                    for item in unique_items:
                        new_base_level = item.base_level
                        new_top_level = self.get_upper_level(new_base_level, item.base_offset)
                        new_wall = Wall.Create(doc, curve, new_base_level.Id, is_structure)
                        new_wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(new_top_level.Id)
                        new_wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(item.base_offset)
                        new_wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(item.top_offset)
                        new_wall.ChangeTypeId(type_id)
                    #delete old element
                    try:
                        doc.Delete(wall.Id)
                    except:
                        pass
            except:
                pass

        
#ISelectionFilter for Door
class Filter_Columns_Wall(ISelectionFilter):
    def AllowElement(self, element):
        if element.Category.Name == "Structural Columns" or isinstance(element, Wall):
             return True
        else:
             return False
       
    def AllowReference(self, reference, position):
        return True

#region Main
class Main ():

    def __init__(self):
        pass

    def main_task(self):
        list_columns = []
        list_walls = []

        try:
            selected = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, Filter_Columns_Wall(), "Select Walls/Columns")
            for r in selected:
                ele = doc.GetElement(r)
                if isinstance(ele, FamilyInstance):
                    index = ele.get_Parameter(BuiltInParameter.SLANTED_COLUMN_TYPE_PARAM).AsInteger()
                    if index == 0:list_columns.append(ele)
                else: list_walls.append(ele)  
        except:
            pass

        try:
            t = Transaction(doc, "split")
            t.Start()

            Utils().split_columns(list_columns)
            Utils().split_walls(list_walls)

            t.Commit()
            
        except Exception as e:
            MessageBox.Show(str(e), "Message")
        
        

if __name__ == "__main__":
    Main().main_task()
#endregion
        

                
    
    




