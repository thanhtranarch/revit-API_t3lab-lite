__author__ ="Tran Tien Thanh"
__title__ = "Wall Type"
__version__ = "1.0.0"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

import System
from System.Windows.Forms import OpenFileDialog, DialogResult
import xlrd
from pyrevit import *
"""--------------------------------------------------"""
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
"""--------------------------------------------------"""
#Get first Wall Type
samplewall=FilteredElementCollector(doc).OfClass(WallType).FirstElement()

# Get first Wall Type
samplewall = FilteredElementCollector(doc).OfClass(WallType).FirstElement()

WallTypeData = namedtuple("WallTypeData", ["name", "typemark", "function", "fire_rate", "layers"])
LayerData = namedtuple("LayerData", ["function", "material", "thickness"])

# Function to get filepath
def get_type_path():
    file_dialog = OpenFileDialog()
    file_dialog.Filter = "Excel Files|*.xlsx;*.xls"
    file_dialog.Title = "Select an Excel File"
    result = file_dialog.ShowDialog()
    if result == DialogResult.OK:
        return file_dialog.FileName
    else:
        return None

# Function to read data from Excel
def read_excel_data(file_path):
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    data = []
    for row_num in range(1, sheet.nrows):  # Skip header row
        row_data = [sheet.cell_value(row_num, col_num) for col_num in range(sheet.ncols)]
        data.append(row_data)
        
    wall_data_list = []
    for walldata in data:
        wall_layers = []
        for i in range(5, len(walldata), 3):
            if walldata[i] != "":
                layer = LayerData(walldata[i], walldata[i + 1], walldata[i + 2])
                wall_layers.append(layer)
        wall_type = WallTypeData(walldata[0], walldata[1], walldata[2], walldata[3], layers=wall_layers)
        wall_data_list.append(wall_type)
    return wall_data_list

# Function to create wall types
def walltype_create(doc,wall_data):
    def get_material_function(function_str):
        mapping = {
            "substrate": MaterialFunctionAssignment.Substrate,
            "structure": MaterialFunctionAssignment.Structure,
            "membrane": MaterialFunctionAssignment.Membrane,
            "finish1": MaterialFunctionAssignment.Finish1,
            "finish2": MaterialFunctionAssignment.Finish2,
        }
        return mapping.get(function_str.lower(), MaterialFunctionAssignment.Structure)
    for wall_type in wall_data:
        # Duplicate the sample wall type
        new_wall_type = samplewall.Duplicate(wall_type.name)
        new_wall_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).Set(wall_type.typemark)
        new_wall_type.get_Parameter(BuiltInParameter.FIRE_RATING).Set(wall_type.fire_rate)

        # Get the structure of the wall type
        wall_structure = new_wall_type.GetCompoundStructure()

        # Clear existing layers
        existing_layers = wall_structure.GetLayers()
        for i in range(len(existing_layers)):
            wall_structure.RemoveLayer(0) 
        # Add new layers
        for layer in wall_type.layers:
            # Convert the thickness to the appropriate unit
            thickness = UnitUtils.ConvertToInternalUnits(float(layer.thickness), DisplayUnitType.DUT_MILLIMETERS)

            # Find material by name
            material = FilteredElementCollector(doc).OfClass(Material).ToElements(). \
                FirstOrDefault(lambda m: m.Name == layer.material)
            
            if material is None:
                raise Exception("Material '{}' not found.".format(layer.material))

            # Determine layer function
            function = get_material_function(layer.function)

            # Create the new layer and add to structure
            new_layer = CompoundStructureLayer(thickness, function)
            new_layer.MaterialId = material.Id
            wall_structure.InsertLayer(wall_structure.LayerCount, new_layer)

        # Set the new structure to the wall type
        new_wall_type.SetCompoundStructure(wall_structure)
        print("Created wall type: ", wall_type.name)

t=Transaction(doc,"Create Wall Type(s)")
t.Start()
file_path=get_type_path()
if file_path:
    wall_data = read_excel_data(file_path)
    wall_type_create(doc, wall_data)



t.Commit()
