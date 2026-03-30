# -*- coding: utf-8 -*-
"""
Tile Counting
......

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""
__author__ ="Tran Tien Thanh"
__title__ = "Tile Counting"

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
import math
from matplotlib import pyplot as plt


# DEFINE VARIABLES
# ==================================================
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# CLASS/FUNCTIONS
# ==================================================
def select_floors():
    sel = uidoc.Selection
    picked = sel.PickObjects(ObjectType.Element, "Select multiple floors")
    floors = [doc.GetElement(p.ElementId) for p in picked if doc.GetElement(p.ElementId).Category.Name == "Floors"]

    # Print the selected floors
    for floor in floors:
        print("Selected Floor: {} (ID: {})".format(floor.Name,floor.Id))
    
    return floors

def get_material_pattern(floor):
    floor_type = floor.FloorType
    material_id = next(iter(floor_type.GetMaterialIds(False)))
    material = doc.GetElement(material_id)
    pattern_id = material.SurfaceForegroundPatternId
    if pattern_id == ElementId.InvalidElementId:
        raise Exception("No surface pattern found on material.")
    pattern_element = doc.GetElement(pattern_id)
    pattern = pattern_element.GetFillPattern()
    return pattern

def layout_tiles(floor, pattern):
    floor_geometry = floor.get_Geometry(Options())
    floor_boundary = [edge.AsCurve() for edge in floor_geometry.Edges]
    tile_width = pattern.GetGrids()[0].Spacing
    tile_height = pattern.GetGrids()[1].Spacing if len(pattern.GetGrids()) > 1 else tile_width
    tiles = []
    for x in range(0, int(floor.Geometry.Area // tile_width)):
        for y in range(0, int(floor.Geometry.Area // tile_height)):
            tile_origin = XYZ(x * tile_width, y * tile_height, 0)
            if is_point_in_floor(tile_origin, floor_boundary):
                tiles.append(tile_origin)
    return tiles

def is_point_in_floor(point, boundary_curves):
    for curve in boundary_curves:
        if not curve.Project(point):
            return False
    return True

def create_tiles(doc, tiles, tile_family_symbol):
    t = Transaction(doc, "Place Tiles")
    t.Start()
    for tile_origin in tiles:
        doc.Create.NewFamilyInstance(tile_origin, tile_family_symbol, StructuralType.NonStructural)
    t.Commit()

def create_tile_chart(tile_count):
    """Create a bar chart for the number of tiles."""
    plt.bar(range(len(tile_count)), tile_count, color='blue')
    plt.xlabel('Floor Index')
    plt.ylabel('Number of Tiles')
    plt.title('Tile Count per Floor')
    plt.show()  # Display the chart

# MAIN SCRIPT
# ==================================================
try:
    floors = select_floors()
    tile_counts = []  # List to store tile counts for each floor
    for floor in floors:
        pattern = get_material_pattern(floor)
        tiles = layout_tiles(floor, pattern)
        create_tiles(doc, tiles, None)  # Assuming tile_family_symbol is defined elsewhere
        tile_counts.append(len(tiles))  # Store the count of tiles for the current floor

    create_tile_chart(tile_counts)  # Call the chart creation function

except Exception as e:
    print("Error: {}".format(e))
