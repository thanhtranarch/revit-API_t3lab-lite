# -*- coding: utf-8 -*-
"""
Tile Counting
Count tiles needed for selected floors based on material surface pattern.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__ = "Tile Counting"

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    Transaction, FillPatternElement, UnitUtils
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, script
import math

logger = script.get_logger()
doc = revit.doc
uidoc = revit.uidoc


def select_floors():
    sel = uidoc.Selection
    picked = sel.PickObjects(ObjectType.Element, "Select multiple floors")
    return [doc.GetElement(p.ElementId) for p in picked
            if doc.GetElement(p.ElementId).Category.Name == "Floors"]


def get_floor_area_sqft(floor):
    param = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)
    if param:
        return param.AsDouble()
    return 0.0


def get_tile_size_from_material(floor):
    try:
        floor_type = floor.FloorType
        mat_ids = floor_type.GetMaterialIds(False)
        if not mat_ids:
            return None, None
        material = doc.GetElement(list(mat_ids)[0])
        pattern_id = material.SurfaceForegroundPatternId
        if pattern_id.IntegerValue == -1:
            return None, None
        pattern_elem = doc.GetElement(pattern_id)
        pattern = pattern_elem.GetFillPattern()
        grids = pattern.GetGrids()
        tile_w = grids[0].Spacing if len(grids) > 0 else None
        tile_h = grids[1].Spacing if len(grids) > 1 else tile_w
        return tile_w, tile_h
    except Exception as ex:
        logger.warning("Could not read tile size from material: {}".format(ex))
        return None, None


try:
    floors = select_floors()
    if not floors:
        TaskDialog.Show("Tile Counting", "No floors selected.")
    else:
        lines = []
        total_tiles = 0
        for floor in floors:
            area_sqft = get_floor_area_sqft(floor)
            tile_w, tile_h = get_tile_size_from_material(floor)
            if tile_w and tile_h and tile_w > 0 and tile_h > 0:
                tile_area = tile_w * tile_h
                count = int(math.ceil(area_sqft / tile_area))
            else:
                count = 0
                lines.append("  [No pattern] {}".format(floor.Name or floor.Id))
                continue
            total_tiles += count
            lines.append("  {} — {:.1f} sqft — {} tiles".format(
                floor.Name or str(floor.Id), area_sqft, count))

        msg = "Tile count per floor:\n" + "\n".join(lines)
        msg += "\n\nTotal tiles: {}".format(total_tiles)
        TaskDialog.Show("Tile Counting", msg)

except Exception as e:
    logger.error("Tile Counting error: {}".format(e))
