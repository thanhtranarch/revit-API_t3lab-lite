"""
Calculator
Perform calculations using the inch unit.
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Calculator\n (Imperial)"


# Import the Revit API and .NET libraries
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *


def convert_to_internal_units(value, unit_type_id):
    internal_value = UnitUtils.ConvertToInternalUnits(value, unit_type_id)
    return internal_value

# Example usage
original_value = 10.0  # Value in original units (e.g., feet, meters, etc.)
unit_type_id = ForgeTypeId.FeetFractionalInches  # You need to provide the correct ForgeTypeId for the original unit type

internal_value = convert_to_internal_units(original_value, unit_type_id)
print("Internal Value:", internal_value)