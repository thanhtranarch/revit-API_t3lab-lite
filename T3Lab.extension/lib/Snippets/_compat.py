# -*- coding: utf-8 -*-
"""
Revit API version-compatibility shims.

Author: Tran Tien Thanh
"""


def eid_value(element_id):
    """Return the integer value of an ElementId, version-safe.

    Revit 2024+ replaced ElementId.IntegerValue with ElementId.Value (Int64).
    Falls back to IntegerValue for Revit 2023 and earlier.
    """
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)          # Revit 2024+ (Int64 -> plain int)
    except Exception:
        try:
            return int(element_id.IntegerValue)   # Revit 2023 and earlier
        except Exception:
            return -1
