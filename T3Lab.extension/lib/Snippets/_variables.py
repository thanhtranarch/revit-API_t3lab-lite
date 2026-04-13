# -*- coding: utf-8 -*-
"""
Variables Snippets

Common variable definitions and constants for Revit scripts.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Variables Snippets"

from Autodesk.Revit.DB import (ViewPlan, ViewSection, View3D, ViewSchedule, View, ViewType, ViewDrafting,
                               DetailLine, DetailCurve, DetailArc, DetailEllipse, DetailNurbSpline,
                               ModelLine, ModelCurve, ModelArc, ModelEllipse, ModelNurbSpline
                               )


ALL_VIEW_TYPES = [ViewPlan, ViewSection, View3D , ViewSchedule, View, ViewDrafting]
LINE_TYPES = [DetailLine, DetailCurve, DetailArc, DetailEllipse, DetailNurbSpline,
              ModelLine, ModelCurve, ModelArc, ModelEllipse, ModelNurbSpline]