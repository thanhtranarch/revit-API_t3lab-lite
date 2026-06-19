import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB
from pyrevit import revit, script

doc = revit.doc

bip = DB.BuiltInParameter.VIEW_TYPE
provider = DB.ParameterValueProvider(DB.ElementId(bip))
evaluator = DB.FilterNumericEquals()

def make_rule(vt):
    return DB.ElementParameterFilter(DB.FilterIntegerRule(provider, evaluator, int(vt)))

from System.Collections.Generic import List
filters = List[DB.ElementFilter]()
filters.Add(make_rule(DB.ViewType.FloorPlan))
filters.Add(make_rule(DB.ViewType.CeilingPlan))
filters.Add(make_rule(DB.ViewType.DraftingView))
filters.Add(make_rule(DB.ViewType.Legend))

or_filter = DB.LogicalOrFilter(filters)

views = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType().WherePasses(or_filter)

for v in views:
    if not v.IsTemplate:
        print(v.Name, v.ViewType)

print("Done")
