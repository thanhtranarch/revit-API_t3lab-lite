import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB
from pyrevit import revit, script

doc = revit.doc

# Let's get one view and check its VIEW_TYPE parameter
view = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType().FirstElement()
if view:
    param = view.get_Parameter(DB.BuiltInParameter.VIEW_TYPE)
    if param:
        print("Param exists! HasValue:", param.HasValue, "AsInteger:", param.AsInteger())
    else:
        print("Param does not exist!")
    
    print("Actual ViewType enum value:", int(view.ViewType))

print("Done")
