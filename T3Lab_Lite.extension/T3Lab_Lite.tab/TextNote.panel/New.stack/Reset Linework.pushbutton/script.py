import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Get the active document
doc = __revit__.ActiveUIDocument.Document

# Get the active view
active_view = doc.ActiveView

# Start a transaction
t = Transaction(doc, "Override Line Graphics")
t.Start()

# Get all lines in the active view
lines_collector = FilteredElementCollector(doc, active_view.Id).OfCategory(BuiltInCategory.OST_Lines).ToElements()

# Get the line pattern id for '<By category>'
line_pattern_id = LinePatternElement.GetSolidPatternId()

# Override the graphics of each line with '<By category>' line pattern
for line in lines_collector:
    overrides = active_view.GetElementOverrides(line.Id)
    overrides.SetProjectionLinePatternId(line_pattern_id)
    active_view.SetElementOverrides(line.Id, overrides)

# Commit the transaction
t.Commit()

# End the transaction
t.Dispose()

# Display a message indicating that line graphics override is complete
TaskDialog.Show("Success", "Line graphics override completed with <By category> line pattern.")