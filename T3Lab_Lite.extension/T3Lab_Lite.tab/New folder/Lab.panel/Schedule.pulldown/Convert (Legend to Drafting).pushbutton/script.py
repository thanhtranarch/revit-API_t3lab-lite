"""
Covert
Covert Schedule to Legend
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__ ="Tran Tien Thanh"
__title__ = "Convert\n (Schedule to Lengend)"



from Autodesk.Revit.UI import TaskDialog, UIApplication, UIDocument
from Autodesk.Revit.DB import *
from pyrevit.framework import List
from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

"""--------------------------------------------------"""
class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes
"""--------------------------------------------------"""
logger = script.get_logger()
uidoc=__revit__.ActiveUIDocument
doc=uidoc.Document
"""--------------------------------------------------"""
# Get legend name to compare:
all_graphviews=revit.query.get_all_views(doc=revit.doc)
all_legend_names=[revit.query.get_name(x) for x in all_graphviews if x.ViewType==DB.ViewType.Legend]
"""--------------------------------------------------"""
# Show Schedule's List -> get a list of selected Schedule:
schedule_views=forms.select_views(
    title="Select Schedules to convert",
    filterfunc=lambda x: x.ViewType== DB.ViewType.Schedule)
# schedule_views=forms.select_schedules() (run lower more )
"""--------------------------------------------------"""
def create_legend_from_schedule(schedule):
    # Start a transaction
    t = DB.Transaction(revit.doc, "Create Legend from Schedule")
    t.Start()

    try:
        # Get the legend view family type
        legend_view_family = DB.ElementId(DB.ViewType.Legend)

        # Create a new legend view
        legend_view = DB.ViewDrafting.Create(revit.doc, legend_view_family)

        # Set view name
        legend_view.Name = "Legend from Schedule"

        # Get the schedule's underlying table
        schedule_table = schedule.GetTableData().GetSectionData(DB.SectionType.Body)

        # Loop through each row in the table
        for i in range(schedule_table.NumberOfRows):
            # Get the element from the specified cell in the row
            cell_id = DB.TableCellId(i, 0)
            element_id = schedule.GetCellText(cell_id)
            if element_id:
                element = revit.doc.GetElement(element_id)

                # Create a legend component for each element in the legend view
                if element:
                    legend_component = DB.LegendComponent.Create(revit.doc, element.Id)
                    legend_view.AddChild(legend_component)

        # Commit the transaction
        t.Commit()

        # Show the legend view
        revit.uidoc.ActiveView = legend_view

    except Exception as ex:
        # Roll back the transaction if there's an exception
        t.RollBack()
        print('Failed to create legend:', ex)

def main():
    # Ensure we have an active document
    if revit.doc:
        active_view = revit.uidoc.ActiveView
        if isinstance(active_view, DB.ViewSchedule):
            result = forms.alert("Convert this schedule to a legend?", title="Convert to Legend", \
                                 options=["Yes", "No"])
            if result == "Yes":
                create_legend_from_schedule(active_view)
        else:
            forms.alert("Please select a schedule view to convert to legend!", title="Error")
    else:
        print('No active document.')

if __name__ == "__main__":
    main()
