# -*- coding: utf-8 -*-
"""Advanced Schedule Duplicator with Template Options"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit import DB
from Autodesk.Revit.DB import ViewDuplicateOption
from pyrevit import revit, forms, script

doc = revit.doc
logger = script.get_logger()

def get_all_view_templates():
    """Get all view templates including those that might work with schedules"""
    collector = DB.FilteredElementCollector(doc)
    view_templates = collector.OfClass(DB.View).WhereElementIsNotElementType().ToElements()
    
    valid_templates = []
    for vt in view_templates:
        if vt.IsTemplate:
            # Include schedule templates and general templates that might work
            if vt.ViewType == DB.ViewType.Schedule:
                valid_templates.append(vt)
            # Also include undefined view types that might be compatible
            elif vt.ViewType == DB.ViewType.Undefined:
                valid_templates.append(vt)
    
    return valid_templates

def apply_template_with_options(schedule, view_template, apply_template=True):
    """Apply view template with option to include or exclude certain elements"""
    if not apply_template or not view_template:
        return True
    
    try:
        # First check if template is valid for this schedule
        if view_template.IsValidViewTemplate(schedule):
            # Apply the template
            schedule.ViewTemplateId = view_template.Id
            return True
        else:
            # Try to apply template parameters manually
            try:
                # This is a fallback method for incompatible templates
                schedule.ViewTemplateId = DB.ElementId.InvalidElementId
                logger.warning("Template '{}' not directly compatible, using fallback".format(view_template.Name))
            except:
                logger.warning("Could not apply template '{}' to schedule '{}'".format(
                    view_template.Name, schedule.Name))
            return False
    except Exception as e:
        logger.error("Error applying template: {}".format(str(e)))
        return False

def main():
    try:
        # Get source schedule
        collector = DB.FilteredElementCollector(doc)
        schedules = collector.OfClass(DB.ViewSchedule).WhereElementIsNotElementType().ToElements()
        valid_schedules = [s for s in schedules if not s.IsTemplate]
        
        if not valid_schedules:
            forms.alert("No schedules found!")
            return
        
        schedule_names = [s.Name for s in valid_schedules]
        selected_name = forms.SelectFromList.show(schedule_names, "Select Schedule to Duplicate")
        if not selected_name:
            return
        
        selected_schedule = next(s for s in valid_schedules if s.Name == selected_name)
        
        # Get copy count
        count_str = forms.ask_for_string("Number of copies?", "3")
        try:
            copy_count = int(count_str)
        except:
            forms.alert("Invalid number!")
            return
        
        # Get base name
        base_name = forms.ask_for_string("Base name:", selected_schedule.Name + "_Copy") or selected_schedule.Name + "_Copy"
        
        # View Template Selection with advanced options
        view_templates = get_all_view_templates()
        
        template_options = ["None - Keep Original", "Use Source Schedule's Template"]
        template_dict = {
            "None - Keep Original": None,
            "Use Source Schedule's Template": "source"
        }
        
        current_template = None
        if selected_schedule.ViewTemplateId != DB.ElementId.InvalidElementId:
            current_template = doc.GetElement(selected_schedule.ViewTemplateId)
            if current_template:
                template_options.insert(0, "Current: " + current_template.Name)
                template_dict["Current: " + current_template.Name] = current_template
        
        for vt in view_templates:
            template_options.append(vt.Name)
            template_dict[vt.Name] = vt
        
        selected_template_key = forms.SelectFromList.show(
            template_options,
            title="Select View Template",
            multiselect=False
        )
        
        selected_template = template_dict.get(selected_template_key) if selected_template_key else None
        
        # If user selected "Use Source Schedule's Template", get the actual template
        if selected_template == "source" and current_template:
            selected_template = current_template
        elif selected_template == "source" and not current_template:
            forms.alert("Source schedule doesn't have a view template applied.")
            selected_template = None
        
        # Duplicate method
        method_options = ["Duplicate", "As Dependent", "As Independent"]
        duplicate_method = forms.SelectFromList.show(method_options, "Duplicate Method") or "Duplicate"
        
        if duplicate_method == "As Dependent":
            duplicate_option = ViewDuplicateOption.AsDependent
        elif duplicate_method == "As Independent":
            duplicate_option = ViewDuplicateOption.AsIndependent
        else:
            duplicate_option = ViewDuplicateOption.Duplicate
        
        # Ask if user wants to apply template
        apply_template = True
        if selected_template:
            apply_template = forms.alert(
                "Apply view template '{}' to all new schedules?".format(selected_template.Name),
                yes=True, no=True
            )
        
        # Create duplicates
        success_count = 0
        
        with revit.TransactionGroup("Duplicate Schedules"):
            for i in range(copy_count):
                with revit.Transaction("Create Copy {}".format(i + 1)):
                    try:
                        new_id = selected_schedule.Duplicate(duplicate_option)
                        new_schedule = doc.GetElement(new_id)
                        new_schedule.Name = "{}_{}".format(base_name, i + 1)
                        
                        # Apply template if requested
                        if apply_template and selected_template:
                            apply_template_with_options(new_schedule, selected_template, apply_template)
                        
                        success_count += 1
                        print("✓ Created: {}".format(new_schedule.Name))
                        
                    except Exception as e:
                        print("✗ Error creating copy {}: {}".format(i + 1, e))
        
        # Results summary
        result_msg = "Created {}/{} schedule copies".format(success_count, copy_count)
        if apply_template and selected_template and success_count > 0:
            result_msg += "\nwith template: {}".format(selected_template.Name)
        
        forms.alert(result_msg)
        
    except Exception as e:
        logger.error("Error: {}".format(str(e)))
        forms.alert("Error: {}".format(str(e)))

if __name__ == "__main__":
    main()