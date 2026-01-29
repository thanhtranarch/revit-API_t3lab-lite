# -*- coding: utf-8 -*-
"""Material Statistics - With Element Finder"""

from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
from collections import defaultdict
import System
from System.Collections.Generic import List

# Get current Revit document
doc = revit.doc
logger = script.get_logger()
output = script.get_output()

def get_all_materials():
    """Get all materials"""
    try:
        collector = DB.FilteredElementCollector(doc)
        materials = collector.OfClass(DB.Material).ToElements()
        return [mat for mat in materials if mat and mat.IsValidObject]
    except Exception as e:
        logger.error("Error getting materials: {}".format(e))
        return []

def get_material_info(material):
    """Get material information"""
    try:
        name = material.Name if material.Name else "Unnamed"
        category = material.MaterialCategory if hasattr(material, 'MaterialCategory') and material.MaterialCategory else "Unknown"
        return name, category
    except:
        return "Unnamed", "Unknown"

def find_elements_by_material(material):
    """Find elements using specific material"""
    if not material or not material.IsValidObject:
        return []
    
    elements_found = []
    material_id = material.Id
    
    try:
        # Categories to search
        categories_to_check = [
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Doors,
            DB.BuiltInCategory.OST_Windows,
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs
        ]
        
        for category in categories_to_check:
            try:
                collector = DB.FilteredElementCollector(doc)
                elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()
                
                for element in elements:
                    try:
                        if hasattr(element, 'GetMaterialIds'):
                            mat_ids = element.GetMaterialIds(False)
                            for mat_id in mat_ids:
                                if mat_id and mat_id.IntegerValue == material_id.IntegerValue:
                                    elements_found.append(element)
                                    break
                    except:
                        continue
            except:
                continue
                
        return elements_found
        
    except Exception as e:
        logger.error("Error finding elements: {}".format(e))
        return []

def show_elements_for_selected_material():
    """Display and select material, then find elements"""
    try:
        # Get all materials
        materials = get_all_materials()
        if not materials:
            forms.alert("No materials found in project!")
            return
        
        # Create selection list with simple class
        class MaterialChoice:
            def __init__(self, material):
                self.material = material
                self.name, self.category = get_material_info(material)
                self.display_name = "{} ({})".format(self.name, self.category)
        
        material_choices = [MaterialChoice(mat) for mat in materials]
        material_choices.sort(key=lambda x: x.display_name)
        
        # Display material selection dialog
        selected_choice = forms.SelectFromList.show(
            material_choices,
            title="Select Material to Find Elements",
            button_name='Find Elements',
            name_attr='display_name'
        )
        
        if not selected_choice:
            forms.alert("No material selected!")
            return
        
        selected_material = selected_choice.material
        material_name = selected_choice.name
        
        # Find elements using this material
        forms.alert("Searching for elements using material '{}'...".format(material_name))
        
        elements = find_elements_by_material(selected_material)
        
        if not elements:
            forms.alert("No elements found using material '{}'".format(material_name))
            return
        
        # Create element list for display
        element_choices = []
        for i, element in enumerate(elements):
            try:
                elem_name = element.Name if element.Name else "Unnamed"
                elem_type = element.GetType().Name
                elem_category = element.Category.Name if element.Category else "Unknown"
                
                display_text = "{}. {} - {} - {}".format(i+1, elem_name, elem_type, elem_category)
                element_choices.append(display_text)
            except:
                element_choices.append("{}. Unknown Element".format(i+1))
        
        # Display element list
        forms.SelectFromList.show(
            element_choices,
            title="Elements using: {}\n(Found: {} elements)".format(material_name, len(elements)),
            button_name='Close'
        )
        
        # Ask user if they want to select elements
        result = forms.alert(
            "Found {} elements using material '{}'\n\nDo you want to select these elements in Revit?".format(len(elements), material_name),
            yes=True,
            no=True
        )
        
        if result:
            # Select elements in Revit
            try:
                element_ids = [element.Id for element in elements]
                
                # Create List[ElementId] properly
                id_list = List[DB.ElementId]()
                for element_id in element_ids:
                    id_list.Add(element_id)
                
                selection = revit.uidoc.Selection
                selection.SetElementIds(id_list)
                forms.alert("Selected {} elements in Revit!".format(len(elements)))
            except Exception as e:
                logger.error("Selection error: {}".format(str(e)))
                forms.alert("Elements found but unable to select: {}".format(str(e)))
        
    except Exception as e:
        logger.error("Error in show_elements: {}".format(e))
        forms.alert("Error: {}".format(e))

def create_material_report():
    """Create simple materials report"""
    try:
        materials = get_all_materials()
        if not materials:
            forms.alert("No materials found")
            return

        # Statistics by category
        category_count = defaultdict(int)
        for material in materials:
            name, category = get_material_info(material)
            category_count[category] += 1

        # Display results
        output.print_md("# **MATERIAL STATISTICS**")
        output.print_md("---")
        output.print_md("**Total Materials:** {}".format(len(materials)))
        output.print_md("**Project:** {}".format(doc.Title))
        
        output.print_md("## **STATISTICS BY CATEGORY**")
        for category, count in sorted(category_count.items()):
            output.print_md("- **{}:** {}".format(category, count))
        
        output.print_md("## **MATERIAL LIST**")
        output.print_md("---")
        
        # Group materials by category
        materials_by_category = defaultdict(list)
        for material in materials:
            name, category = get_material_info(material)
            materials_by_category[category].append((name, material))
        
        # Display each category
        for category in sorted(materials_by_category.keys()):
            output.print_md("### **{}**".format(category))
            materials_list = sorted(materials_by_category[category], key=lambda x: x[0])
            
            for idx, (name, material) in enumerate(materials_list, 1):
                output.print_md("{}. **{}**".format(idx, name))
            
            output.print_md("---")
        
        output.print_md("### **USER GUIDE**")
        output.print_md("- Use 'Find Elements by Material' button to find elements using specific material")
        output.print_md("- Total: {} materials in {} categories".format(len(materials), len(category_count)))
        
    except Exception as e:
        logger.error("Error creating report: {}".format(e))
        forms.alert("Error: {}".format(e))

def main():
    """Main function"""
    try:
        if not doc or doc.IsFamilyDocument:
            forms.alert("Only works in project documents")
            return
        
        # Display selection menu
        choice = forms.CommandSwitchWindow.show(
            ['Create Material Report', 'Find Elements by Material'],
            message='Select function:'
        )
        
        if choice == 'Create Material Report':
            create_material_report()
            forms.alert("Report completed! Check output window.")
        elif choice == 'Find Elements by Material':
            show_elements_for_selected_material()
        else:
            forms.alert("Cancelled")
        
    except Exception as e:
        forms.alert("Error: {}".format(e))

if __name__ == "__main__":
    main()