# -*- coding: utf-8 -*-
"""Material Statistics - With Element Finder (Revit 2024-2027)"""

from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
from collections import defaultdict
from System.Collections.Generic import List

doc = revit.doc
logger = script.get_logger()
output = script.get_output()


def _eid_int(element_id):
    """Get integer value of ElementId, compatible Revit 2024-2027.
    Revit 2025+ removed ElementId.IntegerValue in favor of .Value."""
    try:
        return element_id.Value          # Revit 2025+
    except AttributeError:
        return element_id.IntegerValue   # Revit 2024


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
    except Exception:
        return "Unnamed", "Unknown"


def _element_uses_material(element, material_int):
    """Check if an element (instance or its type) uses the material.
    Checks instance materials, type materials, and compound-structure layers."""
    try:
        # 1) Instance-level materials (includes paint/compound when available)
        if hasattr(element, 'GetMaterialIds'):
            for include_paint in (True, False):
                try:
                    for mid in element.GetMaterialIds(include_paint):
                        if mid and _eid_int(mid) == material_int:
                            return True
                except Exception:
                    pass

        # 2) Type-level materials
        try:
            type_id = element.GetTypeId()
            if type_id and _eid_int(type_id) > 0:
                etype = doc.GetElement(type_id)
                if etype is not None:
                    if hasattr(etype, 'GetMaterialIds'):
                        try:
                            for mid in etype.GetMaterialIds(False):
                                if mid and _eid_int(mid) == material_int:
                                    return True
                        except Exception:
                            pass
                    # Compound structure layers (Wall/Floor/Roof/Ceiling types)
                    if hasattr(etype, 'GetCompoundStructure'):
                        try:
                            cs = etype.GetCompoundStructure()
                            if cs is not None:
                                for layer in cs.GetLayers():
                                    mid = layer.MaterialId
                                    if mid and _eid_int(mid) == material_int:
                                        return True
                        except Exception:
                            pass
        except Exception:
            pass
    except Exception:
        pass
    return False


def find_elements_by_material(material):
    """Find elements using specific material"""
    if not material or not material.IsValidObject:
        return []

    elements_found = []
    material_int = _eid_int(material.Id)

    try:
        categories_to_check = [
            DB.BuiltInCategory.OST_Walls,
            DB.BuiltInCategory.OST_Floors,
            DB.BuiltInCategory.OST_Doors,
            DB.BuiltInCategory.OST_Windows,
            DB.BuiltInCategory.OST_StructuralFraming,
            DB.BuiltInCategory.OST_StructuralColumns,
            DB.BuiltInCategory.OST_Ceilings,
            DB.BuiltInCategory.OST_Roofs,
            DB.BuiltInCategory.OST_StructuralFoundation,
            DB.BuiltInCategory.OST_GenericModel
        ]

        seen_ids = set()
        for category in categories_to_check:
            try:
                collector = DB.FilteredElementCollector(doc)
                elements = collector.OfCategory(category).WhereElementIsNotElementType().ToElements()
                for element in elements:
                    try:
                        eid = _eid_int(element.Id)
                        if eid in seen_ids:
                            continue
                        if _element_uses_material(element, material_int):
                            elements_found.append(element)
                            seen_ids.add(eid)
                    except Exception:
                        continue
            except Exception:
                continue

        return elements_found

    except Exception as e:
        logger.error("Error finding elements: {}".format(e))
        return []


def show_elements_for_selected_material():
    """Display and select material, then find elements"""
    try:
        materials = get_all_materials()
        if not materials:
            forms.alert("No materials found in project!")
            return

        class MaterialChoice(object):
            def __init__(self, material):
                self.material = material
                self.name, self.category = get_material_info(material)
                self.display_name = "{} ({})".format(self.name, self.category)

        material_choices = [MaterialChoice(mat) for mat in materials]
        material_choices.sort(key=lambda x: x.display_name)

        selected_choice = forms.SelectFromList.show(
            material_choices,
            title="Select Material to Find Elements",
            button_name='Find Elements',
            name_attr='display_name'
        )

        if not selected_choice:
            return

        selected_material = selected_choice.material
        material_name = selected_choice.name

        elements = find_elements_by_material(selected_material)

        if not elements:
            forms.alert("No elements found using material '{}'".format(material_name))
            return

        element_choices = []
        for i, element in enumerate(elements):
            try:
                try:
                    elem_name = element.Name if element.Name else "Unnamed"
                except Exception:
                    elem_name = "Unnamed"
                elem_type = element.GetType().Name
                elem_category = element.Category.Name if element.Category else "Unknown"
                display_text = "{}. {} - {} - {} [ID: {}]".format(
                    i + 1, elem_name, elem_type, elem_category, _eid_int(element.Id))
                element_choices.append(display_text)
            except Exception:
                element_choices.append("{}. Unknown Element".format(i + 1))

        forms.SelectFromList.show(
            element_choices,
            title="Elements using '{}' (Found: {})".format(material_name, len(elements)),
            button_name='Close'
        )

        result = forms.alert(
            "Found {} elements using material '{}'.\n\nSelect these elements in Revit?".format(
                len(elements), material_name),
            yes=True, no=True
        )

        if result:
            try:
                id_list = List[DB.ElementId]()
                for element in elements:
                    id_list.Add(element.Id)
                revit.uidoc.Selection.SetElementIds(id_list)
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

        category_count = defaultdict(int)
        for material in materials:
            name, category = get_material_info(material)
            category_count[category] += 1

        output.print_md("# **MATERIAL STATISTICS**")
        output.print_md("---")
        output.print_md("**Total Materials:** {}".format(len(materials)))
        output.print_md("**Project:** {}".format(doc.Title))

        output.print_md("## **STATISTICS BY CATEGORY**")
        for category, count in sorted(category_count.items()):
            output.print_md("- **{}:** {}".format(category, count))

        output.print_md("## **MATERIAL LIST**")
        output.print_md("---")

        materials_by_category = defaultdict(list)
        for material in materials:
            name, category = get_material_info(material)
            materials_by_category[category].append((name, material))

        for category in sorted(materials_by_category.keys()):
            output.print_md("### **{}**".format(category))
            materials_list = sorted(materials_by_category[category], key=lambda x: x[0])
            for idx, (name, material) in enumerate(materials_list, 1):
                output.print_md("{}. **{}**".format(idx, name))
            output.print_md("---")

        output.print_md("### **USER GUIDE**")
        output.print_md("- Use 'Find Elements by Material' to find elements using a specific material")
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

        choice = forms.CommandSwitchWindow.show(
            ['Create Material Report', 'Find Elements by Material'],
            message='Select function:'
        )

        if choice == 'Create Material Report':
            create_material_report()
            forms.alert("Report completed! Check output window.")
        elif choice == 'Find Elements by Material':
            show_elements_for_selected_material()

    except Exception as e:
        forms.alert("Error: {}".format(e))


if __name__ == "__main__":
    main()