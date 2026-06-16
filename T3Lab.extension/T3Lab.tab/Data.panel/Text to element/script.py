# -*- coding: utf-8 -*-
"""Transfer Text Value to Intersecting Elements

This tool transfers the text content from Text Notes to elements 
that intersect with them. Useful for automatically assigning 
room names, area values, or other annotations to model elements.

Author: DQT
Compatible with: Revit 2019+
"""

__title__ = "Text to\nElement"
__doc__ = "Transfer text values to elements that intersect with the text annotation"
__author__ = "DQT"


# =============================================================================
# CONFIGURATION
# =============================================================================

TARGET_CATEGORIES = [
    ("Walls", "OST_Walls"),
    ("Floors", "OST_Floors"),
    ("Ceilings", "OST_Ceilings"),
    ("Roofs", "OST_Roofs"),
    ("Rooms", "OST_Rooms"),
    ("Areas", "OST_Areas"),
    ("Doors", "OST_Doors"),
    ("Windows", "OST_Windows"),
    ("Furniture", "OST_Furniture"),
    ("Generic Models", "OST_GenericModel"),
    ("Columns", "OST_Columns"),
    ("Structural Columns", "OST_StructuralColumns"),
    ("Structural Framing", "OST_StructuralFraming"),
    ("Mechanical Equipment", "OST_MechanicalEquipment"),
    ("Plumbing Fixtures", "OST_PlumbingFixtures"),
    ("Casework", "OST_Casework"),
    ("Detail Items", "OST_DetailComponents"),
]

# Conversion factor: 1 foot = 304.8 mm
MM_TO_FEET = 1.0 / 304.8


# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    # Import everything inside main to avoid load-time errors
    from pyrevit import revit, DB
    from pyrevit import script
    from pyrevit import forms
    
    from Autodesk.Revit.DB import (
        FilteredElementCollector, BuiltInCategory, BuiltInParameter,
        TextNote, Transaction, ElementId, StorageType
    )
    from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
    
    doc = revit.doc
    uidoc = revit.uidoc
    output = script.get_output()
    
    # Selection filter class
    class TextNoteSelectionFilter(ISelectionFilter):
        def AllowElement(self, element):
            return isinstance(element, TextNote)
        def AllowReference(self, reference, position):
            return False
    
    # Helper functions
    def get_category_enum(cat_name):
        return getattr(BuiltInCategory, cat_name, None)
    
    def get_text_content(text_note):
        try:
            return text_note.Text.strip()
        except:
            return ""
    
    def boxes_intersect(bb1, bb2, tolerance_feet=0.5):
        if bb1 is None or bb2 is None:
            return False
        x_overlap = (bb1.Min.X - tolerance_feet <= bb2.Max.X and 
                     bb1.Max.X + tolerance_feet >= bb2.Min.X)
        y_overlap = (bb1.Min.Y - tolerance_feet <= bb2.Max.Y and 
                     bb1.Max.Y + tolerance_feet >= bb2.Min.Y)
        return x_overlap and y_overlap
    
    def get_element_name(element):
        try:
            if element.Name:
                return element.Name
        except:
            pass
        try:
            elem_type = doc.GetElement(element.GetTypeId())
            if elem_type:
                param = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    return param.AsString()
        except:
            pass
        return "Element {}".format(element.Id.IntegerValue)
    
    def get_text_parameters(elements):
        """Get all text-writable parameters from elements (instance + shared)"""
        param_names = set()
        
        for elem in elements[:20]:  # Sample first 20 elements
            try:
                for param in elem.Parameters:
                    try:
                        # Only get text parameters that are writable
                        if param.StorageType == StorageType.String:
                            if not param.IsReadOnly:
                                name = param.Definition.Name
                                if name:
                                    param_names.add(name)
                    except:
                        continue
            except:
                continue
        
        # Sort alphabetically
        sorted_params = sorted(list(param_names))
        return sorted_params
    
    def set_parameter_value(element, param_name, value):
        # Try instance parameter by name
        param = element.LookupParameter(param_name)
        if param and not param.IsReadOnly:
            try:
                if param.StorageType == StorageType.String:
                    param.Set(value)
                    return True
            except:
                pass
        return False
    
    # Check active view
    view = doc.ActiveView
    if view is None:
        forms.alert("Please open a view first.", title="No Active View")
        return
    
    # Step 1: Select text notes
    text_choice = forms.SelectFromList.show(
        ["Select text notes from model", "Use all text notes in current view"],
        title="Step 1: Text Notes",
        button_name="Next",
        multiselect=False
    )
    
    if not text_choice:
        return
    
    text_notes = []
    
    if "Select" in text_choice:
        try:
            picked = uidoc.Selection.PickObjects(
                ObjectType.Element,
                TextNoteSelectionFilter(),
                "Select text notes (ESC to finish)"
            )
            for ref in picked:
                elem = doc.GetElement(ref.ElementId)
                if elem:
                    text_notes.append(elem)
        except:
            return
    else:
        collector = FilteredElementCollector(doc, view.Id).OfClass(TextNote).ToElements()
        text_notes = list(collector)
    
    if not text_notes:
        forms.alert("No text notes found.", title="No Text Notes")
        return
    
    # Step 2: Select category
    category_names = [name for name, cat in TARGET_CATEGORIES]
    selected_cat_name = forms.SelectFromList.show(
        category_names,
        title="Step 2: Select Target Category",
        button_name="Next",
        multiselect=False
    )
    
    if not selected_cat_name:
        return
    
    # Get category enum
    cat_enum_name = None
    for name, cat in TARGET_CATEGORIES:
        if name == selected_cat_name:
            cat_enum_name = cat
            break
    
    category = get_category_enum(cat_enum_name)
    if not category:
        forms.alert("Category not found.", title="Error")
        return
    
    # Get elements
    try:
        collector = FilteredElementCollector(doc, view.Id)\
            .OfCategory(category)\
            .WhereElementIsNotElementType()\
            .ToElements()
        target_elements = list(collector)
    except:
        target_elements = []
    
    if not target_elements:
        forms.alert("No {} found in current view.".format(selected_cat_name), title="No Elements")
        return
    
    # Step 3: Get available parameters from elements
    available_params = get_text_parameters(target_elements)
    
    if not available_params:
        forms.alert(
            "No writable text parameters found for {}.\n\n"
            "Make sure elements have text parameters.".format(selected_cat_name),
            title="No Parameters"
        )
        return
    
    selected_param = forms.SelectFromList.show(
        available_params,
        title="Step 3: Select Target Parameter ({} found)".format(len(available_params)),
        button_name="Next",
        multiselect=False
    )
    
    if not selected_param:
        return
    
    # Step 4: Tolerance in millimeters
    tolerance_str = forms.ask_for_string(
        prompt="Intersection tolerance in millimeters:",
        title="Step 4: Tolerance (mm)",
        default="150"
    )
    
    try:
        tolerance_mm = float(tolerance_str) if tolerance_str else 150.0
    except:
        tolerance_mm = 150.0
    
    # Convert mm to feet for internal calculation
    tolerance_feet = tolerance_mm * MM_TO_FEET
    
    # Find intersections
    transfer_list = []
    
    for tn in text_notes:
        text_content = get_text_content(tn)
        if not text_content:
            continue
        
        text_bb = tn.get_BoundingBox(view)
        if not text_bb:
            continue
        
        for elem in target_elements:
            elem_bb = elem.get_BoundingBox(view)
            if boxes_intersect(text_bb, elem_bb, tolerance_feet):
                transfer_list.append({
                    'text': text_content,
                    'element': elem,
                    'name': get_element_name(elem)
                })
    
    if not transfer_list:
        forms.alert(
            "No intersections found.\n\n"
            "Tips:\n"
            "- Check text overlaps elements\n"
            "- Try larger tolerance (current: {} mm)".format(int(tolerance_mm)),
            title="No Intersections"
        )
        return
    
    # Preview
    preview = "Found {} transfers:\n\n".format(len(transfer_list))
    for i, t in enumerate(transfer_list[:10]):
        txt = t['text'][:20] + "..." if len(t['text']) > 20 else t['text']
        preview += '{}. "{}" -> {}\n'.format(i+1, txt, t['name'])
    if len(transfer_list) > 10:
        preview += "... and {} more\n".format(len(transfer_list) - 10)
    preview += "\nParameter: {}\nTolerance: {} mm\n\nProceed?".format(selected_param, int(tolerance_mm))
    
    if not forms.alert(preview, title="Confirm", yes=True, no=True):
        return
    
    # Transfer
    success = 0
    fail = 0
    
    with revit.Transaction("Transfer Text to Elements"):
        for t in transfer_list:
            if set_parameter_value(t['element'], selected_param, t['text']):
                success += 1
            else:
                fail += 1
    
    # Results
    msg = "Done!\n\nSuccess: {}\nFailed: {}".format(success, fail)
    forms.alert(msg, title="Results")
    
    output.print_md("# Transfer Results")
    output.print_md("**Category:** {}".format(selected_cat_name))
    output.print_md("**Parameter:** {}".format(selected_param))
    output.print_md("**Tolerance:** {} mm".format(int(tolerance_mm)))
    output.print_md("**Success:** {} | **Failed:** {}".format(success, fail))


if __name__ == "__main__":
    main()