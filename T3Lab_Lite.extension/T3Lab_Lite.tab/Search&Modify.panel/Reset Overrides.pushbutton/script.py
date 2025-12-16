from pyrevit import revit, DB

view = revit.active_view
doc = revit.doc
collector = DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElementIds()

override = DB.OverrideGraphicSettings()

def get_element_edges(element, view):
    """Get all edges from an element's geometry"""
    edges = []
    try:
        options = DB.Options()
        options.View = view
        options.ComputeReferences = True
        geom = element.get_Geometry(options)

        if geom:
            for geom_obj in geom:
                if isinstance(geom_obj, DB.GeometryInstance):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        for inst_obj in inst_geom:
                            if isinstance(inst_obj, DB.Solid):
                                for edge in inst_obj.Edges:
                                    edges.append(edge)
                elif isinstance(geom_obj, DB.Solid):
                    for edge in geom_obj.Edges:
                        edges.append(edge)
    except:
        pass
    return edges

with revit.Transaction ("Reset Overrides"):
    for el_id in collector:
        # Reset element overrides
        view.SetElementOverrides(el_id, override)

        # Reset linework overrides and set linework by category
        element = doc.GetElement(el_id)
        if element:
            # Get element's category graphics style
            category = element.Category
            graphics_style_id = DB.ElementId.InvalidElementId

            if category and category.Id.IntegerValue > 0:
                try:
                    # Get the category's graphics style for lines
                    graphics_style_category = doc.Settings.Categories.get_Item(category.Name)
                    if graphics_style_category:
                        subcats = graphics_style_category.SubCategories
                        # Use the parent category's graphics style (by category)
                        graphics_style_id = graphics_style_category.GetGraphicsStyle(DB.GraphicsStyleType.Projection).Id
                except:
                    pass

            edges = get_element_edges(element, view)
            for edge in edges:
                try:
                    if edge.Reference:
                        # Remove line pattern override
                        view.RemoveLinePatternOverride(edge.Reference)

                        # Set linework to by category
                        if graphics_style_id.IntegerValue != -1:
                            doc.SetLineworkGraphicsStyle(edge.Reference, graphics_style_id)
                except:
                    pass