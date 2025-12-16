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

        # Reset linework overrides
        element = doc.GetElement(el_id)
        if element:
            edges = get_element_edges(element, view)
            for edge in edges:
                try:
                    if edge.Reference:
                        view.RemoveLinePatternOverride(edge.Reference)
                except:
                    pass