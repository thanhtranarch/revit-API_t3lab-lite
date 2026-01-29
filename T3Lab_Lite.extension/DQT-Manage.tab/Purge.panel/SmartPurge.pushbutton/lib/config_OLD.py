# -*- coding: utf-8 -*-
"""
Split Floor Tool
Splits a floor with multiple disconnected boundaries into separate individual floors.
Copyright (c) 2025 by Dang Quoc Truong (DQT)
"""

__title__ = "Split Floor"
__author__ = "DQT"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, DB, UI, forms
import math

doc = revit.doc
uidoc = revit.uidoc


def get_curve_loops_from_floor(floor):
    """Extract all curve loops from a floor's sketch"""
    curve_loops = []
    
    # Get the floor's sketch
    sketch_filter = DB.ElementClassFilter(DB.Sketch)
    dependent_elements = floor.GetDependentElements(sketch_filter)
    
    if dependent_elements.Count > 0:
        sketch_id = dependent_elements[0]
        sketch_obj = doc.GetElement(sketch_id)
        
        if sketch_obj:
            # Get all curve arrays from sketch
            profile = sketch_obj.Profile
            
            for curve_array in profile:
                curve_loop = CurveLoop()
                for curve in curve_array:
                    curve_loop.Append(curve)
                curve_loops.append(curve_loop)
    
    return curve_loops


def get_boundary_groups(curve_loops):
    """Group curve loops that are close together (connected boundaries)"""
    if not curve_loops:
        return []
    
    # Calculate centroids for each loop
    centroids = []
    for loop in curve_loops:
        points = []
        for curve in loop:
            points.append(curve.GetEndPoint(0))
        
        # Calculate centroid
        sum_x = sum(p.X for p in points)
        sum_y = sum(p.Y for p in points)
        sum_z = sum(p.Z for p in points)
        centroid = XYZ(sum_x / len(points), sum_y / len(points), sum_z / len(points))
        centroids.append(centroid)
    
    # Group loops by proximity
    groups = []
    used = [False] * len(curve_loops)
    tolerance = 1.0  # 1 foot tolerance for grouping
    
    for i in range(len(curve_loops)):
        if used[i]:
            continue
        
        group = [curve_loops[i]]
        used[i] = True
        
        # Find all loops close to this one
        for j in range(i + 1, len(curve_loops)):
            if used[j]:
                continue
            
            # Check if centroids are close
            distance = centroids[i].DistanceTo(centroids[j])
            if distance < tolerance:
                group.append(curve_loops[j])
                used[j] = True
        
        groups.append(group)
    
    return groups


def check_if_loop_is_inside(inner_loop, outer_loop):
    """Check if inner_loop is inside outer_loop"""
    # Get a point from inner loop
    test_point = None
    for curve in inner_loop:
        test_point = curve.GetEndPoint(0)
        break
    
    if not test_point:
        return False
    
    # Count intersections with a ray from test point
    ray_end = XYZ(test_point.X + 10000, test_point.Y, test_point.Z)
    ray = Line.CreateBound(test_point, ray_end)
    
    intersection_count = 0
    for curve in outer_loop:
        # Simplified intersection check
        try:
            result = curve.Intersect(ray)
            if result == DB.SetComparisonResult.Overlap:
                intersection_count += 1
        except:
            pass
    
    # Odd number of intersections means point is inside
    return intersection_count % 2 == 1


def organize_loops_with_holes(curve_loops):
    """Organize loops into outer boundaries and their holes"""
    if len(curve_loops) == 1:
        return [(curve_loops[0], [])]
    
    # Calculate areas to identify outer vs inner loops
    loop_areas = []
    for loop in curve_loops:
        try:
            area = loop.GetExactLength()  # Approximate using perimeter
            loop_areas.append(area)
        except:
            loop_areas.append(0)
    
    # Sort by area (largest first - likely outer boundaries)
    sorted_indices = sorted(range(len(curve_loops)), key=lambda i: loop_areas[i], reverse=True)
    
    result = []
    used = [False] * len(curve_loops)
    
    for i in sorted_indices:
        if used[i]:
            continue
        
        outer_loop = curve_loops[i]
        holes = []
        used[i] = True
        
        # Find holes for this outer loop
        for j in sorted_indices:
            if used[j]:
                continue
            
            if check_if_loop_is_inside(curve_loops[j], outer_loop):
                holes.append(curve_loops[j])
                used[j] = True
        
        result.append((outer_loop, holes))
    
    return result


def get_floor_point_modifications(floor):
    """Get all point modifications (vertex elevations) from a floor"""
    modifications = []
    
    try:
        # Get the floor's SlabShapeEditor
        slab_shape_editor = floor.SlabShapeEditor
        if slab_shape_editor:
            # Get all vertices
            vertices = slab_shape_editor.SlabShapeVertices
            
            for vertex in vertices:
                # Get vertex position
                position = vertex.Position
                modifications.append({
                    'position': position,
                    'x': position.X,
                    'y': position.Y,
                    'z': position.Z
                })
            
            print("  Found {} point modifications".format(len(modifications)))
    except Exception as e:
        print("  No point modifications found or error: {}".format(str(e)))
    
    return modifications


def create_floor_from_curves(floor_type, level, outer_loop, holes=None):
    """Create a new floor from curve loops (without point modifications)"""
    
    # Get names safely
    try:
        type_name = floor_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    except:
        type_name = "Unknown Type"
    
    try:
        level_name = level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString()
    except:
        level_name = "Unknown Level"
    
    print("  Floor type: {} (ID: {})".format(type_name, floor_type.Id))
    print("  Level: {} (ID: {})".format(level_name, level.Id))
    
    # Count curves
    curve_count = sum(1 for _ in outer_loop)
    print("  Added {} curves to outer boundary".format(curve_count))
    
    # Create the floor using Floor.Create with multiple CurveLoops
    print("  Calling Floor.Create...")
    
    # Create a list of CurveLoops (outer + holes)
    curve_loops = []
    curve_loops.Add(outer_loop)
    
    # Add holes if any
    if holes and len(holes) > 0:
        print("  Adding {} holes to CurveLoop list".format(len(holes)))
        for hole in holes:
            curve_loops.Add(hole)
    
    # Create floor with all loops at once
    new_floor = DB.Floor.Create(doc, curve_loops, floor_type.Id, level.Id)
    print("  Floor created successfully: {}".format(new_floor.Id))
    
    return new_floor


def split_floor(floor):
    """Split a floor with multiple boundaries into separate floors"""
    # Get floor properties
    floor_type_id = floor.GetTypeId()
    floor_type = doc.GetElement(floor_type_id)
    level_id = floor.LevelId
    level = doc.GetElement(level_id)
    
    # Get point modifications from original floor
    print("\nExtracting point modifications from original floor...")
    point_modifications = get_floor_point_modifications(floor)
    
    # Warn user if there are point modifications
    if point_modifications and len(point_modifications) > 0:
        print("\nWARNING: Floor has {} point modifications (vertex height adjustments)".format(
            len(point_modifications)))
        
        result = forms.alert(
            "This floor has {} modified vertices (height adjustments).\n\n".format(len(point_modifications)) +
            "NOTE: These vertex modifications CANNOT be automatically transferred to the split floors.\n\n" +
            "After splitting:\n" +
            "- The new floors will be created with FLAT surfaces\n" +
            "- You will need to manually re-apply vertex modifications using 'Modify Sub Elements'\n\n" +
            "Do you want to continue with the split?",
            title="Point Modifications Warning",
            ok=True,
            cancel=True
        )
        
        if not result:
            print("  User cancelled due to point modifications")
            return None
    
    # Get all curve loops
    curve_loops = get_curve_loops_from_floor(floor)
    
    if len(curve_loops) <= 1:
        print("  Floor has only one boundary - skipping")
        return None
    
    print("\nFound {} curve loops in floor".format(len(curve_loops)))
    
    # Calculate area for each loop
    loop_data = []
    for i, loop in enumerate(curve_loops):
        # Count curves in loop
        curve_count = sum(1 for _ in loop)
        
        # Get approximate area using bounding box
        points = []
        for curve in loop:
            points.append(curve.GetEndPoint(0))
            points.append(curve.GetEndPoint(1))
        
        if points:
            min_x = min(p.X for p in points)
            max_x = max(p.X for p in points)
            min_y = min(p.Y for p in points)
            max_y = max(p.Y for p in points)
            bbox_area = (max_x - min_x) * (max_y - min_y)
            centroid = XYZ(
                sum(p.X for p in points) / len(points),
                sum(p.Y for p in points) / len(points),
                sum(p.Z for p in points) / len(points)
            )
        else:
            bbox_area = 0
            centroid = XYZ(0, 0, 0)
        
        loop_data.append({
            'index': i,
            'loop': loop,
            'curve_count': curve_count,
            'area': bbox_area,
            'centroid': centroid
        })
        
        print("Loop {}: {} curves, area = {:.2f}".format(i, curve_count, bbox_area))
    
    # Sort by area (largest first)
    loop_data.sort(key=lambda x: x['area'], reverse=True)
    
    # Check if loops are inside each other
    print("\nChecking for inside/outside relationships...")
    
    is_hole = [False] * len(loop_data)
    parent_of = [-1] * len(loop_data)
    
    for i in range(len(loop_data)):
        for j in range(len(loop_data)):
            if i == j:
                continue
            
            # Check if loop i is inside loop j
            if check_if_loop_is_inside(loop_data[i]['loop'], loop_data[j]['loop']):
                # Loop i is inside loop j
                # If j is bigger than i, then i might be a hole of j
                if loop_data[j]['area'] > loop_data[i]['area']:
                    is_hole[i] = True
                    parent_of[i] = j
                    print("  Loop {} is inside Loop {} (hole of parent)".format(
                        loop_data[i]['index'], loop_data[j]['index']))
                    break
    
    # Build main boundaries with their holes
    main_boundaries = []
    for i, data in enumerate(loop_data):
        if not is_hole[i]:
            # This is a main boundary, find its holes
            holes = []
            for j in range(len(loop_data)):
                if parent_of[j] == i:
                    holes.append(loop_data[j]['loop'])
            
            main_boundaries.append({
                'loop': data['loop'],
                'holes': holes,
                'index': data['index'],
                'area': data['area'],
                'curve_count': data['curve_count']
            })
    
    print("\nAnalysis:")
    print("  Total loops: {}".format(len(loop_data)))
    print("  Main boundaries: {}".format(len(main_boundaries)))
    print("  Holes: {}".format(sum(is_hole)))
    
    # Show hole assignments
    for idx, mb in enumerate(main_boundaries):
        print("  Main boundary {} has {} hole(s)".format(idx, len(mb['holes'])))
    
    if len(main_boundaries) <= 1:
        print("\nOnly one main boundary found - this is a single floor with holes - skipping split")
        return None
    
    # Create separate floors for each main boundary
    print("\nCreating {} separate floors".format(len(main_boundaries)))
    
    # Start transaction
    t = Transaction(doc, "Split Floor {} into {} Floors".format(floor.Id, len(main_boundaries)))
    t.Start()
    
    try:
        new_floors = []
        
        # Create one floor for each main boundary (with its holes)
        for idx, data in enumerate(main_boundaries):
            print("\nCreating floor {} (area: {:.2f}, {} curves, {} holes)".format(
                idx + 1, data['area'], data['curve_count'], len(data['holes'])))
            try:
                new_floor = create_floor_from_curves(
                    floor_type, 
                    level, 
                    data['loop'], 
                    data['holes'] if len(data['holes']) > 0 else None
                )
                new_floors.append(new_floor)
            except Exception as e:
                print("  WARNING: Failed to create floor {}: {}".format(idx + 1, str(e)))
                # Continue with other floors
        
        # Delete original floor
        doc.Delete(floor.Id)
        
        t.Commit()
        
        return new_floors
        
    except Exception as e:
        t.RollBack()
        import traceback
        print("\n=== ERROR IN SPLIT_FLOOR ===")
        print(traceback.format_exc())
        raise e



def main():
    """Main function"""
    try:
        # Prompt user to select floors
        result = forms.alert(
            "Select multiple floors with disconnected boundaries to split.\n\n"
            "Click OK to start selecting floors.\n"
            "Press ESC or Finish when done.",
            title="Split Floor Tool",
            ok=True,
            cancel=True
        )
        
        if not result:
            return
        
        # Select multiple floors
        selected_floors = []
        try:
            # Use PickObjects for multiple selection
            references = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select floors to split (Press ESC or Finish when done)"
            )
            
            for ref in references:
                element = doc.GetElement(ref.ElementId)
                if isinstance(element, Floor):
                    selected_floors.append(element)
                else:
                    print("Skipping non-floor element: {} (ID: {})".format(
                        element.Category.Name if element.Category else "Unknown",
                        element.Id
                    ))
        except:
            # User cancelled
            return
        
        if not selected_floors:
            forms.alert("No floors selected.", exitscript=True)
        
        print("\n" + "="*60)
        print("SPLIT FLOOR TOOL - Processing {} floor(s)".format(len(selected_floors)))
        print("="*60)
        
        # Process each floor
        total_created = 0
        successful_splits = 0
        failed_splits = 0
        
        for idx, floor in enumerate(selected_floors):
            print("\n" + "-"*60)
            print("Processing Floor {}/{} (ID: {})".format(idx + 1, len(selected_floors), floor.Id))
            print("-"*60)
            
            try:
                new_floors = split_floor(floor)
                if new_floors:
                    total_created += len(new_floors)
                    successful_splits += 1
                    print("SUCCESS: Created {} floors from this split".format(len(new_floors)))
            except Exception as e:
                failed_splits += 1
                print("FAILED: {}".format(str(e)))
                # Continue with next floor
                continue
        
        # Show summary
        print("\n" + "="*60)
        print("SUMMARY")
        print("="*60)
        print("Floors processed: {}".format(len(selected_floors)))
        print("Successful splits: {}".format(successful_splits))
        print("Failed splits: {}".format(failed_splits))
        print("Total new floors created: {}".format(total_created))
        print("="*60)
        
        summary_message = (
            "Split Floor Complete!\n\n"
            "Processed: {} floor(s)\n"
            "Successful: {}\n"
            "Failed: {}\n"
            "Total new floors created: {}"
        ).format(len(selected_floors), successful_splits, failed_splits, total_created)
        
        forms.alert(summary_message, title="Split Floor Summary")
        
    except Exception as e:
        import traceback
        print("\n=== MAIN ERROR ===")
        print(traceback.format_exc())
        forms.alert("Error: {}".format(str(e)), exitscript=True)


if __name__ == "__main__":
    main()