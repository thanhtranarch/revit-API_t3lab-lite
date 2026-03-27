# -*- coding: utf-8 -*-
"""
DWGFamilyHelpers - Helper functions for DWG to Generic Model Family conversion.

Provides:
    get_dwg_filename            - Extract source filename from a DWG ImportInstance
    extract_curves_from_dwg     - Pull all valid curves out of DWG geometry
    calculate_center_point      - Compute average centre from a list of curves
    get_dwg_import_center_point - Bounding-box centre of a DWG ImportInstance
    get_outline_boundary        - Pass-through that returns curves as the profile set
    create_generic_model_family - Build a Generic Model family (extrusion + model lines)
    save_and_load_family        - Save .rfa to a temp folder and load it into the project
"""

import os
import re
import tempfile
import uuid

from Autodesk.Revit.DB import (
    Options, GeometryInstance, Line, Arc,
    CurveArray, CurveArrArray,
    XYZ, Transform, Plane,
    SketchPlane, SaveAsOptions,
    FilteredElementCollector, FamilySymbol,
    Transaction,
)


# ---------------------------------------------------------------------------
# get_dwg_filename
# ---------------------------------------------------------------------------

def get_dwg_filename(dwg_instance, doc, write_log):
    """Return the source DWG filename stored in the element type, or the
    instance name as a fallback.  Returns None if nothing is found."""
    try:
        type_id = dwg_instance.GetTypeId()
        if type_id is not None:
            from Autodesk.Revit.DB import ElementId
            if type_id != ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type is not None and hasattr(elem_type, 'Name') and elem_type.Name:
                    write_log("Found DWG filename from type: {}".format(elem_type.Name))
                    return elem_type.Name
    except Exception as e:
        write_log("Warning getting type name: {}".format(str(e)))

    if hasattr(dwg_instance, 'Name') and dwg_instance.Name:
        write_log("Found DWG filename from instance name: {}".format(dwg_instance.Name))
        return dwg_instance.Name

    return None


# ---------------------------------------------------------------------------
# extract_curves_from_dwg
# ---------------------------------------------------------------------------

def extract_curves_from_dwg(dwg_instance, app, write_log):
    """Extract every valid, bound curve from a DWG ImportInstance.

    Translation from the DWG transform is applied (rotation is intentionally
    skipped to avoid orientation mismatches on imported geometry).

    Returns a list of Curve objects in project coordinates.
    """
    curves = []
    skipped = [0]   # mutable container – IronPython 2 has no 'nonlocal'

    write_log("Extracting curves from DWG geometry...")

    # ---- transform: translation only ----
    transform = dwg_instance.GetTotalTransform()
    origin = transform.Origin
    write_log("DWG transform origin: ({:.9f}, {:.9f}, {:.9f})".format(
        origin.X, origin.Y, origin.Z))
    write_log("SKIPPING ROTATION - using translation only to avoid rotation mismatch")

    # ---- minimum curve length ----
    try:
        min_length = app.ShortCurveTolerance
    except Exception:
        min_length = 0.00256   # ~0.78 mm in feet
    write_log("Revit minimum curve tolerance: {:f} feet ({:.4f} mm)".format(
        min_length, min_length * 304.8))

    # ---- geometry options ----
    opt = Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True

    geom_elem = dwg_instance.get_Geometry(opt)
    if geom_elem is None:
        write_log("No geometry found in DWG instance", "WARN")
        return curves

    translation = Transform.CreateTranslation(origin)

    def _process(geo_obj):
        for item in geo_obj:
            if isinstance(item, GeometryInstance):
                write_log("Processing GeometryInstance (DWG block or layer)...")
                inst_geom = item.GetInstanceGeometry()
                if inst_geom is not None:
                    _process(inst_geom)
            else:
                # Accept any Curve subclass (Line, Arc, NurbSpline, etc.)
                try:
                    from Autodesk.Revit.DB import Curve as _Curve
                    if not isinstance(item, _Curve):
                        continue
                except Exception:
                    pass

                try:
                    translated = item.CreateTransformed(translation)

                    if not translated.IsBound:
                        skipped[0] += 1
                        write_log(
                            "Skipped curve: The input curve is not bound.\n"
                            "Parameter name: curve", "DEBUG")
                        continue

                    if translated.Length < min_length:
                        skipped[0] += 1
                        continue

                    curves.append(translated)

                except Exception as e:
                    skipped[0] += 1
                    write_log("Skipped curve: {}".format(str(e)), "DEBUG")

    _process(geom_elem)

    write_log("Extracted {} valid curves from DWG".format(len(curves)))
    if skipped[0] > 0:
        write_log("Skipped {} curves (too short or invalid)".format(skipped[0]), "WARN")

    return curves


# ---------------------------------------------------------------------------
# calculate_center_point
# ---------------------------------------------------------------------------

def calculate_center_point(curves, write_log):
    """Return the arithmetic mean of all curve endpoints."""
    pts = []
    for curve in curves:
        try:
            pts.append(curve.GetEndPoint(0))
            pts.append(curve.GetEndPoint(1))
        except Exception:
            pass

    if not pts:
        return XYZ(0.0, 0.0, 0.0)

    cx = sum(p.X for p in pts) / float(len(pts))
    cy = sum(p.Y for p in pts) / float(len(pts))
    cz = sum(p.Z for p in pts) / float(len(pts))

    center = XYZ(cx, cy, cz)
    write_log("Center point calculated from {} valid curves: ({:.3f}, {:.3f}, {:.1f})".format(
        len(curves), center.X, center.Y, center.Z))
    return center


# ---------------------------------------------------------------------------
# get_dwg_import_center_point
# ---------------------------------------------------------------------------

def get_dwg_import_center_point(dwg_instance, doc, write_log):
    """Return the centre of the DWG element's axis-aligned bounding box,
    or None if the bounding box cannot be retrieved."""
    try:
        bbox = dwg_instance.get_BoundingBox(None)
        if bbox is not None:
            center = (bbox.Min + bbox.Max) * 0.5
            write_log("DWG import center from bounding box: ({:.3f}, {:.3f}, {:.3f})".format(
                center.X, center.Y, center.Z))
            return center
    except Exception as e:
        write_log("Error getting bounding box center: {}".format(str(e)), "WARN")
    return None


# ---------------------------------------------------------------------------
# get_outline_boundary
# ---------------------------------------------------------------------------

def get_outline_boundary(curves, write_log):
    """Return the input curves as the outline profile set.

    The caller uses this list as the 'profile' for the family extrusion.
    Currently this is a pass-through; it can be extended to compute a convex
    hull or other boundary if needed.
    """
    write_log("Creating outline profile from curves...")
    profile_curves = list(curves)
    write_log("Created {} profile curves from {} input curves".format(
        len(profile_curves), len(curves)))
    return profile_curves


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _find_template(app, write_log):
    """Locate the Generic Model family template (.rft) for the running Revit
    version.  Returns the path string or None."""
    version = app.VersionNumber
    write_log("Revit version: {}".format(version))

    candidates = [
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English\Generic Model.rft",
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\Generic Model.rft",
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English-Imperial\Generic Model.rft",
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English_I\Generic Model.rft",
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\Metric Generic Model.rft",
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English\Metric Generic Model.rft",
    ]

    for tmpl in candidates:
        path = tmpl.format(version)
        write_log("Checking template path: {}".format(path))
        if os.path.exists(path):
            write_log("Found template: {}".format(path))
            return path

    write_log("Generic Model template not found in standard paths", "WARN")
    return None


def _get_xy_bounds(curves):
    """Return (min_x, max_x, min_y, max_y) from all curve endpoints."""
    xs, ys = [], []
    for curve in curves:
        try:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            xs.extend([p0.X, p1.X])
            ys.extend([p0.Y, p1.Y])
        except Exception:
            pass
    if not xs:
        return 0.0, 1.0, 0.0, 1.0
    return min(xs), max(xs), min(ys), max(ys)


def _project_curve_to_z(curve, offset_x, offset_y, z):
    """Return a new curve whose XY is shifted by (-offset_x, -offset_y) and
    whose Z is set to *z*.  Falls back to a straight Line between the two
    translated endpoints for non-line curve types."""
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        np0 = XYZ(p0.X - offset_x, p0.Y - offset_y, z)
        np1 = XYZ(p1.X - offset_x, p1.Y - offset_y, z)

        if isinstance(curve, Line):
            return Line.CreateBound(np0, np1)

        if isinstance(curve, Arc):
            # Re-create the arc with the translated centre
            center = curve.Center
            nc = XYZ(center.X - offset_x, center.Y - offset_y, z)
            return Arc.Create(nc, curve.Radius, curve.GetEndParameter(0),
                              curve.GetEndParameter(1), XYZ.BasisX, XYZ.BasisY)

        # Generic fallback: straight line between translated endpoints
        return Line.CreateBound(np0, np1)

    except Exception:
        return None


# ---------------------------------------------------------------------------
# create_generic_model_family
# ---------------------------------------------------------------------------

def create_generic_model_family(curves, outline_curves, app, write_log, write_error):
    """Create a Generic Model family document containing:
        - A solid extrusion whose footprint is the bounding rectangle of *curves*
        - Model lines on the extrusion top face reproducing the DWG detail

    Returns (fam_doc, has_extrusion).  fam_doc is None on failure.
    """
    write_log("Creating Generic Model family...")

    template_path = _find_template(app, write_log)
    if not template_path:
        write_error("Could not find Generic Model family template")
        return None, False

    try:
        fam_doc = app.NewFamilyDocument(template_path)
        write_log("New family document created")
    except Exception as e:
        write_error("Failed to create family document", e)
        return None, False

    has_extrusion = False

    # ---- calculate centre offset (world coords → family origin) ----
    min_x, max_x, min_y, max_y = _get_xy_bounds(curves)
    cx = (min_x + max_x) / 2.0
    cy = (min_y + max_y) / 2.0

    write_log("Center point calculated from {} valid curves: ({:.3f}, {:.3f}, 0.0)".format(
        len(curves), cx, cy))
    write_log("Centering all geometry at origin by offsetting by: ({:.3f}, {:.3f}, 0.0)".format(
        cx, cy))

    t = Transaction(fam_doc, 'Create Family Geometry')
    t.Start()
    try:
        # ---- find or create the XY sketch plane ----
        sketch_plane = None
        collector = FilteredElementCollector(fam_doc).OfClass(SketchPlane)
        for sp in collector:
            try:
                plane = sp.GetPlane()
                if abs(plane.Normal.Z - 1.0) < 0.001:
                    sketch_plane = sp
                    write_log("Using existing sketch plane for extrusion")
                    break
            except Exception:
                pass

        if sketch_plane is None:
            base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero)
            sketch_plane = SketchPlane.Create(fam_doc, base_plane)
            write_log("Created new sketch plane for extrusion")

        # ---- bounding rectangle extrusion ----
        write_log("Creating solid extrusion from outline...")
        write_log("Using bounding rectangle approach (most reliable for DWG imports)")
        write_log("Creating bounding rectangle from curves...")
        write_log("Bounding box: X({:.3f}, {:.3f}), Y({:.3f}, {:.3f})".format(
            min_x, max_x, min_y, max_y))
        write_log("Applied center offset: ({:.3f}, {:.3f})".format(cx, cy))

        half_w = (max_x - min_x) / 2.0
        half_h = (max_y - min_y) / 2.0

        c1 = XYZ(-half_w, -half_h, 0.0)
        c2 = XYZ( half_w, -half_h, 0.0)
        c3 = XYZ( half_w,  half_h, 0.0)
        c4 = XYZ(-half_w,  half_h, 0.0)

        rect = CurveArray()
        rect.Append(Line.CreateBound(c1, c2))
        rect.Append(Line.CreateBound(c2, c3))
        rect.Append(Line.CreateBound(c3, c4))
        rect.Append(Line.CreateBound(c4, c1))
        write_log("Created bounding rectangle with 4 lines")

        profile = CurveArrArray()
        profile.Append(rect)
        write_log("Created bounding rectangle with 4 curves")

        extrusion_depth = 1.0   # feet
        fam_doc.FamilyCreate.NewExtrusion(True, profile, sketch_plane, extrusion_depth)
        write_log("Created extrusion from bounding rectangle: {} feet".format(extrusion_depth))
        has_extrusion = True

        # ---- sketch plane at top surface (Z = extrusion_depth) ----
        top_plane = Plane.CreateByNormalAndOrigin(
            XYZ.BasisZ, XYZ(0.0, 0.0, extrusion_depth))
        top_sp = SketchPlane.Create(fam_doc, top_plane)
        write_log("Created sketch plane at Z={} (top surface) for model lines".format(
            extrusion_depth))

        # ---- model lines on top surface ----
        write_log("Creating {} model lines on top surface (Z={})...".format(
            len(curves), extrusion_depth))
        created = 0
        for curve in curves:
            projected = _project_curve_to_z(curve, cx, cy, extrusion_depth)
            if projected is None:
                continue
            try:
                fam_doc.FamilyCreate.NewModelCurve(projected, top_sp)
                created += 1
            except Exception:
                pass

        write_log("Created {} model lines on top surface (Z={} feet)".format(
            created, extrusion_depth))
        write_log("Family geometry created successfully")

        t.Commit()

    except Exception as e:
        t.RollBack()
        write_error("Error creating family geometry", e)
        return fam_doc, False

    return fam_doc, has_extrusion


# ---------------------------------------------------------------------------
# save_and_load_family
# ---------------------------------------------------------------------------

def _sanitize_name(dwg_filename):
    """Convert a DWG filename into a valid Revit family name.

    Rules applied (must match the existing log output):
        - Spaces  → underscores
        - Dots    → removed
        - Remaining non-word characters → removed
    """
    name = dwg_filename
    name = name.replace(' ', '_')
    name = name.replace('.', '')
    name = re.sub(r'[^\w]', '', name)
    return name


def save_and_load_family(fam_doc, dwg_filename, doc, write_log, write_error):
    """Save *fam_doc* to a UUID-named temporary folder and load it into *doc*.

    Returns (save_path, family_symbol).
        - save_path    : path of the .rfa file (caller is responsible for cleanup)
        - family_symbol: the FamilySymbol found in the project, or None
    """
    write_log("Saving and loading family document...")
    write_log("DWG source name: {}".format(dwg_filename))

    family_name = _sanitize_name(dwg_filename)
    write_log("Initial family name: {}".format(family_name))
    write_log("Final family name: {}".format(family_name))

    # ---- temp directory ----
    temp_dir = os.path.join(tempfile.gettempdir(), str(uuid.uuid4()))
    try:
        os.makedirs(temp_dir)
    except Exception as e:
        write_error("Could not create temp directory: {}".format(str(e)), e)
        return None, None

    save_path = os.path.join(temp_dir, "{}.rfa".format(family_name))
    write_log("Save path (temp): {}".format(save_path))

    # ---- save ----
    try:
        opts = SaveAsOptions()
        opts.OverwriteExistingFile = True
        fam_doc.SaveAs(save_path, opts)
        write_log("Family saved to temp folder: {}".format(save_path))
    except Exception as e:
        write_error("Error saving family document", e)
        return None, None

    # ---- load into project (requires a transaction) ----
    write_log("Loading family into project...")
    t = Transaction(doc, 'Load DWG Family')
    t.Start()
    try:
        result = doc.LoadFamily(save_path)
        t.Commit()
        if result:
            write_log("Family load successful")
        else:
            write_log("Family already loaded or load returned False", "WARN")
    except Exception as e:
        t.RollBack()
        write_error("Error loading family into project", e)
        return save_path, None

    # ---- locate family symbol ----
    write_log("Searching for family: {}".format(family_name))
    family_symbol = None
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for sym in collector:
        try:
            if sym.Family is not None and sym.Family.Name == family_name:
                write_log("Found family symbol: {}".format(family_name))
                family_symbol = sym
                break
        except Exception:
            pass

    if family_symbol is None:
        write_log("Family symbol not found after load: {}".format(family_name), "WARN")

    return save_path, family_symbol
