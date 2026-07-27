# -*- coding: utf-8 -*-
"""
Bad Geometry Check (faceter / silhouette killers)

Finds the elements whose geometry defeats Revit's faceter — the class of
geometry that makes Revit log

    Zero normal vector for ruled surface.      (RuledSurf.cpp)
    Invalid RuledSurf.                         (Face.cpp)
    Silhouette moved outside face envelope.    (Silhouettes.cpp)
    Shewchuk triangulation failed...           (Face.cpp)
    Problems during faceting
    facetFaceWithSpecifiedAlgorithm is triangulating a long, thin face

and then hard-crash the process (0xc0000005 / "Invalid array ... Arr.cpp")
during PDF or DWG export. That fatal is native and uncatchable, so the point of
this check is to find the culprit BEFORE it is handed to an exporter.

Two safeguards make the scan itself survivable:
  * every element is breadcrumbed to disk before it is probed, so if Revit does
    die mid-scan the marker names the exact element (same trick BatchOut uses
    for exports);
  * the probe is normal-evaluation, not full tessellation. DEEP_PROBE=True
    additionally calls Face.Triangulate(), which is the real crash path — only
    turn it on when a normal scan comes back clean.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__ = "Tran Tien Thanh"
__title__ = "Bad Geometry Check"

# pylint: disable=import-error,invalid-name,broad-except,superfluous-parens
import os
import json
import math
import datetime

from pyrevit import coreutils
from pyrevit import script
from pyrevit import revit, DB, forms

from pyrevit.preflight import PreflightTestCase


# ── Tuning ────────────────────────────────────────────────────────────────
FT_PER_MM = 1.0 / 304.8

# A face thinner than this (mean thickness = 2*Area/Perimeter) is the "long,
# thin face" the journal complains about. Same threshold as the TileLayout
# sliver guard, which fixed the equivalent crash there.
SLIVER_MM = 2.0

# Face area below this is degenerate for all practical purposes (ft^2).
TINY_AREA = 1e-7

# A surface normal shorter than this is the literal "Zero normal vector for
# ruled surface" the faceter chokes on.
ZERO_NORMAL = 1e-9

# UV grid resolution for normal probing (N x N samples per curved face).
UV_SAMPLES = 5

# Call Face.Triangulate() too. This IS the code path that crashes Revit —
# leave it False for a first pass; the breadcrumb will name the element if it
# does bring the session down.
DEEP_PROBE = False

DIAG_FOLDER = os.path.join(os.path.expanduser('~'), 'Documents', 'T3Lab_Diagnostics')
MARKER_FILE = os.path.join(DIAG_FOLDER, '_geometry_scan_marker.json')

# Findings are written here so BatchOut's safe-mode export can hide exactly
# these elements while exporting the sheets they poison.
FINDINGS_FILE = os.path.join(DIAG_FOLDER, '_bad_geometry.json')


# ── ElementId compatibility (Revit 2024+ dropped IntegerValue) ────────────
def eid_value(element_id):
    """Integer value of an ElementId, version-safe across Revit 2022-2026."""
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)              # Revit 2024+
    except Exception:
        try:
            return int(element_id.IntegerValue)   # Revit 2023 and earlier
        except Exception:
            return -1


def elem_name(element):
    """element.Name, IronPython-safe (some subclasses hide the getter)."""
    try:
        return element.Name
    except AttributeError:
        try:
            return DB.Element.Name.GetValue(element)
        except Exception:
            return "<unnamed>"


# ── Crash breadcrumb ──────────────────────────────────────────────────────
def write_marker(view_name, element_id, description):
    """Name the element about to be probed, durably (fsync'd), so a native
    crash still tells us which one was fatal."""
    try:
        if not os.path.exists(DIAG_FOLDER):
            os.makedirs(DIAG_FOLDER)
        payload = {
            'view': view_name,
            'element_id': element_id,
            'element': description,
            'time': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        with open(MARKER_FILE, 'w') as fh:
            json.dump(payload, fh)
            fh.flush()
            try:
                os.fsync(fh.fileno())
            except Exception:
                pass
    except Exception:
        pass


def clear_marker():
    try:
        if os.path.exists(MARKER_FILE):
            os.remove(MARKER_FILE)
    except Exception:
        pass


def read_marker():
    try:
        if os.path.exists(MARKER_FILE):
            with open(MARKER_FILE, 'r') as fh:
                return json.load(fh)
    except Exception:
        pass
    return None


def write_findings(doc, findings):
    """Publish the suspect element ids for BatchOut's safe-mode export to hide.
    Only elements with a hard signature are published — sliver-only hits are
    reported to the user but not auto-hidden, since slivers are common and
    usually harmless."""
    try:
        if not os.path.exists(DIAG_FOLDER):
            os.makedirs(DIAG_FOLDER)
        publishable = []
        for finding in findings:
            joined = " ".join(finding['problems'])
            if ("ZERO NORMAL" in joined
                    or "zero-area" in joined
                    or "Triangulate() failed" in joined
                    or "ComputeNormal failed" in joined):
                publishable.append({
                    'id': finding['id_value'],
                    'category': finding['category'],
                    'name': finding['name'],
                    'view': finding['view'],
                    'problems': finding['problems'],
                })
        payload = {
            'doc': doc.Title,
            'generated': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'elements': publishable,
        }
        with open(FINDINGS_FILE, 'w') as fh:
            json.dump(payload, fh, indent=2)
        return len(publishable)
    except Exception:
        return 0


# ── Geometry probing ──────────────────────────────────────────────────────
def face_perimeter(face):
    """Total edge length of a face, or 0.0 if it cannot be measured."""
    total = 0.0
    try:
        for loop in face.GetEdgesAsCurveLoops():
            for curve in loop:
                total += curve.Length
    except Exception:
        return 0.0
    return total


def probe_face(face):
    """Return a list of problem strings for one face. Never raises."""
    problems = []

    # 1. Degenerate area
    try:
        area = face.Area
        if area <= TINY_AREA:
            problems.append("zero-area face ({:.3e} ft2)".format(area))
            return problems  # nothing else is meaningful
    except Exception as ex:
        problems.append("Face.Area failed: {}".format(ex))
        return problems

    # 2. Sliver: mean thickness = 2 * Area / Perimeter. This is the
    #    "long, thin face" the faceter warns about before dying.
    perimeter = face_perimeter(face)
    if perimeter > 0:
        mean_thickness = 2.0 * area / perimeter
        if mean_thickness < SLIVER_MM * FT_PER_MM:
            problems.append("sliver face (mean thickness {:.3f} mm)".format(
                mean_thickness / FT_PER_MM))

    # 3. Zero / unevaluable surface normals. Planar faces have a constant
    #    normal and never generate silhouettes, so only curved faces matter —
    #    those are exactly what RuledSurf.cpp / Silhouettes.cpp operate on.
    if isinstance(face, DB.PlanarFace):
        return problems

    try:
        bbox = face.GetBoundingBox()
        u0, u1 = bbox.Min.U, bbox.Max.U
        v0, v1 = bbox.Min.V, bbox.Max.V
    except Exception as ex:
        problems.append("no UV bounding box: {}".format(ex))
        return problems

    zero_normals = 0
    failed_normals = 0
    for i in range(UV_SAMPLES):
        for j in range(UV_SAMPLES):
            fu = u0 + (u1 - u0) * (i / float(UV_SAMPLES - 1)) if UV_SAMPLES > 1 else u0
            fv = v0 + (v1 - v0) * (j / float(UV_SAMPLES - 1)) if UV_SAMPLES > 1 else v0
            try:
                normal = face.ComputeNormal(DB.UV(fu, fv))
                length = math.sqrt(normal.X ** 2 + normal.Y ** 2 + normal.Z ** 2)
                if length < ZERO_NORMAL:
                    zero_normals += 1
            except Exception:
                failed_normals += 1

    if zero_normals:
        problems.append(
            "ZERO NORMAL VECTOR at {}/{} sample points -- matches "
            "'Zero normal vector for ruled surface'".format(
                zero_normals, UV_SAMPLES * UV_SAMPLES))
    if failed_normals:
        problems.append("ComputeNormal failed at {}/{} sample points".format(
            failed_normals, UV_SAMPLES * UV_SAMPLES))

    # 4. Optional: the actual crash path.
    if DEEP_PROBE:
        try:
            mesh = face.Triangulate()
            if mesh is None or mesh.NumTriangles == 0:
                problems.append("Triangulate() produced no triangles")
        except Exception as ex:
            problems.append("Triangulate() failed: {}".format(ex))

    return problems


def walk_solids(geometry, depth=0):
    """Yield every Solid inside a GeometryElement, recursing into instances."""
    if geometry is None or depth > 4:
        return
    for obj in geometry:
        try:
            if isinstance(obj, DB.Solid):
                if obj.Faces.Size > 0:
                    yield obj
            elif isinstance(obj, DB.GeometryInstance):
                for solid in walk_solids(obj.GetInstanceGeometry(), depth + 1):
                    yield solid
            elif isinstance(obj, DB.GeometryElement):
                for solid in walk_solids(obj, depth + 1):
                    yield solid
        except Exception:
            continue


def probe_element(element, view):
    """Return a list of problem strings for one element in one view."""
    problems = []
    try:
        opts = DB.Options()
        opts.View = view              # inherits the view's detail level
        opts.ComputeReferences = False
        geometry = element.get_Geometry(opts)
    except Exception as ex:
        return ["get_Geometry failed: {}".format(ex)]

    if geometry is None:
        return problems

    face_count = 0
    for solid in walk_solids(geometry):
        try:
            for face in solid.Faces:
                face_count += 1
                if face_count > 4000:      # runaway guard on huge elements
                    return problems
                for issue in probe_face(face):
                    if issue not in problems:
                        problems.append(issue)
        except Exception as ex:
            problems.append("face iteration failed: {}".format(ex))
    return problems


# ── Target selection ──────────────────────────────────────────────────────
def views_to_scan(doc, output):
    """Views placed on the sheets the user picks. The crash is view-scoped, so
    scanning by view (not whole model) keeps this bounded and relevant."""
    sheets = DB.FilteredElementCollector(doc) \
               .OfClass(DB.ViewSheet) \
               .WhereElementIsNotElementType() \
               .ToElements()
    sheets = [s for s in sheets if not s.IsTemplate]
    if not sheets:
        return []

    picked = forms.SelectFromList.show(
        sorted([SheetOption(s) for s in sheets], key=lambda o: o.name),
        title="Pick the sheet(s) that crash on export",
        multiselect=True,
        button_name="Scan these sheets")
    if not picked:
        return []

    targets = []
    for option in picked:
        sheet = option.sheet
        for vid in sheet.GetAllPlacedViews():
            view = doc.GetElement(vid)
            if view is not None and not view.IsTemplate:
                targets.append((sheet.SheetNumber, view))
    return targets


class SheetOption(object):
    """forms.SelectFromList wrapper so the list shows number + name."""

    def __init__(self, sheet):
        self.sheet = sheet
        self.name = "{} - {}".format(sheet.SheetNumber, elem_name(sheet))

    def __str__(self):
        return self.name


# ── Main check ────────────────────────────────────────────────────────────
def checkModel(doc, output):
    output.print_md("# Bad Geometry Check")
    output.print_md(
        "Hunting the geometry that defeats Revit's faceter and hard-crashes "
        "PDF / DWG export.")

    stale = read_marker()
    if stale:
        output.print_md(
            "> **A previous scan did not finish.** Revit died while probing "
            "element `{}` (id {}) in view `{}` at {}. That element is your "
            "prime suspect.".format(
                stale.get('element'), stale.get('element_id'),
                stale.get('view'), stale.get('time')))
        clear_marker()

    targets = views_to_scan(doc, output)
    if not targets:
        output.print_md("No sheets selected - nothing scanned.")
        return

    output.print_md("Scanning **{}** views. DEEP_PROBE = `{}`.".format(
        len(targets), DEEP_PROBE))

    findings = []
    scanned = 0
    for sheet_number, view in targets:
        view_label = "{} / {}".format(sheet_number, elem_name(view))
        try:
            elements = DB.FilteredElementCollector(doc, view.Id) \
                         .WhereElementIsNotElementType() \
                         .ToElements()
        except Exception as ex:
            output.print_md("- Could not collect from `{}`: {}".format(view_label, ex))
            continue

        for element in elements:
            try:
                category = element.Category.Name if element.Category else "<no category>"
            except Exception:
                category = "<no category>"

            description = "{} : {}".format(category, elem_name(element))
            element_id = eid_value(element.Id)

            write_marker(view_label, element_id, description)
            problems = probe_element(element, view)
            clear_marker()
            scanned += 1

            if problems:
                findings.append({
                    'view': view_label,
                    'id': element.Id,
                    'id_value': element_id,
                    'category': category,
                    'name': elem_name(element),
                    'problems': problems,
                })

    output.print_md("---")
    output.print_md("Probed **{}** elements across **{}** views.".format(
        scanned, len(targets)))

    if not findings:
        output.print_md(
            "**No degenerate geometry found.** Re-run with `DEEP_PROBE = True` "
            "at the top of this file to exercise the actual tessellation path "
            "(that probe can crash Revit - the breadcrumb will name the "
            "element if it does).")
        return

    # Worst first: zero normals are the exact journal signature.
    def severity(finding):
        joined = " ".join(finding['problems'])
        if "ZERO NORMAL" in joined:
            return 0
        if "Triangulate() failed" in joined or "zero-area" in joined:
            return 1
        if "sliver" in joined:
            return 2
        return 3

    findings.sort(key=severity)

    output.print_md("## {} suspect element(s)".format(len(findings)))
    for finding in findings:
        output.print_md("**{}** - {} (id {})".format(
            finding['category'], finding['name'], output.linkify(finding['id'])))
        output.print_md("  - view: `{}`".format(finding['view']))
        for problem in finding['problems']:
            output.print_md("  - {}".format(problem))

    published = write_findings(doc, findings)

    output.print_md("---")
    output.print_md(
        "**Fix order:** elements reporting *ZERO NORMAL VECTOR* first - that is "
        "literally what `RuledSurf.cpp` logs before the crash. In the family, "
        "look for a blend / swept blend whose two profiles collapse to a point, "
        "or a revolve with a zero-radius edge. Rebuilding that one solid "
        "usually clears the whole crash.")
    if published:
        output.print_md(
            "**{} element(s) published to** `{}` - BatchOut safe mode will hide "
            "exactly these while exporting the sheets they poison, then roll the "
            "change back.".format(published, FINDINGS_FILE))


class ModelChecker(PreflightTestCase):
    """
    Find geometry that crashes Revit's faceter during PDF / DWG export.

    Probes every face of every element drawn in the selected sheets' views for
    degenerate ruled surfaces (zero normals), zero-area faces and slivers -
    the geometry behind "Problems during faceting" and the 0xc0000005 export
    crash. Breadcrumbs each element so a crash still identifies the culprit.
    """

    name = "Bad Geometry (faceter crash) Finder"
    author = "Tran Tien Thanh"

    def setUp(self, doc, output):
        pass

    def startTest(self, doc, output):
        timer = coreutils.Timer()
        checkModel(doc, output)
        endtime = timer.get_time()
        print("Scan took " + str(datetime.timedelta(seconds=endtime)))

    def tearDown(self, doc, output):
        pass

    def doCleanups(self, doc, output):
        clear_marker()
