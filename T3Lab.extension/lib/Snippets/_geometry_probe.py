# -*- coding: utf-8 -*-
"""Degenerate-geometry probe, shared by the preflight check and the MCP server.

Finds the faces that kill Revit's faceter during PDF/DWG export: zero-area
faces, slivers, pathological UV domains, zero normals, and singular points
("Singular point in SilSolver::eval" immediately before a 0xc0000005).

This logic used to live only in `checks/badgeometry_check.py`, which is a
pyRevit preflight module — it imports `pyrevit.forms` / `pyrevit.preflight`
and is not on the `lib/` import path, so the MCP server could not reuse it.
The pure probing functions live here; `badgeometry_check.py` imports them, and
so does the `check_bad_geometry` tool.

Nothing here opens a transaction, touches the UI, or reads an ambient document.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

import math

from pyrevit import DB

# ── Tuning ────────────────────────────────────────────────────────────────
FT_PER_MM = 1.0 / 304.8

# A face thinner than this (mean thickness = 2*Area/Perimeter) is the "long,
# thin face" the crash journal complains about. Same threshold as the
# TileLayout sliver guard, which fixed the equivalent crash there.
SLIVER_MM = 2.0

# Face area below this is degenerate for all practical purposes (ft^2).
TINY_AREA = 1e-7

# A surface normal shorter than this is the literal "Zero normal vector for
# ruled surface" the faceter chokes on.
ZERO_NORMAL = 1e-9

# |dP/du x dP/dv| below this is a SINGULAR POINT — the surface has no usable
# normal there, which is exactly what Revit logs before the access violation.
# Checking the derivatives directly is far more reliable than ComputeNormal(),
# which may normalise garbage instead of returning a zero.
SINGULAR_TOL = 1e-9

# Where to sample, as fractions of the face's UV domain. The journal shows the
# silhouette solver blowing up at NEGATIVE uv, just outside the boundary — so a
# plain inside-the-domain grid reports these faces as perfectly clean. Probe the
# exact boundary, step outside it at several orders of magnitude, plus a coarse
# interior grid.
SAMPLE_FRACTIONS = (-0.01, -0.001, -0.0001, 0.0, 0.5, 1.0, 1.0001, 1.001, 1.01)

# A parameter domain more lopsided than this is pathological. The crashing
# faces report envelope ([0,0], [29.535822627398, 1.0]) — a ~29.5:1 domain.
DOMAIN_ASPECT = 20.0

# Runaway guard: stop probing an element after this many faces.
MAX_FACES_PER_ELEMENT = 4000


def surface_kind(face):
    """Underlying surface type name, e.g. 'HermiteSurface' / 'RuledSurface'.
    Falls back to the Face subclass when GetSurface() is unavailable."""
    try:
        return type(face.GetSurface()).__name__
    except Exception:
        try:
            return type(face).__name__
        except Exception:
            return "<unknown surface>"


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


def probe_face(face, deep_probe=False):
    """Return a list of problem strings for one face. Never raises.

    deep_probe additionally calls Face.Triangulate(), which IS the code path
    that crashes Revit — the caller must treat it as dangerous and opt in.
    """
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

    # 3. Zero normals / singular points. Planar faces have a constant normal
    #    and never generate silhouettes, so only curved faces matter.
    if isinstance(face, DB.PlanarFace):
        return problems

    kind = surface_kind(face)

    try:
        bbox = face.GetBoundingBox()
        u0, u1 = bbox.Min.U, bbox.Max.U
        v0, v1 = bbox.Min.V, bbox.Max.V
    except Exception as ex:
        problems.append("no UV bounding box: {}".format(ex))
        return problems

    du_span = abs(u1 - u0)
    dv_span = abs(v1 - v0)

    # 3a. Pathological parameter domain — one direction far longer than the
    #     other is what drives the silhouette solver off the domain.
    if du_span > 0 and dv_span > 0:
        aspect = max(du_span / dv_span, dv_span / du_span)
        if aspect >= DOMAIN_ASPECT:
            problems.append(
                "DEGENERATE UV DOMAIN on {} -- aspect {:.1f}:1 "
                "(envelope u=[{:.3f}, {:.3f}] v=[{:.3f}, {:.3f}])".format(
                    kind, aspect, u0, u1, v0, v1))

    # 3b. Sample the domain edges, just outside them, and a coarse interior.
    zero_normals = 0
    failed_normals = 0
    singular = 0
    singular_uv = None
    outside_only = True

    for fi in SAMPLE_FRACTIONS:
        for fj in SAMPLE_FRACTIONS:
            fu = u0 + (u1 - u0) * fi
            fv = v0 + (v1 - v0) * fj
            inside = (0.0 <= fi <= 1.0) and (0.0 <= fj <= 1.0)
            uv = DB.UV(fu, fv)

            # Singular point: the two surface tangents are parallel or one
            # collapses, so their cross product — the unnormalised normal —
            # vanishes. This is the SilSolver::eval signature.
            try:
                derivs = face.ComputeDerivatives(uv)
                cross = derivs.BasisX.CrossProduct(derivs.BasisY)
                if cross.GetLength() < SINGULAR_TOL:
                    singular += 1
                    if singular_uv is None:
                        singular_uv = (fu, fv, inside)
                    if inside:
                        outside_only = False
            except Exception:
                pass

            try:
                normal = face.ComputeNormal(uv)
                length = math.sqrt(normal.X ** 2 + normal.Y ** 2 + normal.Z ** 2)
                if length < ZERO_NORMAL:
                    zero_normals += 1
            except Exception:
                failed_normals += 1

    total = len(SAMPLE_FRACTIONS) ** 2
    if singular:
        problems.append(
            "SINGULAR POINT on {} at {}/{} sample points, first at "
            "uv=({:.6f}, {:.6f}){} -- matches 'Singular point in "
            "SilSolver::eval'".format(
                kind, singular, total, singular_uv[0], singular_uv[1],
                " JUST OUTSIDE the domain" if outside_only else ""))
    if zero_normals:
        problems.append(
            "ZERO NORMAL VECTOR at {}/{} sample points -- matches "
            "'Zero normal vector for ruled surface'".format(
                zero_normals, total))
    if failed_normals:
        problems.append("ComputeNormal failed at {}/{} sample points".format(
            failed_normals, total))

    # 4. Optional: the actual crash path.
    if deep_probe:
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


def probe_element(element, view, deep_probe=False):
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
                if face_count > MAX_FACES_PER_ELEMENT:
                    return problems
                for issue in probe_face(face, deep_probe=deep_probe):
                    if issue not in problems:
                        problems.append(issue)
        except Exception as ex:
            problems.append("face iteration failed: {}".format(ex))
    return problems
