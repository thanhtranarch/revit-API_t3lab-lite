# -*- coding: utf-8 -*-
"""
Tile Layout
Calculate, optimize, and visualize floor tiling with off-cut nesting.

Workflow:
  1. User selects Floor(s).
  2. Script generates a virtual tile grid over each floor's bounding box.
  3. Sutherland-Hodgman clipping determines full vs cut tiles.
  4. Nesting engine attempts to reuse off-cuts (1A → 1B logic).
  5. DirectShape elements + TextNotes visualize results in a 3D view.
  6. pyRevit output window prints the summary report.

Author: Tran Tien Thanh
"""

__author__ = "Tran Tien Thanh"
__title__ = "Tile Layout"

import os
import math
import csv
import traceback

import clr
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, ElementId,
    XYZ, Line, CurveLoop,
    GeometryCreationUtilities,
    DirectShape,
    View3D, ViewFamilyType, ViewFamily,
    OverrideGraphicSettings, Color,
    TextNote, TextNoteOptions,
    TextNoteType,
    Options, PlanarFace,
    FillPatternElement,
    HorizontalTextAlignment,
)
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

from pyrevit import revit, forms, script

logger = script.get_logger()
doc    = revit.doc
uidoc  = revit.uidoc

# ── Unit conversion ───────────────────────────────────────────────────────────
MM_TO_FT = 1.0 / 304.8

# ── DirectShape thickness (10 mm above floor top face) ───────────────────────
EXTRUDE_H = 10.0 * MM_TO_FT

# ── Area threshold — polygons smaller than this are discarded ─────────────────
MIN_AREA = 1e-9   # ft²

# ── Colour palette for view overrides ────────────────────────────────────────
COL_FULL  = Color(189, 195, 199)   # light grey  — full tiles
COL_CUT   = Color( 39, 174,  96)   # green       — primary cut (xA)
COL_REUSE = Color( 52, 152, 219)   # blue        — reused off-cut (xB/xC)
COL_WASTE = Color(231,  76,  60)   # red         — unused waste

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — PURE 2-D GEOMETRY (no Revit API)
# ─────────────────────────────────────────────────────────────────────────────

class V2(object):
    """Lightweight 2-D vector."""
    __slots__ = ('x', 'y')

    def __init__(self, x, y):
        self.x = float(x)
        self.y = float(y)

    def __repr__(self):
        return "V2({:.5f}, {:.5f})".format(self.x, self.y)

    def __add__(self, o): return V2(self.x + o.x, self.y + o.y)
    def __sub__(self, o): return V2(self.x - o.x, self.y - o.y)
    def __mul__(self, s): return V2(self.x * s,   self.y * s)


def _poly_area(pts):
    """Signed shoelace area; absolute value = area."""
    n, a = len(pts), 0.0
    for i in range(n):
        j = (i + 1) % n
        a += pts[i].x * pts[j].y - pts[j].x * pts[i].y
    return a * 0.5


def poly_area(pts):
    return abs(_poly_area(pts))


def poly_centroid(pts):
    n = len(pts)
    if n == 0:
        return V2(0.0, 0.0)
    cx = cy = area = 0.0
    for i in range(n):
        j = (i + 1) % n
        f = pts[i].x * pts[j].y - pts[j].x * pts[i].y
        cx   += (pts[i].x + pts[j].x) * f
        cy   += (pts[i].y + pts[j].y) * f
        area += f
    area *= 0.5
    if abs(area) < 1e-14:
        return V2(
            sum(p.x for p in pts) / n,
            sum(p.y for p in pts) / n)
    return V2(cx / (6.0 * area), cy / (6.0 * area))


def poly_bbox(pts):
    """Return (xmin, ymin, xmax, ymax)."""
    xs = [p.x for p in pts]
    ys = [p.y for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def ensure_ccw(pts):
    """Reverse polygon if clockwise so SH clips correctly."""
    if _poly_area(pts) < 0:
        return list(reversed(pts))
    return list(pts)


def clean_poly(pts):
    """Remove near-duplicate consecutive vertices; return None if degenerate."""
    if not pts:
        return None
    out = [pts[0]]
    for p in pts[1:]:
        dx, dy = p.x - out[-1].x, p.y - out[-1].y
        if dx * dx + dy * dy > 1e-16:
            out.append(p)
    # close: remove last if equal to first
    if len(out) > 1:
        dx, dy = out[-1].x - out[0].x, out[-1].y - out[0].y
        if dx * dx + dy * dy < 1e-16:
            out = out[:-1]
    return out if len(out) >= 3 else None


def sutherland_hodgman(subject, clip):
    """Clip *subject* polygon against *clip* polygon (Sutherland–Hodgman).

    Both polygons are lists of V2. The clip polygon is assumed to wind CCW.
    The subject is typically convex (tile rectangle).
    """
    def _inside(p, a, b):
        return (b.x - a.x) * (p.y - a.y) - (b.y - a.y) * (p.x - a.x) >= 0.0

    def _isect(p1, p2, p3, p4):
        x1, y1 = p1.x, p1.y
        x2, y2 = p2.x, p2.y
        x3, y3 = p3.x, p3.y
        x4, y4 = p4.x, p4.y
        d = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
        if abs(d) < 1e-15:
            return p2
        t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / d
        return V2(x1 + t * (x2 - x1), y1 + t * (y2 - y1))

    output = list(subject)
    nc = len(clip)
    for i in range(nc):
        if not output:
            return []
        inp = output
        output = []
        a, b = clip[i], clip[(i + 1) % nc]
        for j in range(len(inp)):
            cur, prv = inp[j], inp[j - 1]
            if _inside(cur, a, b):
                if not _inside(prv, a, b):
                    output.append(_isect(prv, cur, a, b))
                output.append(cur)
            elif _inside(prv, a, b):
                output.append(_isect(prv, cur, a, b))
    return output


def rotate_poly(pts, angle_deg, cx=0.0, cy=0.0):
    a = math.radians(angle_deg)
    c, s = math.cos(a), math.sin(a)
    out = []
    for p in pts:
        dx, dy = p.x - cx, p.y - cy
        out.append(V2(cx + dx * c - dy * s, cy + dx * s + dy * c))
    return out


def tile_rect(ox, oy, tw, th):
    """CCW rectangle polygon as list of 4 V2."""
    return [V2(ox, oy), V2(ox + tw, oy),
            V2(ox + tw, oy + th), V2(ox, oy + th)]


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — TILE GRID GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

class TileGrid(object):
    """Generates a virtual grid of tile rectangles covering a floor bbox."""

    def __init__(self, params):
        self.tw    = params['tile_w_ft']
        self.th    = params['tile_h_ft']
        self.jw    = params['joint_ft']
        self.pat   = params['pattern']      # 'grid' | 'staggered'
        self.angle = params['angle_deg']

    def generate(self, floor_pts):
        """Return list of (tile_id, [V2, ...]) covering the floor bounding box."""
        bx0, by0, bx1, by1 = poly_bbox(floor_pts)
        margin = max(self.tw, self.th) * 2.0
        bx0 -= margin; by0 -= margin
        bx1 += margin; by1 += margin

        sx = self.tw + self.jw
        sy = self.th + self.jw

        tiles = []
        tid   = 1
        row   = 0
        y = by0
        while y <= by1 + sy:
            x_off = (self.tw * 0.5) if (self.pat == 'staggered' and row % 2 == 1) else 0.0
            x = bx0 + x_off
            while x <= bx1 + sx:
                pts = tile_rect(x, y, self.tw, self.th)
                if self.angle != 0.0:
                    pts = rotate_poly(pts, self.angle,
                                      x + self.tw * 0.5,
                                      y + self.th * 0.5)
                tiles.append((tid, pts))
                tid += 1
                x += sx
            y += sy
            row += 1
        return tiles


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — DATA MODELS
# ─────────────────────────────────────────────────────────────────────────────

class TilePiece(object):
    """One physical piece: full tile, primary cut, reused off-cut, or waste."""

    def __init__(self, parent_id, sub_id, pts, piece_type):
        self.parent_id  = parent_id   # int  — which original tile number
        self.sub_id     = sub_id      # str  — '' | 'A' | 'B' | 'C' | 'waste'
        self.pts        = pts         # [V2] — polygon to visualise
        self.piece_type = piece_type  # 'full' | 'cut' | 'reuse' | 'waste'
        self.area       = poly_area(pts)
        self.centroid   = poly_centroid(pts)

    @property
    def label(self):
        if self.sub_id and self.sub_id != 'waste':
            return "{}{}".format(self.parent_id, self.sub_id)
        return str(self.parent_id)


class OffCut(object):
    """Leftover piece generated when a tile is cut."""

    def __init__(self, parent_id, tile_pts, inside_area, tile_area):
        self.parent_id   = parent_id
        self.tile_pts    = tile_pts          # full original tile polygon
        self.waste_area  = tile_area - inside_area
        self.used        = False             # True once matched to a gap


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — NESTING ENGINE
# ─────────────────────────────────────────────────────────────────────────────
# Logic overview:
#   Pass 1 — Intersect each virtual tile with the floor boundary.
#             Full tiles → TilePiece('full').
#             Cut tiles  → TilePiece('cut', sub='A')  +  OffCut stored.
#   Pass 2 — (nesting ON) For each cut tile's A-piece, search the OffCut pool
#             for a donor whose waste area can cover the needed inside area.
#             On match: A-piece is relabelled "donor_B" (saved from purchasing).
#   Pass 3 — Remaining unused OffCuts → TilePiece('waste').

class NestingEngine(object):

    def __init__(self, params, floor_pts):
        self.tile_w       = params['tile_w_ft']
        self.tile_h       = params['tile_h_ft']
        self.tile_area    = self.tile_w * self.tile_h
        self.use_nesting  = params['optimize_nesting']
        self.floor_pts    = ensure_ccw(floor_pts)

        self.pieces       = []   # final list of TilePiece
        self.offcuts      = []   # OffCut pool
        self.nesting_log  = []   # human-readable nesting events
        self._sub_cnt     = {}   # parent_id → next sub-letter index (A=0)

    # ── public ────────────────────────────────────────────────────────────────

    def process(self, tiles):
        """Run intersection + nesting; return list of TilePiece."""
        raw = self._intersect_all(tiles)
        self._assign_pieces(raw)
        if self.use_nesting:
            self._nest()
        self._collect_waste()
        return self.pieces

    # ── private ───────────────────────────────────────────────────────────────

    def _intersect_all(self, tiles):
        """
        Return list of dicts:
          { 'tid', 'inside_pts', 'inside_area', 'tile_pts', 'tile_area', 'ratio' }
        Only tiles with any intersection are included.
        """
        results = []
        for tid, tile_pts in tiles:
            clipped = sutherland_hodgman(tile_pts, self.floor_pts)
            clipped = clean_poly(clipped)
            if clipped is None:
                continue
            inside_area = poly_area(clipped)
            if inside_area < MIN_AREA:
                continue
            ta    = poly_area(tile_pts)
            ratio = inside_area / ta if ta > 1e-14 else 0.0
            results.append({
                'tid':          tid,
                'inside_pts':   clipped,
                'inside_area':  inside_area,
                'tile_pts':     tile_pts,
                'tile_area':    ta,
                'ratio':        ratio,
            })
        return results

    def _assign_pieces(self, raw):
        """Create TilePiece objects and populate the OffCut pool."""
        for r in raw:
            if r['ratio'] > 0.9999:
                self.pieces.append(
                    TilePiece(r['tid'], '', r['inside_pts'], 'full'))
            else:
                sub = self._next_sub(r['tid'])   # 'A' for the first cut piece
                self.pieces.append(
                    TilePiece(r['tid'], sub, r['inside_pts'], 'cut'))
                self.nesting_log.append(
                    "Tile {:>4}: piece {:>4}{} at ({:.3f}, {:.3f})  "
                    "area = {:.5f} ft²".format(
                        r['tid'], r['tid'], sub,
                        poly_centroid(r['inside_pts']).x,
                        poly_centroid(r['inside_pts']).y,
                        r['inside_area']))
                self.offcuts.append(
                    OffCut(r['tid'], r['tile_pts'],
                           r['inside_area'], r['tile_area']))

    def _nest(self):
        """
        For each cut piece (type='cut'), check whether a stored OffCut can
        supply that piece's area.  On match the cut piece is relabelled as
        "donor_B/C/..." and the donor OffCut is consumed.
        """
        # Sort offcuts by waste area DESC — largest off-cuts tried first.
        pool = sorted(self.offcuts, key=lambda o: o.waste_area, reverse=True)

        for piece in list(self.pieces):
            if piece.piece_type != 'cut':
                continue
            needed = piece.area

            for oc in pool:
                if oc.used:
                    continue
                if oc.parent_id == piece.parent_id:
                    continue       # don't reuse own offcut for itself
                if oc.waste_area >= needed * 0.90:
                    # Match found — relabel this cut piece as donor's B/C piece.
                    original_label = piece.label   # capture before mutation
                    oc.used = True
                    new_sub = self._next_sub(oc.parent_id)

                    piece.parent_id  = oc.parent_id
                    piece.sub_id     = new_sub
                    piece.piece_type = 'reuse'

                    self.nesting_log.append(
                        "  → piece {} REUSED as {}{} at ({:.3f},{:.3f})".format(
                            original_label, oc.parent_id, new_sub,
                            piece.centroid.x, piece.centroid.y))
                    break

    def _collect_waste(self):
        """Any unused OffCut becomes a waste TilePiece (shown in red)."""
        for oc in self.offcuts:
            if not oc.used:
                # Approximate waste geometry: full tile rect faded at its position
                waste_pts = clean_poly(oc.tile_pts)
                if waste_pts:
                    self.pieces.append(
                        TilePiece(oc.parent_id, 'waste', waste_pts, 'waste'))

    def _next_sub(self, parent_id):
        idx = self._sub_cnt.get(parent_id, 0)
        self._sub_cnt[parent_id] = idx + 1
        return 'ABCDEFGHIJ'[idx] if idx < 10 else str(idx)


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — REVIT FLOOR BOUNDARY EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────

class FloorBoundary(object):

    @staticmethod
    def extract(floor):
        """Return (list_of_V2, z_elevation_ft) for the floor's top PlanarFace."""
        opts = Options()
        opts.ComputeReferences = False
        geom = floor.get_Geometry(opts)

        best_z    = None
        best_loop = None

        for geom_obj in geom:
            if not hasattr(geom_obj, 'Faces'):
                continue
            for face in geom_obj.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                if abs(face.FaceNormal.Z - 1.0) > 0.05:
                    continue              # not a top face
                loops = face.GetEdgesAsCurveLoops()
                if not loops:
                    continue
                pts = []
                for curve in loops[0]:
                    ep = curve.GetEndPoint(0)
                    pts.append(V2(ep.X, ep.Y))
                z = face.Origin.Z
                if best_z is None or z > best_z:
                    best_z    = z
                    best_loop = pts

        if best_loop is None:
            # Fallback: use bounding box
            bb = floor.get_BoundingBox(None)
            if bb:
                mn, mx = bb.Min, bb.Max
                best_z    = mx.Z
                best_loop = [V2(mn.X, mn.Y), V2(mx.X, mn.Y),
                             V2(mx.X, mx.Y), V2(mn.X, mx.Y)]

        return best_loop, best_z


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — REVIT VISUALISER (DirectShape + TextNote)
# ─────────────────────────────────────────────────────────────────────────────

class RevitVisualizer(object):

    _solid_fill_id = None   # cached ElementId of <Solid fill> FillPattern

    def __init__(self, view):
        self.view     = view
        self._resolve_solid_fill()

    # ── public ────────────────────────────────────────────────────────────────

    def draw_piece(self, piece, z_base):
        """Create one DirectShape + TextNote for a TilePiece at z_base."""
        pts = piece.pts
        if not pts or len(pts) < 3:
            return

        loop = self._make_curve_loop(pts, z_base)
        if loop is None:
            return

        try:
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                [loop], XYZ(0, 0, 1), EXTRUDE_H)
        except Exception as ex:
            logger.debug("Extrude failed for {}: {}".format(piece.label, ex))
            return

        ds = DirectShape.CreateElement(
            doc, ElementId(BuiltInCategory.OST_GenericModel))
        ds.SetShape([solid])

        # Colour override
        ogs = OverrideGraphicSettings()
        col = self._piece_colour(piece)
        ogs.SetSurfaceForegroundPatternColor(col)
        ogs.SetSurfaceForegroundPatternVisible(True)
        if RevitVisualizer._solid_fill_id:
            ogs.SetSurfaceForegroundPatternId(RevitVisualizer._solid_fill_id)
        ogs.SetProjectionLineColor(Color(120, 120, 120))
        self.view.SetElementOverrides(ds.Id, ogs)

        # TextNote
        c = piece.centroid
        tn_pos = XYZ(c.x, c.y, z_base + EXTRUDE_H + 0.005)
        self._make_textnote(piece.label, tn_pos)

    # ── private ───────────────────────────────────────────────────────────────

    @staticmethod
    def _make_curve_loop(pts, z):
        loop = CurveLoop()
        n = len(pts)
        for i in range(n):
            p1 = XYZ(pts[i].x,           pts[i].y,           z)
            p2 = XYZ(pts[(i+1)%n].x, pts[(i+1)%n].y, z)
            if p1.DistanceTo(p2) < 1e-6:
                continue
            try:
                loop.Append(Line.CreateBound(p1, p2))
            except Exception:
                pass
        try:
            # validate by querying length
            _ = loop.GetExactLength()
            return loop
        except Exception:
            return None

    def _make_textnote(self, text, position):
        try:
            tn_types = (FilteredElementCollector(doc)
                        .OfClass(TextNoteType).ToElements())
            if not tn_types:
                return
            opts = TextNoteOptions(tn_types[0].Id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Center
            TextNote.Create(doc, self.view.Id, position, text, opts)
        except Exception as ex:
            logger.debug("TextNote failed: {}".format(ex))

    @staticmethod
    def _piece_colour(piece):
        t = piece.piece_type
        if t == 'full':  return COL_FULL
        if t == 'cut':   return COL_CUT
        if t == 'reuse': return COL_REUSE
        return COL_WASTE

    @classmethod
    def _resolve_solid_fill(cls):
        if cls._solid_fill_id is not None:
            return
        for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
            pat = fp.GetFillPattern()
            if pat.IsSolidFill:
                cls._solid_fill_id = fp.Id
                return


def get_or_create_3d_view(view_name="Tile Layout Preview"):
    """Find an existing 3D view by name, or create a new isometric view."""
    for v in FilteredElementCollector(doc).OfClass(View3D):
        if not v.IsTemplate and v.Name == view_name:
            return v
    for vt in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vt.ViewFamily == ViewFamily.ThreeDimensional:
            v3d      = View3D.CreateIsometric(doc, vt.Id)
            v3d.Name = view_name
            return v3d
    raise RuntimeError("No 3D ViewFamilyType found in project.")


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 — REPORT GENERATOR
# ─────────────────────────────────────────────────────────────────────────────

class ReportGenerator(object):

    def __init__(self, pieces, params, nesting_log):
        self.pieces      = pieces
        self.params      = params
        self.log         = nesting_log

    def _tile_area(self):
        return self.params['tile_w_ft'] * self.params['tile_h_ft']

    def summary_text(self):
        full_count   = sum(1 for p in self.pieces if p.piece_type == 'full')
        cut_a_count  = sum(1 for p in self.pieces if p.piece_type == 'cut')
        reuse_count  = sum(1 for p in self.pieces if p.piece_type == 'reuse')
        waste_pieces = [p for p in self.pieces if p.piece_type == 'waste']

        # Tiles to purchase = full tiles + cut tiles whose A-piece was NOT
        # replaced by a reused off-cut.  Each reused piece saves one purchase.
        parents_full = set(p.parent_id for p in self.pieces if p.piece_type == 'full')
        parents_cut  = set(p.parent_id for p in self.pieces if p.piece_type == 'cut')
        total_purchase = len(parents_full) + len(parents_cut)

        tile_area        = self._tile_area()
        purchased_area   = total_purchase * tile_area
        waste_area       = sum(p.area for p in waste_pieces)
        waste_pct        = (waste_area / purchased_area * 100.0
                            if purchased_area > 0 else 0.0)

        tw_mm = self.params['tile_w_ft'] / MM_TO_FT
        th_mm = self.params['tile_h_ft'] / MM_TO_FT
        jw_mm = self.params['joint_ft']  / MM_TO_FT

        lines = [
            "=" * 62,
            "  TILE LAYOUT REPORT",
            "=" * 62,
            "  Tile size     : {:.0f} x {:.0f} mm".format(tw_mm, th_mm),
            "  Joint width   : {:.1f} mm".format(jw_mm),
            "  Pattern       : {}".format(self.params['pattern'].title()),
            "  Angle         : {}°".format(self.params['angle_deg']),
            "  Nesting       : {}".format(
                "ON" if self.params['optimize_nesting'] else "OFF"),
            "-" * 62,
            "  Full tiles         : {:5d}".format(full_count),
            "  Cut tiles (A-pcs)  : {:5d}".format(cut_a_count),
            "  Reused off-cuts    : {:5d}  (tiles saved)".format(reuse_count),
            "  TOTAL TO PURCHASE  : {:5d} tiles".format(total_purchase),
            "",
            "  Waste area         : {:.4f} ft²  ({:.1f}%)".format(
                waste_area, waste_pct),
            "-" * 62,
            "  NESTING LOG",
            "-" * 62,
        ]
        if self.log:
            lines.extend("  " + e for e in self.log)
        else:
            lines.append("  (nesting disabled or no off-cuts generated)")
        lines.append("=" * 62)
        return "\n".join(lines)

    def export_csv(self, filepath):
        with open(filepath, 'wb') as fh:   # 'wb' for IronPython 2 csv compat
            w = csv.writer(fh)
            w.writerow(['Label', 'Type', 'Parent_ID',
                        'Area_ft2', 'Centroid_X', 'Centroid_Y'])
            for p in self.pieces:
                w.writerow([
                    p.label, p.piece_type, p.parent_id,
                    "{:.6f}".format(p.area),
                    "{:.4f}".format(p.centroid.x),
                    "{:.4f}".format(p.centroid.y),
                ])


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 8 — SELECTION FILTER
# ─────────────────────────────────────────────────────────────────────────────

class _FloorFilter(ISelectionFilter):
    def AllowElement(self, elem):
        cat = elem.Category
        return (cat is not None and
                cat.Id.IntegerValue == int(BuiltInCategory.OST_Floors))

    def AllowReference(self, ref, pt):
        return False


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 9 — WPF WINDOW
# ─────────────────────────────────────────────────────────────────────────────

class TileLayoutWindow(forms.WPFWindow):

    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'TileLayout.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        self._selected_floors = []
        self._last_pieces     = []
        self._last_params     = {}
        self._last_log        = []
        self._load_logo()

    # ── Logo ──────────────────────────────────────────────────────────────────
    def _load_logo(self):
        logo_path = os.path.normpath(os.path.join(
            os.path.dirname(__file__),
            '..', '..', '..', 'lib', 'GUI', 'T3Lab_logo.png'))
        if not os.path.exists(logo_path):
            return
        try:
            from System.Windows.Media.Imaging import BitmapImage
            from System import Uri, UriKind
            self.logo_image.Source = BitmapImage(
                Uri(logo_path, UriKind.Absolute))
        except Exception:
            pass

    # ── Chrome buttons ────────────────────────────────────────────────────────
    def minimize_button_clicked(self, sender, args):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        import System.Windows
        ws = System.Windows.WindowState
        self.WindowState = (ws.Normal
                            if self.WindowState == ws.Maximized
                            else ws.Maximized)

    def close_button_clicked(self, sender, args):
        self.Close()

    # ── Select Floors ─────────────────────────────────────────────────────────
    def select_floors_clicked(self, sender, args):
        self.Hide()
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                _FloorFilter(),
                "Select floor(s) for tile layout — press Finish when done")
            self._selected_floors = [doc.GetElement(r.ElementId) for r in refs]
            self.status_text.Text = "{} floor(s) selected.".format(
                len(self._selected_floors))
            self.btn_run.IsEnabled = bool(self._selected_floors)
        except Exception:
            self.status_text.Text = "Selection cancelled."
        finally:
            self.Show()

    # ── Run Layout ────────────────────────────────────────────────────────────
    def run_layout_clicked(self, sender, args):
        if not self._selected_floors:
            TaskDialog.Show("Tile Layout", "Please select floors first.")
            return

        try:
            params = self._read_params()
        except ValueError as exc:
            TaskDialog.Show("Input Error", str(exc))
            return

        self.status_text.Text = "Computing layout…"
        self.btn_run.IsEnabled = False
        self.UpdateLayout()

        try:
            all_pieces = []
            all_logs   = []
            floor_jobs = []    # (floor, pts, z, pieces)

            # ── Geometry pass (no transaction needed) ──────────────────────
            for floor in self._selected_floors:
                pts, z = FloorBoundary.extract(floor)
                if not pts:
                    logger.warning("Could not extract boundary for floor {}"
                                   .format(floor.Id))
                    continue
                grid   = TileGrid(params)
                tiles  = grid.generate(pts)
                engine = NestingEngine(params, pts)
                pieces = engine.process(tiles)
                all_pieces.extend(pieces)
                all_logs.extend(engine.nesting_log)
                floor_jobs.append((floor, pts, z, pieces))

            if not all_pieces:
                self.status_text.Text = "No tile pieces computed — check inputs."
                return

            # ── Revit transaction: create DirectShapes + TextNotes ─────────
            with revit.Transaction("Tile Layout — create preview"):
                view = get_or_create_3d_view()
                for _floor, _pts, z, pieces in floor_jobs:
                    vis = RevitVisualizer(view)
                    for piece in pieces:
                        vis.draw_piece(piece, z)

            # Activate the view
            uidoc.ActiveView = view

            # ── Cache for CSV export ───────────────────────────────────────
            self._last_pieces = all_pieces
            self._last_params = params
            self._last_log    = all_logs
            self.btn_export.IsEnabled = True

            # ── Print report ───────────────────────────────────────────────
            rpt = ReportGenerator(all_pieces, params, all_logs)
            out = script.get_output()
            out.print_md("```\n{}\n```".format(rpt.summary_text()))

            n_full   = sum(1 for p in all_pieces if p.piece_type == 'full')
            n_cut    = sum(1 for p in all_pieces if p.piece_type == 'cut')
            n_reuse  = sum(1 for p in all_pieces if p.piece_type == 'reuse')
            self.status_text.Text = (
                "Done — {} full, {} cut ({}A), {} reused ({}B). "
                "See output window for report.".format(
                    n_full, n_cut, n_cut, n_reuse, n_reuse))

        except Exception as exc:
            TaskDialog.Show(
                "Tile Layout Error",
                "{}\n\n{}".format(exc, traceback.format_exc()))
            self.status_text.Text = "Error — see TaskDialog for details."
        finally:
            self.btn_run.IsEnabled = True

    # ── Export CSV ────────────────────────────────────────────────────────────
    def export_csv_clicked(self, sender, args):
        if not self._last_pieces:
            TaskDialog.Show("Export", "Run the layout first.")
            return
        try:
            from System.Windows.Forms import SaveFileDialog, DialogResult
            dlg            = SaveFileDialog()
            dlg.Title      = "Save Tile Layout CSV"
            dlg.Filter     = "CSV files (*.csv)|*.csv"
            dlg.FileName   = "TileLayout.csv"
            if dlg.ShowDialog() != DialogResult.OK:
                return
            rpt = ReportGenerator(
                self._last_pieces, self._last_params, self._last_log)
            rpt.export_csv(dlg.FileName)
            self.status_text.Text = "CSV saved: {}".format(
                os.path.basename(dlg.FileName))
        except Exception as exc:
            TaskDialog.Show("CSV Export Error", str(exc))

    # ── Input helpers ─────────────────────────────────────────────────────────
    def _read_params(self):
        def _mm(ctrl, name):
            try:
                v = float(ctrl.Text.strip())
            except (ValueError, AttributeError):
                raise ValueError("'{}' is not a valid number.".format(name))
            if v <= 0:
                raise ValueError("'{}' must be greater than zero.".format(name))
            return v * MM_TO_FT

        tw  = _mm(self.txt_tile_w, "Tile Width")
        th  = _mm(self.txt_tile_h, "Tile Height")
        jw  = _mm(self.txt_joint,  "Joint Width")

        try:
            angle = float((self.txt_angle.Text or "0").strip())
        except ValueError:
            raise ValueError("'Layout Angle' is not a valid number.")

        pat = 'grid'
        if self.cmb_pattern.SelectedItem is not None:
            sel = str(self.cmb_pattern.SelectedItem.Content).lower()
            if 'stagger' in sel or 'running' in sel:
                pat = 'staggered'

        optimize = bool(self.chk_nesting.IsChecked)

        return {
            'tile_w_ft':        tw,
            'tile_h_ft':        th,
            'joint_ft':         jw,
            'pattern':          pat,
            'angle_deg':        angle,
            'optimize_nesting': optimize,
        }


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    TileLayoutWindow().ShowDialog()
