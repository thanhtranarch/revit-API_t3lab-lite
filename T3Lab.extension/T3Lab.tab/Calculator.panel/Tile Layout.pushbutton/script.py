# -*- coding: utf-8 -*-
"""
Tile Layout
3-step wizard:
    (1) Extract floor boundaries,
    (2) Choose tile pattern per floor,
    (3) Generate multiple candidate layouts per floor and let the user pick.

All geometry is computed in memory (no physical Revit elements for the tiles
themselves — only temporary DirectShape + TextNote for the chosen layout).

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
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

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

# ── Constants ─────────────────────────────────────────────────────────────────
MM_TO_FT  = 1.0 / 304.8
FT_TO_MM  = 304.8
FT2_TO_M2 = 0.092903

EXTRUDE_H = 10.0 * MM_TO_FT   # DirectShape thickness
MIN_AREA  = 1e-9              # ft² — anything smaller discarded

# Colours for DirectShape overrides
COL_FULL  = Color(189, 195, 199)
COL_CUT   = Color( 39, 174,  96)
COL_REUSE = Color( 52, 152, 219)
COL_WASTE = Color(231,  76,  60)


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 1 — 2-D GEOMETRY (pure Python)
# ═════════════════════════════════════════════════════════════════════════════

class V2(object):
    __slots__ = ('x', 'y')
    def __init__(self, x, y):
        self.x = float(x); self.y = float(y)
    def __repr__(self):
        return "V2({:.4f},{:.4f})".format(self.x, self.y)


def _poly_area_signed(pts):
    n, a = len(pts), 0.0
    for i in range(n):
        j = (i + 1) % n
        a += pts[i].x * pts[j].y - pts[j].x * pts[i].y
    return a * 0.5


def poly_area(pts): return abs(_poly_area_signed(pts))


def poly_centroid(pts):
    n = len(pts)
    if n == 0:
        return V2(0.0, 0.0)
    cx = cy = a = 0.0
    for i in range(n):
        j = (i + 1) % n
        f = pts[i].x * pts[j].y - pts[j].x * pts[i].y
        cx += (pts[i].x + pts[j].x) * f
        cy += (pts[i].y + pts[j].y) * f
        a  += f
    a *= 0.5
    if abs(a) < 1e-14:
        return V2(sum(p.x for p in pts)/n, sum(p.y for p in pts)/n)
    return V2(cx / (6.0 * a), cy / (6.0 * a))


def poly_bbox(pts):
    xs = [p.x for p in pts]; ys = [p.y for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def ensure_ccw(pts):
    return list(reversed(pts)) if _poly_area_signed(pts) < 0 else list(pts)


def clean_poly(pts):
    if not pts: return None
    out = [pts[0]]
    for p in pts[1:]:
        dx, dy = p.x - out[-1].x, p.y - out[-1].y
        if dx*dx + dy*dy > 1e-16: out.append(p)
    if len(out) > 1:
        dx, dy = out[-1].x - out[0].x, out[-1].y - out[0].y
        if dx*dx + dy*dy < 1e-16: out = out[:-1]
    return out if len(out) >= 3 else None


def sutherland_hodgman(subject, clip):
    def _inside(p, a, b):
        return (b.x - a.x)*(p.y - a.y) - (b.y - a.y)*(p.x - a.x) >= 0.0
    def _isect(p1, p2, p3, p4):
        x1,y1 = p1.x,p1.y; x2,y2 = p2.x,p2.y
        x3,y3 = p3.x,p3.y; x4,y4 = p4.x,p4.y
        d = (x1-x2)*(y3-y4) - (y1-y2)*(x3-x4)
        if abs(d) < 1e-15: return p2
        t = ((x1-x3)*(y3-y4) - (y1-y3)*(x3-x4)) / d
        return V2(x1 + t*(x2-x1), y1 + t*(y2-y1))

    out = list(subject); nc = len(clip)
    for i in range(nc):
        if not out: return []
        inp = out; out = []
        a, b = clip[i], clip[(i+1) % nc]
        for j in range(len(inp)):
            cur, prv = inp[j], inp[j-1]
            if _inside(cur, a, b):
                if not _inside(prv, a, b):
                    out.append(_isect(prv, cur, a, b))
                out.append(cur)
            elif _inside(prv, a, b):
                out.append(_isect(prv, cur, a, b))
    return out


def rotate_poly(pts, angle_deg, cx=0.0, cy=0.0):
    a = math.radians(angle_deg); c, s = math.cos(a), math.sin(a)
    return [V2(cx + (p.x-cx)*c - (p.y-cy)*s,
               cy + (p.x-cx)*s + (p.y-cy)*c) for p in pts]


def tile_rect(ox, oy, tw, th):
    return [V2(ox, oy), V2(ox+tw, oy), V2(ox+tw, oy+th), V2(ox, oy+th)]


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 2 — TILE GRID GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

class TileGrid(object):
    """Virtual grid generator; supports origin offset for layout variants."""

    def __init__(self, tile_w, tile_h, joint, pattern, angle_deg,
                 origin_dx=0.0, origin_dy=0.0):
        self.tw, self.th, self.jw = tile_w, tile_h, joint
        self.pat, self.angle = pattern, angle_deg
        self.dx, self.dy = origin_dx, origin_dy

    def generate(self, floor_pts):
        bx0, by0, bx1, by1 = poly_bbox(floor_pts)
        margin = max(self.tw, self.th) * 2.0
        bx0 -= margin; by0 -= margin
        bx1 += margin; by1 += margin

        sx = self.tw + self.jw
        sy = self.th + self.jw

        tiles, tid, row = [], 1, 0
        y = by0 + self.dy
        while y <= by1 + sy:
            x_off = (self.tw * 0.5) if (self.pat == 'staggered' and row % 2 == 1) else 0.0
            x = bx0 + self.dx + x_off
            while x <= bx1 + sx:
                pts = tile_rect(x, y, self.tw, self.th)
                if self.angle != 0.0:
                    pts = rotate_poly(pts, self.angle,
                                      x + self.tw*0.5, y + self.th*0.5)
                tiles.append((tid, pts))
                tid += 1
                x += sx
            y += sy
            row += 1
        return tiles


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 3 — DATA MODELS
# ═════════════════════════════════════════════════════════════════════════════

class TilePiece(object):
    def __init__(self, parent_id, sub_id, pts, piece_type):
        self.parent_id  = parent_id
        self.sub_id     = sub_id
        self.pts        = pts
        self.piece_type = piece_type
        self.area       = poly_area(pts)
        self.centroid   = poly_centroid(pts)

    @property
    def label(self):
        if self.sub_id and self.sub_id != 'waste':
            return "{}{}".format(self.parent_id, self.sub_id)
        return str(self.parent_id)


class OffCut(object):
    def __init__(self, parent_id, tile_pts, inside_area, tile_area):
        self.parent_id  = parent_id
        self.tile_pts   = tile_pts
        self.waste_area = tile_area - inside_area
        self.used       = False


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 4 — NESTING ENGINE
# ═════════════════════════════════════════════════════════════════════════════

class NestingEngine(object):

    def __init__(self, tile_w, tile_h, use_nesting, floor_pts):
        self.tile_w      = tile_w
        self.tile_h      = tile_h
        self.use_nesting = use_nesting
        self.floor_pts   = ensure_ccw(floor_pts)

        self.pieces      = []
        self.offcuts     = []
        self.nesting_log = []
        self._sub_cnt    = {}

    def process(self, tiles):
        raw = self._intersect_all(tiles)
        self._assign_pieces(raw)
        if self.use_nesting:
            self._nest()
        self._collect_waste()
        return self.pieces

    def _intersect_all(self, tiles):
        out = []
        for tid, tp in tiles:
            clipped = clean_poly(sutherland_hodgman(tp, self.floor_pts))
            if clipped is None: continue
            ia = poly_area(clipped)
            if ia < MIN_AREA: continue
            ta = poly_area(tp)
            out.append({'tid': tid, 'inside_pts': clipped,
                        'inside_area': ia, 'tile_pts': tp,
                        'tile_area': ta,
                        'ratio': ia/ta if ta > 1e-14 else 0.0})
        return out

    def _assign_pieces(self, raw):
        for r in raw:
            if r['ratio'] > 0.9999:
                self.pieces.append(TilePiece(r['tid'], '', r['inside_pts'], 'full'))
            else:
                sub = self._next_sub(r['tid'])
                self.pieces.append(TilePiece(r['tid'], sub, r['inside_pts'], 'cut'))
                c = poly_centroid(r['inside_pts'])
                self.nesting_log.append(
                    "Tile {:>4}: {:>4}{} at ({:.2f},{:.2f}) area={:.4f} ft²".format(
                        r['tid'], r['tid'], sub, c.x, c.y, r['inside_area']))
                self.offcuts.append(OffCut(r['tid'], r['tile_pts'],
                                           r['inside_area'], r['tile_area']))

    def _nest(self):
        pool = sorted(self.offcuts, key=lambda o: o.waste_area, reverse=True)
        for piece in list(self.pieces):
            if piece.piece_type != 'cut': continue
            needed = piece.area
            for oc in pool:
                if oc.used or oc.parent_id == piece.parent_id: continue
                if oc.waste_area >= needed * 0.90:
                    original_label = piece.label
                    oc.used = True
                    new_sub = self._next_sub(oc.parent_id)
                    piece.parent_id  = oc.parent_id
                    piece.sub_id     = new_sub
                    piece.piece_type = 'reuse'
                    self.nesting_log.append(
                        "  → {} reused as {}{}".format(
                            original_label, oc.parent_id, new_sub))
                    break

    def _collect_waste(self):
        for oc in self.offcuts:
            if not oc.used:
                wp = clean_poly(oc.tile_pts)
                if wp:
                    self.pieces.append(
                        TilePiece(oc.parent_id, 'waste', wp, 'waste'))

    def _next_sub(self, parent_id):
        idx = self._sub_cnt.get(parent_id, 0)
        self._sub_cnt[parent_id] = idx + 1
        return 'ABCDEFGHIJ'[idx] if idx < 10 else str(idx)


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 5 — LAYOUT OPTION + SCORING + OPTION GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

class LayoutOption(object):
    """One candidate tiling arrangement for a single floor."""

    def __init__(self, option_id, pieces, variant_desc, tile_area_ft2):
        self.option_id   = option_id          # 'A' 'B' 'C' 'D'
        self.pieces      = pieces
        self.variant     = variant_desc       # human-readable variation tag
        self.tile_area   = tile_area_ft2

        self.n_full  = sum(1 for p in pieces if p.piece_type == 'full')
        self.n_cut   = sum(1 for p in pieces if p.piece_type == 'cut')
        self.n_reuse = sum(1 for p in pieces if p.piece_type == 'reuse')

        parents_full = set(p.parent_id for p in pieces if p.piece_type == 'full')
        parents_cut  = set(p.parent_id for p in pieces if p.piece_type == 'cut')
        self.tiles_to_buy = len(parents_full) + len(parents_cut)

        waste_area = sum(p.area for p in pieces if p.piece_type == 'waste')
        purch_area = self.tiles_to_buy * tile_area_ft2
        self.waste_pct = (waste_area / purch_area * 100.0) if purch_area > 0 else 0.0

        self.score = self._compute_score()

    def _compute_score(self):
        """Lower is better. Penalise waste %, reward reused pieces."""
        return self.waste_pct * 10.0 - self.n_reuse * 0.5 + self.n_cut * 0.1


class OptionGenerator(object):
    """Produce N candidate layouts per floor by varying origin + angle."""

    # Shift fractions (of tile size) × angle deltas (°) = variants to try.
    _SHIFT_FRACS = [0.0, 0.25, 0.5, 0.75]
    _ANGLE_DELTAS = [0.0, 45.0]

    def __init__(self, tile_w, tile_h, joint, use_nesting, top_n=4):
        self.tile_w = tile_w
        self.tile_h = tile_h
        self.joint  = joint
        self.use_nesting = use_nesting
        self.top_n  = top_n

    def generate(self, floor_info, pattern, base_angle):
        """Return list of top-N LayoutOption, sorted by score ASC."""
        candidates = []
        tile_area = self.tile_w * self.tile_h

        variants = []
        for fx in self._SHIFT_FRACS:
            for fy in self._SHIFT_FRACS:
                for da in self._ANGLE_DELTAS:
                    variants.append((fx, fy, da))

        for fx, fy, da in variants:
            dx = fx * self.tile_w
            dy = fy * self.tile_h
            angle = base_angle + da
            grid = TileGrid(self.tile_w, self.tile_h, self.joint,
                            pattern, angle, dx, dy)
            tiles = grid.generate(floor_info.pts)
            engine = NestingEngine(self.tile_w, self.tile_h,
                                   self.use_nesting, floor_info.pts)
            pieces = engine.process(tiles)

            desc = "shift {:+.0f}/{:+.0f} mm, angle {:+.0f}°".format(
                dx * FT_TO_MM, dy * FT_TO_MM, da)
            # Option id assigned after sorting
            candidates.append(LayoutOption('?', pieces, desc, tile_area))
            # cache for later
            candidates[-1]._nesting_log = engine.nesting_log

        # Keep the N best
        candidates.sort(key=lambda o: o.score)
        best = candidates[:self.top_n]
        for i, opt in enumerate(best):
            opt.option_id = "ABCDEF"[i] if i < 6 else str(i+1)
        return best


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 6 — FLOOR INFO (wizard-level data model)
# ═════════════════════════════════════════════════════════════════════════════

class FloorInfo(object):
    """Data carried between wizard steps for one floor."""

    def __init__(self, floor_elem, pts, z):
        self.floor = floor_elem
        self.pts   = pts            # [V2] — CCW floor boundary
        self.z     = z              # top-face elevation (ft)

        bx0, by0, bx1, by1 = poly_bbox(pts)
        self.width_ft  = bx1 - bx0
        self.height_ft = by1 - by0
        self.area_ft2  = poly_area(pts)

        # wizard state
        self.options           = []       # [LayoutOption]
        self.chosen_option_id  = None     # 'A' | 'B' | ...


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 7 — REVIT FLOOR BOUNDARY EXTRACTION
# ═════════════════════════════════════════════════════════════════════════════

def extract_floor_boundary(floor):
    opts = Options(); opts.ComputeReferences = False
    geom = floor.get_Geometry(opts)
    best_z, best_loop = None, None
    for g in geom:
        if not hasattr(g, 'Faces'): continue
        for face in g.Faces:
            if not isinstance(face, PlanarFace): continue
            if abs(face.FaceNormal.Z - 1.0) > 0.05: continue
            loops = face.GetEdgesAsCurveLoops()
            if not loops: continue
            pts = [V2(c.GetEndPoint(0).X, c.GetEndPoint(0).Y) for c in loops[0]]
            z = face.Origin.Z
            if best_z is None or z > best_z:
                best_z, best_loop = z, pts
    if best_loop is None:
        bb = floor.get_BoundingBox(None)
        if bb:
            mn, mx = bb.Min, bb.Max
            best_z = mx.Z
            best_loop = [V2(mn.X, mn.Y), V2(mx.X, mn.Y),
                         V2(mx.X, mx.Y), V2(mn.X, mx.Y)]
    return best_loop, best_z


def get_floor_level_name(floor):
    try:
        lvl = doc.GetElement(floor.LevelId)
        return lvl.Name if lvl else "—"
    except Exception:
        return "—"


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 8 — REVIT VISUALISER (DirectShape + TextNote)
# ═════════════════════════════════════════════════════════════════════════════

class RevitVisualizer(object):

    _solid_fill_id = None

    def __init__(self, view):
        self.view = view
        self._resolve_solid_fill()

    def draw_piece(self, piece, z_base):
        if not piece.pts or len(piece.pts) < 3: return
        loop = self._make_curve_loop(piece.pts, z_base)
        if loop is None: return
        try:
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                [loop], XYZ(0,0,1), EXTRUDE_H)
        except Exception as ex:
            logger.debug("Extrude failed for {}: {}".format(piece.label, ex))
            return

        ds = DirectShape.CreateElement(
            doc, ElementId(BuiltInCategory.OST_GenericModel))
        ds.SetShape([solid])

        ogs = OverrideGraphicSettings()
        col = self._piece_colour(piece)
        ogs.SetSurfaceForegroundPatternColor(col)
        ogs.SetSurfaceForegroundPatternVisible(True)
        if RevitVisualizer._solid_fill_id:
            ogs.SetSurfaceForegroundPatternId(RevitVisualizer._solid_fill_id)
        ogs.SetProjectionLineColor(Color(120, 120, 120))
        self.view.SetElementOverrides(ds.Id, ogs)

        c = piece.centroid
        self._make_textnote(piece.label,
                            XYZ(c.x, c.y, z_base + EXTRUDE_H + 0.005))

    @staticmethod
    def _make_curve_loop(pts, z):
        loop = CurveLoop()
        n = len(pts)
        for i in range(n):
            p1 = XYZ(pts[i].x,           pts[i].y,           z)
            p2 = XYZ(pts[(i+1)%n].x, pts[(i+1)%n].y, z)
            if p1.DistanceTo(p2) < 1e-6: continue
            try: loop.Append(Line.CreateBound(p1, p2))
            except Exception: pass
        try:
            _ = loop.GetExactLength()
            return loop
        except Exception:
            return None

    def _make_textnote(self, text, position):
        try:
            tn_types = (FilteredElementCollector(doc)
                        .OfClass(TextNoteType).ToElements())
            if not tn_types: return
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
        if cls._solid_fill_id is not None: return
        for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
            if fp.GetFillPattern().IsSolidFill:
                cls._solid_fill_id = fp.Id
                return


def get_or_create_3d_view(view_name="Tile Layout Preview"):
    for v in FilteredElementCollector(doc).OfClass(View3D):
        if not v.IsTemplate and v.Name == view_name:
            return v
    for vt in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vt.ViewFamily == ViewFamily.ThreeDimensional:
            v3d = View3D.CreateIsometric(doc, vt.Id)
            v3d.Name = view_name
            return v3d
    raise RuntimeError("No 3D ViewFamilyType found in project.")


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 9 — WPF PREVIEW RENDERER (mini canvas drawing of a LayoutOption)
# ═════════════════════════════════════════════════════════════════════════════

# Import WPF drawing primitives lazily to keep pyRevit startup light.
def _wpf():
    from System.Windows import Point as WPoint
    from System.Windows.Media import (
        SolidColorBrush, Color as WColor, PointCollection,
        Pen as WPen,
    )
    from System.Windows.Shapes import Polygon as WPolygon, Polyline as WPolyline
    from System.Windows.Controls import Canvas
    return dict(
        Point=WPoint, Brush=SolidColorBrush, Color=WColor,
        Points=PointCollection, Pen=WPen,
        Polygon=WPolygon, Polyline=WPolyline, Canvas=Canvas)


_PIECE_FILLS = {
    'full':  (189, 195, 199),
    'cut':   ( 39, 174,  96),
    'reuse': ( 52, 152, 219),
    'waste': (231,  76,  60),
}

def render_option_preview(option, floor_pts, canvas_w=220, canvas_h=150):
    """Return a populated WPF Canvas showing the layout option."""
    w = _wpf()
    canvas = w['Canvas'](); canvas.Width = canvas_w; canvas.Height = canvas_h
    canvas.Background = w['Brush'](w['Color'].FromRgb(250, 250, 250))

    # Compute transform: fit floor bbox into canvas with padding
    bx0, by0, bx1, by1 = poly_bbox(floor_pts)
    pad = 6
    fw, fh = bx1 - bx0, by1 - by0
    if fw <= 0 or fh <= 0: return canvas
    sx = (canvas_w - 2*pad) / fw
    sy = (canvas_h - 2*pad) / fh
    s  = min(sx, sy)
    ox = pad + (canvas_w - 2*pad - fw*s) * 0.5
    oy = pad + (canvas_h - 2*pad - fh*s) * 0.5

    def _to_canvas(p):
        # Flip Y so polygon reads correctly (canvas y grows down)
        return w['Point']((p.x - bx0) * s + ox,
                          canvas_h - ((p.y - by0) * s + oy))

    # Draw pieces (skip waste to keep preview tidy)
    for piece in option.pieces:
        if piece.piece_type == 'waste': continue
        pts = piece.pts
        if not pts or len(pts) < 3: continue
        polygon = w['Polygon']()
        pc = w['Points']()
        for v in pts:
            pc.Add(_to_canvas(v))
        polygon.Points = pc
        r, g, b = _PIECE_FILLS[piece.piece_type]
        polygon.Fill = w['Brush'](w['Color'].FromArgb(230, r, g, b))
        polygon.Stroke = w['Brush'](w['Color'].FromRgb(140, 140, 140))
        polygon.StrokeThickness = 0.4
        canvas.Children.Add(polygon)

    # Draw floor outline on top
    outline = w['Polyline']()
    pc = w['Points']()
    for v in list(floor_pts) + [floor_pts[0]]:
        pc.Add(_to_canvas(v))
    outline.Points = pc
    outline.Stroke = w['Brush'](w['Color'].FromRgb(44, 62, 80))
    outline.StrokeThickness = 1.5
    canvas.Children.Add(outline)

    return canvas


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 10 — REPORTING
# ═════════════════════════════════════════════════════════════════════════════

class ReportGenerator(object):

    def __init__(self, chosen_per_floor, params):
        """chosen_per_floor: list of (FloorInfo, LayoutOption, pattern)."""
        self.chosen = chosen_per_floor
        self.params = params

    def summary_text(self):
        tw_mm = self.params['tile_w_mm']
        th_mm = self.params['tile_h_mm']
        jw_mm = self.params['joint_mm']

        lines = [
            "=" * 68,
            "  TILE LAYOUT REPORT",
            "=" * 68,
            "  Tile size     : {:.0f} x {:.0f} mm".format(tw_mm, th_mm),
            "  Joint width   : {:.1f} mm".format(jw_mm),
            "  Nesting       : {}".format(
                "ON" if self.params['optimize_nesting'] else "OFF"),
            "=" * 68,
        ]

        grand_buy = 0
        grand_waste_area = 0.0
        grand_waste_denom = 0.0

        for fi, opt, pat in self.chosen:
            lines.append("")
            lines.append("  FLOOR  {}  ({})".format(fi.floor.Id, pat))
            lines.append("  Option {} — {}".format(opt.option_id, opt.variant))
            lines.append("  " + "-" * 62)
            lines.append("    Full tiles     : {:5d}".format(opt.n_full))
            lines.append("    Cut tiles (A)  : {:5d}".format(opt.n_cut))
            lines.append("    Reused (B/C..) : {:5d}".format(opt.n_reuse))
            lines.append("    TILES TO BUY   : {:5d}".format(opt.tiles_to_buy))
            lines.append("    Waste          : {:5.1f} %".format(opt.waste_pct))

            grand_buy += opt.tiles_to_buy
            grand_waste_area += sum(
                p.area for p in opt.pieces if p.piece_type == 'waste')
            grand_waste_denom += opt.tiles_to_buy * opt.tile_area

            log = getattr(opt, '_nesting_log', [])
            if log:
                lines.append("    Nesting log:")
                lines.extend("      " + e for e in log)

        grand_pct = (grand_waste_area / grand_waste_denom * 100.0
                     if grand_waste_denom > 0 else 0.0)
        lines.extend([
            "",
            "=" * 68,
            "  GRAND TOTAL",
            "    Tiles to buy   : {}".format(grand_buy),
            "    Overall waste  : {:.1f} %".format(grand_pct),
            "=" * 68,
        ])
        return "\n".join(lines)

    def export_csv(self, filepath):
        with open(filepath, 'wb') as fh:
            w = csv.writer(fh)
            w.writerow(['Floor_Id', 'Pattern', 'Option',
                        'Label', 'Type', 'Parent_ID',
                        'Area_ft2', 'Centroid_X', 'Centroid_Y'])
            for fi, opt, pat in self.chosen:
                for p in opt.pieces:
                    w.writerow([
                        str(fi.floor.Id), pat, opt.option_id,
                        p.label, p.piece_type, p.parent_id,
                        "{:.6f}".format(p.area),
                        "{:.4f}".format(p.centroid.x),
                        "{:.4f}".format(p.centroid.y),
                    ])


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 11 — SELECTION FILTER
# ═════════════════════════════════════════════════════════════════════════════

class _FloorFilter(ISelectionFilter):
    def AllowElement(self, e):
        return (e.Category is not None and
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_Floors))
    def AllowReference(self, r, p): return False


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 12 — LIST-VIEW ROW VIEWMODELS (for ListView + ItemsControl binding)
# ═════════════════════════════════════════════════════════════════════════════

class FloorRowVM(object):
    """Row data for the Step 1 ListView (read-only display)."""

    def __init__(self, floor_info, display_index):
        self._fi = floor_info
        self.Name        = "Floor #{}  (id {})".format(
            display_index, floor_info.floor.Id.IntegerValue)
        self.LevelName   = get_floor_level_name(floor_info.floor)
        w_mm = floor_info.width_ft  * FT_TO_MM
        h_mm = floor_info.height_ft * FT_TO_MM
        self.Dimensions  = "{:.0f} × {:.0f}".format(w_mm, h_mm)
        self.AreaM2      = "{:.1f}".format(floor_info.area_ft2 * FT2_TO_M2)
        self.VertexCount = str(len(floor_info.pts))


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 13 — WIZARD WINDOW
# ═════════════════════════════════════════════════════════════════════════════

STEP_BOUNDARIES = 0
STEP_PATTERN    = 1
STEP_CONCEPTS   = 2


class TileLayoutWindow(forms.WPFWindow):

    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'TileLayout.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        # ── wizard state ──
        self._floors = []        # [FloorInfo]
        self._rows   = []        # [FloorRowVM] (parallel to _floors)
        self._params = {}
        self._step   = STEP_BOUNDARIES
        self._selection_buttons = {}   # (floor_idx, option_id) → Border
        self._pattern_ctrls     = []   # [(ComboBox, TextBox)] per floor
        self._applied = False

        self._load_logo()
        self._refresh_step_ui()

    # ── logo ──────────────────────────────────────────────────────────────────
    def _load_logo(self):
        logo_path = os.path.normpath(os.path.join(
            os.path.dirname(__file__),
            '..', '..', '..', 'lib', 'GUI', 'T3Lab_logo.png'))
        if not os.path.exists(logo_path): return
        try:
            from System.Windows.Media.Imaging import BitmapImage
            from System import Uri, UriKind
            self.logo_image.Source = BitmapImage(Uri(logo_path, UriKind.Absolute))
        except Exception: pass

    # ── chrome ────────────────────────────────────────────────────────────────
    def minimize_button_clicked(self, sender, args):
        import System.Windows
        self.WindowState = System.Windows.WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        import System.Windows
        ws = System.Windows.WindowState
        self.WindowState = (ws.Normal if self.WindowState == ws.Maximized
                            else ws.Maximized)

    def close_button_clicked(self, sender, args):
        self.Close()

    # ── step indicator / action-bar refresh ──────────────────────────────────
    def _refresh_step_ui(self):
        """Update step circles, Back/Next button state and labels."""
        from System.Windows.Media import SolidColorBrush, Color as WColor
        active = SolidColorBrush(WColor.FromRgb(52, 152, 219))   # #3498DB
        inactive = SolidColorBrush(WColor.FromRgb(189, 195, 199)) # #BDC3C7
        active_text = SolidColorBrush(WColor.FromRgb(44, 62, 80))
        inactive_text = SolidColorBrush(WColor.FromRgb(127, 140, 141))

        circles = [self.step1_circle, self.step2_circle, self.step3_circle]
        labels  = [self.step1_label,  self.step2_label,  self.step3_label]
        for i, (c, lb) in enumerate(zip(circles, labels)):
            if i <= self._step:
                c.Background = active
                lb.Foreground = active_text
                lb.FontWeight = __import__('System').Windows.FontWeights.SemiBold
            else:
                c.Background = inactive
                lb.Foreground = inactive_text
                lb.FontWeight = __import__('System').Windows.FontWeights.Normal

        self.wizard_tabs.SelectedIndex = self._step
        self.btn_back.IsEnabled = self._step > STEP_BOUNDARIES

        if self._step == STEP_BOUNDARIES:
            self.btn_next.Content = "Next →"
            self.btn_next.IsEnabled = bool(self._floors)
        elif self._step == STEP_PATTERN:
            self.btn_next.Content = "Generate Concepts →"
            self.btn_next.IsEnabled = bool(self._floors)
        else:
            self.btn_next.Content = "Apply to Model ✓"
            self.btn_next.IsEnabled = self._every_floor_has_choice()

    def _every_floor_has_choice(self):
        return all(fi.chosen_option_id is not None for fi in self._floors)

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 1 — floor selection
    # ═════════════════════════════════════════════════════════════════════════

    def select_floors_clicked(self, sender, args):
        self.Hide()
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element, _FloorFilter(),
                "Select floor(s) for tile layout — press Finish when done")
            floors = [doc.GetElement(r.ElementId) for r in refs]
            self._extract_boundaries(floors)
        except Exception:
            self.status_text.Text = "Selection cancelled."
        finally:
            self.Show()
            self._refresh_step_ui()

    def _extract_boundaries(self, floor_elems):
        self._floors = []
        self._rows   = []
        total_area = 0.0
        for f in floor_elems:
            pts, z = extract_floor_boundary(f)
            if not pts:
                logger.warning("No boundary for floor {}".format(f.Id))
                continue
            fi = FloorInfo(f, ensure_ccw(pts), z)
            self._floors.append(fi)
            self._rows.append(FloorRowVM(fi, len(self._rows) + 1))
            total_area += fi.area_ft2

        from System.Collections.ObjectModel import ObservableCollection
        coll = ObservableCollection[object]()
        for r in self._rows: coll.Add(r)
        self.floors_listview.ItemsSource = coll

        self._build_pattern_ui()

        self.floors_count_text.Text = "{} floor(s) — boundaries extracted".format(
            len(self._floors))
        self.floors_total_text.Text = "Total area: {:.1f} m²".format(
            total_area * FT2_TO_M2)
        self.status_text.Text = "Step 1 done — click Next to configure patterns."

    def _build_pattern_ui(self):
        """Populate Step 2 — one row per floor with pattern combo + angle box."""
        import System.Windows.Controls as WC
        import System.Windows as SW
        from System.Windows.Media import SolidColorBrush, Color as WColor

        host = self.pattern_stack
        host.Children.Clear()
        self._pattern_ctrls = []

        row_brush_alt = SolidColorBrush(WColor.FromRgb(250, 250, 250))

        for i, (fi, row) in enumerate(zip(self._floors, self._rows)):
            outer = WC.Border()
            outer.BorderBrush = SolidColorBrush(WColor.FromRgb(236, 240, 241))
            outer.BorderThickness = SW.Thickness(0, 0, 0, 1)
            outer.Padding = SW.Thickness(14, 10, 14, 10)
            if i % 2 == 1: outer.Background = row_brush_alt

            grid = WC.Grid()
            for w in (SW.GridLength(1, SW.GridUnitType.Star),
                      SW.GridLength(200),
                      SW.GridLength(70)):
                cd = WC.ColumnDefinition(); cd.Width = w
                grid.ColumnDefinitions.Add(cd)

            # Label block
            lbl = WC.StackPanel(); lbl.VerticalAlignment = SW.VerticalAlignment.Center
            name_tb = WC.TextBlock()
            name_tb.Text = row.Name
            name_tb.FontSize = 13
            name_tb.FontWeight = SW.FontWeights.SemiBold
            name_tb.Foreground = SolidColorBrush(WColor.FromRgb(44, 62, 80))
            sub_tb = WC.TextBlock()
            sub_tb.Text = "{}  ·  {} m²  ·  {} mm".format(
                row.LevelName, row.AreaM2, row.Dimensions)
            sub_tb.FontSize = 11
            sub_tb.Foreground = SolidColorBrush(WColor.FromRgb(127, 140, 141))
            sub_tb.Margin = SW.Thickness(0, 2, 0, 0)
            lbl.Children.Add(name_tb); lbl.Children.Add(sub_tb)
            WC.Grid.SetColumn(lbl, 0); grid.Children.Add(lbl)

            # Pattern combo
            cmb = WC.ComboBox()
            cmb.Items.Add("Grid (Stacked Bond)")
            cmb.Items.Add("Staggered (Running Bond)")
            cmb.SelectedIndex = 0
            cmb.FontSize = 12
            cmb.Padding = SW.Thickness(8, 5, 8, 5)
            cmb.BorderBrush = SolidColorBrush(WColor.FromRgb(189, 195, 199))
            cmb.Margin = SW.Thickness(0, 0, 10, 0)
            cmb.VerticalAlignment = SW.VerticalAlignment.Center
            WC.Grid.SetColumn(cmb, 1); grid.Children.Add(cmb)

            # Angle box
            txt = WC.TextBox()
            txt.Text = "0"
            txt.FontSize = 12
            txt.Padding = SW.Thickness(8, 6, 8, 6)
            txt.BorderBrush = SolidColorBrush(WColor.FromRgb(189, 195, 199))
            txt.BorderThickness = SW.Thickness(1)
            txt.VerticalAlignment = SW.VerticalAlignment.Center
            txt.ToolTip = "Layout angle (degrees)"
            WC.Grid.SetColumn(txt, 2); grid.Children.Add(txt)

            outer.Child = grid
            host.Children.Add(outer)
            self._pattern_ctrls.append((cmb, txt))

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 2 → generate options
    # ═════════════════════════════════════════════════════════════════════════

    def _generate_concepts(self):
        try:
            params = self._read_params()
        except ValueError as exc:
            TaskDialog.Show("Input Error", str(exc))
            return False
        self._params = params

        self.status_text.Text = "Generating candidate layouts…"
        self.UpdateLayout()

        gen = OptionGenerator(params['tile_w_ft'], params['tile_h_ft'],
                              params['joint_ft'], params['optimize_nesting'],
                              top_n=4)

        for fi, (cmb, txt) in zip(self._floors, self._pattern_ctrls):
            pattern = 'staggered' if cmb.SelectedIndex == 1 else 'grid'
            try:
                base_angle = float((txt.Text or "0").strip())
            except ValueError:
                base_angle = 0.0
            fi.options = gen.generate(fi, pattern, base_angle)
            fi.chosen_option_id = fi.options[0].option_id if fi.options else None
            fi._pattern = pattern

        self._build_concepts_ui()
        self.status_text.Text = (
            "Generated {} options × {} floor(s). Click a card to change the choice."
            .format(sum(len(f.options) for f in self._floors),
                    len(self._floors)))
        return True

    def _build_concepts_ui(self):
        """Populate the Step 3 stack with one section per floor."""
        import System.Windows.Controls as WC
        import System.Windows as SW
        from System.Windows.Media import SolidColorBrush, Color as WColor

        host = self.concepts_host
        host.Children.Clear()
        self._selection_buttons = {}

        for fi_idx, fi in enumerate(self._floors):
            # Section header
            hdr = WC.TextBlock()
            hdr.Text = "{}  —  {} · {} mm".format(
                self._rows[fi_idx].Name,
                self._rows[fi_idx].AreaM2 + " m²",
                self._rows[fi_idx].Dimensions)
            hdr.FontSize = 14
            hdr.FontWeight = SW.FontWeights.SemiBold
            hdr.Foreground = SolidColorBrush(WColor.FromRgb(44, 62, 80))
            hdr.Margin = SW.Thickness(0, 16, 0, 8)
            host.Children.Add(hdr)

            # 4-card grid
            grid = WC.UniformGrid()
            grid.Columns = len(fi.options) if fi.options else 1
            grid.Rows = 1

            for opt in fi.options:
                card = self._build_option_card(fi_idx, fi, opt)
                grid.Children.Add(card)
            host.Children.Add(grid)

        # Paint the default-selected option ("A") for each floor
        for i in range(len(self._floors)):
            self._refresh_selection_highlights(i)

    def _build_option_card(self, fi_idx, fi, opt):
        import System.Windows.Controls as WC
        import System.Windows as SW
        from System.Windows.Media import SolidColorBrush, Color as WColor

        brush_border = SolidColorBrush(WColor.FromRgb(189, 195, 199))
        brush_select = SolidColorBrush(WColor.FromRgb(52, 152, 219))
        brush_text   = SolidColorBrush(WColor.FromRgb(44, 62, 80))
        brush_sub    = SolidColorBrush(WColor.FromRgb(127, 140, 141))

        outer = WC.Border()
        outer.BorderBrush = brush_border
        outer.BorderThickness = SW.Thickness(2)
        outer.CornerRadius = SW.CornerRadius(6)
        outer.Margin = SW.Thickness(6)
        outer.Padding = SW.Thickness(10)
        outer.Background = SolidColorBrush(WColor.FromRgb(255, 255, 255))
        outer.Cursor = SW.Input.Cursors.Hand

        stack = WC.StackPanel()

        # Header: "● Option A — best"
        header = WC.StackPanel(); header.Orientation = WC.Orientation.Horizontal
        dot = WC.TextBlock()
        dot.Text = "●"; dot.FontSize = 16; dot.Foreground = brush_sub
        dot.Margin = SW.Thickness(0, 0, 6, 0)
        name = WC.TextBlock()
        name.Text = "Option {}".format(opt.option_id)
        name.FontSize = 13; name.FontWeight = SW.FontWeights.SemiBold
        name.Foreground = brush_text
        header.Children.Add(dot); header.Children.Add(name)
        stack.Children.Add(header)

        desc = WC.TextBlock()
        desc.Text = opt.variant; desc.FontSize = 10; desc.Foreground = brush_sub
        desc.Margin = SW.Thickness(0, 0, 0, 8)
        stack.Children.Add(desc)

        # Preview canvas
        preview = render_option_preview(opt, fi.pts, canvas_w=190, canvas_h=130)
        preview_host = WC.Border()
        preview_host.Background = SolidColorBrush(WColor.FromRgb(250, 250, 250))
        preview_host.BorderBrush = SolidColorBrush(WColor.FromRgb(236, 240, 241))
        preview_host.BorderThickness = SW.Thickness(1)
        preview_host.CornerRadius = SW.CornerRadius(4)
        preview_host.Child = preview
        preview_host.HorizontalAlignment = SW.HorizontalAlignment.Center
        preview_host.Margin = SW.Thickness(0, 0, 0, 8)
        stack.Children.Add(preview_host)

        # Stat rows
        def _stat_row(k, v, color=None):
            row = WC.StackPanel(); row.Orientation = WC.Orientation.Horizontal
            row.Margin = SW.Thickness(0, 1, 0, 1)
            kt = WC.TextBlock(); kt.Text = k; kt.FontSize = 11
            kt.Foreground = brush_sub; kt.Width = 105
            vt = WC.TextBlock(); vt.Text = v; vt.FontSize = 11
            vt.FontWeight = SW.FontWeights.SemiBold
            vt.Foreground = color or brush_text
            row.Children.Add(kt); row.Children.Add(vt)
            return row

        ok_col    = SolidColorBrush(WColor.FromRgb( 39, 174,  96))
        warn_col  = SolidColorBrush(WColor.FromRgb(231,  76,  60))
        waste_col = ok_col if opt.waste_pct < 10 else warn_col

        stack.Children.Add(_stat_row(
            "Full tiles:", "{}".format(opt.n_full)))
        stack.Children.Add(_stat_row(
            "Cut (A):",    "{}".format(opt.n_cut)))
        stack.Children.Add(_stat_row(
            "Reused (B):", "{}".format(opt.n_reuse), ok_col if opt.n_reuse else brush_text))
        stack.Children.Add(_stat_row(
            "Tiles to buy:", "{}".format(opt.tiles_to_buy)))
        stack.Children.Add(_stat_row(
            "Waste:", "{:.1f} %".format(opt.waste_pct), waste_col))

        outer.Child = stack

        # Click → select this option
        def _on_click(s, e):
            fi.chosen_option_id = opt.option_id
            self._refresh_selection_highlights(fi_idx)
            self._refresh_step_ui()
        outer.PreviewMouseLeftButtonUp += _on_click

        self._selection_buttons[(fi_idx, opt.option_id)] = outer
        return outer

    def _refresh_selection_highlights(self, fi_idx):
        import System.Windows as SW
        from System.Windows.Media import SolidColorBrush, Color as WColor
        selected = SolidColorBrush(WColor.FromRgb(52, 152, 219))
        normal   = SolidColorBrush(WColor.FromRgb(189, 195, 199))
        fi = self._floors[fi_idx]
        for (idx, oid), border in self._selection_buttons.items():
            if idx != fi_idx: continue
            if oid == fi.chosen_option_id:
                border.BorderBrush = selected
                border.BorderThickness = SW.Thickness(2.5)
            else:
                border.BorderBrush = normal
                border.BorderThickness = SW.Thickness(2)

    # ═════════════════════════════════════════════════════════════════════════
    # STEP 3 → apply to model
    # ═════════════════════════════════════════════════════════════════════════

    def _apply_selection(self):
        chosen = []
        for fi in self._floors:
            if fi.chosen_option_id is None: continue
            match = [o for o in fi.options if o.option_id == fi.chosen_option_id]
            if not match: continue
            chosen.append((fi, match[0], getattr(fi, '_pattern', 'grid')))

        if not chosen:
            TaskDialog.Show("Tile Layout", "No selections to apply.")
            return

        try:
            with revit.Transaction("Tile Layout — apply selected concepts"):
                view = get_or_create_3d_view()
                for fi, opt, _pat in chosen:
                    vis = RevitVisualizer(view)
                    for piece in opt.pieces:
                        vis.draw_piece(piece, fi.z)
            uidoc.ActiveView = view
        except Exception as exc:
            TaskDialog.Show("Apply Error",
                "{}\n\n{}".format(exc, traceback.format_exc()))
            return

        # Report
        params_mm = dict(
            tile_w_mm=self._params['tile_w_ft'] * FT_TO_MM,
            tile_h_mm=self._params['tile_h_ft'] * FT_TO_MM,
            joint_mm =self._params['joint_ft']  * FT_TO_MM,
            optimize_nesting=self._params['optimize_nesting'])
        rpt = ReportGenerator(chosen, params_mm)
        out = script.get_output()
        out.print_md("```\n{}\n```".format(rpt.summary_text()))

        self._applied = True
        self._report  = rpt
        self.btn_export.IsEnabled = True
        self.status_text.Text = (
            "Applied {} option(s). DirectShapes created in 'Tile Layout Preview'."
            .format(len(chosen)))

    # ═════════════════════════════════════════════════════════════════════════
    # Action-bar buttons
    # ═════════════════════════════════════════════════════════════════════════

    def next_clicked(self, sender, args):
        if self._step == STEP_BOUNDARIES:
            if not self._floors:
                TaskDialog.Show("Tile Layout", "Select floors first.")
                return
            self._step = STEP_PATTERN
            self._refresh_step_ui()
        elif self._step == STEP_PATTERN:
            if self._generate_concepts():
                self._step = STEP_CONCEPTS
                self._refresh_step_ui()
        elif self._step == STEP_CONCEPTS:
            self._apply_selection()

    def back_clicked(self, sender, args):
        if self._step > STEP_BOUNDARIES:
            self._step -= 1
            self._refresh_step_ui()

    def export_csv_clicked(self, sender, args):
        if not getattr(self, '_report', None):
            TaskDialog.Show("Export", "Apply the layout first.")
            return
        try:
            from System.Windows.Forms import SaveFileDialog, DialogResult
            dlg = SaveFileDialog()
            dlg.Title    = "Save Tile Layout CSV"
            dlg.Filter   = "CSV files (*.csv)|*.csv"
            dlg.FileName = "TileLayout.csv"
            if dlg.ShowDialog() != DialogResult.OK: return
            self._report.export_csv(dlg.FileName)
            self.status_text.Text = "CSV saved: {}".format(
                os.path.basename(dlg.FileName))
        except Exception as exc:
            TaskDialog.Show("CSV Export Error", str(exc))

    # ═════════════════════════════════════════════════════════════════════════
    # Input parsing
    # ═════════════════════════════════════════════════════════════════════════

    def _read_params(self):
        def _mm(ctrl, name):
            try: v = float(ctrl.Text.strip())
            except (ValueError, AttributeError):
                raise ValueError("'{}' is not a valid number.".format(name))
            if v <= 0:
                raise ValueError("'{}' must be greater than zero.".format(name))
            return v * MM_TO_FT
        return dict(
            tile_w_ft = _mm(self.txt_tile_w, "Tile Width"),
            tile_h_ft = _mm(self.txt_tile_h, "Tile Height"),
            joint_ft  = _mm(self.txt_joint,  "Joint Width"),
            optimize_nesting = bool(self.chk_nesting.IsChecked),
        )


# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    TileLayoutWindow().ShowDialog()
