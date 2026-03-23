# -*- coding: utf-8 -*-
"""
Property Line Dialog
Create US property lines in Revit from Lightbox RE parcel API data.

Flow:
1. User enters a US address
2. Tool queries Lightbox API for parcel boundary (GeoJSON polygon)
3. Coordinates are converted from WGS84 lat/lon to Revit feet
4. Property lines are drawn as ModelLines in the active Revit document
"""
__title__ = "Property Line Tool"
__author__ = "T3Lab"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import os
import sys
import clr
import json
import math
import traceback
import threading

# .NET / WPF
clr.AddReference("System")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
from System import Uri, Action
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window, Visibility, Application
from System.Windows.Input import Key
from System.Windows.Media import SolidColorBrush, Color
from System.Windows.Threading import Dispatcher, DispatcherPriority

# IronPython HTTP (urllib2 available in IronPython 2.x)
try:
    import urllib2
    HAS_URLLIB2 = True
except ImportError:
    HAS_URLLIB2 = False

try:
    import urllib
    import urllib.request as urllib_request
    HAS_URLLIB3 = True
except ImportError:
    HAS_URLLIB3 = False

# pyRevit
from pyrevit import revit, DB, forms, script

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
logger = script.get_logger()

# Config file
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".t3lab")
CONFIG_FILE = os.path.join(CONFIG_DIR, "property_line_config.json")

# Lightbox API
LIGHTBOX_BASE = "https://api.lightboxre.com"
LIGHTBOX_PARCELS_ENDPOINT = "/v1/parcels/us"

# Earth radius in feet (for coordinate conversion)
EARTH_RADIUS_FT = 20902231.0

# Zoning / setback endpoint  (appended as: LIGHTBOX_BASE + LIGHTBOX_PARCELS_ENDPOINT + "/{id}/zoning")
LIGHTBOX_ZONING_PATH = "/zoning"


# ╔═╗╔═╗╔╗╔╔═╗╦╔═╗
# ║  ║ ║║║║╠╣ ║║ ╦
# ╚═╝╚═╝╝╚╝╚  ╩╚═╝ CONFIG HELPERS
# ==================================================

def ensure_config_dir():
    if not os.path.exists(CONFIG_DIR):
        os.makedirs(CONFIG_DIR)


def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_config(data):
    try:
        ensure_config_dir()
        existing = load_config()
        existing.update(data)
        with open(CONFIG_FILE, 'w') as f:
            json.dump(existing, f, indent=2)
        return True
    except Exception as ex:
        logger.error("Failed to save config: {}".format(ex))
        return False


# ╔═╗╔═╗╔═╗╦═╗╔╦╗╦╔╗╔╔═╗╔╦╗╔═╗╔═╗
# ║  ║ ║║ ║╠╦╝ ║║║║║║╠═╣ ║ ║╣ ╚═╗
# ╚═╝╚═╝╚═╝╩╚══╩╝╩╝╚╝╩ ╩ ╩ ╚═╝╚═╝ COORDINATE UTILS
# ==================================================

def latlon_to_feet(lat, lon, origin_lat, origin_lon):
    """
    Convert WGS84 lat/lon to Revit internal feet,
    relative to a chosen origin point.

    Uses the equirectangular approximation which is accurate
    for small areas (a few miles), sufficient for property parcels.

    Returns (x_ft, y_ft) in feet where:
      +X = East
      +Y = North
    """
    dlat = math.radians(lat - origin_lat)
    dlon = math.radians(lon - origin_lon)
    cos_lat = math.cos(math.radians(origin_lat))

    x_ft = EARTH_RADIUS_FT * dlon * cos_lat
    y_ft = EARTH_RADIUS_FT * dlat
    return x_ft, y_ft


def compute_centroid(coordinates):
    """Compute centroid of a polygon (list of [lon, lat] pairs)."""
    if not coordinates:
        return 0.0, 0.0
    lons = [c[0] for c in coordinates]
    lats = [c[1] for c in coordinates]
    return sum(lats) / len(lats), sum(lons) / len(lons)


def compute_area_sqft(coordinates):
    """Shoelace formula in lat/lon => approximate area in sqft."""
    if len(coordinates) < 3:
        return 0.0
    # Pick centroid as origin for conversion
    clat, clon = compute_centroid(coordinates)
    pts = [latlon_to_feet(c[1], c[0], clat, clon) for c in coordinates]

    n = len(pts)
    area = 0.0
    for i in range(n):
        j = (i + 1) % n
        area += pts[i][0] * pts[j][1]
        area -= pts[j][0] * pts[i][1]
    return abs(area) / 2.0


def format_area(sqft):
    """Format area as sqft and acres."""
    acres = sqft / 43560.0
    if acres >= 1.0:
        return "{:,.0f} sqft ({:.3f} ac)".format(sqft, acres)
    return "{:,.0f} sqft".format(sqft)


# ╔═╗╔═╗╦  ╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
# ╚═╗║╣ ║  ║╣ ║ ║║ ║║║║╚═╗
# ╚═╝╚═╝╩═╝╚═╝╩ ╩╚═╝╝╚╝╚═╝ LIGHTBOX API
# ==================================================

def http_get(url, headers=None):
    """
    Perform a GET request using urllib2 (IronPython) or urllib.request (CPython).
    Returns (status_code, response_text) or raises an exception.
    """
    headers = headers or {}

    if HAS_URLLIB2:
        req = urllib2.Request(url)
        for k, v in headers.items():
            req.add_header(k, v)
        try:
            response = urllib2.urlopen(req, timeout=15)
            return response.getcode(), response.read()
        except urllib2.HTTPError as e:
            return e.code, e.read()
        except Exception as ex:
            raise ex

    elif HAS_URLLIB3:
        req = urllib_request.Request(url, headers=headers)
        try:
            with urllib_request.urlopen(req, timeout=15) as response:
                return response.status, response.read().decode('utf-8')
        except Exception as e:
            if hasattr(e, 'code'):
                return e.code, str(e)
            raise

    else:
        raise RuntimeError("No HTTP library available (urllib2 / urllib.request)")


def search_parcels(api_key, address, limit=10):
    """
    Query Lightbox parcels API.

    GET /v1/parcels/us?text={address}&limit={limit}
    Header: x-api-key: {api_key}

    Returns list of parcel dicts with keys:
        id, display_address, parcel_id, area_sqft, geometry, county, state
    """
    url = "{}{}/search?text={}&limit={}".format(
        LIGHTBOX_BASE,
        LIGHTBOX_PARCELS_ENDPOINT,
        urllib2.quote(address) if HAS_URLLIB2 else urllib.parse.quote(address),
        limit
    )

    headers = {
        "x-api-key": api_key,
        "Accept": "application/json",
    }

    status, body = http_get(url, headers)

    if isinstance(body, bytes):
        body = body.decode('utf-8', errors='replace')

    if status != 200:
        raise ValueError("Lightbox API error {}: {}".format(status, body[:300]))

    data = json.loads(body)

    parcels = []
    # Lightbox returns { "parcels": [...] } or similar structure
    items = data.get("parcels", data.get("results", data.get("data", [])))

    for item in items:
        try:
            parcel = _parse_parcel(item)
            if parcel:
                parcels.append(parcel)
        except Exception as ex:
            logger.warning("Skipping parcel parse error: {}".format(ex))

    return parcels


def _parse_parcel(item):
    """Parse a raw Lightbox parcel response item into a clean dict."""
    # Lightbox response shape (v1):
    # {
    #   "id": "...",
    #   "attributes": {
    #     "parcelId": "...",
    #     "address": { "oneLine": "...", ... },
    #     "areaAcres": 0.25,
    #     "countyName": "...",
    #     "stateFips": "06",
    #     "stateCode": "CA"
    #   },
    #   "geometry": {
    #     "type": "Polygon" | "MultiPolygon",
    #     "coordinates": [[[lon, lat], ...]]
    #   }
    # }
    attrs = item.get("attributes", item)  # handle both wrapped and flat

    # Try multiple field names used across API versions
    parcel_id = (attrs.get("parcelId") or
                 attrs.get("parcelnumb") or
                 attrs.get("parcel_id") or
                 item.get("id", "N/A"))

    address_raw = attrs.get("address", {})
    if isinstance(address_raw, dict):
        display_address = (address_raw.get("oneLine") or
                           address_raw.get("line1", "") + ", " +
                           address_raw.get("city", ""))
    else:
        display_address = str(address_raw)

    if not display_address.strip():
        display_address = attrs.get("siteAddress", attrs.get("address1", "Unknown"))

    area_acres = (attrs.get("areaAcres") or
                  attrs.get("area_ac") or
                  attrs.get("calc_acreage") or 0.0)
    try:
        area_acres = float(area_acres)
    except (TypeError, ValueError):
        area_acres = 0.0
    area_sqft = area_acres * 43560.0

    geometry = item.get("geometry") or attrs.get("geometry")
    if not geometry:
        return None

    county = attrs.get("countyName", attrs.get("county", ""))
    state = attrs.get("stateCode", attrs.get("state2", ""))

    return {
        "id":              item.get("id", parcel_id),
        "parcel_id":       parcel_id,
        "display_address": display_address.strip(", "),
        "area_sqft":       "{:,.0f}".format(area_sqft) if area_sqft else "N/A",
        "area_sqft_raw":   area_sqft,
        "geometry":        geometry,
        "county":          county,
        "state":           state,
    }


def get_polygon_coords(geometry):
    """
    Extract the outer ring coordinates from a GeoJSON Polygon or MultiPolygon.
    Returns list of [lon, lat] pairs.
    """
    geo_type = geometry.get("type", "")
    coords = geometry.get("coordinates", [])

    if geo_type == "Polygon":
        # coords = [ [[lon,lat],...], [[hole...]] ]
        return coords[0] if coords else []
    elif geo_type == "MultiPolygon":
        # coords = [ [[[lon,lat],...]], ... ]
        # Return the ring with the most points (largest polygon)
        best = []
        for poly in coords:
            if poly and len(poly[0]) > len(best):
                best = poly[0]
        return best
    return []


# ╔═╗╔═╗╔╦╗╔╗ ╔═╗╔═╗╦╔═  ╔═╗╔═╗╦
# ╚═╗║╣  ║ ╠╩╗╠═╣║  ╠╩╗  ╠═╣╠═╝║
# ╚═╝╚═╝ ╩ ╚═╝╩ ╩╚═╝╩ ╩  ╩ ╩╩  ╩ SETBACK / ZONING
# ==================================================

def get_zoning_data(api_key, parcel_id):
    """
    Query Lightbox zoning endpoint for a parcel.

    GET /v1/parcels/us/{parcel_id}/zoning
    Header: x-api-key

    Returns dict:  {"front_ft": float|None, "rear_ft": float|None, "side_ft": float|None}
    Raises ValueError on non-200 response.
    """
    url = "{}{}/{}/zoning".format(LIGHTBOX_BASE, LIGHTBOX_PARCELS_ENDPOINT, parcel_id)
    headers = {"x-api-key": api_key, "Accept": "application/json"}

    status, body = http_get(url, headers)
    if isinstance(body, bytes):
        body = body.decode("utf-8", errors="replace")
    if status != 200:
        raise ValueError("Zoning API error {}: {}".format(status, body[:200]))

    return _parse_zoning(json.loads(body))


def _parse_zoning(data):
    """
    Parse Lightbox zoning response into a clean setback dict (feet).
    Handles nested and flat structures.
    """
    zone = (data.get("zoning") or
            data.get("data") or
            data.get("attributes") or
            data)
    if isinstance(zone, list):
        zone = zone[0] if zone else {}

    def _to_ft(val):
        if val is None:
            return None
        try:
            v = float(val)
            return v if 0 < v <= 500 else None
        except (TypeError, ValueError):
            return None

    front = _to_ft(zone.get("frontSetback") or zone.get("front_setback") or
                   zone.get("front") or zone.get("setback_front"))
    rear  = _to_ft(zone.get("rearSetback")  or zone.get("rear_setback")  or
                   zone.get("rear")  or zone.get("setback_rear"))
    side  = _to_ft(zone.get("sideSetback")  or zone.get("side_setback")  or
                   zone.get("side")  or zone.get("setback_side") or
                   zone.get("sideYardSetback"))

    return {"front_ft": front, "rear_ft": rear, "side_ft": side}


# ── polygon math (pure Python, no shapely) ───────────────────────────────────

def _polygon_signed_area_2d(pts):
    """Signed shoelace area. Positive → CCW."""
    n = len(pts)
    a = 0.0
    for i in range(n):
        j = (i + 1) % n
        a += pts[i][0] * pts[j][1] - pts[j][0] * pts[i][1]
    return a / 2.0


def _line_intersect_2d(p1, p2, p3, p4):
    """Intersection of infinite lines through (p1,p2) and (p3,p4). Returns None if parallel."""
    x1, y1 = p1;  x2, y2 = p2
    x3, y3 = p3;  x4, y4 = p4
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    if abs(denom) < 1e-10:
        return None
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom
    return (x1 + t * (x2 - x1), y1 + t * (y2 - y1))


def inset_polygon_2d(pts, distance):
    """
    Shrink a polygon inward by `distance` feet.

    pts      – open list of (x, y) in Revit feet  (do NOT repeat first point)
    distance – positive offset distance in feet

    Returns a new open list of (x, y) with the same vertex count.
    Raises ValueError if the polygon would collapse (setback too large).
    """
    n = len(pts)
    if n < 3:
        raise ValueError("Need at least 3 vertices for polygon inset")

    # Normalise winding to CCW so that left-of-edge = interior
    if _polygon_signed_area_2d(pts) < 0:
        pts = list(reversed(pts))

    # Build one offset edge per polygon edge
    offset_edges = []
    for i in range(n):
        p1 = pts[i]
        p2 = pts[(i + 1) % n]
        dx = p2[0] - p1[0]
        dy = p2[1] - p1[1]
        length = math.sqrt(dx * dx + dy * dy)
        if length < 1e-10:
            continue
        # Inward (left) unit normal for CCW polygon: (-dy/L, +dx/L)
        nx, ny = -dy / length, dx / length
        offset_edges.append(
            ((p1[0] + nx * distance, p1[1] + ny * distance),
             (p2[0] + nx * distance, p2[1] + ny * distance))
        )

    if len(offset_edges) < 3:
        raise ValueError("Too few valid edges for polygon inset")

    # New vertices = intersection of consecutive offset edges
    m = len(offset_edges)
    result = []
    for i in range(m):
        e1 = offset_edges[i]
        e2 = offset_edges[(i + 1) % m]
        pt = _line_intersect_2d(e1[0], e1[1], e2[0], e2[1])
        result.append(pt if pt is not None else e1[1])

    # Sanity-check: inset area must be at least 1 % of original
    orig_area  = abs(_polygon_signed_area_2d(pts))
    inset_area = abs(_polygon_signed_area_2d(result))
    if orig_area > 0 and inset_area < orig_area * 0.01:
        raise ValueError(
            "Setback distance ({:.1f} ft) is too large: polygon collapsed".format(distance)
        )

    return result


def create_setback_lines_in_revit(doc, coordinates, setback_ft,
                                   elevation_ft=0.0,
                                   origin_mode="Project Base Point"):
    """
    Draw a setback envelope as Model Lines by insetting the parcel boundary.

    Parameters:
        doc          - Revit Document
        coordinates  - list of [lon, lat] from GeoJSON outer ring
        setback_ft   - inset distance in feet (must be > 0)
        elevation_ft - Z elevation in feet
        origin_mode  - "Project Base Point" | "Survey Point" | "World Origin (0,0,0)"

    Returns the number of line segments created.
    """
    if setback_ft <= 0:
        raise ValueError("Setback distance must be greater than zero")

    # Convert GeoJSON → feet, relative to polygon centroid
    centroid_lat, centroid_lon = compute_centroid(coordinates)
    pts_2d = []
    for c in coordinates:
        lon, lat = c[0], c[1]
        x_ft, y_ft = latlon_to_feet(lat, lon, centroid_lat, centroid_lon)
        pts_2d.append((x_ft, y_ft))

    # Drop closing duplicate if present
    if (len(pts_2d) > 1 and
            abs(pts_2d[0][0] - pts_2d[-1][0]) < 0.001 and
            abs(pts_2d[0][1] - pts_2d[-1][1]) < 0.001):
        pts_2d = pts_2d[:-1]

    # Inset the polygon
    inset_2d = inset_polygon_2d(pts_2d, setback_ft)

    # Determine insertion origin (same logic as property lines)
    if origin_mode == "Survey Point":
        offset = get_survey_point(doc)
    elif origin_mode == "Project Base Point":
        offset = get_project_base_point(doc)
    else:
        offset = DB.XYZ(0, 0, 0)

    # Convert to Revit XYZ and close the loop
    revit_pts = [
        DB.XYZ(p[0] + offset.X, p[1] + offset.Y, elevation_ft + offset.Z)
        for p in inset_2d
    ]
    revit_pts.append(revit_pts[0])

    count = 0
    with DB.Transaction(doc, "Create Setback Envelope") as t:
        t.Start()
        count = _create_model_lines_from_pts(doc, revit_pts, elevation_ft + offset.Z)
        t.Commit()

    return count


# ╦═╗╔═╗╦  ╦╦╔╦╗  ╔═╗╦═╗╔═╗╔═╗╔╦╗╦╔═╗╔╗╔
# ╠╦╝║╣ ╚╗╔╝║ ║   ║  ╠╦╝║╣ ╠═╣ ║ ║║ ║║║║
# ╩╚═╚═╝ ╚╝ ╩ ╩   ╚═╝╩╚═╚═╝╩ ╩ ╩ ╩╚═╝╝╚╝ REVIT CREATION
# ==================================================

def get_project_base_point(doc):
    """Get the project base point in Revit internal feet."""
    collector = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_ProjectBasePoint
    ).WhereElementIsNotElementType().ToElements()
    if collector:
        bp = collector[0]
        loc = bp.Location
        if hasattr(loc, 'Point'):
            return loc.Point
    return DB.XYZ(0, 0, 0)


def get_survey_point(doc):
    """Get the survey point in Revit internal feet."""
    collector = DB.FilteredElementCollector(doc).OfCategory(
        DB.BuiltInCategory.OST_SharedBasePoint
    ).WhereElementIsNotElementType().ToElements()
    if collector:
        sp = collector[0]
        loc = sp.Location
        if hasattr(loc, 'Point'):
            return loc.Point
    return DB.XYZ(0, 0, 0)


def create_property_lines_in_revit(doc, coordinates, elevation_ft=0.0,
                                   line_category="Property Lines",
                                   origin_mode="Project Base Point"):
    """
    Create property boundary lines in the Revit document.

    Parameters:
        doc           - Revit Document
        coordinates   - list of [lon, lat] from GeoJSON outer ring
        elevation_ft  - Z elevation in feet
        line_category - "Property Lines" | "Model Lines" | "Detail Lines"
        origin_mode   - where to place the centroid

    Returns number of lines created.
    """
    if len(coordinates) < 2:
        raise ValueError("Need at least 2 coordinates to create lines")

    # Compute centroid for coordinate origin
    centroid_lat, centroid_lon = compute_centroid(coordinates)

    # Convert all coords to Revit XYZ (feet)
    revit_pts = []
    for c in coordinates:
        lon, lat = c[0], c[1]
        x_ft, y_ft = latlon_to_feet(lat, lon, centroid_lat, centroid_lon)
        revit_pts.append(DB.XYZ(x_ft, y_ft, elevation_ft))

    # Determine insertion offset (project base / survey / world origin)
    if origin_mode == "Survey Point":
        offset = get_survey_point(doc)
    elif origin_mode == "Project Base Point":
        offset = get_project_base_point(doc)
    else:
        offset = DB.XYZ(0, 0, 0)

    # Translate points to chosen origin
    revit_pts = [DB.XYZ(pt.X + offset.X, pt.Y + offset.Y, pt.Z + offset.Z)
                 for pt in revit_pts]

    # Close the loop: first == last
    if revit_pts[0].DistanceTo(revit_pts[-1]) > 0.001:
        revit_pts.append(revit_pts[0])

    lines_created = 0

    with DB.Transaction(doc, "Create Property Lines") as t:
        t.Start()

        if line_category == "Property Lines":
            # Use Revit's native PropertyLine element
            lines_created = _create_native_property_lines(doc, revit_pts)
        elif line_category == "Detail Lines":
            lines_created = _create_detail_lines(doc, revit_pts)
        else:
            # Default: Model Lines
            lines_created = _create_model_lines(doc, revit_pts, elevation_ft)

        t.Commit()

    return lines_created


def _create_native_property_lines(doc, pts):
    """Create Revit PropertyLine elements (site category)."""
    count = 0
    try:
        curve_loop = DB.CurveLoop()
        for i in range(len(pts) - 1):
            start = pts[i]
            end = pts[i + 1]
            if start.DistanceTo(end) < 0.001:
                continue
            line = DB.Line.CreateBound(start, end)
            curve_loop.Append(line)

        # PropertyLine.Create requires a CurveLoop and a site topo/level
        # Try the simplest overload first
        if hasattr(DB, 'PropertyLine'):
            prop_line = DB.PropertyLine.Create(doc, curve_loop)
            count = 1
        else:
            # Fallback to model lines if PropertyLine not available
            count = _create_model_lines_from_pts(doc, pts, pts[0].Z)
    except Exception as ex:
        logger.warning("PropertyLine creation failed, falling back to ModelLine: {}".format(ex))
        count = _create_model_lines_from_pts(doc, pts, pts[0].Z)
    return count


def _create_model_lines(doc, pts, elevation_ft):
    """Create ModelLine elements on a horizontal sketch plane."""
    return _create_model_lines_from_pts(doc, pts, elevation_ft)


def _create_model_lines_from_pts(doc, pts, elevation_ft):
    """Internal: create model lines from a list of XYZ points."""
    count = 0
    try:
        # Build sketch plane at the given elevation
        normal = DB.XYZ.BasisZ
        origin = DB.XYZ(0, 0, elevation_ft)
        plane = DB.Plane.CreateByNormalAndOrigin(normal, origin)
        sketch_plane = DB.SketchPlane.Create(doc, plane)

        for i in range(len(pts) - 1):
            start = pts[i]
            end = pts[i + 1]
            if start.DistanceTo(end) < 0.001:
                continue
            line = DB.Line.CreateBound(start, end)
            doc.Create.NewModelCurve(line, sketch_plane)
            count += 1
    except Exception as ex:
        logger.error("Model line creation error: {}".format(ex))
        raise
    return count


def _create_detail_lines(doc, pts):
    """Create DetailLine elements in the active view."""
    count = 0
    active_view = doc.ActiveView

    # Detail lines only work in 2D views
    if active_view.ViewType not in [DB.ViewType.FloorPlan,
                                     DB.ViewType.CeilingPlan,
                                     DB.ViewType.Section,
                                     DB.ViewType.Elevation,
                                     DB.ViewType.Detail]:
        logger.warning("Active view is not a 2D view. Switching to Model Lines.")
        return _create_model_lines_from_pts(doc, pts, pts[0].Z)

    for i in range(len(pts) - 1):
        start = pts[i]
        end = pts[i + 1]
        if start.DistanceTo(end) < 0.001:
            continue
        line = DB.Line.CreateBound(start, end)
        doc.Create.NewDetailCurve(active_view, line)
        count += 1
    return count


# ╔╦╗╦╔═╗╦  ╔═╗╔═╗
#  ║║║╠═╣║  ║ ║║ ╦
# ═╩╝╩╩ ╩╩═╝╚═╝╚═╝ WPF DIALOG
# ==================================================

class ParcelItem(object):
    """Data object for ListView binding."""
    def __init__(self, data):
        self.id              = data["id"]
        self.parcel_id       = data["parcel_id"]
        self.display_address = data["display_address"]
        self.area_sqft       = data["area_sqft"]
        self.area_sqft_raw   = data["area_sqft_raw"]
        self.geometry        = data["geometry"]
        self.county          = data["county"]
        self.state           = data["state"]


class PropertyLineDialog(forms.WPFWindow):
    """Main WPF dialog for Property Line Tool."""

    def __init__(self):
        # Build absolute path so forms.WPFWindow finds the XAML regardless of
        # which script calls this class (avoids the IronPython absolute-URI bug
        # that occurs with Application.LoadComponent + file:// URIs)
        xaml_path = os.path.join(os.path.dirname(__file__), "PropertyLine.xaml")
        forms.WPFWindow.__init__(self, xaml_path)

        self._selected_parcel = None
        self._parcels = []
        self._zoning_data = None

        # Load saved API key
        config = load_config()
        saved_key = config.get("lightbox_api_key", "")
        if saved_key:
            self.txt_api_key.Text = saved_key
            self._update_api_status(True, "API key loaded from config")

    # ───────────────────────────────────── GUI EVENTS

    def header_drag(self, sender, e):
        from System.Windows.Input import MouseButtonState
        from System.Windows.Window import DragMove
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def btn_close_Click(self, sender, e):
        self.Close()

    def btn_save_key_Click(self, sender, e):
        api_key = self.txt_api_key.Text.strip()
        if not api_key:
            self._update_api_status(False, "API key cannot be empty")
            return
        if save_config({"lightbox_api_key": api_key}):
            self._update_api_status(True, "API key saved successfully")
        else:
            self._update_api_status(False, "Failed to save API key")

    def txt_address_KeyDown(self, sender, e):
        if e.Key == Key.Return:
            self.btn_search_Click(sender, e)

    def btn_search_Click(self, sender, e):
        address = self.txt_address.Text.strip()
        if not address:
            self._set_status("Please enter an address to search.", error=True)
            return

        api_key = self.txt_api_key.Text.strip()
        if not api_key:
            self._set_status("Please enter your Lightbox API key first.", error=True)
            return

        self._set_status("Searching for parcels...", busy=True)
        self.btn_search.IsEnabled = False

        # Run in background thread to avoid blocking UI
        def search_thread():
            try:
                parcels = search_parcels(api_key, address)
                self.Dispatcher.Invoke(
                    DispatcherPriority.Normal,
                    Action(lambda: self._on_search_complete(parcels))
                )
            except Exception as ex:
                error_msg = str(ex)
                self.Dispatcher.Invoke(
                    DispatcherPriority.Normal,
                    Action(lambda: self._on_search_error(error_msg))
                )

        t = threading.Thread(target=search_thread)
        t.daemon = True
        t.start()

    def _on_search_complete(self, parcels):
        self.btn_search.IsEnabled = True
        self._parcels = parcels

        if not parcels:
            self._set_status("No parcels found for this address. Try a more specific address.")
            self.lv_parcels.Visibility = Visibility.Collapsed
            self.border_no_results.Visibility = Visibility.Visible
            return

        # Populate ListView
        self.lv_parcels.Items.Clear()
        for p in parcels:
            self.lv_parcels.Items.Add(ParcelItem(p))

        self.lv_parcels.Visibility = Visibility.Visible
        self.border_no_results.Visibility = Visibility.Collapsed

        self._set_status("Found {} parcel(s). Select one to continue.".format(len(parcels)))

    def _on_search_error(self, error_msg):
        self.btn_search.IsEnabled = True
        self._set_status("Search error: {}".format(error_msg), error=True)
        logger.error("Lightbox API search error: {}".format(error_msg))

    def lv_parcels_SelectionChanged(self, sender, e):
        item = self.lv_parcels.SelectedItem
        if not item:
            self._selected_parcel = None
            self.btn_create.IsEnabled = False
            self.grp_parcel_details.Visibility = Visibility.Collapsed
            self.grp_setback.Visibility = Visibility.Collapsed
            return

        self._selected_parcel = item
        self._zoning_data = None
        self._show_parcel_details(item)
        self.btn_create.IsEnabled = True

        # Reset setback section for the newly selected parcel
        self.txt_setback_front.Text = ""
        self.txt_setback_rear.Text  = ""
        self.txt_setback_side.Text  = ""
        self.txt_zoning_status.Text = "Click 'Fetch Zoning' to auto-fill setback values from the Lightbox API."
        self.txt_zoning_status.Foreground = SolidColorBrush(Color.FromRgb(136, 136, 136))
        self.grp_setback.Visibility = Visibility.Visible

    def _show_parcel_details(self, item):
        self.grp_parcel_details.Visibility = Visibility.Visible

        self.txt_detail_id.Text = item.parcel_id or "N/A"
        self.txt_detail_address.Text = item.display_address or "N/A"
        self.txt_detail_county.Text = item.county or "N/A"
        self.txt_detail_state.Text = item.state or "N/A"

        # Area
        raw = item.area_sqft_raw
        if raw and raw > 0:
            self.txt_detail_area.Text = format_area(raw)
        else:
            self.txt_detail_area.Text = "N/A"

        # Vertex count
        coords = get_polygon_coords(item.geometry)
        self.txt_detail_vertices.Text = "{} vertices".format(len(coords))

    # ───────────────────────────────────── ZONING / SETBACK

    def btn_fetch_zoning_Click(self, sender, e):
        if not self._selected_parcel:
            return

        api_key = self.txt_api_key.Text.strip()
        if not api_key:
            self.txt_zoning_status.Text = "No API key configured."
            self.txt_zoning_status.Foreground = SolidColorBrush(Color.FromRgb(255, 107, 107))
            return

        self.txt_zoning_status.Text = "Fetching zoning data..."
        self.txt_zoning_status.Foreground = SolidColorBrush(Color.FromRgb(255, 197, 61))
        self.btn_fetch_zoning.IsEnabled = False

        parcel_id = self._selected_parcel.id

        def zoning_thread():
            try:
                zoning = get_zoning_data(api_key, parcel_id)
                self.Dispatcher.Invoke(
                    DispatcherPriority.Normal,
                    Action(lambda: self._on_zoning_complete(zoning))
                )
            except Exception as ex:
                err = str(ex)
                self.Dispatcher.Invoke(
                    DispatcherPriority.Normal,
                    Action(lambda: self._on_zoning_error(err))
                )

        t = threading.Thread(target=zoning_thread)
        t.daemon = True
        t.start()

    def _on_zoning_complete(self, zoning):
        self.btn_fetch_zoning.IsEnabled = True
        self._zoning_data = zoning

        parts = []
        if zoning.get("front_ft") is not None:
            self.txt_setback_front.Text = "{:.1f}".format(zoning["front_ft"])
            parts.append("F:{:.0f}ft".format(zoning["front_ft"]))
        if zoning.get("rear_ft") is not None:
            self.txt_setback_rear.Text = "{:.1f}".format(zoning["rear_ft"])
            parts.append("R:{:.0f}ft".format(zoning["rear_ft"]))
        if zoning.get("side_ft") is not None:
            self.txt_setback_side.Text = "{:.1f}".format(zoning["side_ft"])
            parts.append("S:{:.0f}ft".format(zoning["side_ft"]))

        if parts:
            self.txt_zoning_status.Text = "Zoning loaded: {}".format(" | ".join(parts))
            self.txt_zoning_status.Foreground = SolidColorBrush(Color.FromRgb(78, 201, 176))
        else:
            self.txt_zoning_status.Text = "No setback data found in zoning response. Enter values manually."
            self.txt_zoning_status.Foreground = SolidColorBrush(Color.FromRgb(255, 197, 61))

    def _on_zoning_error(self, error_msg):
        self.btn_fetch_zoning.IsEnabled = True
        self.txt_zoning_status.Text = "Zoning fetch failed: {}".format(error_msg)
        self.txt_zoning_status.Foreground = SolidColorBrush(Color.FromRgb(255, 107, 107))
        logger.warning("Zoning API error: {}".format(error_msg))

    def _get_min_setback(self):
        """Return the smallest positive value among the three setback fields, or None."""
        values = []
        for txt in (self.txt_setback_front, self.txt_setback_rear, self.txt_setback_side):
            raw = txt.Text.strip()
            if not raw:
                continue
            try:
                v = float(raw)
                if v > 0:
                    values.append(v)
            except ValueError:
                pass
        return min(values) if values else None

    def btn_create_Click(self, sender, e):
        if not self._selected_parcel:
            self._set_status("No parcel selected.", error=True)
            return

        doc = revit.doc
        if not doc:
            self._set_status("No active Revit document.", error=True)
            return

        # Get options
        try:
            elevation_ft = float(self.txt_elevation.Text.strip() or "0")
        except ValueError:
            elevation_ft = 0.0

        # ComboBox selected item text
        line_cat_item = self.cmb_line_type.SelectedItem
        line_cat = line_cat_item.Content if line_cat_item else "Property Lines"

        origin_item = self.cmb_origin.SelectedItem
        origin_mode = origin_item.Content if origin_item else "Project Base Point"

        # Get coordinates
        coords = get_polygon_coords(self._selected_parcel.geometry)
        if len(coords) < 3:
            self._set_status("Invalid geometry: not enough coordinates.", error=True)
            return

        # Check if setback envelope was requested
        draw_setback = bool(self.chk_draw_setback.IsChecked)
        setback_ft = self._get_min_setback() if draw_setback else None
        if draw_setback and not setback_ft:
            self._set_status(
                "Draw Setback is checked but no setback values are set. "
                "Enter values manually or click Fetch Zoning.", error=True
            )
            return

        self._set_status("Creating property lines in Revit...", busy=True)
        self.btn_create.IsEnabled = False

        try:
            count = create_property_lines_in_revit(
                doc, coords, elevation_ft, line_cat, origin_mode
            )
            logger.info("Property lines created: {} segments".format(count))

            setback_count = 0
            if draw_setback and setback_ft:
                try:
                    setback_count = create_setback_lines_in_revit(
                        doc, coords, setback_ft, elevation_ft, origin_mode
                    )
                    logger.info("Setback envelope created: {} segments ({} ft)".format(
                        setback_count, setback_ft))
                except Exception as sb_ex:
                    logger.warning("Setback creation failed: {}".format(sb_ex))
                    self._set_status(
                        "Property lines created ({} segs) but setback failed: {}".format(
                            count, sb_ex), error=True
                    )
                    return

            if setback_count:
                msg = ("Done! {} property line seg(s) + {} setback seg(s) "
                       "({:.1f} ft inset) for: {}".format(
                           count, setback_count, setback_ft,
                           self._selected_parcel.display_address))
            else:
                msg = "Done! Created {} property line segment(s) for: {}".format(
                    count, self._selected_parcel.display_address)

            self._set_status(msg, success=True)

        except Exception as ex:
            self._set_status("Error creating lines: {}".format(ex), error=True)
            logger.error("Property line creation failed: {}".format(traceback.format_exc()))
        finally:
            self.btn_create.IsEnabled = True

    # ───────────────────────────────────── HELPERS

    def _update_api_status(self, ok, msg):
        self.txt_api_status.Text = msg
        if ok:
            self.txt_api_status.Foreground = SolidColorBrush(Color.FromRgb(78, 201, 176))  # teal
        else:
            self.txt_api_status.Foreground = SolidColorBrush(Color.FromRgb(255, 107, 107))  # red

    def _set_status(self, msg, error=False, success=False, busy=False):
        self.txt_status.Text = msg
        if error:
            color = Color.FromRgb(255, 107, 107)
            dot_color = Color.FromRgb(255, 107, 107)
            label = "Error"
        elif success:
            color = Color.FromRgb(78, 201, 176)
            dot_color = Color.FromRgb(78, 201, 176)
            label = "Done"
        elif busy:
            color = Color.FromRgb(255, 197, 61)
            dot_color = Color.FromRgb(255, 197, 61)
            label = "Working..."
        else:
            color = Color.FromRgb(136, 136, 136)
            dot_color = Color.FromRgb(136, 136, 136)
            label = "Idle"

        self.txt_status.Foreground = SolidColorBrush(color)
        self.dot_status.Fill = SolidColorBrush(dot_color)
        self.txt_status_label.Text = label


# ╔═╗╦ ╦╔═╗╦  ╦╔═╗
# ╚═╗╠═╣║ ║║  ║╚═╗
# ╚═╝╩ ╩╚═╝╩═╝╩╚═╝ PUBLIC ENTRY POINT
# ==================================================

def show_property_line_dialog():
    """Show the Property Line dialog and return when closed."""
    try:
        dlg = PropertyLineDialog()
        dlg.ShowDialog()
    except Exception as ex:
        logger.error("Failed to open Property Line dialog: {}".format(ex))
        logger.error(traceback.format_exc())
        forms.alert(
            "Property Line Tool error:\n{}".format(ex),
            title="Property Line Tool"
        )
