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


class PropertyLineDialog(Window):
    """Main WPF dialog for Property Line Tool."""

    def __init__(self):
        # Load XAML
        xaml_path = os.path.join(os.path.dirname(__file__), "PropertyLine.xaml")
        Application.LoadComponent(self, Uri(xaml_path))

        self._selected_parcel = None
        self._parcels = []

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
        t.IsBackground = True
        t.Start()

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
            return

        self._selected_parcel = item
        self._show_parcel_details(item)
        self.btn_create.IsEnabled = True

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

        self._set_status("Creating property lines in Revit...", busy=True)
        self.btn_create.IsEnabled = False

        try:
            count = create_property_lines_in_revit(
                doc, coords, elevation_ft, line_cat, origin_mode
            )
            self._set_status(
                "Done! Created {} property line segment(s) for: {}".format(
                    count, self._selected_parcel.display_address
                ),
                success=True
            )
            logger.info("Property lines created: {} segments".format(count))
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
