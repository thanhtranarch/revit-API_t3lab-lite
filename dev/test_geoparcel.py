# -*- coding: utf-8 -*-
"""
CPython 3 test harness for lib/Snippets/_geoparcel.py — the worldwide address
to property-boundary lookup behind the Property Line tool.

Run:  python3 dev/test_geoparcel.py
Exit code 0 = all pass. No external test framework — plain asserts, mirroring
dev/test_assistant_routing.py conventions.

Every network call is stubbed, so the suite is deterministic and offline. What
it locks down:

    geometry    area / centroid / containment maths on [lon, lat] rings
    encoding    Vietnamese addresses survive URL encoding as UTF-8 (the
                IronPython trap that silently mangles diacritics)
    ranking     a district-sized landuse polygon can never outrank a
                plot-sized boundary, even when it contains the point
    dedup       the same footprint returned by both Nominatim and Overpass,
                with different vertex order, is listed once
    fallback    an address with nothing mapped still yields geometry, flagged
                approximate
"""
from __future__ import unicode_literals

import json
import math
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, "T3Lab.extension", "lib"))

from Snippets import _geoparcel as gp   # noqa: E402

FAILURES = []


def check(name, condition, detail=""):
    if condition:
        print("  ok   {}".format(name))
    else:
        print("  FAIL {} {}".format(name, detail))
        FAILURES.append(name)


# ── fixtures ─────────────────────────────────────────────────────────────────

def square_ring(lat, lon, size_m):
    """A north-aligned square of *size_m*, as a [lon, lat] ring."""
    return gp._square_ring(lat, lon, size_m)


def overpass_way(way_id, ring, tags):
    """Overpass 'out geom' shape: closed way, first node repeated last."""
    geometry = [{"lon": c[0], "lat": c[1]} for c in ring]
    geometry.append(dict(geometry[0]))
    return {"type": "way", "id": way_id, "tags": tags, "geometry": geometry}


class StubHTTP(object):
    """Replaces gp.http_request; serves canned bodies and records the calls."""

    def __init__(self, nominatim=None, overpass=None, overpass_status=200):
        self.nominatim = nominatim if nominatim is not None else []
        self.overpass = overpass if overpass is not None else []
        self.overpass_status = overpass_status
        self.calls = []

    def __call__(self, url, data=None, headers=None, timeout=25):
        self.calls.append((url, data))
        if "nominatim" in url:
            return 200, json.dumps(self.nominatim)
        if "interpreter" in url:
            if self.overpass_status != 200:
                return self.overpass_status, "<html>nope</html>"
            return 200, json.dumps({"elements": self.overpass})
        return 404, ""


def with_stub(stub, func):
    original = gp.http_request
    gp.http_request = stub
    try:
        return func()
    finally:
        gp.http_request = original


# ── geometry ─────────────────────────────────────────────────────────────────

def test_geometry():
    print("\n[geometry]")
    ring = square_ring(10.0, 106.0, 100.0)          # 100 m square
    area = gp.polygon_area_m2(ring)
    check("100m square measures ~10,000 m2", abs(area - 10000.0) < 60.0,
          "got {:.1f}".format(area))

    lat, lon = gp.polygon_centroid(ring)
    check("centroid returns to the centre",
          abs(lat - 10.0) < 1e-9 and abs(lon - 106.0) < 1e-9,
          "got {}, {}".format(lat, lon))

    check("centre point is inside", gp.point_in_polygon(106.0, 10.0, ring))
    check("far point is outside", not gp.point_in_polygon(106.5, 10.5, ring))

    check("sqft conversion is consistent",
          abs(gp.polygon_area_sqft(ring) - area * gp.SQM_TO_SQFT) < 1e-6)

    closed = list(ring) + [list(ring[0])]
    check("close_ring drops the repeated vertex",
          len(gp.close_ring(closed)) == len(ring))

    check("ring_key ignores vertex order",
          gp.ring_key(ring) == gp.ring_key(ring[2:] + ring[:2]),
          "keys differ")

    check("ring_key separates distinct plots",
          gp.ring_key(ring) != gp.ring_key(square_ring(10.01, 106.01, 100.0)))

    check("area formats metric first",
          gp.format_area_dual(10000.0).startswith("929 m"),
          gp.format_area_dual(10000.0))
    check("zero area reads N/A", gp.format_area_dual(0) == "N/A")


# ── encoding ─────────────────────────────────────────────────────────────────

def test_encoding():
    print("\n[encoding]")
    quoted = gp.url_quote("268 Lý Thường Kiệt, Quận 10")
    check("Vietnamese encodes as UTF-8 percent escapes",
          "%C3%BD" in quoted and "%20" in quoted, quoted)
    check("no raw non-ascii survives quoting",
          all(ord(ch) < 128 for ch in quoted))

    stub = StubHTTP(nominatim=[])
    with_stub(stub, lambda: gp.geocode("Đà Nẵng, Việt Nam"))
    check("geocode URL carries the encoded address",
          stub.calls and "%C4%90" in stub.calls[0][0],
          stub.calls[0][0] if stub.calls else "no call")


# ── classification and ranking ───────────────────────────────────────────────

def test_demojibake():
    print("\n[demojibake]")
    # The exact corruption seen in Revit: UTF-8 read back as Latin-1.
    broken = "Thành phố Hồ Chí Minh".encode("utf-8").decode("latin-1")
    check("Vietnamese is repaired",
          gp.demojibake(broken) == "Thành phố Hồ Chí Minh", broken)

    for label, good in (
            ("Thai", "กรุงเทพมหานคร"),
            ("Arabic", "شارع الملك فهد"),
            ("Greek", "Λεωφόρος Κηφισίας"),
            ("Cyrillic", "Тверская улица"),
            ("Japanese", "東京都千代田区"),
            ("Korean", "서울특별시 강남구"),
            ("Hindi", "नई दिल्ली")):
        mangled = good.encode("utf-8").decode("latin-1")
        check("{} is repaired".format(label),
              gp.demojibake(mangled) == good, mangled)

    # Text that is legitimately Latin-1 must be left exactly as it is.
    for good in ("Café de Flore", "Ångström", "Müller Straße",
                 "Plaça de Catalunya", "São Paulo", "Ísafjörður"):
        check("'{}' is left alone".format(good), gp.demojibake(good) == good,
              gp.demojibake(good))

    check("ASCII is left alone",
          gp.demojibake("10 Downing Street") == "10 Downing Street")
    check("already-correct unicode is left alone",
          gp.demojibake("Thành phố") == "Thành phố")
    check("empty input is safe", gp.demojibake("") == "")
    check("None is safe", gp.demojibake(None) is None)

    # _u is the funnel every API string passes through
    check("_u repairs on the way through", gp._u(broken) == "Thành phố Hồ Chí Minh")
    check("_u still handles None", gp._u(None) == "")


def test_metes_and_bounds():
    print("\n[metes and bounds]")
    # A 100 m square walked counter-clockwise from the south-west corner:
    # east, north, west, south.
    ring = square_ring(10.0, 106.0, 100.0)
    segs = gp.metes_and_bounds(ring)
    check("one segment per side and back to the start", len(segs) == 4,
          len(segs))

    lengths = [s["length_m"] for s in segs]
    check("every side measures ~100 m",
          all(abs(l - 100.0) < 0.6 for l in lengths),
          ["{:.2f}".format(l) for l in lengths])

    bearings = [s["bearing"] for s in segs]
    check("the four sides read due E, N, W, S",
          bearings == ['N 90°00\'00" E', 'N 0°00\'00" E',
                       'N 90°00\'00" W', 'S 0°00\'00" W'],
          bearings)
    check("due east and due west stay symmetric",
          gp.format_bearing(90.0) == 'N 90°00\'00" E' and
          gp.format_bearing(270.0) == 'N 90°00\'00" W',
          (gp.format_bearing(90.0), gp.format_bearing(270.0)))
    check("due north and due south stay symmetric",
          gp.format_bearing(0.0) == 'N 0°00\'00" E' and
          gp.format_bearing(180.0) == 'S 0°00\'00" W',
          (gp.format_bearing(0.0), gp.format_bearing(180.0)))

    # Bearings must round-trip: reading the quadrant text back must land on
    # the azimuth it came from.
    def bearing_to_azimuth(text):
        ns, rest = text[0], text[2:]
        ew = text[-1]
        deg, rest = rest.split("°")
        minutes, rest = rest.split("'")
        seconds = rest.split('"')[0]
        angle = int(deg) + int(minutes) / 60.0 + int(seconds) / 3600.0
        if ns == "N" and ew == "E":
            return angle
        if ns == "S" and ew == "E":
            return 180.0 - angle
        if ns == "S" and ew == "W":
            return 180.0 + angle
        return 360.0 - angle

    round_trip_ok = True
    for az in (0.0, 12.5125, 45.0, 89.9, 90.0, 120.0, 179.5, 180.0,
               200.0, 269.0, 270.0, 300.0, 359.5):
        back = bearing_to_azimuth(gp.format_bearing(az)) % 360.0
        if abs(back - az) > 0.001 and abs(back - az - 360.0) > 0.001:
            round_trip_ok = False
            print("       {} -> {} -> {}".format(az, gp.format_bearing(az), back))
    check("every bearing round-trips to its azimuth", round_trip_ok)

    check("azimuths run clockwise from north",
          [round(s["azimuth_deg"]) for s in segs] == [90, 0, 270, 180],
          [round(s["azimuth_deg"]) for s in segs])

    check("feet track metres",
          all(abs(s["length_ft"] - s["length_m"] * 3.280839895) < 1e-6
              for s in segs))

    # Quadrant conversion across all four quadrants
    check("45 deg is NE", gp.format_bearing(45.0) == 'N 45°00\'00" E')
    check("135 deg is SE", gp.format_bearing(135.0) == 'S 45°00\'00" E')
    check("225 deg is SW", gp.format_bearing(225.0) == 'S 45°00\'00" W')
    check("315 deg is NW", gp.format_bearing(315.0) == 'N 45°00\'00" W')
    check("wraps past 360", gp.format_bearing(405.0) == 'N 45°00\'00" E')
    check("minutes and seconds are carried",
          gp.format_bearing(12.5125) == 'N 12°30\'45" E',
          gp.format_bearing(12.5125))

    text = gp.format_metes_and_bounds(ring, area_sqft=107639.1)
    check("the table lists a row per segment",
          all(b in text for b in bearings), text[:200])
    check("the table totals the perimeter", "PERIMETER" in text)
    check("the table states the area", "ENCLOSED AREA" in text)
    check("the table warns that bearings are true north",
          "TRUE north" in text)

    check("a degenerate ring yields nothing",
          gp.metes_and_bounds([[0, 0], [0, 0]]) == [])
    check("and formats as a readable message",
          "No boundary segments" in gp.format_metes_and_bounds([[0, 0]]))


def test_classify():
    print("\n[classify]")
    label, rank = gp._classify({"boundary": "cadastral"})
    check("cadastral parcel ranks first", rank == 0 and "Cadastral" in label,
          label)

    label, rank = gp._classify({"building": "yes"})
    check("building footprint ranks last of the known kinds", rank == 5, label)
    check("building=yes drops the redundant detail",
          label == "Building footprint", label)

    label, _ = gp._classify({"building": "commercial"})
    check("building detail is kept when meaningful",
          label == "Building footprint (commercial)", label)

    label, rank = gp._classify({"place": "house"})
    check("unknown tags name themselves rather than reading 'Mapped area'",
          label == "Place (house)", label)

    check("landuse outranks a fence",
          gp._classify({"landuse": "residential"})[1] <
          gp._classify({"barrier": "fence"})[1])


def test_ranking_and_dedup():
    print("\n[ranking + dedup]")
    lat, lon = 10.0, 106.0
    plot = square_ring(lat, lon, 40.0)              # 1,600 m2, contains point
    district = square_ring(lat, lon, 600.0)         # 360,000 m2, also contains
    neighbour = square_ring(lat + 0.0003, lon, 40.0)   # ~33 m away

    stub = StubHTTP(overpass=[
        overpass_way(1, district, {"landuse": "residential"}),
        overpass_way(2, neighbour, {"building": "yes"}),
        overpass_way(3, plot, {"landuse": "residential"}),
    ])
    place = {"lat": lat, "lon": lon, "address": {"country": "Vietnam"},
             "display_name": "test", "osm_type": "node", "osm_id": "1",
             "geojson": {}, "boundingbox": [], "name": ""}

    results = with_stub(stub, lambda: gp.nearby_boundaries(place, limit=8))
    kinds = [r["boundary_kind"] for r in results]
    check("something came back", bool(results), kinds)
    check("the plot-sized boundary ranks first",
          results and abs(results[0]["area_m2_raw"] - 1600.0) < 50.0,
          kinds)
    check("the district polygon is labelled and demoted",
          any("district scale" in k for k in kinds) and
          "district scale" not in kinds[0], kinds)
    check("the district polygon ranks below the neighbour too",
          kinds.index([k for k in kinds if "district scale" in k][0]) ==
          len(kinds) - 1, kinds)
    check("Overpass way ids are reported as ways",
          all(r["parcel_id"].startswith("OSM way ") for r in results),
          [r["parcel_id"] for r in results])

    # Same footprint, different starting vertex: must not be listed twice.
    stub = StubHTTP(overpass=[overpass_way(3, plot[2:] + plot[:2],
                                           {"landuse": "residential"})])
    results = with_stub(stub, lambda: gp.nearby_boundaries(
        place, limit=8, exclude={gp.ring_key(plot)}))
    check("a shape already shown by the geocoder is not repeated",
          results == [], [r["boundary_kind"] for r in results])


def test_area_cap():
    print("\n[area cap]")
    lat, lon = 1.3, 103.8
    huge = square_ring(lat, lon, 3000.0)            # 9,000,000 m2
    tiny = square_ring(lat, lon, 1.5)               # ~2 m2
    stub = StubHTTP(overpass=[overpass_way(1, huge, {"landuse": "industrial"}),
                              overpass_way(2, tiny, {"building": "yes"})])
    place = {"lat": lat, "lon": lon, "address": {}, "display_name": "x",
             "osm_type": "node", "osm_id": "1", "geojson": {},
             "boundingbox": [], "name": ""}
    results = with_stub(stub, lambda: gp.nearby_boundaries(place, limit=8))
    check("a whole district and a doorstep are both rejected",
          results == [], [r["boundary_kind"] for r in results])


# ── end to end (stubbed network) ─────────────────────────────────────────────

NOMINATIM_HIT = [{
    "lat": "10.7769", "lon": "106.7009",
    "display_name": "268 Lý Thường Kiệt, Quận 10, Thành phố Hồ Chí Minh",
    "name": "OCB",
    "address": {"country": "Việt Nam", "state": "Thành phố Hồ Chí Minh",
                "suburb": "Phường 14"},
    "geojson": {"type": "Polygon",
                "coordinates": [gp._square_ring(10.7769, 106.7009, 30.0)]},
    "boundingbox": ["10.7768", "10.7770", "106.7008", "106.7010"],
    "osm_type": "way", "osm_id": "12345",
    "category": "building", "type": "commercial",
}]

NOMINATIM_POINT_ONLY = [{
    "lat": "21.0278", "lon": "105.8342",
    "display_name": "Somewhere unmapped, Việt Nam",
    "name": "",
    "address": {"country": "Việt Nam"},
    "geojson": {},
    "boundingbox": ["21.0277", "21.0279", "105.8341", "105.8343"],
    "osm_type": "node", "osm_id": "999",
    "category": "place", "type": "house",
}]


def test_primary_phase():
    print("\n[primary phase]")
    stub = StubHTTP(nominatim=NOMINATIM_HIT)
    places, results = with_stub(
        stub, lambda: gp.primary_boundaries("268 Lý Thường Kiệt", limit=8))
    check("the geocoder polygon is returned without touching Overpass",
          len(results) == 1 and
          all("interpreter" not in url for url, _ in stub.calls),
          [c[0] for c in stub.calls])
    check("it is classified from the geocoder category",
          results[0]["boundary_kind"] == "Building footprint (commercial)",
          results[0]["boundary_kind"])
    check("country survives as unicode",
          results[0]["country"] == "Việt Nam", results[0]["country"])
    check("the subtitle carries kind and area",
          "Building footprint" in results[0]["subtitle"] and
          "m²" in results[0]["subtitle"], results[0]["subtitle"])
    check("centroid is reported", abs(results[0]["lat"] - 10.7769) < 1e-6)
    check("nothing is flagged approximate", not results[0]["is_approximate"])

    stub = StubHTTP(nominatim=[])
    try:
        with_stub(stub, lambda: gp.primary_boundaries("nowhere at all"))
        check("an ungeocodable address raises", False, "no exception")
    except ValueError as ex:
        check("an ungeocodable address raises a helpful ValueError",
              "No location found" in str(ex), str(ex))


def test_fallback():
    print("\n[fallback]")
    stub = StubHTTP(nominatim=NOMINATIM_POINT_ONLY, overpass=[])
    places, results = with_stub(
        stub, lambda: gp.primary_boundaries("Somewhere unmapped"))
    check("a point-only match yields no polygon yet", results == [])

    more = with_stub(stub, lambda: gp.nearby_boundaries(
        places[0], limit=8, include_fallback=True))
    check("an approximate plot is offered instead of nothing", len(more) == 1,
          more)
    check("and it is flagged approximate", more and more[0]["is_approximate"])
    check("the approximate plot sits at the geocoded point",
          more and abs(more[0]["lat"] - 21.0278) < 1e-3)

    without = with_stub(stub, lambda: gp.nearby_boundaries(
        places[0], limit=8, include_fallback=False))
    check("the fallback stays opt-in", without == [])


def test_overpass_failure_is_survivable():
    print("\n[overpass failure]")
    stub = StubHTTP(nominatim=NOMINATIM_HIT, overpass_status=406)
    place = {"lat": 10.7769, "lon": 106.7009, "address": {},
             "display_name": "x", "osm_type": "way", "osm_id": "1",
             "geojson": {}, "boundingbox": [], "name": ""}
    results = with_stub(stub, lambda: gp.nearby_boundaries(place, limit=8))
    check("a dead mirror returns an empty list rather than raising",
          results == [])
    tried = len([url for url, _ in stub.calls if "interpreter" in url])
    check("failover stops after OVERPASS_MAX_TRIES mirrors",
          tried == gp.OVERPASS_MAX_TRIES, "tried {}".format(tried))


def test_overpass_query_shape():
    print("\n[overpass query]")
    query = gp._overpass_query(10.5, 106.5, 80)
    check("query is valid Overpass QL", query.endswith(");out geom;"), query)
    check("query asks for geometry", "out geom" in query)
    check("query filters to plot-like features",
          '["landuse"]' in query and '["building"]' in query)
    check("query radius is honoured", "around:80," in query, query)


def main():
    test_geometry()
    test_encoding()
    test_demojibake()
    test_metes_and_bounds()
    test_classify()
    test_ranking_and_dedup()
    test_area_cap()
    test_primary_phase()
    test_fallback()
    test_overpass_failure_is_survivable()
    test_overpass_query_shape()

    print("")
    if FAILURES:
        print("FAILED ({}): {}".format(len(FAILURES), ", ".join(FAILURES)))
        return 1
    print("all _geoparcel tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
