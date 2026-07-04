# CATEGORY OVERLAY — Lighting Fixture

> Appended AFTER the general core prompt above. All core rules still apply. This section
> adds Lighting-specific guidance. Set `"family_category": "Lighting Fixture"`. Output ONLY
> the JSON object.

## Part inventory (typical — adapt to the real object; model MORE parts for a finished look)
- Canopy / mounting plate / base (the part that fixes to ceiling, wall or floor).
- **Backplate boss / mount knuckle** — the small round fitting where the arm leaves the plate (add it; a bare tube stabbing a flat plate looks crude).
- Arm / stem / neck / chain / cord (connects mount to body — usually a graceful **curved** swing arm).
- **Neck / socket cup** — the fitting where the arm meets the shade (a short slightly-wider cylinder; hides the raw tube end).
- Body / housing / socket.
- Shade / diffuser / globe (drum, cone, dome, sphere).
- **Finial / bottom knob**, top diffuser ring, lamp/bulb, decorative rings.
- A typical wall sconce is **6–9 parts**, not 3–4 — the extra fittings are what make it read as a real lamp instead of sticks + a can.

## Placement & anchor
- **Ceiling-mounted** (pendant, chandelier, flush): ceiling is the XY plane at **Z = 0**;
  the fixture hangs **downward** (negative Z). Anchor = canopy at Z=0.
- **Wall-mounted** (sconce, swing-arm): wall face at **Y = 0**; mounting plate on the wall
  (y 0 → -t), body projecting to negative Y. Anchor = wall plate.
- **Table/floor lamp**: base at **Z = 0**, growing up. Anchor = base.
- Chain the whole assembly: Arm → Plate/Canopy, Body → Arm, Shade → Body, Finial → top.

## Typical dimensions (mm)
- Drum/empire shade dia 200–450, height 180–300. Wall sconce projection 200–350; plate
  120×200. Pendant drop 300–1500 (rod/cord). Arm/stem tube radius 6–12.

## Arm shape — reliable AND rounded (READ — build it this way)
Build the swing arm from **axis-aligned `Cylinder` segments** (a vertical rise → a horizontal reach → a short drop), and put a small **sphere knuckle** at each bend so the corners are rounded, not raw right angles. This is the combination that **always builds** in Revit.
- **Every arm `Cylinder` must be axis-aligned** — its `start` and `end` differ in **exactly one** of X/Y/Z (pure vertical, or pure horizontal). An axis-aligned `Cylinder` builds as a rock-solid Extrusion. A *diagonal* `Cylinder` (or a `Sweep`) relies on Revit's sweep engine, which **fails in this Revit build** — avoid both for the arm.
- **Sphere knuckle** at each bend = a `Revolution` of a semicircle centred on the shared corner point, radius ≈ 1.3× the tube radius. It hides the hard corner and reads as a brass ball-joint.
- Chain: each segment's `start` = the previous segment's `end`; drop a sphere on every shared corner.
- **Fittings**: keep a **backplate boss** (short `Cylinder`) where the arm leaves the plate and a **neck/socket cup** (short wider `Cylinder`) where the arm meets the shade.
- ⚠️ Do **not** use `Sweep` or a diagonal `Cylinder` here until told the sweep engine works — they get skipped and the arm disappears.

## Category pitfalls
- **Never leave the arm floating or built along the wrong axis** — chain plate → boss → arm → socket → shade → finial, each touching the previous.
- Shades: drum = Extrusion of a `Circle`; tapered/empire = **Blend** (both loops on the sketch
  plane, heights via offsets); dome/bowl = **Revolution** of an `Arc3P`; globe = Revolution of
  a semicircle. Never taper a Blend to a near-zero top to fake a cone — use Revolution.
- **Rim / edge trim (do this — a bare extruded drum looks like a sharp-edged can).** Add a thin
  **rim ring at the top edge AND the bottom edge** of the shade so the outline reads as a framed
  shade, not a solid cylinder. Model each ring as a **torus** = a `Revolution` of a small `Circle`
  offset from the shade's vertical axis: profile `Circle` centred at `[0, shade_y + shade_radius, rim_z]`
  with a small radius (≈ shade_radius/25, e.g. 6 mm), revolved a full turn about the vertical axis
  through the shade centre. Put one at the bottom `rim_z` and one at the top. (Optional: also make
  the drum a thin-walled tube via `inner_loops` for a fabric-shade look, but keep the finial/socket
  attached to a solid part, not the thin wall.)
- Keep the whole fixture centered and derive every height from one shared table.


## Worked example — detailed wall swing-arm lamp (reliable jointed arm: Cylinders + sphere knuckles + rimmed drum)

```json
{
  "family_name": "T3Lab_WallSwingLamp",
  "family_category": "Lighting Fixture",
  "_plan": {
    "parts": [
      "Plate",
      "Boss",
      "Riser",
      "Elbow1",
      "Reach",
      "Elbow2",
      "Socket",
      "Shade",
      "RimBottom",
      "RimTop",
      "Finial"
    ],
    "heights": {
      "Plate": "z 200-340 wall",
      "Boss": "knuckle z 270",
      "Riser": "z 270-482 @ y -30",
      "Elbow1": "ball joint z 480",
      "Reach": "y -30..-250 @ z 480",
      "Elbow2": "ball joint at reach end",
      "Socket": "z 480-510 @ y -250",
      "Shade": "drum z 330-510",
      "RimBottom": "rim z 330",
      "RimTop": "rim z 510",
      "Finial": "knob"
    },
    "connections": {
      "Boss": "on plate front y=-15",
      "Riser": "start = boss end",
      "Elbow1": "on riser top / reach corner",
      "Reach": "start = riser top",
      "Elbow2": "on reach end",
      "Socket": "start = reach end / elbow2",
      "Shade": "top z510 = socket bottom",
      "RimBottom": "shade bottom edge",
      "RimTop": "shade top edge",
      "Finial": "shade base"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Plate",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -60.0,
            0.0,
            200.0
          ],
          "end": [
            60.0,
            0.0,
            200.0
          ]
        },
        {
          "type": "Line",
          "start": [
            60.0,
            0.0,
            200.0
          ],
          "end": [
            60.0,
            0.0,
            340.0
          ]
        },
        {
          "type": "Line",
          "start": [
            60.0,
            0.0,
            340.0
          ],
          "end": [
            -60.0,
            0.0,
            340.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -60.0,
            0.0,
            340.0
          ],
          "end": [
            -60.0,
            0.0,
            200.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -15.0
    },
    {
      "type": "Cylinder",
      "id": "Boss",
      "attaches_to": "Plate",
      "is_solid": true,
      "start": [
        0.0,
        -15.0,
        270.0
      ],
      "end": [
        0.0,
        -30.0,
        270.0
      ],
      "radius": 18.0
    },
    {
      "type": "Cylinder",
      "id": "Riser",
      "attaches_to": "Boss",
      "is_solid": true,
      "start": [
        0.0,
        -30.0,
        270.0
      ],
      "end": [
        0.0,
        -30.0,
        482.0
      ],
      "radius": 8.0
    },
    {
      "type": "Revolution",
      "id": "Elbow1",
      "attaches_to": "Riser",
      "is_solid": true,
      "sketch_plane_x": 0.0,
      "profile": [
        {
          "type": "Arc3P",
          "start": [
            0.0,
            -30.0,
            491.0
          ],
          "end": [
            0.0,
            -30.0,
            469.0
          ],
          "mid": [
            0.0,
            -19.0,
            480.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            -30.0,
            469.0
          ],
          "end": [
            0.0,
            -30.0,
            491.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        -30.0,
        469.0
      ],
      "axis_end": [
        0.0,
        -30.0,
        491.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    },
    {
      "type": "Cylinder",
      "id": "Reach",
      "attaches_to": "Elbow1",
      "is_solid": true,
      "start": [
        0.0,
        -30.0,
        480.0
      ],
      "end": [
        0.0,
        -250.0,
        480.0
      ],
      "radius": 8.0
    },
    {
      "type": "Revolution",
      "id": "Elbow2",
      "attaches_to": "Reach",
      "is_solid": true,
      "sketch_plane_x": 0.0,
      "profile": [
        {
          "type": "Arc3P",
          "start": [
            0.0,
            -250.0,
            491.0
          ],
          "end": [
            0.0,
            -250.0,
            469.0
          ],
          "mid": [
            0.0,
            -239.0,
            480.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            -250.0,
            469.0
          ],
          "end": [
            0.0,
            -250.0,
            491.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        -250.0,
        469.0
      ],
      "axis_end": [
        0.0,
        -250.0,
        491.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    },
    {
      "type": "Cylinder",
      "id": "Socket",
      "attaches_to": "Elbow2",
      "is_solid": true,
      "start": [
        0.0,
        -250.0,
        480.0
      ],
      "end": [
        0.0,
        -250.0,
        510.0
      ],
      "radius": 14.0
    },
    {
      "type": "Extrusion",
      "id": "Shade",
      "attaches_to": "Socket",
      "is_solid": true,
      "sketch_plane_z": 330.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            -250.0,
            330.0
          ],
          "radius": 150.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 180.0
    },
    {
      "type": "Revolution",
      "id": "RimBottom",
      "attaches_to": "Shade",
      "is_solid": true,
      "sketch_plane_x": 0.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            -100.0,
            330.0
          ],
          "radius": 6.0
        }
      ],
      "axis_start": [
        0.0,
        -250.0,
        324.0
      ],
      "axis_end": [
        0.0,
        -250.0,
        360.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    },
    {
      "type": "Revolution",
      "id": "RimTop",
      "attaches_to": "Shade",
      "is_solid": true,
      "sketch_plane_x": 0.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            -100.0,
            510.0
          ],
          "radius": 6.0
        }
      ],
      "axis_start": [
        0.0,
        -250.0,
        504.0
      ],
      "axis_end": [
        0.0,
        -250.0,
        540.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    },
    {
      "type": "Revolution",
      "id": "Finial",
      "attaches_to": "Shade",
      "is_solid": true,
      "sketch_plane_x": 0.0,
      "profile": [
        {
          "type": "Arc3P",
          "start": [
            0.0,
            -250.0,
            345.0
          ],
          "end": [
            0.0,
            -250.0,
            315.0
          ],
          "mid": [
            0.0,
            -235.0,
            330.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            -250.0,
            315.0
          ],
          "end": [
            0.0,
            -250.0,
            345.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        -250.0,
        315.0
      ],
      "axis_end": [
        0.0,
        -250.0,
        345.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    }
  ]
}
```
