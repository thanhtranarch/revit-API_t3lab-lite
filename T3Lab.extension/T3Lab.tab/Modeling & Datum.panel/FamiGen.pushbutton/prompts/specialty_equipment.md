# CATEGORY OVERLAY — Specialty Equipment

> Appended AFTER the general core prompt above. Specialty Equipment is a broad catch-all for
> appliances, gym/medical/kitchen/lab equipment and similar. All core rules apply. Set
> `"family_category": "Specialty Equipment"`. Output ONLY the JSON object.

## Part inventory (typical — adapt to the real object)
- Main body / cabinet / chassis.
- Doors, drawers, lids, control panels, displays, knobs/handles.
- Base / feet / castors, spouts / nozzles / trays, racks.

## Placement & anchor
- Usually **floor-standing** (base/feet at Z = 0) or **counter/wall-mounted**. Anchor = the
  part contacting the host surface.
- Panels, doors, knobs and trays chain to the body via `attaches_to` with a verified 1–2 mm
  3-axis overlap.

## Typical dimensions (mm)
- Appliance bodies 500–900 wide; taller units up to 2000+. Control panels 200–500;
  knobs/handles 20–120.

## Category pitfalls
- Decompose honestly with the core Step-1 method — don't merge everything into one box.
- Round parts (drums, bowls, nozzles) → Revolution / Extrusion of a `Circle`, not boxes.
- Vents, grilles and slots → `inner_loops`. Keep every added part connected (no orphans).


## Worked example (Appliance (body + door + vertical bar handle)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_Appliance",
  "family_category": "Specialty Equipment",
  "_plan": {
    "parts": [
      "Body",
      "Door",
      "Handle"
    ],
    "heights": {
      "Body": "z 0-850",
      "Door": "z 20-830",
      "Handle": "z 400-700"
    },
    "connections": {
      "Door": "on body front face y=-300",
      "Handle": "on door front face y=-320"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Body",
      "is_solid": true,
      "sketch_plane_z": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -300.0,
            -300.0,
            0.0
          ],
          "end": [
            300.0,
            -300.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            300.0,
            -300.0,
            0.0
          ],
          "end": [
            300.0,
            300.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            300.0,
            300.0,
            0.0
          ],
          "end": [
            -300.0,
            300.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -300.0,
            300.0,
            0.0
          ],
          "end": [
            -300.0,
            -300.0,
            0.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 850.0
    },
    {
      "type": "Extrusion",
      "id": "Door",
      "attaches_to": "Body",
      "is_solid": true,
      "sketch_plane_y": -300.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -290.0,
            -300.0,
            20.0
          ],
          "end": [
            290.0,
            -300.0,
            20.0
          ]
        },
        {
          "type": "Line",
          "start": [
            290.0,
            -300.0,
            20.0
          ],
          "end": [
            290.0,
            -300.0,
            830.0
          ]
        },
        {
          "type": "Line",
          "start": [
            290.0,
            -300.0,
            830.0
          ],
          "end": [
            -290.0,
            -300.0,
            830.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -290.0,
            -300.0,
            830.0
          ],
          "end": [
            -290.0,
            -300.0,
            20.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -20.0
    },
    {
      "type": "Extrusion",
      "id": "Handle",
      "attaches_to": "Door",
      "is_solid": true,
      "sketch_plane_z": 400.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            -250.0,
            -330.0,
            400.0
          ],
          "radius": 12.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 300.0
    }
  ]
}
```
