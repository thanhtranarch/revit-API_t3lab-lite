# CATEGORY OVERLAY — Casework

> Appended AFTER the general core prompt above. All core rules (schema, forms, curve
> segments, `_plan`/`attaches_to` connectivity, failure modes, checklist) still apply.
> This section adds Casework-specific guidance. Set `"family_category": "Casework"`.
> Output ONLY the JSON object. Core **Example 8** is a base cabinet — reuse its *method*
> (not its literal coordinates).

## Part inventory (typical — adapt to the real object)
- Carcass box (sides + top + bottom + back — usually one Extrusion for a simple unit).
- Toe-kick / plinth (recessed base).
- Door fronts and/or drawer fronts (one slab each, on the carcass front face).
- Handles / pulls / knobs (on each front).
- Countertop (base cabinets), shelves (visible), light rail / valance (wall cabinets).

## Placement & anchor
- **Base cabinets**: carcass bottom at **Z = 0** (above the toe-kick), floor-standing.
  Anchor = carcass.
- **Wall (upper) cabinets**: mount against the wall — back face at **Y = 0**, body
  projecting to negative Y; anchor = carcass back / wall face.
- Doors/drawer fronts sit ON the carcass front face (shared plane), projecting forward
  1–2 mm past it; pulls sit ON the front face of their door. Chain with `attaches_to`:
  Door → Carcass, Pull → Door, Toe-kick → Carcass.

## Typical dimensions (mm)
- Base cabinet: width 300–900, height 720 (carcass) + 100 toe-kick, depth 560–600,
  countertop +40. Wall cabinet: height 300–900, depth 300–350.
- Door/drawer front thickness 18–20; reveal gap ~2–3 around fronts. Pull projection 25–40.

## Category pitfalls
- Model the **fronts as separate slabs** on the carcass face, not merged into the box —
  otherwise doors/drawers are invisible.
- Fronts and pulls must overlap the face they sit on (verify all 3 axes) — floating pulls
  are the classic casework failure.
- Openings/glazed doors → punch with `inner_loops` on the front slab, don't model a void box.


## Worked example (Wall cabinet (wall-mounted anchor, carcass + door + pull)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_WallCabinet",
  "family_category": "Casework",
  "_plan": {
    "parts": [
      "Carcass",
      "Door",
      "Pull"
    ],
    "heights": {
      "Carcass": "z 0-720",
      "Door": "z 20-700",
      "Pull": "z 344-376"
    },
    "connections": {
      "Door": "back face y=-320 on carcass front face y=-320",
      "Pull": "base y=-338 on door front face y=-338"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Carcass",
      "is_solid": true,
      "sketch_plane_z": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -300.0,
            -320.0,
            0.0
          ],
          "end": [
            300.0,
            -320.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            300.0,
            -320.0,
            0.0
          ],
          "end": [
            300.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            300.0,
            0.0,
            0.0
          ],
          "end": [
            -300.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -300.0,
            0.0,
            0.0
          ],
          "end": [
            -300.0,
            -320.0,
            0.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 720.0
    },
    {
      "type": "Extrusion",
      "id": "Door",
      "attaches_to": "Carcass",
      "is_solid": true,
      "sketch_plane_y": -320.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -290.0,
            -320.0,
            20.0
          ],
          "end": [
            290.0,
            -320.0,
            20.0
          ]
        },
        {
          "type": "Line",
          "start": [
            290.0,
            -320.0,
            20.0
          ],
          "end": [
            290.0,
            -320.0,
            700.0
          ]
        },
        {
          "type": "Line",
          "start": [
            290.0,
            -320.0,
            700.0
          ],
          "end": [
            -290.0,
            -320.0,
            700.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -290.0,
            -320.0,
            700.0
          ],
          "end": [
            -290.0,
            -320.0,
            20.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -18.0
    },
    {
      "type": "Extrusion",
      "id": "Pull",
      "attaches_to": "Door",
      "is_solid": true,
      "sketch_plane_y": -338.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            250.0,
            -338.0,
            360.0
          ],
          "radius": 16.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -22.0
    }
  ]
}
```
