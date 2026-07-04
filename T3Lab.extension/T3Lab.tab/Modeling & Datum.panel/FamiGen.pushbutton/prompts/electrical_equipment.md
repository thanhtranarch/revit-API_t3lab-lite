# CATEGORY OVERLAY — Electrical Equipment

> Appended AFTER the general core prompt above. All core rules apply. Set
> `"family_category": "Electrical Equipment"`. Output ONLY the JSON object.

## Part inventory (typical — adapt to the real object)
- Enclosure / box / cabinet (main body).
- Faceplate / cover / door.
- Switches, buttons, dials, meters, indicator lights (small parts on the faceplate).
- Mounting plate / bracket, conduit knockouts / stubs, terminal lugs.

## Placement & anchor
- **Wall-mounted** (panelboard, disconnect, outlet box): back/mounting plate at **Y = 0**,
  body projecting to negative Y. Anchor = mounting plate/box back.
- **Floor-mounted** (switchgear, floor panel): base at **Z = 0**. Anchor = base.
- Faceplate sits ON the enclosure front face; buttons/dials sit ON the faceplate. Chain with
  `attaches_to` and verify 3-axis overlap.

## Typical dimensions (mm)
- Panelboard 300–600 wide × 400–900 tall × 100–200 deep. Outlet/switch box 70–120.
- Buttons/dials 15–40; meters 40–100.

## Category pitfalls
- Mostly prismatic → Extrusions; round gauges/dials → Extrusion of a `Circle`.
- Vents/louvres and knockouts → `inner_loops`, not separate void boxes.
- Keep small controls overlapping the faceplate (no floating buttons).


## Worked example (Wall panel box (enclosure + faceplate + switch)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_PanelBox",
  "family_category": "Electrical Equipment",
  "_plan": {
    "parts": [
      "Enclosure",
      "Faceplate",
      "Switch"
    ],
    "heights": {
      "Enclosure": "z 0-600",
      "Faceplate": "z 20-580",
      "Switch": "z 290-330"
    },
    "connections": {
      "Faceplate": "on enclosure front face y=-150",
      "Switch": "on faceplate front y=-165"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Enclosure",
      "is_solid": true,
      "sketch_plane_z": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -200.0,
            -150.0,
            0.0
          ],
          "end": [
            200.0,
            -150.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            200.0,
            -150.0,
            0.0
          ],
          "end": [
            200.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            200.0,
            0.0,
            0.0
          ],
          "end": [
            -200.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -200.0,
            0.0,
            0.0
          ],
          "end": [
            -200.0,
            -150.0,
            0.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 600.0
    },
    {
      "type": "Extrusion",
      "id": "Faceplate",
      "attaches_to": "Enclosure",
      "is_solid": true,
      "sketch_plane_y": -150.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -195.0,
            -150.0,
            20.0
          ],
          "end": [
            195.0,
            -150.0,
            20.0
          ]
        },
        {
          "type": "Line",
          "start": [
            195.0,
            -150.0,
            20.0
          ],
          "end": [
            195.0,
            -150.0,
            580.0
          ]
        },
        {
          "type": "Line",
          "start": [
            195.0,
            -150.0,
            580.0
          ],
          "end": [
            -195.0,
            -150.0,
            580.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -195.0,
            -150.0,
            580.0
          ],
          "end": [
            -195.0,
            -150.0,
            20.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -15.0
    },
    {
      "type": "Extrusion",
      "id": "Switch",
      "attaches_to": "Faceplate",
      "is_solid": true,
      "sketch_plane_y": -165.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -25.0,
            -165.0,
            290.0
          ],
          "end": [
            25.0,
            -165.0,
            290.0
          ]
        },
        {
          "type": "Line",
          "start": [
            25.0,
            -165.0,
            290.0
          ],
          "end": [
            25.0,
            -165.0,
            330.0
          ]
        },
        {
          "type": "Line",
          "start": [
            25.0,
            -165.0,
            330.0
          ],
          "end": [
            -25.0,
            -165.0,
            330.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -25.0,
            -165.0,
            330.0
          ],
          "end": [
            -25.0,
            -165.0,
            290.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -20.0
    }
  ]
}
```
