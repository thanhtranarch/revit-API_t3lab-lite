# CATEGORY OVERLAY — Columns

> Appended AFTER the general core prompt above. All core rules still apply. This section
> adds Column-specific guidance. Set `"family_category": "Columns"`. Output ONLY the JSON
> object.

## Part inventory (typical — adapt to the real object)
- Shaft (the main vertical body — round, square, or fluted).
- Base / plinth (wider foot at the bottom).
- Capital (wider head at the top).
- Optional: necking rings, entasis (slight bulge), fluting details.

## Placement & anchor
- Vertical, **centered on X=0, Y=0**. Bottom at **Z = 0**, growing upward. Anchor = base/plinth.
- Stack strictly by Z: plinth (z 0→h1) → base moulding → shaft (z h1→h2) → capital (z h2→top).
  Each part overlaps the previous 1–2 mm in Z; base and capital are wider than the shaft.
  Chain with `attaches_to`: Shaft → Base, Capital → Shaft.

## Typical dimensions (mm)
- Shaft dia (round) or side (square) 200–600. Height per storey 3000–4000 (or as stated).
- Base/capital 1.2–1.6× the shaft width, height 100–400 each. Plinth square, height 80–200.

## Category pitfalls
- Round shafts → **Revolution** of the half-silhouette (add a subtle `Arc3P` entasis if the
  column visibly bulges), or a simple Extrusion of a `Circle` for a plain cylinder.
- Keep everything concentric on the Z axis — an off-center base/capital reads as broken.
- Fluting → many thin `inner_loops`/voids around the shaft, or leave as a plain shaft if not
  essential; don't approximate a round shaft with a box.


## Worked example (Column (base + round shaft + capital, stacked)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_Column",
  "family_category": "Columns",
  "_plan": {
    "parts": [
      "Base",
      "Shaft",
      "Capital"
    ],
    "heights": {
      "Base": "z 0-150",
      "Shaft": "z 148-3000",
      "Capital": "z 2998-3150"
    },
    "connections": {
      "Shaft": "bottom laps base top",
      "Capital": "bottom laps shaft top"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Base",
      "is_solid": true,
      "sketch_plane_z": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -250.0,
            -250.0,
            0.0
          ],
          "end": [
            250.0,
            -250.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            250.0,
            -250.0,
            0.0
          ],
          "end": [
            250.0,
            250.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            250.0,
            250.0,
            0.0
          ],
          "end": [
            -250.0,
            250.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -250.0,
            250.0,
            0.0
          ],
          "end": [
            -250.0,
            -250.0,
            0.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 150.0
    },
    {
      "type": "Extrusion",
      "id": "Shaft",
      "attaches_to": "Base",
      "is_solid": true,
      "sketch_plane_z": 148.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            0.0,
            148.0
          ],
          "radius": 180.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 2852.0
    },
    {
      "type": "Extrusion",
      "id": "Capital",
      "attaches_to": "Shaft",
      "is_solid": true,
      "sketch_plane_z": 2998.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -250.0,
            -250.0,
            2998.0
          ],
          "end": [
            250.0,
            -250.0,
            2998.0
          ]
        },
        {
          "type": "Line",
          "start": [
            250.0,
            -250.0,
            2998.0
          ],
          "end": [
            250.0,
            250.0,
            2998.0
          ]
        },
        {
          "type": "Line",
          "start": [
            250.0,
            250.0,
            2998.0
          ],
          "end": [
            -250.0,
            250.0,
            2998.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -250.0,
            250.0,
            2998.0
          ],
          "end": [
            -250.0,
            -250.0,
            2998.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 152.0
    }
  ]
}
```
