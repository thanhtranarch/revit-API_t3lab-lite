# CATEGORY OVERLAY — Mechanical Equipment

> Appended AFTER the general core prompt above. All core rules apply. Set
> `"family_category": "Mechanical Equipment"`. Output ONLY the JSON object.

## Part inventory (typical — adapt to the real object)
- Main housing / unit body (AHU, FCU, pump, chiller, fan, VAV box).
- Duct / pipe connections (round or rectangular collars/stubs).
- Grilles / louvres / fan guards, access panels.
- Base / feet / mounting rails, motor, brackets.

## Placement & anchor
- **Floor-mounted**: base/feet at **Z = 0**, anchor = base. **Ceiling-hung**: mounting
  brackets at the top (Z = top), body below. **Wall-mounted**: plate at **Y = 0**.
- Duct/pipe collars attach to the housing faces; grilles cover openings on a face. Chain with
  `attaches_to`, verify 3-axis overlap.

## Typical dimensions (mm)
- Units vary widely (300–3000). Round duct dia 100–600; rectangular duct 200–800 wide.
- Feet/base height 50–150.

## Category pitfalls
- Round duct connections/pipes → Extrusion of a `Circle` on the face's perpendicular plane;
  bent pipe runs → orthogonal Extrusion-of-`Circle` runs (core Example 7).
- Grilles/fan guards → face Extrusion with `inner_loops` slots, not modelled bar-by-bar.
- Keep collars/stubs overlapping the housing (no floating connections).


## Worked example (Fan-coil unit (feet + housing + round duct collar)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_FanCoilUnit",
  "family_category": "Mechanical Equipment",
  "_plan": {
    "parts": [
      "Feet",
      "Housing",
      "DuctCollar"
    ],
    "heights": {
      "Feet": "z 0-100",
      "Housing": "z 98-700",
      "DuctCollar": "z ~400, out front"
    },
    "connections": {
      "Housing": "sits on feet",
      "DuctCollar": "on housing front face y=-300"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Feet",
      "is_solid": true,
      "sketch_plane_z": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -350.0,
            -280.0,
            0.0
          ],
          "end": [
            350.0,
            -280.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            350.0,
            -280.0,
            0.0
          ],
          "end": [
            350.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            350.0,
            0.0,
            0.0
          ],
          "end": [
            -350.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -350.0,
            0.0,
            0.0
          ],
          "end": [
            -350.0,
            -280.0,
            0.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 100.0
    },
    {
      "type": "Extrusion",
      "id": "Housing",
      "attaches_to": "Feet",
      "is_solid": true,
      "sketch_plane_z": 98.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -350.0,
            -300.0,
            98.0
          ],
          "end": [
            350.0,
            -300.0,
            98.0
          ]
        },
        {
          "type": "Line",
          "start": [
            350.0,
            -300.0,
            98.0
          ],
          "end": [
            350.0,
            0.0,
            98.0
          ]
        },
        {
          "type": "Line",
          "start": [
            350.0,
            0.0,
            98.0
          ],
          "end": [
            -350.0,
            0.0,
            98.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -350.0,
            0.0,
            98.0
          ],
          "end": [
            -350.0,
            -300.0,
            98.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 602.0
    },
    {
      "type": "Extrusion",
      "id": "DuctCollar",
      "attaches_to": "Housing",
      "is_solid": true,
      "sketch_plane_y": -300.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            -300.0,
            400.0
          ],
          "radius": 90.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": -120.0
    }
  ]
}
```
