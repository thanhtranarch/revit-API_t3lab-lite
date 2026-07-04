# CATEGORY OVERLAY — Generic Model

> Appended AFTER the general core prompt above. Generic Model is the **catch-all** template —
> use it for anything that doesn't fit a specific category. All core rules apply unchanged.
> Set `"family_category": "Generic Model"`. Output ONLY the JSON object.

## Guidance
- There is no fixed part inventory — decompose the actual object with the core Step-1 method.
- **Placement**: floor/tabletop objects → bottom at Z=0; wall-mounted → wall face at Y=0,
  body at negative Y; ceiling-mounted → ceiling at Z=0, body downward (negative Z).
- **Anchor** = whichever part contacts the host surface; every other part chains to it via
  `attaches_to` with a verified 1–2 mm 3-axis overlap.
- Pick the form per part from the core SHAPE → FORM recipe table; don't force everything into
  boxes. This is the right category when you simply need faithful geometry without category
  parameters.


## Worked example (Sign plate on a post (Extrusion + connectivity)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_SignPost",
  "family_category": "Generic Model",
  "_plan": {
    "parts": [
      "Post",
      "Plate"
    ],
    "heights": {
      "Post": "z 0-300",
      "Plate": "z 300-450"
    },
    "connections": {
      "Plate": "bottom z=300 meets post top z=300, on the post"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Post",
      "is_solid": true,
      "sketch_plane_z": 0.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            0.0,
            0.0
          ],
          "radius": 15.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 300.0
    },
    {
      "type": "Extrusion",
      "id": "Plate",
      "attaches_to": "Post",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -120.0,
            0.0,
            300.0
          ],
          "end": [
            120.0,
            0.0,
            300.0
          ]
        },
        {
          "type": "Line",
          "start": [
            120.0,
            0.0,
            300.0
          ],
          "end": [
            120.0,
            0.0,
            450.0
          ]
        },
        {
          "type": "Line",
          "start": [
            120.0,
            0.0,
            450.0
          ],
          "end": [
            -120.0,
            0.0,
            450.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -120.0,
            0.0,
            450.0
          ],
          "end": [
            -120.0,
            0.0,
            300.0
          ]
        }
      ],
      "extrusion_start": -8.0,
      "extrusion_end": 8.0
    }
  ]
}
```
