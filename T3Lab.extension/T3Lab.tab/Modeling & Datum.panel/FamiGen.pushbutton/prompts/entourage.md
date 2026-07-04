# CATEGORY OVERLAY — Entourage

> Appended AFTER the general core prompt above. Entourage = people, plants/trees, vehicles and
> scene props used for context. All core rules apply. Set `"family_category": "Entourage"`.
> Output ONLY the JSON object.

## Guidance
- These objects are organic/complex — model **simplified massing**, not fine detail.
- **Trees/plants**: trunk = Revolution/Extrusion of a `Circle`; canopy = Revolution of an
  `Arc3P` (dome) or a Blend; anchor = trunk base at Z=0.
- **People**: block/mass approximation — a base/legs mass → torso mass → head (sphere via
  Revolution). Keep it upright, feet at Z=0.
- **Vehicles**: body as Extrusions/Blends, wheels as Extrusions of a `Circle` on the side plane;
  anchor = the chassis/wheels contacting Z=0.

## Placement & anchor
- Ground objects: base/feet/wheels at **Z = 0**. Anchor = that base; chain all masses upward
  with `attaches_to` and a verified 1–2 mm 3-axis overlap.

## Category pitfalls
- Don't over-model — a few clean masses read better than many disconnected fragments.
- Use `Spline`/`Arc3P` for organic silhouettes revolved or extruded; avoid faceting curves
  into straight lines. Keep every mass connected to the body.


## Worked example (Tree (trunk + revolved canopy blob)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_Tree",
  "family_category": "Entourage",
  "_plan": {
    "parts": [
      "Trunk",
      "Canopy"
    ],
    "heights": {
      "Trunk": "z 0-1800",
      "Canopy": "z 1700-3200"
    },
    "connections": {
      "Canopy": "base z 1700 laps trunk top (to 1800)"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Trunk",
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
          "radius": 90.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 1800.0
    },
    {
      "type": "Revolution",
      "id": "Canopy",
      "attaches_to": "Trunk",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            1700.0
          ],
          "end": [
            80.0,
            0.0,
            1700.0
          ]
        },
        {
          "type": "Arc3P",
          "start": [
            80.0,
            0.0,
            1700.0
          ],
          "end": [
            0.0,
            0.0,
            3200.0
          ],
          "mid": [
            750.0,
            0.0,
            2450.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            3200.0
          ],
          "end": [
            0.0,
            0.0,
            1700.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        0.0,
        1700.0
      ],
      "axis_end": [
        0.0,
        0.0,
        3200.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    }
  ]
}
```
