# CATEGORY OVERLAY — Site

> Appended AFTER the general core prompt above. All core rules apply. Set
> `"family_category": "Site"`. Output ONLY the JSON object.

## Part inventory (typical — adapt to the real object)
- Site furnishings and objects: bollards, benches, planters, posts, signage, fences, bins,
  light poles, manholes, tree grates, retaining/curb elements.

## Placement & anchor
- **Ground-mounted**: base at **Z = 0**, growing upward. Center on X=0, Y=0. Anchor = base/foot.
- Multi-part objects (a bench = slats + frame + legs; a lamp post = base + shaft + head) chain
  every part to the anchor via `attaches_to` with a verified 1–2 mm 3-axis overlap.

## Typical dimensions (mm)
- Bollard dia 100–300, height 600–1000. Bench length 1200–1800, seat z ~450. Light pole
  height 3000–6000. Planter 400–1000.

## Category pitfalls
- Round bollards/posts/poles → Extrusion of a `Circle` or Revolution, not boxes.
- Keep parts connected and grounded (base at Z=0) — floating site objects read as broken.
- Curved paths/rails → `Arc3P`; repeated pickets → keep them overlapping their rails.


## Worked example (Bollard (base + post + domed cap)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_Bollard",
  "family_category": "Site",
  "_plan": {
    "parts": [
      "Base",
      "Post",
      "Cap"
    ],
    "heights": {
      "Base": "z 0-40",
      "Post": "z 38-900",
      "Cap": "z 895-950"
    },
    "connections": {
      "Post": "bottom laps base",
      "Cap": "dome on post top"
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
          "type": "Circle",
          "center": [
            0.0,
            0.0,
            0.0
          ],
          "radius": 130.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 40.0
    },
    {
      "type": "Extrusion",
      "id": "Post",
      "attaches_to": "Base",
      "is_solid": true,
      "sketch_plane_z": 38.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            0.0,
            38.0
          ],
          "radius": 90.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 862.0
    },
    {
      "type": "Revolution",
      "id": "Cap",
      "attaches_to": "Post",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            895.0
          ],
          "end": [
            90.0,
            0.0,
            895.0
          ]
        },
        {
          "type": "Arc3P",
          "start": [
            90.0,
            0.0,
            895.0
          ],
          "end": [
            0.0,
            0.0,
            950.0
          ],
          "mid": [
            64.0,
            0.0,
            935.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            950.0
          ],
          "end": [
            0.0,
            0.0,
            895.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        0.0,
        895.0
      ],
      "axis_end": [
        0.0,
        0.0,
        950.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    }
  ]
}
```
