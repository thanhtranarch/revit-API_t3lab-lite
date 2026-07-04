# CATEGORY OVERLAY — Furniture

> Appended AFTER the general core prompt above. All core rules (schema, forms, curve
> segments, `_plan`/`attaches_to` connectivity, failure modes, checklist) still apply.
> This section adds Furniture-specific guidance. Set `"family_category": "Furniture"`.
> Output ONLY the JSON object.

## Part inventory (typical — adapt to the real object)
- Tables/desks: top → apron/rails → 4 legs → stretchers → optional drawers/pulls.
- Chairs/stools: seat → 4 legs (or base column + foot) → aprons → back stiles → crest/back rails → arms → stretchers.
- Beds: headboard → footboard → side rails → slats/platform → legs.
- Sofas/upholstered: base frame → seat cushions → back cushions → arms → feet.
- Shelving: uprights/sides → shelves → back panel → plinth.

## Placement & anchor
- **Floor-standing**: bottom of the object at **Z = 0**, growing upward. Center on X=0, Y=0.
- **Anchor part** = whatever touches the floor (the legs or the base). Every other part
  chains to it via `attaches_to` with a verified 1–2 mm overlap in all three axes:
  legs → into the top/apron at their top face; stretchers → into two legs; back stiles →
  into the seat and the crest rail.

## Typical dimensions (mm)
- Table/desk top surface z 720–750, top thickness 18–40. Dining chair seat z 430–470,
  back top ~900–1000. Bar stool seat z 650–750. Coffee table z 400–450.
- Square leg 40–60; round leg dia 40–80 (use **Revolution**, not a box).

## Category pitfalls
- Round/turned legs, spindles, balusters → **Revolution** of the half-silhouette (core
  Example 4 / 13), NOT an extruded square.
- Legs must actually **reach** the underside of the top/apron — the most common furniture
  bug is legs stopping short (floating). Derive leg `extrusion_end` from the top's underside z.
- Rounded/soft edges on tops and cushions → `Arc3P` corners, not sharp rectangles.


## Worked example (Pedestal stool (Extrusion + Revolution, stacked & connected)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_PedestalStool",
  "family_category": "Furniture",
  "_plan": {
    "parts": [
      "Base",
      "Column",
      "Seat"
    ],
    "heights": {
      "Base": "z 0-62",
      "Column": "z 60-442",
      "Seat": "z 440-470"
    },
    "connections": {
      "Column": "bottom z 60 inside base (0-62), centred",
      "Seat": "bottom z 440 laps column top (to 442); seat covers column"
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
          "radius": 140.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 62.0
    },
    {
      "type": "Revolution",
      "id": "Column",
      "attaches_to": "Base",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            60.0
          ],
          "end": [
            40.0,
            0.0,
            60.0
          ]
        },
        {
          "type": "Line",
          "start": [
            40.0,
            0.0,
            60.0
          ],
          "end": [
            40.0,
            0.0,
            442.0
          ]
        },
        {
          "type": "Line",
          "start": [
            40.0,
            0.0,
            442.0
          ],
          "end": [
            0.0,
            0.0,
            442.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            442.0
          ],
          "end": [
            0.0,
            0.0,
            60.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        0.0,
        60.0
      ],
      "axis_end": [
        0.0,
        0.0,
        442.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    },
    {
      "type": "Extrusion",
      "id": "Seat",
      "attaches_to": "Column",
      "is_solid": true,
      "sketch_plane_z": 440.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            0.0,
            0.0,
            440.0
          ],
          "radius": 180.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 30.0
    }
  ]
}
```
