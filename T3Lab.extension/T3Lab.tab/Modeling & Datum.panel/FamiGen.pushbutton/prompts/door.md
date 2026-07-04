# CATEGORY OVERLAY — Door

> Appended AFTER the general core prompt above. All core rules still apply. This section
> adds Door-specific guidance. Set `"family_category": "Door"`. Output ONLY the JSON object.
> Core **Example 11** (panel with a framed opening) is the base recipe.

## Part inventory (typical — adapt to the real object)
- Frame: two jambs + head (+ optional threshold/sill) — a rectangular ring.
- Leaf / panel (single or double). Panelled leaves: stiles + rails + raised/flat panels.
- Glazing / vision light (thin pane set inside an opening).
- Hardware: lever/knob, escutcheon, hinges, optional kick plate.

## Placement & anchor
- The wall face is the **XZ plane**; the door's thickness runs along **Y** (through the wall).
  Width along X (centered on X=0), height along Z from **Z = 0** up.
- **Do NOT model the wall** — only the door assembly. Anchor = the frame.
- Leaf sits inside the frame; glazing sits inside the leaf opening; lever sits on the leaf
  face projecting out along Y. Chain with `attaches_to`: Leaf → Frame, Glazing → Leaf,
  Lever → Leaf.

## Typical dimensions (mm)
- Single leaf: width 700–1000, height 2100–2400, leaf thickness 40–45.
- Frame profile 45–70 wide, frame depth = wall-ish 100–200. Vision light inset 80–150 from edges.
- Lever projection 55–65; escutcheon 40–60.

## Category pitfalls
- Build the leaf on a vertical plane (`sketch_plane_y = 0`) and extrude the thickness along Y;
  don't lay it flat.
- Glazed lights and louvres → `inner_loops` through the leaf, then (optionally) a thin glass
  Extrusion filling the opening — not a separate void box.
- Panelled doors: model stiles/rails as one framed slab with `inner_loops` recesses, or as
  separate overlapping slabs — keep every added panel overlapping the leaf (verify 3 axes).


## Worked example (Glazed door (Extrusion + inner_loops + lever)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_GlazedDoor",
  "family_category": "Door",
  "_plan": {
    "parts": [
      "Leaf",
      "Glazing",
      "Lever"
    ],
    "heights": {
      "Leaf": "z 0-2100",
      "Glazing": "z 1098-1902",
      "Lever": "z ~1050"
    },
    "connections": {
      "Glazing": "laps the leaf opening edges (2 mm) inside leaf depth",
      "Lever": "base on leaf front face y=45"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Leaf",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -450.0,
            0.0,
            0.0
          ],
          "end": [
            450.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            450.0,
            0.0,
            0.0
          ],
          "end": [
            450.0,
            0.0,
            2100.0
          ]
        },
        {
          "type": "Line",
          "start": [
            450.0,
            0.0,
            2100.0
          ],
          "end": [
            -450.0,
            0.0,
            2100.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -450.0,
            0.0,
            2100.0
          ],
          "end": [
            -450.0,
            0.0,
            0.0
          ]
        }
      ],
      "inner_loops": [
        [
          {
            "type": "Line",
            "start": [
              -350.0,
              0.0,
              1100.0
            ],
            "end": [
              350.0,
              0.0,
              1100.0
            ]
          },
          {
            "type": "Line",
            "start": [
              350.0,
              0.0,
              1100.0
            ],
            "end": [
              350.0,
              0.0,
              1900.0
            ]
          },
          {
            "type": "Line",
            "start": [
              350.0,
              0.0,
              1900.0
            ],
            "end": [
              -350.0,
              0.0,
              1900.0
            ]
          },
          {
            "type": "Line",
            "start": [
              -350.0,
              0.0,
              1900.0
            ],
            "end": [
              -350.0,
              0.0,
              1100.0
            ]
          }
        ]
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 45.0
    },
    {
      "type": "Extrusion",
      "id": "Glazing",
      "attaches_to": "Leaf",
      "is_solid": true,
      "sketch_plane_y": 15.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -352.0,
            15.0,
            1098.0
          ],
          "end": [
            352.0,
            15.0,
            1098.0
          ]
        },
        {
          "type": "Line",
          "start": [
            352.0,
            15.0,
            1098.0
          ],
          "end": [
            352.0,
            15.0,
            1902.0
          ]
        },
        {
          "type": "Line",
          "start": [
            352.0,
            15.0,
            1902.0
          ],
          "end": [
            -352.0,
            15.0,
            1902.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -352.0,
            15.0,
            1902.0
          ],
          "end": [
            -352.0,
            15.0,
            1098.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 15.0
    },
    {
      "type": "Extrusion",
      "id": "Lever",
      "attaches_to": "Leaf",
      "is_solid": true,
      "sketch_plane_y": 45.0,
      "profile": [
        {
          "type": "Circle",
          "center": [
            380.0,
            45.0,
            1050.0
          ],
          "radius": 14.0
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 60.0
    }
  ]
}
```
