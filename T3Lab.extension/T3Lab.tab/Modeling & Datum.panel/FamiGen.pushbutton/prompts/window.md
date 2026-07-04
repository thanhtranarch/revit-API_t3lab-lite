# CATEGORY OVERLAY — Window

> Appended AFTER the general core prompt above. All core rules still apply. This section
> adds Window-specific guidance. Set `"family_category": "Window"`. Output ONLY the JSON
> object. Core **Example 11** (panel + framed opening) is the base recipe for frame + glass.

## Part inventory (typical — adapt to the real object)
- Outer frame (rectangular ring: 2 jambs + head + sill).
- Sash(es) — the openable/fixed sub-frame(s).
- Glazing pane(s) — thin sheet(s) inside the sash/frame opening.
- Mullions / muntins / transom bars dividing lights.
- Sill / stool, optional apron.

## Placement & anchor
- The wall face is the **XZ plane**; window depth runs along **Y**. Width along X
  (centered on X=0), height along Z. **Do NOT model the wall.** Anchor = the outer frame.
- Glass sits inside the frame opening (thin, centered in the frame depth); mullions sit on
  top of the glass plane spanning the opening. Chain: Sash → Frame, Glass → Sash/Frame,
  Mullion → Frame.

## Typical dimensions (mm)
- Width 600–1800, height 900–1500 (sill height set by host, model from Z=0 up).
- Frame profile 45–70 wide, frame/wall depth 100–200. Glass thickness 6–24. Mullion 30–50.

## Category pitfalls
- Frame = rectangle with a rectangular `inner_loops` opening (the glass hole), extruded
  along Y — not four separate bars unless you need a profiled section.
- Glass = a thin Extrusion filling the opening; keep it inside the frame depth (verify Y range).
- Muntins/mullions are thin — model as slim Extrusions overlapping the frame; keep them
  connected (no floating bars).


## Worked example (Window (frame ring + glass + mullion)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_Window",
  "family_category": "Window",
  "_plan": {
    "parts": [
      "Frame",
      "Glass",
      "Mullion"
    ],
    "heights": {
      "Frame": "z 0-1200",
      "Glass": "z 38-1162",
      "Mullion": "z 40-1160"
    },
    "connections": {
      "Glass": "laps frame opening edges inside frame depth",
      "Mullion": "spans opening, touches frame head and sill"
    }
  },
  "geometry": [
    {
      "type": "Extrusion",
      "id": "Frame",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -600.0,
            0.0,
            0.0
          ],
          "end": [
            600.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Line",
          "start": [
            600.0,
            0.0,
            0.0
          ],
          "end": [
            600.0,
            0.0,
            1200.0
          ]
        },
        {
          "type": "Line",
          "start": [
            600.0,
            0.0,
            1200.0
          ],
          "end": [
            -600.0,
            0.0,
            1200.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -600.0,
            0.0,
            1200.0
          ],
          "end": [
            -600.0,
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
              -560.0,
              0.0,
              40.0
            ],
            "end": [
              560.0,
              0.0,
              40.0
            ]
          },
          {
            "type": "Line",
            "start": [
              560.0,
              0.0,
              40.0
            ],
            "end": [
              560.0,
              0.0,
              1160.0
            ]
          },
          {
            "type": "Line",
            "start": [
              560.0,
              0.0,
              1160.0
            ],
            "end": [
              -560.0,
              0.0,
              1160.0
            ]
          },
          {
            "type": "Line",
            "start": [
              -560.0,
              0.0,
              1160.0
            ],
            "end": [
              -560.0,
              0.0,
              40.0
            ]
          }
        ]
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 120.0
    },
    {
      "type": "Extrusion",
      "id": "Glass",
      "attaches_to": "Frame",
      "is_solid": true,
      "sketch_plane_y": 55.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -562.0,
            55.0,
            38.0
          ],
          "end": [
            562.0,
            55.0,
            38.0
          ]
        },
        {
          "type": "Line",
          "start": [
            562.0,
            55.0,
            38.0
          ],
          "end": [
            562.0,
            55.0,
            1162.0
          ]
        },
        {
          "type": "Line",
          "start": [
            562.0,
            55.0,
            1162.0
          ],
          "end": [
            -562.0,
            55.0,
            1162.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -562.0,
            55.0,
            1162.0
          ],
          "end": [
            -562.0,
            55.0,
            38.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 10.0
    },
    {
      "type": "Extrusion",
      "id": "Mullion",
      "attaches_to": "Frame",
      "is_solid": true,
      "sketch_plane_y": 50.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            -20.0,
            50.0,
            40.0
          ],
          "end": [
            20.0,
            50.0,
            40.0
          ]
        },
        {
          "type": "Line",
          "start": [
            20.0,
            50.0,
            40.0
          ],
          "end": [
            20.0,
            50.0,
            1160.0
          ]
        },
        {
          "type": "Line",
          "start": [
            20.0,
            50.0,
            1160.0
          ],
          "end": [
            -20.0,
            50.0,
            1160.0
          ]
        },
        {
          "type": "Line",
          "start": [
            -20.0,
            50.0,
            1160.0
          ],
          "end": [
            -20.0,
            50.0,
            40.0
          ]
        }
      ],
      "extrusion_start": 0.0,
      "extrusion_end": 20.0
    }
  ]
}
```
