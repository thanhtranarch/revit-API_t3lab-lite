# CATEGORY OVERLAY — Plumbing Fixture

> Appended AFTER the general core prompt above. All core rules still apply. This section
> adds Plumbing-specific guidance. Set `"family_category": "Plumbing Fixture"`. Output ONLY
> the JSON object.

## Part inventory (typical — adapt to the real object)
- Basin / bowl / tub / sink body (usually round or oval — a Revolution or Blend).
- Spout / faucet body / neck (often a bent tube).
- Handle(s) / lever(s) / cross-tops.
- Pedestal / base / apron, or counter-mount rim.
- Trap / drain, escutcheon / wall flange (wall-mounted).

## Placement & anchor
- **Floor-mounted** (toilet, pedestal basin, tub): base at **Z = 0**, anchor = base/foot.
- **Wall-mounted** (wall basin, wall faucet): wall face at **Y = 0**, escutcheon/flange on the
  wall, body projecting to negative Y; anchor = the wall flange/plate.
- **Counter-mounted** (vessel basin, deck faucet): base at the counter z, anchor = base ring.
- Spout → attaches into the faucet body; handles → into the body/deck. Verify the spout
  actually reaches the body (3-axis overlap) — a floating spout is the classic failure.

## Typical dimensions (mm)
- Basin 450–650 wide, bowl depth 150–200. Toilet ~700 long × 380 wide, seat z ~400.
- Bathtub 1500–1800 long. Faucet spout tube radius 8–15; spout reach 100–200; handle r 15–25.

## Category pitfalls
- Round bowls, basins, tub shells → **Revolution** (of an `Arc3P` section for the curved
  wall), not a box or a taper-to-zero Blend.
- Bent spouts/necks → orthogonal Extrusion-of-`Circle` runs (core Example 7) for right-angle
  bends, or a Sweep with an `Arc3P` path for a smooth gooseneck (core Example 12).
- Hollow bowls → outer Revolution + inner void / `inner_loops`.


## Worked example (Pedestal basin (two Revolutions, stacked)) - connected, with `_plan` + `attaches_to`

```json
{
  "family_name": "T3Lab_PedestalBasin",
  "family_category": "Plumbing Fixture",
  "_plan": {
    "parts": [
      "Pedestal",
      "Basin"
    ],
    "heights": {
      "Pedestal": "z 0-800",
      "Basin": "z 790-900"
    },
    "connections": {
      "Basin": "base z 790 laps pedestal top (to 800), centred"
    }
  },
  "geometry": [
    {
      "type": "Revolution",
      "id": "Pedestal",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            0.0
          ],
          "end": [
            160.0,
            0.0,
            0.0
          ]
        },
        {
          "type": "Arc3P",
          "start": [
            160.0,
            0.0,
            0.0
          ],
          "end": [
            120.0,
            0.0,
            800.0
          ],
          "mid": [
            110.0,
            0.0,
            400.0
          ]
        },
        {
          "type": "Line",
          "start": [
            120.0,
            0.0,
            800.0
          ],
          "end": [
            0.0,
            0.0,
            800.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            800.0
          ],
          "end": [
            0.0,
            0.0,
            0.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        0.0,
        0.0
      ],
      "axis_end": [
        0.0,
        0.0,
        800.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    },
    {
      "type": "Revolution",
      "id": "Basin",
      "attaches_to": "Pedestal",
      "is_solid": true,
      "sketch_plane_y": 0.0,
      "profile": [
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            790.0
          ],
          "end": [
            280.0,
            0.0,
            790.0
          ]
        },
        {
          "type": "Arc3P",
          "start": [
            280.0,
            0.0,
            790.0
          ],
          "end": [
            40.0,
            0.0,
            900.0
          ],
          "mid": [
            230.0,
            0.0,
            815.0
          ]
        },
        {
          "type": "Line",
          "start": [
            40.0,
            0.0,
            900.0
          ],
          "end": [
            0.0,
            0.0,
            900.0
          ]
        },
        {
          "type": "Line",
          "start": [
            0.0,
            0.0,
            900.0
          ],
          "end": [
            0.0,
            0.0,
            790.0
          ]
        }
      ],
      "axis_start": [
        0.0,
        0.0,
        790.0
      ],
      "axis_end": [
        0.0,
        0.0,
        900.0
      ],
      "start_angle": 0.0,
      "end_angle": 6.283185307
    }
  ]
}
```
