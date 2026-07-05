# FamiGen — "From JSON" Improvement Roadmap

> Standalone roadmap for the FamiGen **From JSON** feature (photo-feedback loop).
> Not part of the `gd<N> revitapi` phase tables — do not wire into `README.md`'s
> "Lịch thực hiện". Source files:
> - Prompts (external AI → JSON): `T3Lab.extension/T3Lab.tab/Modeling & Datum.panel/FamiGen.pushbutton/prompts/<slug>.md`
>   — **7 self-contained per-category files** (`SYSTEM_PROMPT.md` retired 2026-07-05, see below)
> - Parser (JSON → Revit API): `T3Lab.extension/lib/GUI/FamiGenDialog.py`
>   (`_generate_json_family`, `_json_curve`, `_json_loop`, `_heal_loop`, …)

## Guiding principle
**The shared core (the identical block at the top of all 7 prompt files) must stay
category-general**: teach transferable *techniques*, not per-object templates.
Category-specific knowledge (case libraries, playbooks, worked examples) lives ONLY in
each file's `# <CAT>-SPECIFIC GUIDANCE` section. The 7 cores are **byte-identical** and
there is no sync script — apply core edits with a replace-in-all-files script, then
verify equality.

## Done
- [x] **WS2 — Rebalance examples (de-bias).** Removed lighting-specific framing:
      Example 6 Lotus lamp → Ceiling Medallion (Generic Model); Example 7 wall-lamp
      arm → generic tube; part-inventory `wall lantern` → `wall-mounted faucet`.
      Examples now span Furniture / Generic Model / Casework / Door, **0 Lighting**.
- [x] **WS3 — Orthogonal tube recipe.** Bent / right-angle tubes = axis-aligned
      Extrusion-of-`Circle` runs overlapping ~2 mm at each elbow (Example 7), NOT a
      bent Sweep (parser only allows X/Y/Z sketch planes → diagonal Sweeps fragile).
      Added Step 5 rule + 3-row RECIPES split (straight / bent / curved tube).
- [x] **More examples (14 total).** Added E8 assembly/connectivity (Casework),
      E9 dome (Revolution+Arc3P), E10 sphere (Revolution semicircle), E11 panel
      with framed opening (Extrusion+inner_loops, Door), E12 curved Sweep (Arc3P
      path), E13 organic Revolution (Spline), E14 Ellipse extrusion. All 4 form
      types + all 6 curve types now demonstrated. Added a "generalize, don't copy"
      header to the EXAMPLES section.

- [x] **WS1 — Mandatory "connection pass" in the prompt** (2026-07-04, triggered by
      a worse-than-before disconnected lamp result: parts built fine but floated).
      Step 4 now requires an anchor part + per-part `"attaches_to"` key (parser
      ignores it — reasoning-forcing only) + arithmetic 3-axis range-overlap check;
      orphan parts declared a bug; hard-failure #7 rewritten; checklist gained a
      connectivity walk + a "never reuse example coordinates" rule; E7/E8 demo the
      check.
- [x] **WS6 — `"_plan"` scratch key** (2026-07-04). Root key the parser ignores;
      the AI must write part inventory + height table + connection table there
      before geometry. Demonstrated in Example 8.

## Backlog (proposed, prioritized)
- [ ] **WS4 — Disconnected-part warning in the parser.** After build, a light
      bounding-box check flags any solid that doesn't touch/overlap another and
      surfaces it in the result dialog (currently only API *failures* are reported,
      not built-but-floating parts). Complements WS1: prompt-side prevention can
      still be ignored by the AI; parser-side detection catches it.
- [ ] **WS5 — Arbitrary sketch planes in the parser.** Accept `sketch_plane` by
      `normal + origin` so genuinely diagonal geometry (angled legs, sloped panels,
      true diagonal arms) builds without axis-aligned approximation. Larger parser
      change — do after WS1/WS4 raise the base build rate.

## Done — split the prompt by family category (2026-07-04)
Shipped the per-category split: **select family type first, then Copy Prompt** copies
`general core + that category's overlay`.

- **Shared core** = `SYSTEM_PROMPT.md` (unchanged, still fully usable standalone).
- **Per-category overlays** = `FamiGen.pushbutton/prompts/<slug>.md`, one per Revit
  family category (13 files: generic_model, door, window, furniture, plumbing_fixture,
  electrical_equipment, mechanical_equipment, specialty_equipment, casework, columns,
  lighting_fixture, site, entourage). Each adds that category's part inventory,
  placement/anchor convention, typical dimensions and pitfalls, and references the
  relevant core example.
- **UI**: `FamiGen.xaml` JSON action bar gained a "Family type" `ComboBox`
  (`json_category_combo`, `T3ComboBoxSmall`) left of Copy Prompt; Tip updated.
- **Wiring**: `FamiGenDialog.py` — `_PROMPTS_DIR` const, `_init_json_panel()` fills the
  combo from `CATEGORY_TEMPLATES`, `_category_slug()`/`_overlay_path()` map name→file,
  `copy_prompt_clicked()` concatenates `core + overlay` (falls back to core-only if an
  overlay is missing, so the general prompt always works).
- **Slug rule**: `re.sub(r'[^a-z0-9]+','_', name.lower()).strip('_')` — keep overlay
  filenames in sync if `CATEGORY_TEMPLATES` changes.
- **Constraint kept:** core alone still builds any family; overlays are an enhancement.

### Overlay worked examples (2026-07-04)
Every one of the 13 overlays now ends with a **complete, connected worked JSON example**
(with `_plan` + per-part `attaches_to`), validated: valid JSON, exactly one anchor, and a
geometric 3-axis bounding-box check confirming 0 disconnected parts across all 13. The
lighting overlay's example is the user's **wall swing-arm lamp built correctly** (plate →
orthogonal arm → drop → drum shade as an Extrusion, all lapping).

### `Cylinder` primitive (2026-07-04) — kills the wrong-axis rod failure
Real test of the wall lamp came out worse: the AI put the vertical riser on `sketch_plane_x`
and the horizontal arm on `sketch_plane_z`, so both extruded along the wrong axis (Extrusion
pushes perpendicular to its plane) — rods scattered. Added a `Cylinder` geometry type
(`start`, `end`, `radius`): the parser derives the sketch plane & direction (axis-aligned →
Extrusion of a circle; diagonal → perpendicular-circle Sweep), so the AI can't mismatch plane
vs axis. Prompt now: `Cylinder` is the required form for every straight round rod/arm/leg/spout;
chain `Cylinder`s sharing endpoints for bent arms; Example 7 + the lighting overlay example
rewritten to `Cylinder`; new hard-failure #10 documents the wrong-axis trap.

### Sweep bug fix (2026-07-04)
`NewCurveLoopsProfile` lives on `Application.Create`, not `Application` — every Sweep (and
the diagonal-`Cylinder` fallback) was calling `fam_doc.Application.NewCurveLoopsProfile(...)`
and dying with "'Application' object has no attribute 'NewCurveLoopsProfile'". Fixed all 3
call sites to `fam_doc.Application.Create.NewCurveLoopsProfile(...)`. (Sweeps had never
actually built before this.)

### Parser hardening (2026-07-04)
- **Blend → Extrusion fallback**: a Blend whose two loops are congruent (same area +
  centroid, e.g. a drum shade circle→identical circle) now auto-builds as an Extrusion
  instead of dying with NewBlend "internal error code 1". Helpers `_loops_congruent()` /
  `_loop_centroid_uv()` in `FamiGenDialog.py`; the core prompt's Blend section now also
  tells the AI a same-size top/base is an Extrusion, not a Blend.

## Done — self-contained prompts + soft-form detail upgrade (2026-07-05)

### Restructure: 13 core+overlay → 7 self-contained files
`SYSTEM_PROMPT.md` deleted; each `prompts/<slug>.md` is now a **fully self-contained**
system prompt (schema, forms, curves, failure modes, checklist + category guidance).
Category set trimmed to the 7 in active use: generic_model, furniture, casework,
plumbing_fixture, mechanical_equipment, specialty_equipment, lighting_fixture (door,
window, columns, electrical_equipment, site, entourage dropped). `copy_prompt_clicked()`
copies the single selected file; `_init_json_panel()` lists only categories whose file
exists, so the dropdown self-syncs with the folder.

### SOFT-FORM RECIPES in the shared core (user request: sofa-photo puffiness)
- New core section, 7 numbered recipes: 1 capsule-profile Extrusion (puffy pads),
  2 channel-tufting rows (8–12 mm neighbour overlap), 3 rounded-corner slab (`Arc3P`
  fillets, mid at `0.293*r`), 4 dome/sphere Revolution, 5 soft-taper Blend, 6 rolled
  edge (prefer a bolster tube over the fragile Sweep), 7 `Spline` organic silhouette.
- **Bulge depth** dial (offset of `Arc3P` `mid` off its chord; upholstery 15–30 % of the
  smaller cross dimension) + new `_plan.shapes` key (parser-ignored, forces per-part
  bulge reasoning). METHOD / CHECKLIST / HARD FAILURE all cross-reference it ("soft part
  emitted as a sharp box = fidelity bug").
- furniture.md: **Upholstery playbook** (seat tubes = X-Z capsule extruded along Y; back
  tubes = Y-Z side-silhouette capsule extruded along X so tops/fronts come out rounded)
  + new 11-part **channel-tufted sofa worked example** (plinth + 2 rolled arms + 4 seat
  tubes + 4 back tubes). Statically verified: loops closed, points on-plane, exactly one
  anchor, 3-axis bbox overlap ≥ 1 mm on every `attaches_to` pair.
- Every other category got a "Soft-form applications" section.

### Object case libraries (per category)
Each file's short "Part inventory" was replaced by an **Object case library** — ~9–11
common object types with per-part form/recipe, typical mm and anchor/attach chain:
furniture (tables, office chair, bed, wardrobe, bar stool…), casework (base/wall/tall,
island, vanity, reception desk…), plumbing (toilets, tubs, urinal, sinks, faucet…),
lighting (pendant, chandelier, track, floor lamp…), mechanical (AHU, VRF, cassette,
pump, tank, fan…), specialty (fridge, washer, hood, treadmill, lockers…), generic
(planter, bollard, curtain…). Teaching point added: **in-plane rotation is free** (star
bases / radial outlines = plan-profile Extrusions) — only sketch PLANES are limited to
X/Y/Z.

### Stale-warning fix (lighting) — diagonal `Cylinder`
lighting_fixture.md claimed a diagonal `Cylinder` relies on the sweep engine and fails;
the parser actually builds a diagonal `Cylinder` as a **`NewRevolution` about its own
axis** (`FamiGenDialog.py` ≈2900 — the old Sweep fallback was replaced). Warning
softened: diagonal Cylinder acceptable where a slanted straight run is unavoidable
(verify once in Revit); `Sweep` stays banned for arms. The chandelier radial-arm case
depends on this — see pending tests.

## Testing note
Validate generality on a **diverse basket** (chair, cabinet, faucet, door, lamp) —
not just the current lamp case — and record `built/total` + floating-part count per
category. The prompt is only "general" if it holds across categories.

Pending Revit smoke tests (2026-07-05 batch — user runs them, no self-invented results):
- [ ] Paste the `T3Lab_ChannelSofa` worked example (furniture.md) → expect 11/11 parts,
      visible per-tube puffiness and creases between channels.
- [ ] Diagonal `Cylinder` probe: one Cylinder `start [0,0,0]` → `end [200,200,300]`,
      radius 10 → expect a correctly-oriented slanted rod (this decides whether the
      chandelier radial-arm recipe stays or gets pulled).
- [ ] End-to-end on 2–3 categories (photo → Copy Prompt → external AI → paste JSON):
      check the AI follows the case library and fills `_plan.shapes`.
