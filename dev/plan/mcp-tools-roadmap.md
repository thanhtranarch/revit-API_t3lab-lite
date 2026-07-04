# MCP Tools — Expansion Roadmap

> Standalone roadmap for growing the Revit MCP tool surface exposed to AI clients
> (Claude Desktop / Cursor) via `core/server.py`.
> **Not** part of the `gd<N> revitapi` phase tables — do not wire into
> `README.md`'s "Lịch thực hiện".
>
> Source of truth (single file, 3 edit sites per tool):
> - Schema registry: `T3Lab.extension/lib/core/server.py` → `_register_default_tools()`
> - Write-tool set:   `T3Lab.extension/lib/core/server.py` → `_WRITE_TOOLS` frozenset
> - Implementation:   `T3Lab.extension/lib/core/server.py` → `_execute_tool_in_context()`

## Guiding principles
- **Parity with the GUI, not duplication of it.** Each new tool should let the AI
  drive a capability the plugin already ships as a pushbutton (Schedule, Param,
  Split, SheetGen…), reusing the same Snippets/logic where possible — not invent
  new behaviour.
- **Read tools run on the HTTP worker thread; write tools MUST be in `_WRITE_TOOLS`**
  so they marshal onto Revit's main thread via the ExternalEvent. A write tool left
  out of the set fails with "transaction outside API context".
- **Every write tool = Transaction + try/RollBack + explicit guard errors.** Never
  silently mutate/flatten data (see the `split_curve` fix: slice the source curve,
  don't rebuild endpoints).
- **Units:** MCP boundary is **meters**; convert with `M2FT = 3.28084` on the way in,
  `* 0.3048` on the way out. IDs via `eid_value(...)`.
- After adding tools, users must **Reload pyRevit → reopen T3Lab Assistant → restart
  the `t3lab-revit` connector** for `tools/list` to refresh.

## Status — ALL SHIPPED (2026-07-04)
All 3 tiers implemented in `core/server.py` in one pass (schema + `_WRITE_TOOLS`
routing + `_execute_tool_in_context` impl), alongside shared helpers
`_bic_map()`, `_apply_param_value()`, `_parse_color()`. `py_compile` clean;
schema↔impl parity verified for every tool. Read tools (`get_schedule_data`,
`audit_model`) correctly excluded from `_WRITE_TOOLS`. **Pending live smoke test in
Revit** after a pyRevit reload (see the checklist at the bottom). Tier-3 tools
(`create_project_parameter`, `room_to_floor`, `purge_unused`) are best-effort with
version fallbacks and heavy guards — verify first in a scratch model.

## Current coverage (baseline)
~61 tools + `split_curve` (shipped). Strong on: read/query, basic create
(level/wall/grid/room/framing/view/sheet/text), basic modify (move/copy/rotate/
delete/color/tag walls+rooms/rename/param get-set), data store/query, PDF export.
Gaps map cleanly onto the panels below.

---

## 🟢 Tier 1 — High value, compact API, low risk (build first)

- [x] **`bulk_set_parameter`** (W) — parity `ManaPara`.
  Set one parameter across MANY elements matched by category + optional
  name/value filter. Args: `category`, `parameter_name`, `value`,
  `filter_parameter?`, `filter_value?`, `limit?`. Returns per-element
  success/skip counts. *The single highest-leverage AI tool — bulk edits.*
- [x] **`get_schedule_data`** (R) — read a `ViewSchedule` as a JSON table
  (header row + data rows). Args: `schedule_name?` or `schedule_id?`.
  API: `ViewSchedule.GetTableData()` / `TableSectionData`, or export-to-string.
- [x] **`create_schedule`** (W) — create a schedule for a category with a field
  list. Args: `category`, `fields[]`, `name?`. API: `ViewSchedule.CreateSchedule`
  + `ScheduleDefinition.AddField`.
- [x] **`select_elements`** (W) — make `ai_element_filter` actionable: set the
  Revit selection to the filtered set. Args: same filter shape as
  `ai_element_filter`. API: `uidoc.Selection.SetElementIds`.
- [x] **`tag_elements`** (W) — generic tagging by category in the active view
  (generalise `tag_all_walls` / `tag_all_rooms`). Args: `category`, `tag_type?`,
  `leader?`. API: `IndependentTag.Create`.
- [x] **`create_dimension`** (W) — aligned dimension between two references
  (or a grid run). Args: `reference_element_ids[]` or `grid_ids[]`, plus a
  placement line. API: `doc.Create.NewDimension`, `ReferenceArray`.

## 🟡 Tier 2 — View / Sheet & Datum automation

- [x] **`duplicate_view`** (W) — args: `view_id`, `mode` (plain / with_detailing /
  as_dependent), `name?`. API: `View.Duplicate(ViewDuplicateOption)`.
- [x] **`apply_view_template`** (W) — args: `view_id`, `template_name`.
  API: `View.ViewTemplateId`.
- [x] **`create_view_filter`** (W) — rule-based `ParameterFilterElement` + apply to
  a view with overrides/visibility. Args: `name`, `categories[]`, `rules[]`,
  `view_id?`, `hide?`/`color?`.
- [x] **`place_views_on_sheets`** (W) — parity `SheetGen`: auto-lay unplaced views
  onto sheets. Args: `view_ids[]`, `title_block?`, `per_sheet?`.
  API: `Viewport.Create`.
- [x] **`export_dwg`** / **`export_image`** (W) — siblings of `export_sheets_pdf`.
  API: `doc.Export(... DWGExportOptions)`, `doc.ExportImage(ImageExportOptions)`.
- [x] **`split_element`** (W) — split a wall / beam / other `LocationCurve` element
  at a point or ratio, reusing the `split_curve` Clone+MakeBound approach on the
  location curve. Args: `element_id`, `at_ratio?` or `at_point{x,y}`.
- [x] **`join_geometry`** (W) — parity `AutoJoin`: join/unjoin two elements.
  Args: `element_id_a`, `element_id_b`, `unjoin?`.
  API: `JoinGeometryUtils.JoinGeometry`.

## 🔴 Tier 3 — Complex / handle with care

- [x] **`create_project_parameter`** (W) — bind a project/shared parameter to
  categories. Args: `name`, `type`, `categories[]`, `group?`, `instance?`.
  API: shared-parameter file + `doc.ParameterBindings`. *Fiddly; needs a shared
  param file path.*
- [x] **`room_to_floor`** (W) — parity `RoomToFloor`: build a floor from a room's
  boundary. Args: `room_id`, `floor_type?`. API: room boundary loops → `Floor.Create`.
- [x] **`purge_unused`** (W) — remove unused families/types/view templates.
  Args: `dry_run?`, `categories?`. Guard heavily; default `dry_run=True`.
- [x] **`audit_model`** (R) — parity `ModelAuditor`: richer report than
  `get_model_health` (naming, warnings by type, imported CAD, groups, in-place
  families, purgeable count).
- [x] **`create_workset`** (W) — args: `name`. API: `Workset.Create` (workshared
  models only).

---

## How to add one tool (checklist per tool)
1. **Schema** in `_register_default_tools()` — add a `'<name>': {...}` block under
   the matching section comment; keep param descriptions AI-facing and in meters.
2. **Thread routing** — if it opens a `Transaction`, add `'<name>'` to `_WRITE_TOOLS`.
3. **Implementation** — add `elif tool_name == '<name>':` in
   `_execute_tool_in_context()`, before the final `return {'error': 'Tool not
   implemented'}`. Follow the existing pattern: resolve `doc/uidoc`, validate args,
   `Transaction(...) → Start/Commit`, `RollBack` on exception, return a dict with
   `success` + ids (or `error`).
4. **Syntax** — `python -m py_compile T3Lab.extension/lib/core/server.py`.
5. **Live test** — Reload pyRevit, reopen Assistant, restart connector, then call
   the tool from the AI client against a real element.

## Notes / decisions
- Prefer **`Clone()` + `MakeBound(raw_param)`** for any geometry slicing (all curve
  types preserved) over rebuilding endpoints — established by `split_curve`.
- Bulk/write tools should cap element counts (e.g. `limit`, max 200) to keep the
  120s ExternalEvent budget safe.
