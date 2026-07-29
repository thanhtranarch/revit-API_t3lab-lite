# T3Lab — pyRevit Extension for Autodesk Revit

[![Revit Version](https://img.shields.io/badge/Revit-2020%2B-blue.svg)](https://www.autodesk.com/products/revit/overview)
[![pyRevit](https://img.shields.io/badge/pyRevit-4.8%2B-orange.svg)](https://github.com/eirannejad/pyRevit)
[![IronPython](https://img.shields.io/badge/IronPython-2.7-lightgrey.svg)](https://ironpython.net/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

T3Lab is a BIM automation and intelligence framework for Autodesk Revit, built on the **T3Lab Master Architecture System (MAS)**. It bridges traditional BIM workflows with AI assistance and batch automation.

Tools are consolidated rather than scattered: instead of dozens of single-purpose buttons, related workflows are merged into unified `Mana*` managers (ManaAnno, ManaSelect, ManaSheets, …) — one window, tabbed modes, shared state.

---

## Architecture: T3Lab MAS 3.0

Three layers form a self-sustaining ecosystem for architectural intelligence:

| Layer | Description |
|-------|-------------|
| **Intelligence** | T3Lab Assistant with NLU engine, multi-provider LLM routing (Claude, OpenAI, DeepSeek, Ollama, LM Studio) and RAG over Revit API knowledge |
| **Execution** | Ribbon-integrated tools organized by discipline across 7 panels |
| **Data Fabric** | MCP server bridge for external agents; Vercel cloud API for family metadata; hybrid local/cloud storage |

---

## Ribbon: T3Lab Tab

The tab exposes **7 panels**. Buttons marked *(DQT)* were developed in collaboration with Dang Quoc Truong.

### Standard

| Tool | Description |
|------|-------------|
| **Auto Work** | Automation recorder & player — quick click at a fixed coordinate, or record and replay a full mouse sequence with timing. |
| **UI Showcase** | Reference window for the T3Lab Lumina design standard (palette, typography, buttons, inputs). |

### Annotation & Select

| Tool | Description |
|------|-------------|
| **Mana Anno** | Unified Find / Remove / Rename manager for Dimensions and Text Notes. *(DQT)* |
| **Auto Dimension** | Automatic dimension chains for walls, columns, doors, lifts and grids in the active or a chosen view. |
| **Mana DWG** | CAD import and CAD link manager — list, rename and delete DWG imports/links. *(DQT)* |
| **Mana Select** | Consolidated selection manager: Quick Select by parameter/text, Select Similar by type/family/category, and linked-element selection. |

### Modeling & Datum

| Tool | Description |
|------|-------------|
| **CAD to BIM** (pulldown) | **CAD to Elements** (map DWG layers → Walls / Floors / Beams), **Point Cloud to Model** (Scan-to-BIM wizard detecting walls, floors, ceilings, doors, windows, columns, stairs, roofs), **Room To Floor**, **Door Threshold**, **Image to Drafting**, **Text to Element**. |
| **Property Line** | Build a closed property-line loop from Lightbox parcel data with computed bearings and distances. |
| **Tile Layout** | 3-step wizard: extract floor boundaries, pick a tile pattern per floor, generate and place a tiled layout. |
| **Element Adjust** (pulldown) | **Auto Join** (rule-based joining, Shift+Click for defaults) *(DQT)*, **Split Elements** at levels, **Wall Cut Profile** from linked-model intersections, **Auto Adj Base Offset**. |
| **FamiGen** | Family generator — from CAD blocks (DWG → .rfa), from a JSON schema, or from built-in batch presets. |
| **Mana Fami** | Family manager — browse by category, search/filter, and load families from disk. *(DQT)* |

### Views & Sheets

| Tool | Description |
|------|-------------|
| **BatchOut** | Batch export sheets to PDF, DWG, NWD and IFC with revision tracking and advanced options. |
| **Mana Views** | View manager — rename views, batch rename, apply and update view templates. |
| **Mana Sheets** | Sheet manager — Excel sync, view placement, sheet sets, re-numbering. |
| **SheetGen** | Generate floor-plan views from a room list via a WPF selection interface. |

### Data & IFC-SG

| Tool | Description |
|------|-------------|
| **Mana Sched** | Schedule manager — export to Excel with formatting, import values back, duplicate schedules. |
| **Mana Para** | Parameter manager — transfer values by rule, Text-to-Element assignment, values-to-filled-region. |
| **Mana Contains** | Spatial containment — find elements inside Rooms/Areas/Spaces/Zones/Masses/Scope Boxes, push container values down or aggregate element data up. |
| **BCF Reader** | Modeless BCF issue browser (IFC Delta Viewer exports) — click an issue to navigate the view. |
| **Foundation Volume** | Write computed volume of Structural Foundations into a chosen shared parameter. |
| **IFC-SG Suite** | Subtype Assigner (Excel mapping → IFC Export Class & Predefined Type) + Compliance Checker against CORENET X rules. |

### Standards & Settings

| Tool | Description |
|------|-------------|
| **Mana Styles** | Fill patterns, line styles, line patterns and visual colour-splashing in one window. |
| **Mana Workset** | Enable worksharing, create/delete/purge worksets, generate workset view filters. |
| **Mana Loca** | Modeless element location editor — read and edit XYZ in a grid, commit in one transaction. |
| **Model Auditor** | Consolidated model health check, warnings, in-place models and material audit. |

### Support

| Tool | Description |
|------|-------------|
| **T3Lab Assistant** | Natural-language AI assistant — drive T3Lab tools via Vietnamese/English chat. |
| **PDF Import** | Import PDF pages into selected Revit views sequentially. |
| **Assistant Tools** (stack) | **MCP Control** (start/stop the MCP server, connection settings), **LLMs Setting** (provider, model, API key), **Feedback**. |
| **UI Theme & Tabs** (stack) | **BG Theme** (HSV picker with eyedropper, gradient 3D backgrounds, Light/Dark UI for Revit 2024+), **Mana Tabs** (hide/show ribbon tabs), **Ribbon Names** (shorten/restore tab names). |
| **Cloud Links** (stack) | Autodesk Forma, Autodesk Health, Bluebeam Status. |

---

## Project Structure

```
t3lab-revit-api/
├── T3Lab.extension/
│   ├── T3Lab.tab/
│   │   ├── Standard.panel/
│   │   ├── Annotation & Select.panel/
│   │   ├── Modeling & Datum.panel/
│   │   ├── Views & Sheets.panel/
│   │   ├── Data & IFC-SG.panel/
│   │   ├── Standards & Settings.panel/
│   │   └── Support.panel/
│   ├── lib/
│   │   ├── GUI/                    # WPF dialogs (XAML + Python classes)
│   │   │   ├── Tools/              # 50+ tool window views (.xaml)
│   │   │   └── Resources/          # Shared WPF styles (WPF_styles.xaml)
│   │   ├── Intelligence/           # AI engine: NLU, routing, RAG, LLM providers, skills
│   │   ├── Services/               # Exporters, MCP service, spell checker, tool discovery
│   │   ├── Selection/              # Element selection helpers
│   │   ├── Renaming/               # Renaming engine classes
│   │   ├── Snippets/               # 19 reusable Revit API code snippets
│   │   ├── Utils/                  # CAD/family helpers
│   │   ├── config/                 # Settings, project store, user profile
│   │   ├── core/                   # MCP server, ExternalEvent bridge, registry, paths
│   │   └── ui/                     # Button states, settings dialog
│   ├── checks/                     # Model checker script validations
│   ├── commands/                   # Standalone command scripts
│   ├── hooks/                      # pyRevit event hooks
│   └── startup.py                  # Extension startup
├── api/                            # Cloud serverless functions (family metadata)
├── dev/                            # Dev utilities, audits, plans, tests
├── docs/                           # Documentation
└── scripts/                        # Reload, cache-clearing, icon generation
```

---

## Library Highlights

### `lib/Intelligence/`
- `t3lab_assistant.py` — main AI assistant engine
- `t3lab_agent.py` / `agent_loop.py` — agentic tool-calling loop
- `nlu_engine.py` — Natural Language Understanding for Vietnamese / English
- `routing.py` / `llm_router.py` — testable routing ladder and provider selection
- `claude_provider.py`, `openai_provider.py`, `deepseek_provider.py`,
  `ollama_provider.py`, `lmstudio_provider.py` — pluggable LLM backends
  ([setting flow](docs/assistant-llms-setting-flow.md))
- `rag_processor.py` — Retrieval-Augmented Generation over Revit API context
- `skills_engine.py` — instruction packs that activate on a request
- `skill_installer.py` — installs Claude-format skills from a GitHub repo link
  ([docs](docs/assistant-skills-from-github.md))

### `lib/core/`
`server.py` and `bridge.py` implement the thread-safe local MCP server (dynamic port allocation from `48884`) and the ExternalEvent bridge that marshals agent calls onto the Revit API thread.

### `lib/Snippets/`
19 reusable IronPython patterns covering annotations, bounding boxes, context managers, unit conversion, element manipulation, Excel integration, filtered element collectors, filters, groups, lines, graphic overrides, revisions, selection, sheets, text and views.

### `lib/GUI/Resources/WPF_styles.xaml`
Single source of truth for all shared button styles (T3Lab Lumina design system). Propagated to every tool XAML with `python3 dev/sync_wpf_styles.py`.

### `checks/`
| Script | Description |
|--------|-------------|
| `modelchecker_check.py` | Comprehensive model quality checks |
| `modelchecker_Warnings_check.py` | Warning-specific validation |
| `refplanes_check.py` | Reference plane validation |
| `schedules_not_on_sheet_check.py` | Identifies schedules not placed on sheets |
| `badgeometry_check.py` | Detects problematic element geometry |

---

## Setup & Installation

1. Clone this repository into your pyRevit extensions folder:
   ```
   %APPDATA%\pyRevit\Extensions\T3Lab.extension
   ```
2. Ensure **pyRevit 4.8+** is installed.
3. Reload pyRevit — the **T3Lab** tab will appear in the Revit ribbon.
4. Configure the LLM provider and API key under **Support → Assistant Tools → LLMs Setting**.

---

## Development

| Command | Purpose |
|---------|---------|
| `python3 dev/sync_wpf_styles.py` | Propagate shared button styles to all tool XAML files |
| `python3 dev/sync_wpf_styles.py --check` | Verify all tool XAML files match the master styles |
| `python3 dev/audit_tools.py --quiet` | Audit pushbutton bundles and script structure |
| `python3 dev/audit_ui.py --quiet` | Audit XAML files against the Lumina UI standard |
| `scripts/clear_pyrevit_cache.ps1` | Clear pyRevit compiled cache |
| `scripts/fix_pyrevit_reload.ps1` | Fix pyRevit reload issues |

- UI design standard: `.claude/rules/ui-design-standard.md` (T3Lab Lumina — deep slate `#0F172A` + accent blue `#3B82F6`)
- Contributor / agent guide: [`AGENTS.md`](AGENTS.md) · Design notes: [`DESIGN.md`](DESIGN.md)
- Debug & QA plans: `dev/plan/`

---

## License

Released under the [MIT License](LICENSE) — © 2026 Tran Tien Thanh (T3Lab).

Autodesk® and Revit® are registered trademarks of Autodesk, Inc. This project is not affiliated with or endorsed by Autodesk.

---

## Author
**Tran Tien Thanh** — Architect & BIM Developer
- [trantienthanh909@gmail.com](mailto:trantienthanh909@gmail.com)
- [T3Lab.Space](https://t3lab.space)

---
*Empowering BIM with Intelligence.*
