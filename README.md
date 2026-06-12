# T3Lab — pyRevit Extension for Autodesk Revit

[![Revit Version](https://img.shields.io/badge/Revit-2020%2B-blue.svg)](https://www.autodesk.com/products/revit/overview)
[![pyRevit](https://img.shields.io/badge/pyRevit-4.8%2B-orange.svg)](https://github.com/eirannejad/pyRevit)
[![IronPython](https://img.shields.io/badge/IronPython-2.7-lightgrey.svg)](https://ironpython.net/)

T3Lab is an advanced BIM Automation and Intelligence framework for Autodesk Revit, built on the **T3Lab Master Architecture System (MAS)**. It bridges traditional BIM workflows with modern AI and batch automation.

---

## Architecture: T3Lab MAS 3.0

Three layers form a self-sustaining ecosystem for architectural intelligence:

| Layer | Description |
|-------|-------------|
| **Intelligence** | T3Lab Assistant with NLU engine, local LLM (Ollama), and RAG processing |
| **Execution** | Ribbon-integrated tools organized by discipline (Annotation, Project, Export) |
| **Data Fabric** | Vercel cloud API for family management and metadata; hybrid local/cloud storage |

---

## Tools by Panel

### AI Connection Panel

| Tool | Description |
|------|-------------|
| **T3Lab Assistant** | Natural language AI assistant — type commands in Vietnamese or English to control Revit tools, powered by Claude API or a local Ollama LLM |
| **Start MCP** | Start the local MCP server for AI-to-Revit communication |
| **Stop MCP** | Stop the running MCP server |
| **Settings** | Configure API keys and AI backend (Claude API or local Ollama) |

Supported backends: **Claude API** (Anthropic key required) · **Local LLM** via [Ollama](https://ollama.com/) — recommended models: `qwen2.5`, `llama3.2`, `phi3:mini`

---

### Annotation Panel

#### Graphic

| Tool | Description |
|------|-------------|
| **Auto Dimension** | Auto-dimension walls, structural/architectural columns, and grids in the current view; supports Plan, Section, Elevation, and Detail views with configurable offsets and style selection |
| **Save Grids** | Save current grid head and tail positions for later restoration |
| **Restore Grids** | Restore selected grid heads and tails to their saved positions |
| **Restore All Grids** | Restore all grid heads and tails to their saved positions in all views |
| **Reset Overrides** | Reset all by-element graphic overrides and linework in the active view |
| **DWG Management** | List, rename, and delete CAD imports and CAD links from a single interface |

#### SmartAlign

| Tool | Description |
|------|-------------|
| **Align Top** | Align selected elements to the topmost edge |
| **Align Center (V)** | Align selected elements to their vertical center |
| **Align Bottom** | Align selected elements to the bottom edge |
| **Align Left** | Align selected elements to the leftmost edge |
| **Align Center (H)** | Align selected elements to their horizontal center |
| **Align Right** | Align selected elements to the rightmost edge |
| **Distribute Horizontal** | Evenly distribute selected elements with equal horizontal spacing |
| **Distribute Vertical** | Evenly distribute selected elements with equal vertical spacing |

#### Text

| Tool | Description |
|------|-------------|
| **Annotation Manager** | Unified window with tabs for finding, deleting, and auto-renaming Dimension and Text Note types and instances |
| **Dim Text** | View and edit prefix, suffix, and value overrides on selected dimension elements |
| **Upper All Text** | Convert view names, sheet title block parameters, text notes, and dimension overrides to uppercase |

---

### Cloud Panel

| Tool | Description |
|------|-------------|
| **ACC Platform** | Quick link to Autodesk Construction Cloud |
| **B360 Health** | Quick link to BIM 360 / ACC service health status |
| **Bluebeam Health** | Quick link to Bluebeam service health status |

---

### Export Panel

| Tool | Description |
|------|-------------|
| **BatchOut** | Batch export sheets to PDF, DWG, NWD (Navisworks), and IFC formats with sheet filtering, custom naming patterns, revision tracking, progress tracking, and combined PDF support |

---

### Project Panel

#### Create

| Tool | Description |
|------|-------------|
| **Point Cloud to Model** | Scan-to-BIM: auto-detect Walls, Floors, Ceilings, Doors, Windows, Columns, Stairs, and Roof from a point cloud scan and create all elements in one click |
| **CAD to Beam** | Create structural beams from imported CAD files — pairs parallel lines to find centerlines and widths, then places beams at the correct level and Z-offset |
| **Room to Floor** | Create architectural or structural floors from selected room boundaries |
| **Door Threshold** | Create threshold floor elements at the base of selected doors with automatic dimension matching |
| **Image to Drafting** | Create a new Drafting View and import an image into it |
| **Property Line** | Create US property lines from Lightbox parcel data |
| **Create Plan Views** | Batch-generate individual floor plan views for each room with custom naming and template assignment |

#### Family Work

| Tool | Description |
|------|-------------|
| **Load Family** | Browse and load Revit families from local disk or cloud library; supports category filtering and batch loading |
| **Bulk Family Export** | Scan imported DWG/DXF files for block definitions and export each block as a separate `.rfa` family file |
| **JSON to Family** | Generate fully parametric Revit families (Extrusion, Sweep, Revolve, Blend, Void) from a structured JSON schema |

#### Workset

| Tool | Description |
|------|-------------|
| **Workset Manager** | List, rename, and manage user worksets; remove unused worksets via a checklist interface |
| **Central File** | Quick access to sync-to-central and central file worksharing workflows |
| **Tile Layout** | 3-step wizard to extract floor boundaries, choose a tile pattern, and place a tiled view arrangement on the active sheet |

#### Areas

| Tool | Description |
|------|-------------|
| **Room to Area** | Convert room boundaries to area boundaries automatically in the active area plan |
| **Tag Area Opening** | Auto-tag all area openings in the active view |
| **Opening Assign Values** | Map room or area parameter data onto filled region elements for color-filled area diagrams |

#### Other

| Tool | Description |
|------|-------------|
| **Auto Join** | Automatically join intersecting elements by configurable category rules; supports saving/loading rule presets |

---

### Review Panel

| Tool | Description |
|------|-------------|
| **Location Manager** | List and adjust element locations in the current view or by level; modeless — stays open while you work |

---

### Support Panel

| Tool | Description |
|------|-------------|
| **Auto Work** | Automate repetitive actions — Quick Click simulates automated clicks at fixed coordinates; Record & Replay records and replays mouse action sequences |
| **Send Feedback** | Write and send feedback or suggestions directly to the T3Lab team by email |

---

## Project Structure

```
t3lab-revit-api/
├── T3Lab.extension/
│   ├── T3Lab.tab/
│   │   ├── AI Connection.panel/
│   │   ├── Annotation.panel/
│   │   ├── Cloud.panel/
│   │   ├── Export.panel/
│   │   ├── Project.panel/
│   │   ├── Review.panel/
│   │   └── Support.panel/
│   ├── lib/
│   │   ├── GUI/                    # WPF dialogs (XAML + Python classes)
│   │   │   ├── Tools/              # All .xaml window files
│   │   │   └── Resources/          # Shared WPF styles (WPF_styles.xaml)
│   │   ├── Intelligence/           # AI engine: NLU, RAG, local LLM, assistant
│   │   ├── Renaming/               # Find & replace base classes
│   │   ├── Selection/              # Element selection utilities
│   │   ├── Services/               # BatchOut executor, tool discovery
│   │   ├── Snippets/               # 19 reusable Revit API code patterns
│   │   ├── Utils/                  # CAD/family conversion helpers
│   │   ├── config/                 # Settings, tool registry JSON
│   │   ├── core/                   # MCP server integration
│   │   └── ui/                     # Button state & settings UI
│   ├── checks/                     # Model quality check scripts
│   └── commands/                   # Standalone command scripts
├── api/                            # Vercel serverless functions (cloud families)
├── dev/                            # Dev utilities (sync_wpf_styles.py)
├── docs/                           # API learning guide, cloud loader docs
├── scripts/                        # Cache-clearing and pyRevit reload helpers
└── config/                         # Extension configuration (extensions.json)
```

---

## Library Highlights

### `lib/Intelligence/`
- `t3lab_assistant.py` — main AI assistant engine
- `nlu_engine.py` — Natural Language Understanding for Vietnamese / English
- `rag_processor.py` — Retrieval-Augmented Generation for Revit API context
- `local_llm.py` — Offline LLM support via Ollama

### `lib/Snippets/`
19 reusable IronPython patterns covering: annotations, bounding boxes, context managers, unit conversion, element manipulation, Excel integration, filtered element collectors, groups, lines, graphics overrides, revisions, selection, sheets, text, views, and more.

### `lib/GUI/Resources/WPF_styles.xaml`
Single source of truth for all shared button styles (T3Lab Terra design system). Propagated to every tool XAML with `python3 dev/sync_wpf_styles.py`.

### `checks/`
| Script | Description |
|--------|-------------|
| `modelchecker_check.py` | Comprehensive model quality checks |
| `modelchecker_Warnings_check.py` | Warning-specific validation |
| `refplanes_check.py` | Reference plane validation |
| `schedules_not_on_sheet_check.py` | Identifies schedules not placed on sheets |

---

## Setup & Installation

1. Clone this repository into your pyRevit extensions folder:
   ```
   %APPDATA%\pyRevit\Extensions\T3Lab.extension
   ```
2. Ensure **pyRevit 4.8+** is installed.
3. Reload pyRevit — the **T3Lab** tab will appear in the Revit ribbon.
4. Configure AI settings under **AI Connection → Settings**.

---

## Development

| Script | Purpose |
|--------|---------|
| `dev/sync_wpf_styles.py` | Propagate shared button styles to all tool XAML files |
| `dev/sync_wpf_styles.py --check` | Verify all tool XAML files match the master styles |
| `scripts/clear_pyrevit_cache.ps1` | Clear pyRevit compiled cache |
| `scripts/fix_pyrevit_reload.ps1` | Fix pyRevit reload issues |

UI design standard: `.claude/rules/ui-design-standard.md` (T3Lab Terra palette — deep teal `#0F766E` + amber `#F59E0B`)

---

## Author
**Tran Tien Thanh** — Architect & BIM Developer
- [trantienthanh909@gmail.com](mailto:trantienthanh909@gmail.com)
- [T3Lab.Space](https://t3lab.space)

---
*Empowering BIM with Intelligence.*
