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
| **T3Lab Assistant** | Natural language AI assistant — control Revit via Vietnamese/English chat, powered by Claude API or Ollama. |
| **Start MCP** | Starts the thread-safe local MCP server with dynamic port allocation (starting at `48884`) and pyRevit routes integration. |
| **Stop MCP** | Stops the running MCP server and deactivates pyRevit routes. |
| **Settings** | Configure API keys, backend options, and copy dynamically-updating Claude Desktop/Cline JSON config snippets. |

---

### Annotation Panel

#### Graphic Stack
- **Auto Dimension**: Consolidated wall/column grid dimensioning (collaborative tool by T3Lab & Dang Quoc Truong).
- **Snap Dimension**: Quick-snap dimensioning tool for model alignments (from Dang Quoc Truong).
- **Grids** (Pulldown): Save Grids, Restore Grids, Restore All Grids.
- **Reset Overrides**: Reset visual and graphic overrides in the active view.

#### SmartAlign Stack
- Smart-alignment buttons: Align Top/Bottom/Left/Right, Align Center H/V, Distribute H/V.

#### Smart Selection (Pulldown)
- Material Select, Quick Element Select, Select Linked, and Select (Category/Family/Type).

#### Text Stack
- **Annotation Manager**: Consolidated Dimension and Text Note type and instance managers (collaborative tool by T3Lab & Dang Quoc Truong).
- **Text & Tagging** (Pulldown): Text Note Type tools.
- **Renumbering**: Renumber elements sequentially.
- **Tag Checker**: Tag checkers and auto-openings tagging.

#### Other
- **DWG Management**: Unified CAD import and link manager with sheet selection capabilities (collaborative tool by T3Lab & Dang Quoc Truong).

---

### Project Panel

- **Auto Join**: Rule-based automatic elements joiner (collaborative tool by T3Lab & Dang Quoc Truong).
- **Room to Area**: Convert room boundaries to area elements (collaborative tool by T3Lab & Dang Quoc Truong).
- **Create** (Stack):
  - **Create Elements** (Pulldown): CAD to Beam, CAD to Wall, CAD to Floor, Point Cloud to Model, Door Threshold, Room to Floor, Property Line, Image to Drafting.
  - **Datum** (Pulldown): Save/Restore Levels and Grids.
- **Element Adjust** (Pulldown): Split elements, Wall Cut Profile, Wall Adjust Base.
- **Family Work** (Stack): Load Family, CAD to Family (DWG block exporter), JSON to Family, Family Management (collaborative tool by T3Lab & Dang Quoc Truong).
- **Workset** (Stack): Workset Manager, Central File (Sync).

---

### Views & Sheets Panel

- **View Manager** (Pulldown): ViewManager Advanced, ViewTemplate, Create Room Plan.
- **Sheet Manager** (Pulldown): SheetManager Advanced, Sheet re-number, Tile Layout.
- **Linked Element Box**: Generate 3D Section Box around selected linked elements.

---

### Data Panel

- **Excel Schedules** (Pulldown): Schedule Export/Import Pro, Schedule Copy.
- **Parameter Tools** (Pulldown): Transfer Parameters, Text to Element, Values to Filled Region.
- **Model Audits** (Pulldown): Room Data Collector, Foundation Volume.
- **BCF Reader**: BIM Collaboration Format Reader.
- **IFC-SG Submission** (Pulldown): Parameter Loader, Auto Assign, Manual Assign, IFCSG Subtype Definer, IFCSG Checker.

---

### Settings Panel

- **Family Audit**: Clean, purge, and size-audit families in the active model.
- **Style Editors** (Pulldown): Line Style Edit, Line Pattern, Hatching.
- **Parameter Config**: ParaManager (advanced parameter configuration).
- **Color Splasher**: Audits elements visually by mapping colors to parameters.
- **Model Health** (Pulldown): ModelChecker, HealthCheck, Warnings, In-Place Models check, Location Manager, Material List.

---

### Support Panel

- **BatchOut**: Bulk PDF, DWG, NWD, and IFC exporter with revision tracking.
- **UI Customizer** (Pulldown): Background Theme, Ribbon Names, Tab Manager.
- **Cloud Links** (Pulldown): Autodesk Forma, Autodesk Health, Bluebeam Status.
- **Auto Work**: Automation recorder & player.
- **Help & Feedback**: Send Feedback, Documentation.

---

## Project Structure

```
t3lab-revit-api/
├── T3Lab.extension/
│   ├── T3Lab.tab/
│   │   ├── AI Connection.panel/
│   │   ├── Annotation.panel/
│   │   ├── ViewsSheets.panel/
│   │   ├── Data.panel/
│   │   ├── Project.panel/
│   │   ├── Settings.panel/
│   │   └── Support.panel/
│   ├── lib/
│   │   ├── GUI/                    # WPF dialogs (XAML + Python classes)
│   │   │   ├── Tools/              # Consolidated window views
│   │   │   └── Resources/          # Shared WPF styles (WPF_styles.xaml)
│   │   ├── Intelligence/           # AI engine: NLU, RAG, Ollama local LLM
│   │   ├── Renaming/               # Renaming engine classes
│   │   ├── Selection/              # Element selection helpers
│   │   ├── Services/               # Exporters & tool discovery
│   │   ├── Snippets/               # 19 reusable Revit API code snippets
│   │   ├── Utils/                  # CAD/family helpers
│   │   ├── config/                 # Configurations, tool registry
│   │   ├── core/                   # MCP Server & ExternalEvents bridge
│   │   └── ui/                     # Button states and general settings
│   ├── checks/                     # Model checker script validations
│   └── commands/                   # Standalone command scripts
├── api/                            # Cloud serverless functions
├── dev/                            # Dev utilities (sync_wpf_styles.py)
├── docs/                           # Documentation
└── scripts/                        # Reload and cache-clearing scripts
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
Single source of truth for all shared button styles (T3Lab Lumina design system). Propagated to every tool XAML with `python3 dev/sync_wpf_styles.py`.

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

UI design standard: `.claude/rules/ui-design-standard.md` (T3Lab Lumina palette — deep slate `#0F172A` + accent blue `#3B82F6`)

---

## Author
**Tran Tien Thanh** — Architect & BIM Developer
- [trantienthanh909@gmail.com](mailto:trantienthanh909@gmail.com)
- [T3Lab.Space](https://t3lab.space)

---
*Empowering BIM with Intelligence.*
