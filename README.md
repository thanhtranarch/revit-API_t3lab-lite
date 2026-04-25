# T3Lab Lite — pyRevit Extension for Autodesk Revit

A lightweight IronPython/pyRevit extension that adds productivity tools for annotation, export, project management, and AI-assisted automation directly inside Autodesk Revit.

**Author:** Tran Tien Thanh
**Contact:** trantienthanh909@gmail.com
**LinkedIn:** [linkedin.com/in/sunarch7899](https://linkedin.com/in/sunarch7899/)

---

## Requirements

- Autodesk Revit 2020 or later
- [pyRevit](https://github.com/eirannejad/pyRevit) 4.8+

---

## Installation

1. Install [pyRevit](https://github.com/eirannejad/pyRevit)
2. Clone or download this repository
3. Copy the `T3Lab.extension` folder to your pyRevit extensions directory
   - Default path: `%APPDATA%\pyRevit\Extensions\`
4. Reload pyRevit (`pyRevit > Reload`)

---

## Tools

All tools are accessible from the **T3Lab** tab in the Revit ribbon.

### Cloud Panel

| Tool | Description |
|------|-------------|
| **ACC Platform** | Quick link to Autodesk Construction Cloud |
| **B360 Health** | Quick link to BIM 360 platform health status |
| **Bluebeam Health** | Quick link to Bluebeam service health status |

---

### Project Panel

#### Workset

| Tool | Description |
|------|-------------|
| **Workset Manager** | List, rename, and manage user worksets; remove unused worksets via a checklist |
| **Workset Views** | Generate dedicated 3D views per workset for isolation and coordination review |
| **Central File** | Quick access to sync-to-central and central file worksharing workflows |
| **Tile Layout** | 3-step wizard to extract floor boundaries, choose a tile pattern, and place a tiled view arrangement on the active sheet |

#### Family Work

| Tool | Description |
|------|-------------|
| **Load Family** | Browse and load Revit families from local disk; supports category filtering and batch loading |
| **Bulk Family Export** | Scan imported DWG/DXF files for block definitions and export each block as a separate `.rfa` family file |
| **JSON to Family** | Generate fully parametric Revit families from a structured JSON schema inside an open Family Document |

#### Create

| Tool | Description |
|------|-------------|
| **Wall Type Manager** | Browse, filter, and edit wall type properties; supports bulk updates and type duplication |
| **Property Line** | Create US property lines from Lightbox parcel data |
| **Create Plan Views** | Batch-generate individual floor plan views for each room with custom naming and template assignment |

#### Areas

| Tool | Description |
|------|-------------|
| **Room to Area** | Convert room boundaries to area boundaries automatically in the active area plan |
| **Tag Area Opening** | Auto-tag all area openings in the active view |
| **Opening Assign Values** | Map room or area parameter data onto filled region elements for color-filled area diagrams |

#### Other

| Tool | Description |
|------|-------------|
| **Align Positions** | Snap surrounding element distances to clean multiples of 5 or 10 mm relative to a reference Grid, Wall, or Column |

---

### Annotation Panel

#### Annotation Manager

| Tool | Description |
|------|-------------|
| **Annotation Manager** | Unified window with tabs for managing Dimensions and Text Notes — find, delete, and auto-rename types and instances |

#### Grids

| Tool | Description |
|------|-------------|
| **Save Grids** | Save current grid head and tail positions for later restoration |
| **Restore Grids** | Restore selected grid heads and tails to their saved positions |
| **Restore All Grids** | Restore all grid heads and tails to their saved positions |

#### SmartAlign

| Tool | Description |
|------|-------------|
| **Align Left / Center / Right** | Align selected elements to the leftmost, center, or rightmost edge |
| **Align Top / Center / Bottom** | Align selected elements to the top, center, or bottom edge |
| **Distribute Horizontal** | Evenly distribute selected elements with equal horizontal spacing |
| **Distribute Vertical** | Evenly distribute selected elements with equal vertical spacing |

#### Text

| Tool | Description |
|------|-------------|
| **Dim Text** | View and edit prefix, suffix, and value overrides on selected dimension elements |
| **Adjust TextNote** | Batch-edit TextNote content, type, and formatting with search-and-replace support |
| **Upper All Text** | Convert view names, sheet title block parameters, text notes, and dimension overrides to uppercase |

#### Other

| Tool | Description |
|------|-------------|
| **Reset Overrides** | Reset all graphic overrides on selected elements in the active view |

---

### Export Panel

| Tool | Description |
|------|-------------|
| **BatchOut** | Batch export sheets to PDF, DWG, NWD (Navisworks), and IFC formats with sheet filtering, custom naming patterns, revision tracking, and combined PDF support |

---

### AI Connection Panel

| Tool | Description |
|------|-------------|
| **T3Lab Assistant** | Natural language AI assistant — type commands in Vietnamese or English to control Revit tools |
| **Start MCP** | Start the local MCP server for AI-to-Revit communication |
| **Stop MCP** | Stop the running MCP server |
| **Settings** | Configure API keys and AI backend (Claude API or local Ollama) |

The AI assistant supports two backends:
- **Claude API** — requires an Anthropic API key
- **Local LLM** — uses [Ollama](https://ollama.com/) for fully offline inference (recommended models: `qwen2.5`, `llama3.2`, `phi3:mini`)

---

### Support Panel

| Tool | Description |
|------|-------------|
| **Send Feedback** | Write and send feedback or suggestions directly to the T3Lab team by email |

---

## Project Structure

```
T3Lab.extension/
├── T3Lab.tab/               # Ribbon tab with all tools
│   ├── Cloud.panel/
│   ├── Project.panel/
│   ├── Annotation.panel/
│   ├── Export.panel/
│   ├── AI Connection.panel/
│   └── Support.panel/
├── checks/                  # Model quality check scripts
├── commands/                # Standalone command scripts
└── lib/                     # Shared libraries
    ├── GUI/                 # WPF dialogs (XAML + Python)
    ├── Renaming/            # Find & replace base classes
    ├── Selection/           # Element selection utilities
    ├── Snippets/            # Reusable API code snippets
    ├── config/              # Settings management
    ├── core/                # MCP server & tool registry
    └── ui/                  # Button state & settings UI
```

---

## Known Issues

### pyRevit Reload Error

If you encounter an `IOError` when reloading pyRevit (file locking issue):

```powershell
# Run as Administrator
PowerShell -ExecutionPolicy Bypass -File scripts/fix_pyrevit_reload.ps1
```

See [`PYREVIT_RELOAD_FIX.md`](PYREVIT_RELOAD_FIX.md) for details.

---

## License

For other issues, please open an issue on GitHub.
