<<<<<<< HEAD
# T3Lab — pyRevit Extension for Autodesk Revit
=======
# T3Lab Revit API
>>>>>>> 5a2304818d72c88bc612843bab9f21796296fe70

[![Revit Version](https://img.shields.io/badge/Revit-2020%2B-blue.svg)](https://www.autodesk.com/products/revit/overview)
[![pyRevit](https://img.shields.io/badge/pyRevit-4.8%2B-orange.svg)](https://github.com/eirannejad/pyRevit)
[![Architecture](https://img.shields.io/badge/Architecture-T3Lab%20MAS%203.0-green.svg)](#architecture)

T3Lab Revit API is an advanced BIM Automation and Intelligence framework for Autodesk Revit. Built upon the **T3Lab Master Architecture System (MAS)**, it bridges the gap between traditional BIM workflows and modern Artificial Intelligence.

---

## Architecture: T3Lab MAS 3.0

The framework is organized into three distinct layers, creating a self-sustaining ecosystem for architectural intelligence:

### 1. Intelligence Layer (Agent-Core)
*   **T3Lab Assistant**: A natural language bridge allowing users to query and command Revit using Vietnamese or English.
*   **Local Intelligence**: Support for offline LLMs (via Ollama) ensuring data privacy and high-speed local inference.

### 2. Execution Layer (Tool-Stack)
*   **Production Tools**: Specialized panels for Annotation, Project Management, and Data Export.
*   **Discipline Modules**: Advanced logic for Architecture, Structure, and Coordination.
*   **Automation Wizards**: Multi-step wizards (like BatchOut) that handle complex task sequences with minimal user input.

### 3. Data Fabric Layer (Cloud-Connect)
*   **Cloud API**: Vercel-hosted backend for distributed family loading and metadata management.
*   **Hybrid Storage**: Seamless switching between local library files and cloud-hosted BIM content.

---

## Key Modules

### AI Connection & Control
*   **T3Lab Assistant**: Context-aware AI that understands Revit terminology and executes complex commands.

### BatchOut: The Export Master
*   A unified export engine for PDF, DWG, NWD, and IFC.
*   Features custom naming patterns, revision-aware filtering, and automatic sheet set generation.

### Project & Annotation Intelligence
*   **Workset Management**: Automated isolation and coordination views.
*   **SmartAlign**: Geometry-aware alignment and distribution logic.
*   **Family Synthesis**: Generate parametric families from structured JSON data.

---

## The Evolution Loop

<<<<<<< HEAD
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
| **Load Family** | Browse and load Revit families from local disk or **Cloud (Vercel)**; supports category filtering, batch loading, and auto-download |
| **Bulk Family Export** | Scan imported DWG/DXF files for block definitions and export each block as a separate `.rfa` family file |
| **JSON to Family** | Generate fully parametric Revit families from a structured JSON schema inside an open Family Document |

#### Create

| Tool | Description |
|------|-------------|
| **CAD to Beam** | Create structural beams from CAD lines; features **AI-assisted dimension detection** from nearby Revit TextNotes |
| **Property Line** | Create US property lines from Lightbox parcel data |
| **Create Plan Views** | Batch-generate individual floor plan views for each room with custom naming and template assignment |

#### Areas

| Tool | Description |
|------|-------------|
| **Room to Area** | Convert room boundaries to area boundaries automatically in the active area plan |
| **Tag Area Opening** | Auto-tag all area openings in the active view |
| **Opening Assign Values** | Map room or area parameter data onto filled region elements for color-filled area diagrams |

#### CAD

| Tool | Description |
|------|-------------|
| **Revit Beam From CAD** | Automatically create structural beams by pairing parallel lines in a selected CAD layer to find centerlines and widths |

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
=======
T3Lab Revit API is designed for **Self-Evolution**. By utilizing specialized agents (QA-Agent, Tool-Builder, UI-Agent), the framework allows for:
1.  **Automated Error Detection**: Identifying BIM inconsistencies via the `checks` module.
2.  **Rapid Tool Iteration**: Building new pushbutton scripts through the `tool-builder` agent.
3.  **Knowledge Accumulation**: Storing project-specific wisdom in a local knowledge graph.
>>>>>>> 5a2304818d72c88bc612843bab9f21796296fe70

---

## Project Structure

<<<<<<< HEAD
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
├── lib/                     # Shared libraries
    ├── GUI/                 # WPF dialogs (XAML + Python)
    ├── Renaming/            # Find & replace base classes
    ├── Selection/           # Element selection utilities
    ├── Snippets/            # Reusable API code snippets
    ├── config/              # Settings management
    ├── core/                # MCP server & tool registry
    └── ui/                  # Button state & settings UI
=======
```bash
t3lab-revit-api/
├── T3Lab.extension/          # Main pyRevit Extension
│   ├── T3Lab.tab/            # Ribbon Panels (AI, Annotation, Project, Export)
│   ├── lib/                  # Framework Core (GUI, NLU, RAG, Utils)
│   ├── checks/               # Quality Assurance Scripts
│   └── commands/             # Standalone Automation Commands
├── api/                      # Vercel Serverless Functions
├── scripts/                  # Maintenance & Environment Setup
├── requirements.txt          # Framework Dependencies
└── vercel.json               # Cloud Infrastructure Config
>>>>>>> 5a2304818d72c88bc612843bab9f21796296fe70
```

---

## Setup & Installation

1.  Clone this repository to: `%APPDATA%\pyRevit\Extensions\T3Lab.extension`
2.  Ensure **pyRevit 4.8+** is installed and linked.
3.  Reload pyRevit to initialize the **T3Lab** tab.
4.  Configure AI settings in the **AI Connection > Settings** panel.

---

## Author
**Tran Tien Thanh**
Architect & BIM Developer
- [trantienthanh909@gmail.com](mailto:trantienthanh909@gmail.com)
- [T3Lab.Space](https://t3lab.space)

---
*Empowering BIM with Intelligence.*
