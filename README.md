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
3. Copy the `T3Lab_Lite.extension` folder to your pyRevit extensions directory
   - Default path: `%APPDATA%\pyRevit\Extensions\`
4. Reload pyRevit (`pyRevit > Reload`)

---

## Tools

All tools are accessible from the **T3Lab Lite** tab in the Revit ribbon.

### Annotation Panel

#### Dimension
| Tool | Description |
|------|-------------|
| **Find Dim** | Find and select dimensions by type or value in the active view |
| **Remove Dim** | Remove selected or all dimensions from the active view |
| **Rename Dim** | Rename dimension types using find and replace |

#### Text
| Tool | Description |
|------|-------------|
| **Dim Text** | Edit dimension text overrides on selected dimensions |
| **Upper Dim Text** | Convert dimension text overrides to uppercase |
| **Save Grids** | Save current grid head and tail positions for later restoration |
| **Restore Grids** | Restore selected grid heads and tails to their saved positions |
| **Restore All Grids** | Restore all grid heads and tails to their saved positions |

#### Text Note
| Tool | Description |
|------|-------------|
| **Find Text** | Find and select text notes by content in the active view |
| **Remove Text** | Remove selected or filtered text notes from the active view |
| **Rename Text** | Rename text note types using find and replace |

#### Other
| Tool | Description |
|------|-------------|
| **Reset Overrides** | Reset all graphic overrides on selected elements in the active view |

---

### Export Panel

| Tool | Description |
|------|-------------|
| **Batch Out** | Batch export sheets to PDF, DWG, DWF, DGN, IFC, NWD and image formats with sheet filtering and bilingual (Vietnamese/English) support |

---

### Project Panel

| Tool | Description |
|------|-------------|
| **Load Family** | Load Revit families from local folders with category-based browsing and search (cloud feature currently disabled) |
| **Para Sync** | Synchronize parameter values between selected elements |
| **Property Line** | Create and manage property lines from survey data |
| **Workset** | Manage and assign worksets to selected elements |

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

## Project Structure

```
T3Lab_Lite.extension/
├── T3Lab_Lite.tab/          # Ribbon tab with all tools
│   ├── Annotation.panel/
│   ├── Export.panel/
│   ├── Project.panel/
│   └── AI Connection.panel/
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
