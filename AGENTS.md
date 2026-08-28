# T3Lab pyRevit Extension

pyRevit extension for Revit automation.
Framework: IronPython 2.7 + WPF + Revit API

## Rules
- **UI standard — one source, no exceptions:** `pyRevit UI Design System/T3LAB_UI_STANDARD.md` + `pyRevit UI Design System/T3Lab.Styles.xaml`. Every colour, size, margin and control style comes from `{StaticResource T3.*}`; a tool XAML never defines its own brush, style or hex.
- **Every new tool/script follows `.claude/rules/new-tool-standard.md`** — file layout, the 12 XAML rules, the `script.py` frame, and the pre-commit checklist.
- **The old standards are dead** (2026-08-28): Lumina, Revit-native, Terra v2, Kinetix. Their rule files, the showcase and the Lumina audit/sync scripts were deleted — see git history if you need the old wording.
- UI gate: `python3 dev/audit_t3.py --quiet` (`--legacy` lists migration debt). Static gate: `python3 dev/audit_tools.py --quiet`.
- XAML files go in `T3Lab.extension/lib/GUI/Tools/`
- Python dialog classes stay in `T3Lab.extension/lib/GUI/`
- Keep Revit API logic separate from WPF/UI code
- **Path Portability Rule**: All file paths in agent definitions and documentation must be relative to the repository workspace (e.g. `T3Lab.extension/...`) to ensure portability.

## Agents

Spawn the appropriate agent based on the task:

| Task | Agent |
|------|-------|
| Create or modify WPF windows / XAML | `@ui-agent` |
| Revit API logic, transactions, collectors | `@revit-api-agent` |
| Build a new pushbutton end-to-end | `@tool-builder-agent` |
| Review or test completed code | `@qa-agent` |
| Standardize script.py structure to BatchOut frame | `@script-frame-agent` |
| Daily UI/UX governance cycle (audit · score · standardize · track) | `@ui-governance-agent` |

Agent definitions: `.Codex/agents/`

## Skills

| Skill | Purpose |
|-------|---------|
| `.claude/skills/wpf-pattern.md` | Python WPF window class boilerplate |
| `.claude/skills/xaml-templates.md` | XAML snippets (T3) for all UI components |

## Quick Reference

| Resource | Path |
|----------|------|
| Bilingual VI/EN analysis | `T3Lab.extension/lib/Intelligence/language/` |
| Graph agent layer | `T3Lab.extension/lib/Intelligence/graph/` |
| Assistant architecture doc | `docs/assistant-graph-architecture.md` |
| **UI standard (the only one)** | `pyRevit UI Design System/T3LAB_UI_STANDARD.md` |
| **UI stylesheet — 82 `T3.*` keys** | `pyRevit UI Design System/T3Lab.Styles.xaml` |
| **Rule for every new tool** | `.claude/rules/new-tool-standard.md` |
| **UI governance routine** | `docs/ui-governance/` (start at `README.md`) |
| All XAML files | `T3Lab.extension/lib/GUI/Tools/` |
| Logo asset | `T3Lab.extension/lib/GUI/T3Lab_logo.png` |
| Example XAML (simple) | `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml` |
| Example XAML (wizard nav) | `T3Lab.extension/lib/GUI/Tools/ExportManagerTest.xaml` |
| Python dialogs | `T3Lab.extension/lib/GUI/` (FamilyLoaderDialog.py, etc.) |
| Snippets | `T3Lab.extension/lib/Snippets/` |

## Folder Layout

```
T3Lab.extension/
├── T3Lab.tab/          ← ribbon panels and pushbutton scripts
├── lib/
│   ├── GUI/
│   │   ├── Tools/      ← ALL .xaml files live here
│   │   ├── Resources/  ← shared WPF styles (WPF_styles.xaml)
│   │   ├── forms.py    ← WPF helpers
│   │   ├── WPF_Base.py
│   │   ├── *Dialog.py  ← Python WPF dialog classes
│   │   └── T3Lab_logo.png
│   ├── Snippets/       ← reusable Revit API helpers
│   ├── Renaming/       ← renaming tool library
│   └── ...
├── checks/             ← model checker scripts
└── commands/           ← command scripts
```

## Example Workflow: New Tool

```
You: "Build a new WallType manager tool"
         ↓
Codex reads AGENTS.md → spawns @tool-builder-agent
    ├── @ui-agent    → creates lib/GUI/Tools/WallTypeManager.xaml
    └── @revit-api-agent → implements WallType logic in script.py
         ↓
@qa-agent reviews output
         ↓
Files placed in correct folders ✅
```
