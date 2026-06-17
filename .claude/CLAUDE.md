# T3Lab pyRevit Extension

pyRevit extension for Revit automation.
Framework: IronPython 2.7 + WPF + Revit API

## Rules
- Always follow `.claude/rules/ui-design-standard.md` for any UI work (T3Lab Lumina design system, utilizing Hanken Grotesk and ultra-thin scrollbars)
- XAML files go in `T3Lab.extension/lib/GUI/Tools/`
- Python dialog classes stay in `T3Lab.extension/lib/GUI/`
- Keep Revit API logic separate from WPF/UI code
- Shared button styles live in `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` and are propagated into every tool XAML with `python3 dev/sync_wpf_styles.py` (`--check` to verify) — never hand-edit the marked style block inside a tool XAML
- **Path Portability Rule**: All file paths in agent definitions and documentation must be relative to the repository workspace (e.g. `T3Lab.extension/...`) to ensure portability.

## UI-Frozen Files (DO NOT MODIFY UI)

The following XAML files are **UI-locked** — their visual design is finalized and must never be altered by any UI sweep, agent, or style sync operation. Logic/script changes are still allowed, but the XAML UI must remain untouched:

| File | Reason |
|------|--------|
| `T3Lab.extension/lib/GUI/Tools/DWGManagement.xaml` | Finalized custom design — UI locked |
| `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml` | Finalized custom design (BatchOut) — UI locked |

**All agents** (`@ui-agent`, `@ui-police-agent`, `@tool-builder-agent`, `@script-frame-agent`) must skip these files entirely during any UI-related task. Do not run `sync_wpf_styles.py` against them. Do not include them in bulk XAML audits.

## Agents

Spawn the appropriate agent based on the task:

| Task | Agent |
|------|-------|
| Create or modify WPF windows / XAML | `@ui-agent` |
| Revit API logic, transactions, collectors | `@revit-api-agent` |
| Build a new pushbutton end-to-end | `@tool-builder-agent` |
| Review or test completed code | `@qa-agent` |
| Standardize script.py structure to BatchOut frame | `@script-frame-agent` |
| Audit & fix ALL XAML files for UI consistency | `@ui-police-agent` |

Agent definitions: `.claude/agents/`

## Skills

| Skill | Purpose |
|-------|---------|
| `.claude/skills/wpf-pattern.md` | Python WPF window class boilerplate |
| `.claude/skills/xaml-templates.md` | XAML snippets for all UI components |

## Quick Reference

| Resource | Path |
|----------|------|
| Canonical UI | `.claude/standard/UIStandardShowcase.xaml` |
| All XAML files | `T3Lab.extension/lib/GUI/Tools/` |
| Shared styles | `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` |
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
Claude reads CLAUDE.md → spawns @tool-builder-agent
    ├── @ui-agent    → creates lib/GUI/Tools/WallTypeManager.xaml
    └── @revit-api-agent → implements WallType logic in script.py
         ↓
@qa-agent reviews output
         ↓
Files placed in correct folders ✅
```
