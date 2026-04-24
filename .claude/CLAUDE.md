# T3Lab pyRevit Extension

pyRevit extension for Revit automation.
Framework: IronPython 2.7 + WPF + Revit API

## Rules
- Always follow `.claude/rules/ui-design-standard.md` for any UI work
- XAML files go in `T3Lab.extension/lib/GUI/Tools/`
- Python dialog classes stay in `T3Lab.extension/lib/GUI/`
- Keep Revit API logic separate from WPF/UI code

## Agents

Spawn the appropriate agent based on the task:

| Task | Agent |
|------|-------|
| Create or modify WPF windows / XAML | `@ui-agent` |
| Revit API logic, transactions, collectors | `@revit-api-agent` |
| Build a new pushbutton end-to-end | `@tool-builder-agent` |
| Review or test completed code | `@qa-agent` |

Agent definitions: `.claude/agents/`

## Skills

| Skill | Purpose |
|-------|---------|
| `.claude/skills/wpf-pattern.md` | Python WPF window class boilerplate |
| `.claude/skills/xaml-templates.md` | XAML snippets for all UI components |

## Quick Reference

| Resource | Path |
|----------|------|
| Canonical UI | `T3Lab.extension/T3Lab.tab/Export.panel/BatchOut.pushbutton/` |
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
