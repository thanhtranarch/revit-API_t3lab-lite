# T3Lab Lite -- Development Guide

All new pyrevit tool windows **MUST** follow the **BatchOut UI design language** (white light theme).

## Project Instructions

Detailed instructions are organized in the `.claude/` folder:

- **`.claude/rules/ui-design-standard.md`** - UI design rules, color palette, and new tool checklist
- **`.claude/docs/wpf-window-templates.md`** - XAML templates for window structure, buttons, DataGrid, info box
- **`.claude/docs/python-wpf-pattern.md`** - Python WPF Window class pattern

## Quick Reference

| Resource | Path |
|----------|------|
| Canonical UI | `T3Lab.extension/T3Lab.tab/Export.panel/BatchOut.pushbutton/` |
| All XAML files | `T3Lab.extension/lib/GUI/Tools/` |
| Shared styles | `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` |
| Logo asset | `T3Lab.extension/lib/GUI/T3Lab_logo.png` |
| Example XAML | `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml` |
| Python dialogs | `T3Lab.extension/lib/GUI/` (FamilyLoaderDialog.py, etc.) |

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
