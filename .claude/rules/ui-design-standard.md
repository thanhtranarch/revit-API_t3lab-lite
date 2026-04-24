# UI Design Standard

All new pyrevit tool windows **MUST** follow the **BatchOut UI design language** (white light theme).

## Design Reference Files
- **Canonical UI**: `T3Lab.extension/T3Lab.tab/Export.panel/BatchOut.pushbutton/`
- **Shared styles**: `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml`
- **Logo asset**: `T3Lab.extension/lib/GUI/T3Lab_logo.png`
- **All XAML files**: `T3Lab.extension/lib/GUI/Tools/`
- **Example XAML**: `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml`

## Color Palette

| Token        | Hex       | Usage                                   |
|-------------|-----------|------------------------------------------|
| Primary blue | `#3498DB` | Primary buttons, T3Lab brand, accents    |
| Hover blue   | `#2980B9` | Primary button hover                     |
| Dark text    | `#2C3E50` | Headings, labels, main text              |
| Gray text    | `#7F8C8D` | Secondary text, subtitles, icons         |
| Border       | `#BDC3C7` | Input borders, dividers, separators      |
| Light bg     | `#ECF0F1` | Secondary buttons, DataGrid headers      |
| Hover light  | `#D5DBDB` | Secondary button hover                   |
| Row hover    | `#EBF5FB` | DataGrid row hover                       |
| Row select   | `#D6EAF8` | DataGrid selected row                    |
| Info bg      | `#E8F4F8` | Tip / info boxes background              |
| Danger red   | `#E74C3C` | Delete/destructive buttons               |
| Danger hover | `#C0392B` | Danger button hover                      |
| Success green| `#27AE60` | Apply/confirm action buttons             |
| White        | `White`   | Window background, cards, inputs         |

## Checklist for New Tools

When creating a new pushbutton with a WPF UI:

- [ ] `Window.Background="White"`, `ResizeMode="CanResizeWithGrip"`
- [ ] `WindowChrome` with `CaptionHeight="64"`, `UseAeroCaptionButtons="False"`
- [ ] Title bar: 64px, white, T3Lab logo + brand name + tool name + subtitle
- [ ] Minimize / Maximize / Close chrome buttons with correct styles
- [ ] All button styles defined (`PrimaryButton`, `SecondaryButton`, `DangerButton`, `SuccessButton`)
- [ ] DataGrid with `#ECF0F1` headers, row hover `#EBF5FB`, selected `#D6EAF8`
- [ ] Status bar: `#FAFAFA` background, `#7F8C8D` text
- [ ] Font: `Segoe UI` throughout
- [ ] `_load_logo()` called in `__init__`
- [ ] `minimize_button_clicked`, `maximize_button_clicked`, `close_button_clicked` implemented

## Template References

See `.claude/docs/` for full XAML and Python templates:
- `wpf-window-templates.md` - Window structure, button styles, DataGrid, info box XAML
- `python-wpf-pattern.md` - Python WPF Window class pattern (includes SCRIPT_DIR/EXT_DIR/XAML_FILE constants)
