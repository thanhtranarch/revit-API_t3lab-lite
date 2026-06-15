---
name: ui-police-agent
description: UI consistency enforcer for T3Lab pyRevit tools. Audits ALL XAML files in lib/GUI/Tools/ against the T3Lab Lumina design standard, produces a violation report grouped by severity, then fixes every violation automatically. Use this agent when you want a full codebase-wide UI audit or when onboarding new tools that may not have been reviewed yet.
tools: Read, Edit, Write, Glob, Grep
---

# UI Police Agent — Codebase-Wide XAML Auditor & Fixer

## Mission

Scan every `.xaml` file in `T3Lab.extension/lib/GUI/Tools/`, compare it against the T3Lab Lumina design standard, report all violations with file + line number, then patch each file so it complies.

**Authoritative standard**: `.claude/rules/ui-design-standard.md`
**Canonical reference**: `.claude/standard/UIStandardShowcase.xaml`

---

## Step 1 — Discover All Files

Glob `T3Lab.extension/lib/GUI/Tools/*.xaml`.

Immediately classify each file by variant:
- **Variant A** — root is `<Window …>` → apply all rules below.
- **Variant B** — root is `<Grid …>` → apply only the Copyright rule. Skip all Window/WindowChrome/TitleBar/StatusBar checks (these 3 files are intentionally borderless: `FamilyLoader.xaml`, `FamilyLoaderCloud.xaml`, `ParameterSelector.xaml`).

---

## Step 2 — Audit Rules (Variant A only unless noted)

Read each file fully. Check every rule below. Record every violation as:

```
[SEVERITY] FileName.xaml  — rule description  (line N if known)
```

Severity levels:
- **CRITICAL** — will cause a runtime crash or render completely wrong
- **MAJOR** — visible design inconsistency, brand-breaking
- **MINOR** — small deviation that degrades polish

### 2.1 Window Shell
| Check | Expected | Severity if wrong |
|-------|----------|-------------------|
| `Background` on `<Window>` | `"White"` | MAJOR |
| `ResizeMode` on `<Window>` | `"CanResizeWithGrip"` | MINOR |
| `FontFamily` on `<Window>` | `"Hanken Grotesk"` (or fallback `"Inter"`) | MAJOR |

### 2.2 WindowChrome
The `<WindowChrome>` **must** use the multi-line form with **all five** attributes present:
- `CaptionHeight="64"`
- `ResizeBorderThickness="5"`
- `GlassFrameThickness="0"`
- `CornerRadius="8"`
- `UseAeroCaptionButtons="False"`

If any attribute is missing → **MAJOR**.
If written as a single-line self-closing tag (which silently drops CornerRadius and GlassFrameThickness) → **MAJOR**.

### 2.3 Title Bar
| Check | Expected | Severity |
|-------|----------|----------|
| Row 0 height | `64` (or `Height="64"` on RowDefinition) | MAJOR |
| Background | `White` | MINOR |
| "T3Lab" brand label font size | `11` | MINOR |
| "T3Lab" brand label foreground | `#0F172A` | MAJOR |
| "T3Lab" brand label font weight | `Bold` | MINOR |
| Tool name font size | `18` | MINOR |
| Tool name foreground | `#0F172A` | MAJOR |
| Separator thickness | `1` px, color `#E2E8F0` | MINOR |
| Subtitle font size | `10` | MINOR |
| Subtitle font style | `Italic` | MINOR |
| Subtitle foreground | `#64748B` | MINOR |
| Bottom border of title row | `BorderThickness="0,0,0,1"`, `BorderBrush="#E2E8F0"` | MAJOR |

### 2.4 Window Control Buttons (Min / Max / Close)

Each of the three buttons must use an **embedded `<TextBlock>` child** — never `Content="…"`.

| Button | Correct glyph | Wrong glyph (reject) | Style key |
|--------|--------------|----------------------|-----------|
| Minimize | `&#xE921;` | `&#x2212;` (minus) | `WinCtrlButton` |
| Maximize | `&#xE922;` | `&#x25A1;` (square) | `WinCtrlButton` |
| Close    | `&#xE8BB;` | literal `X`         | `CloseButton` |

On each TextBlock child:
- `FontFamily="Segoe MDL2 Assets"` — CRITICAL if missing/wrong
- `FontSize="10"` — MINOR
- NO `Foreground` attribute (must inherit from button style) — MAJOR if present

`WindowChrome.IsHitTestVisibleInChrome="True"` must be on the containing `<StackPanel>` — CRITICAL.

Click handlers: `minimize_button_clicked`, `maximize_button_clicked`, `close_button_clicked` — CRITICAL if missing.

### 2.5 Status Bar (last content row)
| Check | Expected | Severity |
|-------|----------|----------|
| Background | `#F8FAFC` | MAJOR |
| BorderBrush | `#E2E8F0` | MAJOR |
| BorderThickness | `"0,1,0,0"` | MAJOR |
| Padding | `"14,8"` | MINOR |

### 2.6 Copyright TextBlock (Variant A and Variant B)
Exactly **one** copyright TextBlock per file. It must be the **last child** of the root `<Grid>`. The snippet must match verbatim:

```xml
<!-- Copyright added automatically -->
<TextBlock Text="© Copyright by T3Lab" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,14,8" Foreground="#F59E0B" FontSize="11" IsHitTestVisible="False" Panel.ZIndex="999"/>
```

Violations:
- Missing entirely → CRITICAL
- `&#169;` instead of literal `©` → MAJOR
- Has `Grid.Row` / `Grid.RowSpan` attribute → MAJOR
- Embedded inside a Border/StackPanel (not a direct child of root Grid) → MAJOR
- Two or more copyright blocks in the same file → MAJOR

### 2.7 Font Consistency
Search the whole file for:
- `FontFamily="Manrope"` → MAJOR (replace with `Hanken Grotesk` or fallback `Inter`)
- `FontFamily="Segoe UI"` → MAJOR (replace with `Hanken Grotesk` or fallback `Inter`; exception: `Segoe MDL2 Assets` on glyph TextBlocks is allowed)

### 2.8 Forbidden Elements
- `<Image Source="…T3Lab_logo.png"/>` → CRITICAL (logo was removed; delete the element and any surrounding container added only to hold the logo)

### 2.9 Button Style Usage
If any `<Button>` has an inline `Background`, `Foreground`, or color attribute that duplicates a named style → MAJOR. Buttons must reference `PrimaryButton`, `SecondaryButton`, `SuccessButton`, `DangerButton`, `AccentButton`, `WinCtrlButton`, or `CloseButton`.

If a named style is referenced but **not defined** in `Window.Resources` → CRITICAL.

### 2.10 DataGrid Style (if file contains a DataGrid)
| Check | Expected | Severity |
|-------|----------|----------|
| `Background` | `"White"` | MINOR |
| `BorderBrush` | `"#E2E8F0"` | MINOR |
| `AlternatingRowBackground` | `"#F8FAFC"` | MINOR |
| `FontFamily` | `"Hanken Grotesk"` (or fallback `"Inter"`) | MAJOR |
| Header `Background` | `"#F8FAFC"` | MINOR |
| Header `Foreground` | `"#0F172A"` | MINOR |

---

## Step 3 — Report

After auditing all files, produce a structured report.

---

## Step 4 — Fix All Violations

After printing the report, fix every violation in every file using the Edit tool.

### Fixing guidelines

**WindowChrome — single-line → multi-line**
Replace with the full multi-line form including `GlassFrameThickness="0"` and `CornerRadius="8"`.

**Wrong minimize glyph**
Replace `Content="&#x2212;"` or `Content="-"` on the minimize button with the canonical TextBlock child form.

**Wrong maximize glyph**
Replace `Content="&#x25A1;"` or `Content="□"` with the canonical TextBlock child form.

**Font: Manrope / Segoe UI**
`replace_all=true` for `FontFamily="Manrope"` → `FontFamily="Hanken Grotesk"` (or fallback `Inter`) and `FontFamily="Segoe UI"` → `FontFamily="Hanken Grotesk"` (or fallback `Inter`). Do NOT touch `Segoe MDL2 Assets`.

**Missing/wrong copyright**
If missing: add the verbatim snippet as the last child of root `<Grid>`.
If embedded in a Border: extract it out to the root Grid level and remove the `Grid.Row` attribute.
If using `&#169;`: replace with literal `©`.

**Logo image**
Delete the entire `<Image Source="…T3Lab_logo.png" …/>` element.

**Status bar background**
Replace with `Background="#F8FAFC"`.

**Style not defined in Window.Resources**
If a named button style key is referenced but not in `Window.Resources`, add the complete style definition (copy verbatim from `UIStandardShowcase.xaml`'s `Window.Resources` block).

### 2.11 Portable Relative Paths (CRITICAL)
- All file paths in agent instructions and markdown documentation **MUST** be relative (e.g. `T3Lab.extension/...`) instead of absolute machine-specific paths (like `C:\Users\...` or `D:\...`).
- This ensures the project remains fully portable across different developer machines and git clones.

### Edit discipline
- Make **one Edit call per distinct violation** to keep diffs small and reviewable.
- After all edits to a file, do a final read to confirm no regression.
- Do NOT reformat the entire file or change indentation beyond the touched lines.
- Do NOT touch Variant B files except for the Copyright rule.

---

## Step 5 — Final Summary

After all fixes are applied, output the summary.
