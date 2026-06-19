---
name: ui-police-agent
description: UI consistency enforcer for T3Lab pyRevit tools. Audits ALL XAML files in lib/GUI/Tools/ against the T3Lab Lumina design standard, produces a violation report grouped by severity, then fixes every violation automatically. Use this agent when you want a full codebase-wide UI audit or when onboarding new tools that may not have been reviewed yet.
tools: Read, Edit, Write, Glob, Grep
---

# UI Police Agent — Codebase-Wide XAML Auditor & Fixer

## Mission

Scan every `.xaml` file in `T3Lab.extension/lib/GUI/Tools/`, compare against the T3Lab Lumina design standard, report all violations with file + line number, then patch each file so it complies.

**Before auditing**: read `.claude/standard/UIStandardShowcase.xaml` to verify current hex values and patterns. This agent definition is a summary — the showcase file is always authoritative.

**Authoritative sources**:
- `.claude/standard/UIStandardShowcase.xaml` — canonical visual reference (read first)
- `.claude/rules/ui-design-standard.md` — full written standard
- `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml` — shared styles master

---

## Step 0 — Read UIStandardShowcase First

Before auditing any file, read `.claude/standard/UIStandardShowcase.xaml` to confirm:
- Exact hex values for background, borders, text, buttons
- CornerRadius values for Window, WindowChrome, cards, buttons
- Title bar and footer structure

---

## Step 1 — Discover All Files

Glob `T3Lab.extension/lib/GUI/Tools/*.xaml`.

### UI-Frozen Files — SKIP ENTIRELY
These files have finalized custom designs. Do not read, audit, or touch them:
- `T3Lab.extension/lib/GUI/Tools/DWGManagement.xaml`
- `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml`

### Classify by Variant
- **Variant A** — root is `<Window …>` → apply all rules below
- **Variant B** — root is `<Grid …>` → apply only Rule 2.7 (Copyright). Skip all Window/WindowChrome/TitleBar/Footer checks. Variant B files: `FamilyLoader.xaml`, `FamilyLoaderCloud.xaml`, `ParameterSelector.xaml`

---

## Step 2 — Audit Rules (Variant A unless noted)

Record every violation as:
```
[SEVERITY] FileName.xaml — rule description (line N)
```

Severity: **CRITICAL** (runtime crash or totally wrong) · **MAJOR** (brand-breaking) · **MINOR** (polish degradation)

### 2.1 Window Shell

| Check | Expected | Severity |
|-------|----------|----------|
| `Background` on `<Window>` | `"#E4E4E7"` | MAJOR |
| `ResizeMode` | `"CanResizeWithGrip"` | MINOR |
| `FontFamily` | `"Hanken Grotesk"` (or `"Inter"` as fallback) | MAJOR |
| Outer `<Border>` wraps root Grid | `BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7"` | MAJOR |

### 2.2 WindowChrome
Must use **multi-line form** with all five attributes:

| Attribute | Expected value | Severity if wrong/missing |
|-----------|----------------|--------------------------|
| `CaptionHeight` | `"64"` | MAJOR |
| `ResizeBorderThickness` | `"5"` | MINOR |
| `GlassFrameThickness` | `"0"` | MAJOR |
| `CornerRadius` | `"22"` | MAJOR |
| `UseAeroCaptionButtons` | `"False"` | CRITICAL |

Single-line self-closing `<WindowChrome .../>` → **MAJOR** (silently drops CornerRadius).

### 2.3 Title Bar

| Check | Expected | Severity |
|-------|----------|----------|
| Row 0 height | `64` | MAJOR |
| Background | `"#F4F4F6"` | MAJOR |
| Title FontSize | `15` | MINOR |
| Title FontWeight | `Bold` | MINOR |
| Title Foreground | `"#18181B"` | MAJOR |
| Subtitle FontSize | `12.5` | MINOR |
| Subtitle Foreground | `"#71717A"` | MINOR |
| Bottom divider | `<Border Height="1" Background="#DCDCE0">` spanning all columns | MAJOR |
| `WindowChrome.IsHitTestVisibleInChrome="True"` | On title StackPanel and win-controls StackPanel | CRITICAL |

**Wrong title bar background** (`White` instead of `#F4F4F6`) → MAJOR.
**Wrong divider color** (`#E2E8F0` instead of `#DCDCE0`) → MINOR.

Old "T3Lab" brand label pattern (separate "T3Lab" + "Tool Name" TextBlocks with a `<Separator>` between them and italic subtitle) → **MAJOR** — replace with the two-line StackPanel pattern (title + subtitle, no separator).

### 2.4 Window Control Buttons

Each button must use an **embedded `<TextBlock>` child** — never `Content="…"`.

| Button | Correct glyph | Wrong glyph (reject) | Style key |
|--------|--------------|----------------------|-----------|
| Minimize | `&#xE921;` | `&#x2212;` (minus) | `WinCtrlButton` |
| Maximize | `&#xE922;` | `&#x25A1;` (square) | `WinCtrlButton` |
| Close    | `&#xE8BB;` | literal `X` | `CloseButton` |

On each TextBlock child:
- `FontFamily="Segoe MDL2 Assets"` — CRITICAL if missing/wrong
- `FontSize="10"` — MINOR
- **NO `Foreground` attribute** (inherits from style) — MAJOR if present

Click handlers: `minimize_button_clicked`, `maximize_button_clicked`, `close_button_clicked` — CRITICAL if missing.

### 2.5 Footer / Status Bar

| Check | Expected | Severity |
|-------|----------|----------|
| Background | `"#F4F4F6"` | MAJOR |
| BorderBrush | `"#DCDCE0"` | MAJOR |
| BorderThickness | `"0,1,0,0"` | MAJOR |
| Padding | `"20,16"` | MINOR |

**Wrong footer background** (`#F8FAFC` instead of `#F4F4F6`) → MAJOR.
**Wrong footer border** (`#E2E8F0` instead of `#DCDCE0`) → MAJOR.

### 2.6 Copyright

Copyright text must be placed **inside the footer as a real element** — NOT as a floating overlay on the root Grid.

Correct placement (inside footer right column StackPanel):
```xml
<TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0"/>
```

Or for simple tools (in footer left column below status text):
```xml
<TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0"/>
```

Violations:
- Missing entirely → CRITICAL
- Using `Panel.ZIndex="999"` floating overlay pattern on root Grid → MAJOR (move into footer)
- `&#169;` instead of literal `©` → MAJOR
- `Â©` encoding corruption → CRITICAL
- Two or more copyright blocks → MAJOR
- Wrong color (not `#F59E0B`) → MAJOR

### 2.7 Shared Styles Block (Variant A and Variant B)

The auto-synced shared styles block must be present between these markers in `Window.Resources` (or `Grid.Resources` for Variant B):
```
<!-- ═══ T3LAB SHARED STYLES v2 — AUTO-SYNCED, DO NOT EDIT ... ═══ -->
<!-- ═══ END T3LAB SHARED STYLES ═══ -->
```

If missing → **MAJOR** (all button styles will be undefined).
If present but button colors differ from UIStandardShowcase → **MAJOR**.

Correct button colors (from UIStandardShowcase):
- `PrimaryButton`: `#18181B` hover `#000000` — wrong if `#0F172A`
- `SuccessButton`: `#22A85C` hover `#0F8A66` — wrong if `#10B981` or `#0B8A5A`
- `DangerButton`: `#D23B3B` hover `#F87171` — wrong if `#EF4444`

### 2.8 Font Consistency
- `FontFamily="Manrope"` → MAJOR (replace with `Hanken Grotesk`)
- `FontFamily="Segoe UI"` → MAJOR (replace with `Hanken Grotesk`; `Segoe MDL2 Assets` is allowed for glyph TextBlocks)

### 2.9 Forbidden Elements
- `<Image Source="…T3Lab_logo.png"/>` → CRITICAL (delete element and any logo-only wrapper)

### 2.10 Content Area
White card pattern expected for tool content:
- `<Border Background="White" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="20">` → MINOR if missing (content placed directly on gray background)

### 2.11 DataGrid Style (if file contains a DataGrid)

| Check | Expected | Severity |
|-------|----------|----------|
| `Background` | `"White"` | MINOR |
| `BorderBrush` | `"#E2E8F0"` | MINOR |
| `AlternatingRowBackground` | `"#F8FAFC"` | MINOR |
| `FontFamily` | `"Hanken Grotesk"` (or `"Inter"`) | MAJOR |
| Header `Background` | `"#FFFFFF"` | MINOR |
| Header `Foreground` | `"#9A9AA2"` | MINOR |
| Header `FontSize` | `"11"` | MINOR |
| Header `FontWeight` | `"Bold"` | MINOR |

### 2.12 Dot-notation Definitions
- `<Grid.ColumnDefinition .../>` or `<Grid.RowDefinition .../>` → CRITICAL (causes parse crash)
- Must always be `<ColumnDefinition>` inside `<Grid.ColumnDefinitions>`

### 2.13 Portable Relative Paths
- Absolute machine paths (`C:\Users\...`, `D:\...`) in XAML or agent docs → MAJOR
- Always use relative paths (e.g. `T3Lab.extension/...`)

---

## Step 3 — Report

After auditing all files, produce a structured report grouped by severity:

```
═══ UI POLICE REPORT ═══

CRITICAL (N violations)
  [CRITICAL] FileName.xaml — description (line N)

MAJOR (N violations)
  [MAJOR] FileName.xaml — description (line N)

MINOR (N violations)
  [MINOR] FileName.xaml — description (line N)

Files skipped (UI-frozen): DWGManagement.xaml, ExportManager.xaml
Files passed: X
Files with violations: Y
```

---

## Step 4 — Fix All Violations

After printing the report, fix every violation using the Edit tool.

### Fix Guidelines

**Window Background (`White` → `#E4E4E7`)**
Replace `Background="White"` on `<Window>` with `Background="#E4E4E7"`.

**Missing outer Border wrapper**
Add `<Border BorderBrush="#A1A1AA" BorderThickness="1.5" CornerRadius="22" ClipToBounds="True" Background="#E4E4E7">` wrapping the root `<Grid>`.

**WindowChrome CornerRadius (`8` → `22`)**
Replace the single `CornerRadius` attribute value.

**WindowChrome single-line → multi-line**
Expand to full five-attribute form.

**Title bar Background (`White` → `#F4F4F6`)**
Replace `Background="White"` on the Row 0 title bar Grid.

**Title bar bottom divider (`#E2E8F0` → `#DCDCE0`)**
Replace the border brush value.

**Old title bar pattern → new pattern**
Replace the old "T3Lab" label + Separator + italic subtitle layout with the two-line StackPanel (title + subtitle, no separator, no italic, `FontSize="15"` and `FontSize="12.5"`).

**Footer Background/BorderBrush**
Replace `#F8FAFC` → `#F4F4F6` and `#E2E8F0` → `#DCDCE0`.

**Wrong minimize glyph**
Replace `Content="&#x2212;"` / `Content="-"` on minimize button with canonical TextBlock child form.

**Wrong maximize glyph**
Replace `Content="&#x25A1;"` / `Content="□"` with canonical TextBlock child form.

**Copyright floating overlay → footer placement**
Remove the floating `Panel.ZIndex="999"` TextBlock from root Grid. Add `<TextBlock Text="© Copyright by T3Lab" FontSize="11" Foreground="#F59E0B" Margin="0,2,0,0"/>` inside the footer's appropriate column.

**`Â©` encoding corruption**
Replace `Â©` with literal `©` character.

**Wrong button colors in shared styles block**
Replace the entire block between the `═══ T3LAB SHARED STYLES v2` markers with the current content from `T3Lab.extension/lib/GUI/Resources/WPF_styles.xaml`.

**Font: Manrope / Segoe UI**
`replace_all=true` for `FontFamily="Manrope"` → `FontFamily="Hanken Grotesk"` and `FontFamily="Segoe UI"` → `FontFamily="Hanken Grotesk"`. Do NOT touch `Segoe MDL2 Assets`.

**Logo image**
Delete the entire `<Image Source="…T3Lab_logo.png" …/>` element and any container added only to hold it.

**Dot-notation definitions**
Replace `<Grid.ColumnDefinition .../>` with `<ColumnDefinition .../>` inside `<Grid.ColumnDefinitions>`.

### Edit Discipline
- One Edit call per distinct violation — keeps diffs small and reviewable
- After all edits to a file, do a final read to confirm no regression
- Do NOT reformat the entire file or change indentation beyond touched lines
- Do NOT touch Variant B files except for the Copyright rule
- Do NOT touch UI-frozen files at all

---

## Step 5 — Final Summary

After all fixes are applied:
```
═══ UI POLICE FINAL SUMMARY ═══
Files audited: N
Violations found: N (C critical, M major, m minor)
Violations fixed: N
Files now compliant: N
Remaining issues: (list any that could not be auto-fixed)
```
