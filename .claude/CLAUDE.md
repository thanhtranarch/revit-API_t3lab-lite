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

## Daily Debug Workflow — lệnh `gd# revitapi`

Kế hoạch debug & UI consistency theo ngày nằm ở `dev/plan/`. **Nguồn trạng thái duy nhất**: bảng "Lịch thực hiện" trong `dev/plan/README.md` + các checkbox `- [ ]`/`- [x]` trong từng file plan.

Khi user chat **`gd<N> revitapi`** (ví dụ `gd1 revitapi`), thực hiện đúng quy trình sau:

1. **Đọc `dev/plan/README.md`** để xác định trạng thái GĐ<N>:
   | Lệnh | Giai đoạn | Ngày | File plan |
   |------|-----------|------|-----------|
   | `gd1 revitapi` | Sửa lỗi tĩnh & dọn dẹp | 1–2 | `dev/plan/phase-1-cleanup-fixes.md` |
   | `gd2 revitapi` | UI Consistency (Lumina) | 3–4 | `dev/plan/phase-2-ui-consistency.md` |
   | `gd3 revitapi` | Debug chức năng theo panel | 5–11 | `dev/plan/panel-1..6-*.md` |
   | `gd4 revitapi` | Regression & tổng kết | 12 | mục cuối `dev/plan/README.md` |
2. **Nếu GĐ<N> đã hoàn thành** (mọi ngày của nó ✅): báo "GĐ<N> đã thực hiện xong" + tóm tắt kết quả và commit liên quan. KHÔNG làm lại.
3. **Nếu chưa hoàn thành**: báo "Đang thực hiện GĐ<N>" rồi thực hiện **ngày ⬜ kế tiếp** của giai đoạn đó theo đúng checklist trong file plan. Việc bắt buộc cần Revit (Ngày 4 visual QA, toàn bộ GĐ3 smoke test) thì không tự bịa kết quả — xuất checklist cho user tự test trong Revit, và tick kết quả dựa trên phản hồi user báo lại.
4. **Sau khi làm xong phần việc**: tick `- [x]` trong file plan, cập nhật bảng tiến độ README (⬜ → 🔄/✅), chạy regression:
   `python3 dev/audit_tools.py --quiet` · `python3 dev/audit_ui.py --quiet` · `python3 dev/sync_wpf_styles.py --check`
   rồi commit + push (kèm cả file plan đã tick).
5. **Luôn kết thúc câu trả lời** bằng danh sách các GĐ còn chưa thực hiện (và GĐ đang dở nếu có).

Lỗi phát sinh ngoài checklist → ghi vào mục "Phát sinh" của file plan tương ứng, không sửa lan man ngoài phạm vi ngày.

## UI-Frozen Files (DO NOT MODIFY UI)

The following XAML files are **UI-locked** — their visual design is finalized and must never be altered by any UI sweep, agent, or style sync operation. Logic/script changes are still allowed, but the XAML UI must remain untouched:

| File | Reason |
|------|--------|
| `T3Lab.extension/lib/GUI/Tools/DWGManagement.xaml` | Finalized custom design — UI locked |
| `T3Lab.extension/lib/GUI/Tools/ExportManager.xaml` | Finalized custom design (BatchOut) — UI locked |
| `T3Lab.extension/lib/GUI/Tools/T3LabAssistant.xaml` | Chat surface, not a tool dialog — Revit's own UI greys + Revit light/dark theme instead of Lumina. See `docs/assistant-revit-ui.md` |

> **T3LabAssistant.xaml — do NOT re-apply the Lumina palette.** Its colours come
> from `lib/GUI/RevitTheme.py` and follow Revit's own UI theme; every surface is
> bound with `{DynamicResource T3Theme<Token>}`. Hardcoding hexes back into it,
> or replacing the nested `<ScrollViewer.Resources>` thin-scrollbar overrides,
> breaks dark mode. The copyright amber `#F59E0B` and the shared style block are
> unchanged and still audited.

**All agents** (`@ui-agent`, `@ui-police-agent`, `@tool-builder-agent`, `@script-frame-agent`) must skip these files entirely during any UI-related task. Do not run `sync_wpf_styles.py` against them. Do not include them in bulk XAML audits.

> **Note (2026-07-20):** a Pause/Stop panel was briefly added to `ExportManager.xaml` for a chunked BatchOut export, then **reverted** — the chunked ExternalEvent re-queue crashed Revit mid-export. Both `ExportManager.xaml` and `BatchOut.pushbutton/script.py` are back at their pre-refactor state (commit `984eef0`) and remain UI-locked. Do not re-attempt BatchOut pause without a Revit-tested design — see the notes in the progress/pause rollout memory.

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
