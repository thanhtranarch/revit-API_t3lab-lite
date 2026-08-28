# CHANGE HISTORY — lịch sử các cycle

> Entry mới nhất ở **trên cùng**. Mỗi cycle một entry, đủ 11 trường
> (xem `06-baseline-tracking.md` §8). Không sửa entry cũ — sai thì thêm entry đính chính.

---

## Cycle 0b — 2026-08-28 — DỌN DẸP + GATE T3

```
DATE:                   2026-08-28
CYCLE:                  0b (dọn dẹp luật cũ + dựng gate T3; vẫn chưa audit tool nào)
TOOLS AUDITED:          0
ISSUES FOUND:           1 (false positive của chính audit_t3.py — xem dưới)
ISSUES FIXED:           1
ISSUES REMAINING:       0
SCORE CHANGES:          — (chưa có điểm nào)
REGRESSIONS:            —
SYSTEM-WIDE PATTERNS:   —
DESIGN SYSTEM GAPS:     #1 và #3 vẫn mở; #2 vẫn mở. Q1 đã đóng.
NEXT PRIORITIES:        Q4 quét 3 lỗi P0 tiềm ẩn · Q2 audit sâu 6 tool đầu ·
                        Q3 migrate 2 file (bị GAP #1 chặn)
COMMIT:                 <điền sha sau khi commit>
```

**Thêm `dev/audit_t3.py`** — gate của chuẩn T3, thay `dev/audit_ui.py`. Chia file làm
**T3-DECLARED** (soi đầy đủ, vi phạm ⇒ `exit 1`) và **LEGACY** (chỉ đếm). Nhờ vậy gate
xanh ngay hôm nay dù 0/51 file đạt chuẩn, và siết dần theo từng file được migrate.

**Thêm `.claude/rules/new-tool-standard.md`** — luật cho **mọi tool/script mới**:
bố cục file, 12 luật XAML, khung `script.py` với 9 luật (transaction, rollback, không
nuốt exception, nạp grid trong `__init__`…), checklist trước khi commit.

**Xoá 13 file của chuẩn đã bỏ:** `.claude/rules/ui-design-standard.md` ·
`.claude/standard/UIStandardShowcase.xaml` · `.claude/docs/wpf-window-templates.md` ·
`.claude/agents/ui-police-agent.md` · `.codex/agents/ui-police-agent.toml` ·
`dev/audit_ui.py` · `dev/sync_wpf_styles.py` · và 6 script migration one-shot
(`fix_ui.py`, `migrate_xaml_colors.py`, `add_sidebar_nav.py`, `export_palette.py`,
`mute_colors.js`, `refine_palette.js`). Không file code sống nào phụ thuộc chúng —
đã kiểm tra trước khi xoá; nội dung còn trong lịch sử git.

**Viết lại theo T3:** `.claude/agents/ui-agent.md` · `.claude/skills/xaml-templates.md`
· `.codex/agents/ui-agent.toml` · `.claude/CLAUDE.md` · `AGENTS.md` · `README.md` ·
`.claude/settings.local.json` (bỏ 12 permission trỏ script đã xoá).

**Issue tự tìm thấy và đã sửa:** bản đầu của `audit_t3.py` báo false positive
`thiếu HorizontalScrollBarVisibility` / `thiếu EnableRowVirtualization` cho `DataGrid`
dùng `Style="{StaticResource T3.DataGrid}"` — style đó đã set sẵn hai property bằng
`Setter`. Đã sửa: nhận diện `T3.DataGrid` / `T3.ListBox` / `T3.LogBox` là đã tuân thủ.

**Kiểm chứng:** XAML cố ý sai → bắt đủ 18 vi phạm, `exit 1`. XAML dựng từ đúng snippet
trong `xaml-templates.md` → `exit 0`. Tài liệu và gate khớp nhau.

---

## Cycle 0 — 2026-08-28 — BOOTSTRAP

```
DATE:                   2026-08-28
CYCLE:                  0 (bootstrap — thiết lập hệ thống governance, chưa audit)
TOOLS AUDITED:          0
ISSUES FOUND:           0 issue cấp tool
ISSUES FIXED:           0
ISSUES REMAINING:       0
SCORE CHANGES:          — (chưa có điểm nào)
REGRESSIONS:            —
SYSTEM-WIDE PATTERNS:   51/51 tool đang dùng hệ UI đã bị bỏ (Lumina) — toàn bộ là
                        legacy debt. 38–58 màu hex cứng mỗi file.
DESIGN SYSTEM GAPS:     #1 vị trí deploy T3Lab.Styles.xaml chưa chốt (CHẶN mọi migration)
                        #2 thiếu Style cho <Window> theo size class S/M/L
                        #3 copyright © T3Lab không có trong chuẩn T3 — giữ hay bỏ?
NEXT PRIORITIES:        Q1 viết dev/audit_t3.py (P1) · Q4 quét 3 lỗi P0 tiềm ẩn ·
                        Q2 audit sâu 6 tool đầu · Q3 migrate 2 file (bị GAP #1 chặn)
COMMIT:                 <điền sha sau khi commit>
```

**Nội dung cycle 0 — thiết lập, không sửa tool nào:**

- Chốt `pyRevit UI Design System/` là **chuẩn duy nhất**; bỏ Lumina, Revit-native,
  Terra v2, Kinetix.
- Dựng bộ tài liệu governance `docs/ui-governance/` (01→08 + BASELINE + CHANGELOG +
  PRIORITY_QUEUE + ROUTINE_PROMPT).
- Dựng agent `.claude/agents/ui-governance-agent.md`.
- Kiểm kê 54 XAML vào `BASELINE.md` — **không chấm điểm**, vì chấm điểm mà chưa audit
  là bịa số.

**Phát hiện quan trọng nhất của cycle 0:** `dev/audit_ui.py` gác chuẩn đã bị bỏ và sẽ
báo fail đúng những file được migrate cho đúng. Phải thay trước khi migration bắt đầu —
xem `PRIORITY_QUEUE.md` Q1.
