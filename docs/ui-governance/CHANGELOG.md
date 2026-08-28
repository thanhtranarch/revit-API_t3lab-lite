# CHANGE HISTORY — lịch sử các cycle

> Entry mới nhất ở **trên cùng**. Mỗi cycle một entry, đủ 11 trường
> (xem `06-baseline-tracking.md` §8). Không sửa entry cũ — sai thì thêm entry đính chính.

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
