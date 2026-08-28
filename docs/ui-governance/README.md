# pyRevit UI/UX Governance — bộ tài liệu routine

Hệ thống quản trị chất lượng UI/UX cho toàn bộ pyRevit extension của T3Lab: audit,
chấm điểm, chuẩn hoá, phát hiện regression, và cải thiện liên tục theo chu kỳ.

**Chuẩn duy nhất:** `pyRevit UI Design System/` — mọi hệ UI khác trong repo đã bị bỏ.

---

## Bắt đầu từ đâu

| Bạn muốn | Mở file |
|----------|---------|
| Dán prompt vào Routine của Claude | **`ROUTINE_PROMPT.md`** |
| Hiểu agent làm gì | `../../.claude/agents/ui-governance-agent.md` |
| Biết luật thiết kế là gì | `../../pyRevit UI Design System/T3LAB_UI_STANDARD.md` |
| Viết một tool/script MỚI | `../../.claude/rules/new-tool-standard.md` |
| Chạy gate UI | `python3 dev/audit_t3.py --quiet` |
| Biết cái gì được sửa, cái gì không | `08-design-system-authority.md` |
| Xem điểm hiện tại của từng tool | `BASELINE.md` |
| Xem cycle sau làm gì | `PRIORITY_QUEUE.md` |

## Bản đồ file

```
pyRevit UI Design System/           ← CHUẨN DUY NHẤT (nguồn, không sửa khi chưa duyệt)
├── T3LAB_UI_STANDARD.md               luật: token, type, spacing, 10 luật bố cục, P1-P5
└── T3Lab.Styles.xaml                  82 resource key T3.*

.claude/
├── rules/new-tool-standard.md      ← luật cho MỌI tool/script mới
├── skills/xaml-templates.md           snippet T3 cho từng control
├── agents/ui-agent.md                 agent viết XAML
└── agents/ui-governance-agent.md   ← định nghĩa agent governance

dev/audit_t3.py                     ← gate UI (--quiet / --legacy / --file)

docs/ui-governance/
├── README.md                       ← file này
├── ROUTINE_PROMPT.md               ← prompt tổng hợp cho Routine
├── 01-scan.md                         phạm vi quét, chọn tool cho cycle
├── 02-audit-checklist.md              7 nhóm kiểm tra A-G
├── 03-scoring.md                      rubric 0-100 + trần điểm cứng
├── 04-severity-and-issues.md          P0-P4, format issue 9 trường, hàng đợi
├── 05-improvement-and-safety.md       luật sửa · an toàn Revit · 7 bước verify
├── 06-baseline-tracking.md            baseline · trend · regression · change history
├── 07-report-template.md              format report
├── 08-design-system-authority.md   ← ĐỌC TRƯỚC KHI SỬA XAML
├── BASELINE.md                     ← dữ liệu sống: điểm từng tool
├── CHANGELOG.md                    ← dữ liệu sống: lịch sử cycle
└── PRIORITY_QUEUE.md               ← dữ liệu sống: hàng đợi
```

Ba file "dữ liệu sống" được agent cập nhật mỗi cycle. Bảy file `0N-*.md` là quy trình,
chỉ đổi khi quy trình đổi.

## Một cycle làm gì

```
SCAN → AUDIT → SCORE → IDENTIFY → PRIORITIZE → IMPROVE → VERIFY → COMPARE → LOG → REPORT
```

Ngân sách: audit sâu 5–8 tool · in-place fix ≤10 issue · migrate ≤2 file · ≤1 đề xuất
cấp hệ thống. Hết ngân sách thì đẩy phần dư vào `PRIORITY_QUEUE.md`.

## Trạng thái hiện tại (cycle 0)

| Chỉ số | Giá trị |
|--------|---------|
| Tổng XAML | 54 |
| Đạt chuẩn T3 | **0** |
| Còn LEGACY | 51 |
| LOCKED (ngoài phạm vi) | 3 |
| Overall Score | — (chưa tool nào được chấm) |

Hai việc chặn cần user quyết trước khi migration bắt đầu — xem `PRIORITY_QUEUE.md`
(Q1 — thay gate bằng `dev/audit_t3.py` — đã xong 2026-08-28):

1. **GAP #1** — `T3Lab.Styles.xaml` nằm ngoài `T3Lab.extension/` nên pyRevit không nạp
   được lúc runtime. Chưa chốt vị trí deploy thì chưa migrate được file nào.
2. **GAP #3** — dòng copyright `© Copyright by T3Lab` có ở cả 54 file nhưng không có
   trong chuẩn T3. Giữ hay bỏ?
