# T3Lab — Master Plan: Debug & UI Consistency (thực hiện theo ngày)

> Bắt đầu: 2026-07-02 · Branch làm việc: tạo branch mới cho từng giai đoạn từ `main`
> Nguồn dữ liệu: review tĩnh 41 tool + UI audit 76 XAML (xem `dev/DEBUG_PLAN.md`)

## Cách dùng plan này

1. Mỗi ngày mở đúng file của ngày đó (bảng lịch bên dưới), làm theo checklist, tick `- [x]` trực tiếp vào file.
2. Cuối mỗi ngày chạy bộ kiểm tra regression rồi commit (kể cả file plan đã tick):
   ```bash
   python3 dev/audit_tools.py --quiet      # audit code tĩnh
   python3 dev/audit_ui.py --quiet         # audit UI Lumina
   python3 dev/sync_wpf_styles.py --check  # style block đồng bộ
   ```
3. Gặp lỗi mới → ghi vào mục **"Phát sinh"** cuối file panel tương ứng, đừng sửa lan man ngoài phạm vi ngày.
4. Cập nhật bảng tiến độ bên dưới (cột Trạng thái).

## 4 Giai đoạn

| Giai đoạn | Nội dung | File plan | Số ngày |
|-----------|----------|-----------|---------|
| **1. Sửa lỗi tĩnh & dọn dẹp** | 4 bug đã xác nhận (F1–F4) + archive dead code | `phase-1-cleanup-fixes.md` | 2 |
| **2. UI Consistency (Lumina)** | Chuẩn hoá copyright 73 file, palette sót, verify trực quan | `phase-2-ui-consistency.md` | 4 |
| **3. Debug chức năng theo panel** | Smoke test 41 tool trong Revit, chia 6 panel | `panel-1..6-*.md` | 7 |
| **4. Regression & tổng kết** | Chạy lại toàn bộ audit, chốt báo cáo | mục cuối file này | 1 |

## Lịch thực hiện 14 ngày

| Ngày | Giai đoạn | Việc | File | Trạng thái |
|------|-----------|------|------|-----------|
| 1 | GĐ1 | Sửa F1 FindReplace · F2 CreateFromRooms · F3 UIShowcase | `phase-1-cleanup-fixes.md` §Ngày 1 | ⬜ |
| 2 | GĐ1 | Archive 15 XAML mồ côi + 6 dialog chết · chạy lại audit | `phase-1-cleanup-fixes.md` §Ngày 2 | ⬜ |
| 3 | GĐ2 | Viết `dev/fix_copyright.py` · fix copyright Panel 1 + Panel 3 · palette T3LabAssistant | `phase-2-ui-consistency.md` §Ngày 3 | ⬜ |
| 4 | GĐ2 | Fix copyright Panel 2 (16 file — nhiều nhất) | `phase-2-ui-consistency.md` §Ngày 4 | ⬜ |
| 5 | GĐ2 | Fix copyright Panel 4 + 5 + 6 + XAML dùng chung | `phase-2-ui-consistency.md` §Ngày 5 | ⬜ |
| 6 | GĐ2 | Mở từng cửa sổ trong Revit, verify trực quan + screenshot | `phase-2-ui-consistency.md` §Ngày 6 | ⬜ |
| 7 | GĐ3 | Debug Panel: Annotation & Select (4 tool) | `panel-1-annotation-select.md` | ⬜ |
| 8 | GĐ3 | Debug Panel: Modeling & Datum — nhóm Create (7 tool) | `panel-2-modeling-datum.md` §Ngày 8 | ⬜ |
| 9 | GĐ3 | Debug Panel: Modeling & Datum — nhóm Adjust/Family (7 tool) | `panel-2-modeling-datum.md` §Ngày 9 | ⬜ |
| 10 | GĐ3 | Debug Panel: Views & Sheets (4 tool, trọng tâm BatchOut) | `panel-3-views-sheets.md` | ⬜ |
| 11 | GĐ3 | Debug Panel: Data & IFC-SG (6 tool) | `panel-4-data-ifcsg.md` | ⬜ |
| 12 | GĐ3 | Debug Panel: Standards & Settings (4 tool) | `panel-5-standards-settings.md` | ⬜ |
| 13 | GĐ3 | Debug Panel: Support + Standard (9 tool) | `panel-6-support-standard.md` | ⬜ |
| 14 | GĐ4 | Regression toàn bộ + tổng kết | mục dưới | ⬜ |

> Nguyên tắc: **không nhảy giai đoạn**. GĐ1 dọn dead code trước để GĐ2 không tốn công sửa UI cho file sắp xoá; GĐ2 chốt UI trước để GĐ3 smoke test không bị nhiễu bởi thay đổi XAML.

## Giai đoạn 4 — Ngày 14: Regression & tổng kết

- [ ] `python3 dev/audit_tools.py --quiet` → exit 0
- [ ] `python3 dev/audit_ui.py --quiet` → exit 0
- [ ] `python3 dev/sync_wpf_styles.py --check` → 0 lệch
- [ ] Mọi bảng panel: không còn ô ⬜ (hoặc có ghi chú lý do skip)
- [ ] Tổng hợp mục "Phát sinh" của 6 file panel → tạo danh sách issue còn lại, xếp ưu tiên
- [ ] Cập nhật `dev/DEBUG_PLAN.md` trạng thái cuối cùng
- [ ] Screenshot bộ UI đã chuẩn hoá lưu vào `dev/plan/screenshots/` (nếu có)

## Quy ước trạng thái

- ⬜ chưa làm · 🔄 đang làm · ✅ xong · ❌ fail (kèm ghi chú) · ⏭️ skip có lý do

## Định nghĩa "xong" cho 1 tool (GĐ3)

1. Mở từ ribbon không lỗi; chrome window (drag/min/max/close/resize) hoạt động.
2. Happy path chạy đúng trên model test; Ctrl+Z revert đúng 1 bước.
3. Edge case (không chọn gì / model rỗng / cancel) ra thông báo thân thiện, không stacktrace.
4. Không exception bị nuốt trong pyRevit log khi thao tác bình thường.
5. Tick ✅ vào bảng panel + ghi chú nếu có hành vi lạ.
