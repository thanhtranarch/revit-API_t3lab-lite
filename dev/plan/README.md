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
| **2. UI Consistency (Lumina)** | Sửa 2 outlier T3LabAssistant (palette + copyright) + verify trực quan | `phase-2-ui-consistency.md` | 2 |
| **3. Debug chức năng theo panel** | Smoke test 41 tool trong Revit, chia 6 panel | `panel-1..6-*.md` | 7 |
| **4. Regression & tổng kết** | Chạy lại toàn bộ audit, chốt báo cáo | mục cuối file này | 1 |

> Ghi chú UI: chuẩn copyright là **footer (status bar), bên trái** — codebase vốn đã theo pattern này ở 74/76 file; `ui-design-standard.md` đã được cập nhật khớp thực tế (2026-07-02), nên GĐ2 chỉ còn 2 ngày.

## Lịch thực hiện 12 ngày

| Ngày | Giai đoạn | Việc | File | Trạng thái |
|------|-----------|------|------|-----------|
| 1 | GĐ1 | Sửa F1 FindReplace · F2 CreateFromRooms · F3 UIShowcase | `phase-1-cleanup-fixes.md` §Ngày 1 | ✅ |
| 2 | GĐ1 | Archive 15 XAML mồ côi + 6 dialog chết · chạy lại audit | `phase-1-cleanup-fixes.md` §Ngày 2 | ✅ |
| 3 | GĐ2 | Sửa T3LabAssistant.xaml (palette + copyright) + AssistantPane.xaml → audit UI sạch | `phase-2-ui-consistency.md` §Ngày 3 | ✅ |
| 4 | GĐ2 | Mở từng cửa sổ trong Revit, verify trực quan + screenshot | `phase-2-ui-consistency.md` §Ngày 4 | ✅ **Hoàn thành 2026-07-03** — user xác nhận "tất cả UI đều okay". Panel 1–3 xác nhận qua nhiều screenshot thực tế trước đó (gray-gutter/layout/copyright đã sửa); Panel 4–6 xác nhận đợt này. Screenshot file riêng skip có lý do (QA trực tiếp trong Revit thay vì lưu ảnh). ⚠️ Ngoại lệ: `audit_ui.py` còn 1 lỗi biết trước chưa sửa (IFCSG copyright trùng lặp x2, task nền riêng) — chưa phải audit sạch 100% |
| 5 | GĐ3 | Debug Panel: Annotation & Select (4 tool) | `panel-1-annotation-select.md` | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung "các tool hiện tại đều đã hoạt động tốt" (crash ManaSelect + trang trắng ManaAnno đã sửa 2026-07-03 trước đó) |
| 6 | GĐ3 | Debug Panel: Modeling & Datum — nhóm Create (7 tool) | `panel-2-modeling-datum.md` §Ngày 6 | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung; các fix trước đó: CADToElements (07-03), DoorThreshold width/thickness (07-04), ImageToDrafting/TextToElement/TileLayout UI+path (07-03). Ngoài scope còn lại: DoorThreshold trên curtain wall/wall nghiêng |
| 7 | GĐ3 | Debug Panel: Modeling & Datum — nhóm Adjust/Family (7 tool) | `panel-2-modeling-datum.md` §Ngày 7 | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung. FamiGen "From JSON" prompt v2 (7 file tự chứa + SOFT-FORM RECIPES + case library) shipped cùng ngày; 3 probe chi tiết theo dõi riêng ở `famigen-from-json-roadmap.md` |
| 8 | GĐ3 | Debug Panel: Views & Sheets (4 tool, trọng tâm BatchOut) | `panel-3-views-sheets.md` | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung (UI ManaSheets/ManaViews đã xác nhận riêng 2026-07-03) |
| 9 | GĐ3 | Debug Panel: Data & IFC-SG (6 tool) | `panel-4-data-ifcsg.md` | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung (pre-flight tĩnh sạch 2026-07-04: 3 fix trang trắng còn nguyên, print flood đã xoá) |
| 10 | GĐ3 | Debug Panel: Standards & Settings (4 tool) | `panel-5-standards-settings.md` | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung (đóng theo xác nhận, không có phiên review tĩnh riêng) |
| 11 | GĐ3 | Debug Panel: Support + Standard (9 tool) | `panel-6-support-standard.md` | ✅ **Hoàn thành 2026-07-05** — user xác nhận chung (đóng theo xác nhận, không có phiên review tĩnh riêng) |
| 12 | GĐ4 | Regression toàn bộ + tổng kết | mục dưới | ⬜ |

> Nguyên tắc: **không nhảy giai đoạn**. GĐ1 dọn dead code trước để GĐ2 không tốn công sửa UI cho file sắp xoá; GĐ2 chốt UI trước để GĐ3 smoke test không bị nhiễu bởi thay đổi XAML.

## Giai đoạn 4 — Ngày 12: Regression & tổng kết

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

---

## Roadmap độc lập theo tool — ĐÃ ĐÓNG TOÀN BỘ (2026-07-05)

> User xác nhận 2026-07-05: **toàn bộ 4 roadmap độc lập đã thực hiện xong**; 4 file roadmap
> đã xóa cho gọn theo yêu cầu (nội dung đầy đủ vẫn còn trong lịch sử git — xem commit xóa).

| Roadmap | Kết cục |
|---------|---------|
| T3Lab Assistant — Agentic Upgrade | ✅ Đóng — code GĐ A/B/C đã ship trước đó; mục mở cuối cùng (D3 smoke test) xác nhận qua phản hồi chung "các tool hoạt động tốt" + xác nhận đóng roadmap 2026-07-05 |
| AutoDimension — Cải thiện | ✅ Đóng — GĐ A đã ship & hoạt động tốt; backlog cải tiến (~16 mục GĐ B/C/D: idempotent, dim mặt tường host, core-layer thật, stacking manager, EQ, trục xiên, section mở rộng, preset/preview…) **đóng không triển khai** theo quyết định user — nếu cần thì mở roadmap mới |
| FamiGen — From JSON | ✅ Đóng — prompt v2 (7 file tự chứa + SOFT-FORM RECIPES + Object case library + ví dụ sofa 11 part) shipped, smoke test xác nhận chung 2026-07-05; backlog WS4 (cảnh báo part rời rạc) / WS5 (sketch plane tùy ý) **đóng không triển khai** |
| MCP Tools — Expansion | ✅ Đóng — toàn bộ mục đã tick từ trước, smoke test xác nhận chung 2026-07-05 |
