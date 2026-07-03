# Panel 3 — Views & Sheets (Ngày 8)

> 4 tool nhưng có **BatchOut (3764 loc)** — tool xuất bản lớn nhất, dành nửa ngày riêng.
> Model test: file có ~5 sheet (titleblock chuẩn), vài view plan/section/3D, 1 schedule.
> Máy test: cần PDF printer + đường dẫn export ghi được (BatchOut phụ thuộc máy).

## Tool 1/4 — ManaSheets
Chain: launcher → `ManaSheetsDialog.py` (583 loc) → `ManaSheets.xaml`

- [ ] Mở cửa sổ, list đủ sheet
- [ ] Batch rename/renumber → đúng, Ctrl+Z OK
- [ ] Sheet đang mở làm active view → thao tác vẫn đúng
- [ ] Số sheet trùng nhau khi renumber → validate (Revit cấm duplicate number)
- [ ] Ghi chú:

## Tool 2/4 — SheetGen ⚠️
Chain: self-contained (1262 loc, **7 transaction / 0 rollback**) → `SheetGen.xaml`

- [ ] Gen sheet từ template → sheet + viewport đúng vị trí
- [ ] Titleblock thiếu trong model → thông báo, không crash
- [ ] **Fail giữa batch** (ví dụ sheet number trùng ở sheet thứ 3/5) → kiểm tra state: các sheet trước đó có bị dở dang không (0 rollback là rủi ro chính) — nếu để lại rác, tạo fix bọc try/except + RollBack
- [ ] Ctrl+Z sau batch → xác định phải undo mấy lần (7 transaction rời → có thể cần TransactionGroup, ghi nhận để cải thiện)
- [ ] Ghi chú:

## Tool 3/4 — ManaViews
Chain: launcher → `ManaViewsDialog.py` (689 loc) → `ManaViews.xaml` · dialog phụ: `AdvancedViewManagerDialog.py` → `AdvancedViewManager.xaml` + `AdvancedViewManagerBatchRename.xaml`

- [ ] Mở cửa sổ, list view đúng phân loại
- [ ] Batch thao tác view (rename/duplicate/delete tuỳ chức năng) → Ctrl+Z OK
- [ ] View đang active / view đặt trên sheet → cảnh báo hợp lý trước khi xoá
- [ ] Mở AdvancedViewManager dialog phụ → không lỗi; batch rename trong đó hoạt động
- [ ] Ghi chú:

## Tool 4/4 — BatchOut ⚠️ trọng tâm ngày
Chain: self-contained (3764 loc, **30 bare except**) → `ExportManager.xaml` (**UI-LOCKED**) · dialog phụ: `ParameterSelectorDialog.py` → `ParameterSelector.xaml` · optional: `lib/Intelligence/api_learner.py`, `api_updater.py`

Chuẩn bị:
- [ ] Kiểm tra khối import Intelligence (script.py:62-67): thêm log tạm in giá trị `HAS_API_LEARNER` khi khởi động — nếu `False` trên máy có đủ file thì import đang fail ngầm, truy nguyên nhân
- [ ] Chèn traceback log tạm vào các bare except quanh flow export chính

Test:
- [ ] Mở wizard, đi đủ các step (hidden TabItem navigation) — Back/Next đúng
- [ ] Chọn 2 sheet → export PDF → file ra đúng thư mục, đúng tên theo naming pattern
- [ ] Naming pattern: mở ParameterSelector dialog (nút chọn param) → không lỗi, pattern chèn đúng vào textbox
- [ ] Export DWG cùng 2 sheet → OK
- [ ] Export NWC (nếu có trong scope) → OK hoặc thông báo thiếu Navisworks exporter
- [ ] Đường dẫn output không tồn tại / không có quyền ghi → thông báo rõ, không nuốt lỗi
- [ ] Máy không có PDF printer phù hợp → thông báo rõ
- [ ] Queue nhiều định dạng 1 lượt → thứ tự chạy + progress đúng, Pause/Stop hoạt động
- [ ] Ghi chú:

## Chốt ngày 8
- [ ] Gỡ log tạm · cập nhật bảng + README
- [ ] Commit: `fix(panel-views-sheets): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| ManaSheets | ⬜ | |
| SheetGen | ⬜ | |
| ManaViews | ⬜ | |
| BatchOut | ⬜ | |

## Phát sinh

_(ghi lỗi mới tại đây)_
