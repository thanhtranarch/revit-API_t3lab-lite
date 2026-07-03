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
| ManaSheets | 🔄 | UI đã user xác nhận OK trong Revit (2026-07-03) — mở được, không crash, layout đúng. Còn thiếu test functional (batch rename/renumber, Ctrl+Z, edge case, Excel Import/Export) |
| SheetGen | ⬜ | |
| ManaViews | 🔄 | UI đã user xác nhận OK trong Revit (2026-07-03) — mở được, không lỗi, layout đúng. Còn thiếu test functional (batch thao tác view, Ctrl+Z, AdvancedViewManager) |
| BatchOut | ⬜ | |

## Phát sinh

- **2026-07-03 — ManaSheets, ManaViews không mở được**: `EXT_DIR` trong `script.py` đi lên thừa 1 cấp thư mục (`os.path.dirname` x4 thay vì x3 cần thiết cho `T3Lab.tab/<Panel>.panel/<Tool>.pushbutton/`), vượt quá `T3Lab.extension` tới tận thư mục gốc repo → `XAML_FILE` trỏ sai, tool không mở được (IOError). Phát hiện qua quét hệ thống sau khi user báo lỗi tương tự ở Tile Layout (thiếu 1 cấp, chiều ngược lại). Đã sửa cả 2 file, verify lại path trỏ đúng file `.xaml` thật. User xác nhận mở được cả 2 tool trong Revit (2026-07-03).
- **2026-07-03 — ManaSheets crash Revit khi mở** (không chỉ lỗi path ở trên — bug thứ 2, nghiêm trọng hơn): User báo "crash khi mở ManaSheets". Truy vết qua Revit journal `journal.0064.txt` (12:08:18) thấy exception thật: `A TwoWay or OneWayToSource binding cannot work on the read-only property 'IsPlaceholder' of type 'Autodesk.Revit.DB.ViewSheet'`, xảy ra ngay khi DataGrid bind dữ liệu lúc mở cửa sổ — exception này nổ trên WPF dispatcher thread nên **không** bị bắt bởi try/except Python trong `show_sheet_manager()`, Revit hiện TaskDialog "could not complete the external command" rồi crash thật (access violation ngay sau đó, kèm `.dmp`). Nguyên nhân: `ManaSheets.xaml` dòng ~1384, cột `DataGridCheckBoxColumn Header="Placeholder"` bind thẳng `element.IsPlaceholder` (property Revit API chỉ đọc) — `IsReadOnly="True"` trên cột chỉ khoá UI edit, không đổi Binding Mode mặc định (TwoWay cho DataGridCheckBoxColumn), nên WPF vẫn cố gắn TwoWay binding vào property không có setter → nổ ngay khi attach binding. Đã sửa: thêm `Mode=OneWay` vào Binding. Đã quét toàn bộ `DataGridCheckBoxColumn` trong `lib/GUI/Tools/` — chỉ ManaSheets.xaml bind thẳng property Revit API, các tool khác đều bind property Python thường (`is_selected`, `IsSelected`, `is_checked`) nên không dính bug này. User xác nhận mở lại OK, không còn crash (2026-07-03).
- **2026-07-03 — ManaSheets Excel Import/Export chưa hoạt động (phát hiện, chưa sửa — ngoài phạm vi bug crash)**: `ManaSheetsDialog.py` gọi `self.excel_service.export_sheets(path, data)` / `.import_sheets(path)`, nhưng `excel_service.py` chỉ định nghĩa `export_sheets_to_excel(sheet_models, filepath)` / `import_sheets_from_excel(filepath)` — tên method lệch, thứ tự tham số lệch (dialog truyền path trước, service nhận sheet_models trước), và `export_sheets_to_excel` kỳ vọng list `SheetModel` object (`.sheet_number` attribute) trong khi dialog truyền list dict. Import cũng lệch: service trả dict không có key `"id"` nhưng dialog cần `row.get("id")` để map về sheet. Kết quả: nút "Excel Export/Import" sẽ luôn báo lỗi/import 0 dòng, không crash (được bọc try/except) nhưng tính năng không dùng được. Cần sửa riêng, nằm ngoài scope bug crash-on-open lần này.
- **2026-07-03 — ManaViews báo lỗi ngay khi mở** (sau khi fix ManaSheets, user báo tiếp lỗi tương tự ở ManaViews): User xác nhận qua trả lời trực tiếp: hiện error dialog nhưng Revit KHÔNG đóng/crash thật (khác cơ chế với bug ManaSheets ở trên — đây là exception Python thường, bị bắt bởi try/except ngoài `show_view_manager()`, không phải lỗi WPF dispatcher không bắt được). Không tìm thấy bằng chứng trong Revit journal 2026 (phiên đang chạy mở/đóng ManaViews nhiều lần thành công) lẫn Windows WER — nghĩa là lỗi xảy ra ngay ở bước Python trước khi cửa sổ kịp hiện, không phải native crash. Rà soát `ManaViewsDialog.py`: `_load_views_data()` (dùng khi build tab Views) lọc bỏ `ViewType.ProjectBrowser/SystemBrowser/Undefined/Internal` TRƯỚC khi đụng vào property của view (các pseudo-view này được biết là ném exception khi đọc property). Nhưng `_get_all_templates_names()` — được gọi TRƯỚC `_load_views_data()` ngay trong `__init__`, để populate dropdown "View Template" — lại KHÔNG áp filter này và không có try/except nào, chỉ có `if view.IsTemplate: templates.append(view.Name)` trần trụi trên toàn bộ collector `OfClass(View).WhereElementIsNotElementType()`. Nếu `.IsTemplate` ném lỗi trên 1 trong các pseudo-view đó → toàn bộ `ViewManagerWindow()` construction fail → chỉ bị bắt bởi try/except ngoài cùng `show_view_manager()` → đúng triệu chứng "lỗi ngay khi mở, Revit vẫn sống". Đã sửa: thêm cùng bộ lọc ViewType + try/except vào `_get_all_templates_names()` (`ManaViewsDialog.py` dòng ~282), khớp pattern phòng thủ đã dùng ở `_load_views_data()` và `_load_templates_data()`. Lưu ý: không có traceback thật để xác nhận 100% đây là dòng gây lỗi (không như bug ManaSheets ở trên, nơi đọc được exception message thật từ journal) — đây là fix dựa trên review code tìm ra lỗ hổng phòng thủ cụ thể, logic khớp hoàn toàn với triệu chứng. User xác nhận mở lại OK, không còn lỗi (2026-07-03).
- **2026-07-03 — cả 2 tool mở chậm + trang đầu tiên trống, phải bấm lại mới hiện** (sau khi fix 2 bug ở trên, user báo tiếp): 2 vấn đề riêng biệt trong cùng lần test, đã sửa cả hai.
  1. **Trang đầu trống**: cả `ManaSheets.xaml`/`ManaViews.xaml` dùng pattern "headerless tab" — RadioButton điều hướng (`nav_sheets`/`nav_views`) có `IsChecked="True"` sẵn trong XAML, và event `Checked` của nó được wire tới `_on_tab_changed` (hàm set `tab_control.SelectedIndex`) NHƯNG việc wire xảy ra SAU khi XAML đã parse xong — tức sự kiện `Checked` của radio button đã checked sẵn đã bắn ra và bị bỏ lỡ trước khi handler kịp gắn vào. Kết quả: `tab_control.SelectedIndex` không bao giờ được set tường minh ở lần mở đầu, nên `ContentPresenter ContentSource="SelectedContent"` trong `TabControl.Template` không refresh nội dung tab đầu — cho tới khi user bấm sang tab khác rồi bấm lại (lúc đó `Checked` event mới thực sự bắn và set `SelectedIndex`). Đã sửa: thêm `self.tab_control.SelectedIndex = 0` tường minh ở cuối `__init__` (sau khi load data) cho cả `ManaSheetsDialog.py` và `ManaViewsDialog.py`.
  2. **Mở chậm (chỉ ManaViews/AdvancedViewManager)**: `EnhancedViewItem.__init__` (`core/advanced_view_manager.py`) gọi 4 hàm (`_get_sheet_count`, `_get_title_on_sheet`, `_get_sheet_number`, `_get_sheet_name`), MỖI hàm tự chạy riêng 1 `FilteredElementCollector(doc).OfClass(Viewport)` và quét TOÀN BỘ viewport trong model để tìm viewport khớp view hiện tại → với N view và P viewport là 4×N×P lần quét, chậm rõ rệt trên model lớn. Đã sửa: thêm `build_viewport_map(doc)` build 1 lần duy nhất map `{view_id: [viewports]}`, truyền vào `EnhancedViewItem(view, doc, viewport_map)`, 4 hàm trên đọc từ `self._viewports` (đã lọc sẵn) thay vì quét lại. Áp dụng ở cả 2 nơi gọi `EnhancedViewItem`: `ManaViewsDialog.py` và `AdvancedViewManagerDialog.py`. User xác nhận hết trang trắng + tốc độ ổn (2026-07-03).
- **2026-07-03 — ManaSheets: xoá cột Placeholder, dời Summary Cards, bỏ label ở tab Renumber** (yêu cầu UI trực tiếp từ user, không phải bug):
  - Xoá cột `DataGridCheckBoxColumn Header="Placeholder"` khỏi `sheets_grid` (Tab 1) theo yêu cầu user — giữ nguyên dropdown filter "Placeholder Only/Non-Placeholder Only" ở sidebar vì đó là tính năng khác, không nằm trong yêu cầu.
  - Tab 1 (Sheets List): user muốn 4 thẻ Summary (TOTAL SHEETS/SELECTED/CATEGORIES/PENDING CHANGES) chuyển từ trên cùng xuống ngay phía trên Actions Footer. Đã đổi thứ tự `Grid.RowDefinitions` (DataGrid&Filters lên `Row=0` height `*`, Cards Summary xuống `Row=1` height `Auto`, Actions Footer giữ `Row=2`) và đổi `Grid.Row` tương ứng trên 2 phần tử.
  - Tab 2 (Renumber): xoá `TextBlock Text="SELECT SHEETS TO RENUMBER"` phía trên preview grid theo yêu cầu user; đơn giản hoá luôn `Grid.RowDefinitions` 2 hàng (Auto label + `*` grid) thành 1 Border duy nhất vì chỉ còn 1 phần tử.
  - Ghi chú: XAML hiện tại có vài chỗ đã bị sửa ở phiên khác/ngoài phạm vi hội thoại này (checkbox column dùng `DataGridTemplateColumn` + event `sheets_row_checkbox_changed`, `CornerRadius="0"` áp dụng cho panel filter/DataGrid) — không đụng vào, chỉ thực hiện đúng 3 thay đổi user yêu cầu ở trên.
  - User xác nhận UI OK trong Revit (2026-07-03).
- **2026-07-03 — ManaViews: dời Summary Cards ở cả 2 tab, đổi thứ tự nút footer Tab 1** (yêu cầu UI trực tiếp từ user, không phải bug):
  - Tab 1 (Views): 4 thẻ Summary (TOTAL/SELECTED/TYPES/FILTERS ACTIVE) dời từ trên cùng xuống ngay phía trên Actions Footer — cùng cách làm đã áp dụng ở ManaSheets (đổi `Grid.RowDefinitions`: DataGrid&Filters lên `Row=0` height `*`, Cards Summary xuống `Row=1` height `Auto`, Actions Footer giữ `Row=2`).
  - Tab 2 (Templates): 4 thẻ Summary (TOTAL TEMPLATES/SELECTED/IN USE/UNUSED) dời tương tự xuống trên Actions Footer.
  - Tab 1 footer: đổi thứ tự 2 nút bên trái từ "Excel Export/Import, Refresh List" → "Refresh List, Excel Export/Import".
  - User xác nhận UI OK trong Revit (2026-07-03).
- **2026-07-03 — ParameterSelector.xaml (dialog phụ của BatchOut): nút "Ok" không có Style** (phát hiện khi quét toàn bộ codebase tìm bug button-không-style tương tự ManaContains, xem `panel-4-data-ifcsg.md` Phát sinh): `button_ok` set thẳng `Background="#0F172A" Foreground="White" BorderThickness="0"` bằng tay thay vì dùng `Style="{StaticResource PrimaryButton}"` sẵn có trong file (Variant B, style block nằm trong `Grid.Resources`) → nút hiện góc vuông (default WPF chrome) dù màu nền/chữ đúng, lệch với các nút khác trong hệ thống Lumina đều bo góc. Đã sửa: thay bằng `Style="{StaticResource PrimaryButton}"`, bỏ 3 thuộc tính set tay trùng lặp với style. Verify `sync_wpf_styles.py --check` 0 lệch. Chưa được user xác nhận trực quan trong Revit.
