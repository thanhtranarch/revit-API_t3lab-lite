# Panel 4 — Data & IFC-SG (Ngày 9)

> 6 tool. Chuẩn bị: file BCF 2.1 mẫu (kèm 1 file BCF hỏng để test), model có foundation giao nhau, room + element, schedule, shared parameter file .txt.

## Tool 1/6 — IFC-SG
Chain: launcher → `IFCSGDialog.py` (2236 loc) → `IFCSG.xaml` + `SubtypeDefinerColMap.xaml` · config: `IFC-SG.pushbutton/configs/`

- [ ] Mở suite, load config mặc định từ `configs/` → không lỗi parse JSON
- [ ] Apply mapping param IFC-SG lên model mẫu → param ghi đúng element
- [ ] Config thiếu field / JSON hỏng → thông báo rõ file nào lỗi
- [ ] Mở SubtypeDefiner column mapping dialog → hoạt động
- [ ] Ctrl+Z revert việc ghi param
- [ ] Ghi chú:

## Tool 2/6 — Foundation Volume
Chain: self-contained (410 loc, 1 transaction) → `FoundationVolume.xaml`

- [ ] Foundation rời nhau → volume đúng (đối chiếu tay 1 khối đơn giản)
- [ ] Foundation giao nhau → boolean không tính trùng phần giao
- [ ] Element không có solid geometry → skip, không crash
- [ ] Ghi chú:

## Tool 3/6 — ManaContains
Chain: launcher → `ManaContainsDialog.py` (2204 loc) → `ManaContains.xaml`

- [ ] Element nằm trong room → nhận diện đúng
- [ ] Element ngoài room / trên biên → hành vi nhất quán
- [ ] Element trong **link model** → hoạt động hoặc thông báo giới hạn
- [ ] Model lớn → đo thời gian, có progress không treo
- [x] Ghi chú: **2026-07-03 (theo screenshot user)** — pyRevit output console bị flood bởi `print("[ManaContains] Skipped spatial element {}: {}"...)` lặp lại liên tục (mỗi room/space bị lỗi in ra 1 dòng). Đã xoá cả 3 chỗ `print()` trong file (2 chỗ "Skipped spatial element" ở `_t1_load_sp`/`_t2_load_spatial`, 1 chỗ "Failed to load" ở `_safe_call`) — except block vẫn skip phần tử lỗi như cũ, chỉ không in ra console nữa. **⚠️ Chưa điều tra nguyên nhân gốc**: log cho thấy rất nhiều spatial element liên tiếp bị skip cùng lỗi "Index was outside the bounds of the array" khi khởi tạo `Tab1SpatialItem`/`Tab2SpatialData` — nghi ngờ đây là bug thật (có thể liên quan Revit 2026 API) khiến danh sách room/space trong tool bị thiếu dữ liệu đáng kể, không chỉ là spam console. Cần điều tra riêng ở GĐ3 khi test functional với model thật.

## Tool 4/6 — ManaSched
Chain: launcher → `ManaSchedDialog.py` (1601 loc) → `ManaSched.xaml`

- [ ] List schedule đúng
- [ ] Sửa field/filter/sort qua tool → schedule cập nhật, Ctrl+Z OK
- [ ] Key schedule + schedule thường → cả 2 loại xử lý đúng
- [ ] Schedule đang mở làm active view → không conflict
- [ ] Ghi chú:

## Tool 5/6 — ManaPara
Chain: launcher → `ManaParaDialog.py` (2548 loc) → `ManaPara.xaml`

- [ ] Add shared param từ file .txt → binding đúng category
- [ ] Remove param → confirm trước khi xoá
- [ ] File shared param bị lock/không tồn tại → thông báo rõ
- [ ] Param ghi vào element instance vs type → đúng target
- [ ] Ghi chú:

## Tool 6/6 — BCF Reader
Chain: self-contained (3041 loc, 4 transaction) → `BCFReader.xaml`

- [ ] Mở file BCF 2.1 mẫu → list topic/comment/viewpoint đúng, hiển thị unicode tiếng Việt OK
- [ ] Apply viewpoint → camera view Revit khớp hướng nhìn BCF
- [ ] BCF hỏng (zip lỗi / XML thiếu) → thông báo, không crash
- [ ] Element highlight theo GUID trong BCF → đúng element (test cả GUID không tồn tại trong model)
- [ ] Ghi chú:

## Chốt ngày 9
- [ ] Cập nhật bảng + README · Commit: `fix(panel-data-ifcsg): <mô tả>`

| Tool | Trạng thái | Ghi chú |
|------|-----------|---------|
| IFC-SG | ⬜ | Bug trang trắng khi mở (xem Phát sinh) đã sửa 2026-07-03, chưa test functional |
| Foundation Volume | ⬜ | |
| ManaContains | ⬜ | Bug trang trắng khi mở (xem Phát sinh) đã sửa 2026-07-03, chưa test functional |
| ManaSched | ⬜ | Gray-gutter + footer height fix + đồng bộ header/cột/chiều cao bảng Tab1↔Tab2 (xem Phát sinh) đã sửa 2026-07-03, chưa test functional |
| ManaPara | ⬜ | Bug trang trắng khi mở (xem Phát sinh) đã sửa 2026-07-03, chưa test functional |
| BCF Reader | ⬜ | |

## Phát sinh

- **2026-07-03 — ManaContains, ManaPara (cùng lúc với ManaAnno ở `panel-1-annotation-select.md`): mở lên không hiện bảng, phải bấm tab khác rồi quay lại mới hiện**: cùng root cause đã tìm/sửa ở ManaSheets/ManaViews/ManaAnno — headerless TabControl (`ContentPresenter ContentSource="SelectedContent"`) điều hướng bằng RadioButton sidebar có `IsChecked="True"` sẵn trong XAML, sự kiện `Checked` của nút mặc định bắn trước khi code Python kịp gắn handler, nên `TabControl.SelectedIndex` không bao giờ được set tường minh ở lần mở đầu. Đã sửa bằng 1 dòng ở cuối `__init__`:
  - `ManaContainsDialog.py`: thêm `self.tab_control.SelectedIndex = 0` (nav mặc định `nav_contains`, khớp `tab_contains_item`).
  - `ManaParaDialog.py`: thêm `self.tab_main.SelectedIndex = 0` (nav mặc định `btn_nav_browse`, khớp TabItem đầu = Browse Parameters).
  - Chi tiết đầy đủ (3 tool cùng bug, cùng lần sửa) xem `panel-1-annotation-select.md` Phát sinh. Đã verify `ast.parse` OK + `audit_tools.py --quiet` clean + `sync_wpf_styles.py --check` 0 lệch. Chưa được user xác nhận mở lại OK trong Revit.
- **2026-07-03 — ManaContains: ~10 button không có Style, hiện dạng nút Windows mặc định** (yêu cầu UI trực tiếp từ user, không phải bug logic): user thích cách các nút đã style trong `ManaContains.xaml` (bo góc, font Hanken Grotesk theo shared style) và muốn áp dụng thống nhất. Rà soát phát hiện 10 `<Button>` trong chính file này hoàn toàn thiếu `Style="{StaticResource ...}"` (Select All/Clear ở Tab 1 spatial + category, All/None/Invert ở Tab 2 spatial + category, Select All/Select None ở Tab 2 results) → render nút Windows mặc định (góc vuông, font hệ thống), lệch hẳn so với các nút khác trong cùng file đã dùng `SecondaryButton`/`PrimaryButton`/`SuccessButton`/`AccentButton`. Đã sửa: gắn `Style="{StaticResource SecondaryButton}"` cho cả 10 nút, giữ nguyên `Height`/`Width` nhỏ gọn sẵn có (20-24px), thêm `Padding="0"` + `FontSize` phù hợp (10.5-11) để không bị bể layout khi Style áp `Padding="14,7"` mặc định. Quét toàn bộ 52 file XAML còn lại (trừ 2 file UI-locked) tìm bug tương tự — chỉ thấy thêm 1 chỗ thật ở `ParameterSelector.xaml` (dialog phụ của BatchOut, xem `panel-3-views-sheets.md` Phát sinh); 2 chỗ khác (`FamiGen.xaml` nút xoá tìm kiếm dạng icon trong suốt, `FoundationVolume.xaml` nút `CloseButton` bị `Visibility="Collapsed"` chết) không phải bug — không sửa. Verify `sync_wpf_styles.py --check` 0 lệch + `audit_tools.py --quiet` clean. Chưa được user xác nhận trực quan trong Revit.
- **2026-07-03 — IFC-SG Suite: cùng bug trang trắng khi mở, phát hiện riêng sau khi user báo tiếp**: cùng root cause hệt trên — `IFCSGDialog.py` wire `self.btn_tab_assigner.Checked += self._on_tab_changed` / `self.btn_tab_checker.Checked += self._on_tab_changed` (dòng ~1009-1010) trong khi `btn_tab_assigner` đã có `IsChecked="True"` sẵn trong `IFCSG.xaml`, nên sự kiện `Checked` bắn trước khi handler kịp gắn → `main_tab_control.SelectedIndex` không bao giờ được set tường minh lúc mở. Đã sửa: thêm `self.main_tab_control.SelectedIndex = 0` ở cuối `__init__` (nav mặc định `btn_tab_assigner`, khớp `TabItem` đầu = Subtype Assigner). Verify `ast.parse` OK + `audit_tools.py --quiet` clean (không đụng XAML). Chưa được user xác nhận mở lại OK trong Revit.
- **2026-07-03 — IFC-SG: 10 nút vẫn giữ `Height="32"` cũ, chưa ăn theo chuẩn compact mới (28px)** (user tự phát hiện sau khi GĐ2 đã báo "tất cả UI đều okay"): việc thu nhỏ chiều cao nút app-wide (40→28) trước đó chỉ ảnh hưởng các nút KHÔNG có `Height` set tay riêng (ăn theo default của style dùng chung). `IFCSG.xaml` có 10 nút set tay `Height="32"` từ trước (độc lập với style chung): `btnLoadExcel`, `btnAutoAssign`, `btnApply`, `btnApplyAll`, `btnImportXML`, `btnImportExcel`, `btnSaveConfig`, `btnDeleteConfig`, `btnRunCheck`, `btnExportExcel` → các nút này KHÔNG tự co theo, vẫn to hơn 7 nút khác trong cùng file (`btnExpandAll`/`btnCollapseAll`/`btnFilterAll`/`btnFilterFail`/`btnFilterWarn`/`btnFilterPass`/`btnSelectAllFailed`) vốn đã sẵn `Height="28"` từ trước — gây lệch kích thước nút ngay trong cùng 1 tool. Đã sửa cả 10 nút từ `Height="32"` → `"28"` để đồng nhất. Đã kiểm tra thêm dialog phụ `SubtypeDefinerColMap.xaml` (SubtypeDefiner column mapping) — 2 nút `btnCancel`/`btnOK` không set `Height` riêng nên đã tự ăn theo default 28px từ trước, không cần sửa. Verify: `audit_tools.py --quiet` clean, `sync_wpf_styles.py --check` 53/53 khớp. Chưa được user xác nhận trực quan trong Revit.
- **2026-07-03 — ManaSched: gray gutter quanh vùng nội dung + footer action-bar cao dư** (yêu cầu UI trực tiếp từ user):
  - Gray gutter: `TabControl x:Name="tab_main"` (Grid.Row="1"/Col="1") có `Background="Transparent"` → nền xám cửa sổ lộ ra qua khe hở, cùng lớp bug đã gặp nhiều lần. Đã sửa: bọc `TabControl` trong `Border Background="#F8FAFC"` full-bleed, giống pattern chuẩn.
  - Footer cao dư: Border "Footer: Actions full-width" (Tab 0 — Excel Link, chứa nút Preview/Export to Excel/Import from Excel/Apply Changes + dòng status) có `Padding="16,8"` — hơi rộng so với nút giờ chỉ còn 28px cao (sau lần thu nhỏ app-wide). Đã giảm `Padding` xuống `"16,6"` và margin-top của dòng status text từ `0,6,0,0` xuống `0,4,0,0` cho gọn hơn.
  - Verify: XML well-formed, `audit_tools.py --quiet` clean, `sync_wpf_styles.py --check` 53/53 khớp. Chưa được user xác nhận trực quan trong Revit.
- **2026-07-03 (tiếp) — ManaSched: đồng bộ kích thước panel "Select Schedules" giữa Tab 1 (Excel Link) và Tab 2 (Duplicator), bỏ gap 2 bên bảng** (yêu cầu UI trực tiếp từ user): Tab 1 dùng `ColumnDefinition Width="310"` cho cột trái trong khi Tab 2 dùng `"300"` → lệch 10px giữa 2 tab cùng 1 khái niệm UI ("Select Schedules"). Đồng thời Tab 1 có `Margin="0,0,8,0"`/`"0,0,8,8"` (phải 8px) trên search-pill, DataGrid border và selection-bar — tạo khoảng hở bên phải trước divider dọc — trong khi Tab 2 không có margin nào (đã flush 2 bên sẵn). Đã sửa: đổi `Width="310"` → `"300"` khớp Tab 2; bỏ toàn bộ `,8,` (margin phải) trên 3 chỗ ở Tab 1 (search pill, DataGrid border, selection bar) để bảng + search + nút giờ chạm sát cả 2 cạnh trái/phải của cột, đồng nhất với Tab 2. Verify: XML well-formed, `audit_tools.py --quiet` clean, `sync_wpf_styles.py --check` 53/53 khớp. Chưa được user xác nhận trực quan trong Revit.
