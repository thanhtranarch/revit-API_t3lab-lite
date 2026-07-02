# Giai đoạn 1 — Sửa lỗi tĩnh & Dọn dẹp (Ngày 1–2)

> Không cần Revit. Làm trên branch riêng, ví dụ `fix/phase-1-cleanup`.
> Căn cứ: findings F1–F4 trong `dev/DEBUG_PLAN.md` + kết quả `dev/audit_tools.py`.

## Ngày 1 — Sửa 3 bug đã xác nhận

### F1 — `FindReplace.py` thiếu handler `button_run` 🔴
`lib/GUI/Tools/FindReplace.xaml:1181` khai báo `Click="button_run"` → XamlParseException khi mở.

- [x] Đọc `lib/GUI/FindReplace.py` xác định hành vi Run mong muốn (class có `run = False` → pattern: set `self.run = True` rồi `self.Close()`)
- [x] Thêm method:
  ```python
  def button_run(self, sender, e):
      self.run = True
      self.Close()
  ```
- [x] Kiểm tra các handler khác trong `FindReplace.xaml` (`TextChanged`, v.v.) đều có method — chỉ có 4 handler tổng cộng (minimize/maximize/close từ `WPF_Base.my_WPF` + `button_run` mới thêm), không còn thiếu
- [x] Chạy `python3 dev/audit_tools.py --quiet` → hết dòng FindReplace

### F2 — `CreateFromRooms.py` thiếu 4 handler 🔴
Thiếu: `NumberValidationTextBox`, `UIe_ItemChecked`, `button_run`, `text_filter_updated`. Không nơi nào import → dead code.

- [x] Xác nhận lần cuối không ai dùng: `grep -rn "CreateFromRooms" --include="*.py" T3Lab.extension` → chỉ xuất hiện trong chính `CreateFromRooms.py` (class def + đường dẫn XAML của chính nó), 0 tham chiếu ngoài
- [x] **Quyết định**: archive (khuyến nghị) — dead code, không tool nào import
- [ ] Chuyển `lib/GUI/CreateFromRooms.py` + `lib/GUI/Tools/CreateFromRooms.xaml` vào `_archive/` (thực hiện ở Ngày 2 cùng đợt archive XAML mồ côi/dialog chết)
- [ ] ~~Nếu giữ → viết đủ 4 handler~~ (không áp dụng — đã quyết định archive)

### F3 — UIShowcase load XAML ngoài extension 🟠
`lib/GUI/UIShowcaseDialog.py:33` trỏ `[repo]/.claude/standard/UIStandardShowcase.xaml` — hỏng khi deploy extension rời repo.

- [x] Copy `.claude/standard/UIStandardShowcase.xaml` → `lib/GUI/Tools/UIStandardShowcase.xaml`
- [x] Sửa `UIShowcaseDialog.py`: ưu tiên bản trong `lib/GUI/Tools/`, fallback về `.claude/standard/` (giữ tương thích repo dev)
- [x] Lưu ý: `.claude/standard/` vẫn là bản canonical cho design reference — đã thêm comment nêu rõ khi sửa chuẩn UI phải cập nhật cả 2; `dev/sync_wpf_styles.py` đã tự loại trừ `UIStandardShowcase.xaml` khỏi vòng kiểm tra style-block đồng bộ (giữ nguyên hành vi cũ)
- [x] Chạy `python3 dev/audit_ui.py` — file mới copy pass, không bị flag trong danh sách vi phạm

### Chốt ngày 1
- [x] `python3 dev/audit_tools.py --quiet` — chỉ còn cảnh báo orphan XAML + `CreateFromRooms` (xử lý Ngày 2, đã quyết định archive)
- [x] Commit: `fix: add missing XAML handlers, make UIShowcase XAML path deploy-safe`

---

## Ngày 2 — Archive dead code (F4)

> **Cần xác nhận trước khi xoá hẳn.** Phương án an toàn: chuyển vào `T3Lab.extension/_archive/` (pyRevit không quét thư mục không có đuôi `.tab/.panel/.pushbutton`, và `_archive` nằm ngoài `lib` nên không vào `sys.path`). Sau 1–2 tháng không ai cần thì xoá thật.

### 2a. XAML mồ côi (không Python nào tham chiếu) — 15 file

- [ ] `lib/GUI/Tools/AutoClick.xaml`
- [ ] `lib/GUI/Tools/BulkFamilyExport.xaml`
- [ ] `lib/GUI/Tools/InPlaceModelMain.xaml`
- [ ] `lib/GUI/Tools/InPlaceModelRename.xaml`
- [ ] `lib/GUI/Tools/InPlaceModelSettings.xaml`
- [ ] `lib/GUI/Tools/JSONtoFamily.xaml`
- [ ] `lib/GUI/Tools/LocationManager_backup.xaml`
- [ ] `lib/GUI/Tools/ModelChecker.xaml`
- [ ] `lib/GUI/Tools/ParaSync.xaml`
- [ ] `lib/GUI/Tools/ParameterLoaderCatPicker.xaml`
- [ ] `lib/GUI/Tools/ParameterLoaderMain.xaml`
- [ ] `lib/GUI/Tools/ParameterLoaderMapper.xaml`
- [ ] `lib/GUI/Tools/TransferPara.xaml`
- [ ] `lib/GUI/Tools/ConvertGridline.xaml` (chỉ DatumManagerDialog chết tham chiếu)
- [ ] `lib/GUI/Tools/ConvertLevel.xaml` (nt)

**Ngoại lệ — giữ lại:** `ExportManagerTest.xaml` đang được `.claude/CLAUDE.md` dẫn làm ví dụ wizard nav. Giữ nguyên, chỉ thêm comment đầu file `<!-- REFERENCE ONLY - not wired to any tool -->`.
- [ ] Thêm comment reference vào `ExportManagerTest.xaml`

### 2b. Dialog chết trong `lib/GUI/` — 6 module + XAML đi kèm

| Module | XAML kéo theo | Ghi chú |
|--------|----------------|---------|
| - [ ] `DatumManagerDialog.py` | `DatumManager.xaml` | Còn trỏ tới pushbutton `Datum.pulldown/...` **không tồn tại** |
| - [ ] `FamilyBatchCreatorDialog.py` | `FamilyBatchCreator.xaml` | Không ai import |
| - [ ] `OpeningAssignValuesDialog.py` | `OpeningAssignValues.xaml` | Không ai import |
| - [ ] `SheetHubDialog.py` | `SheetHub.xaml` | Không ai import (ManaSheets dùng ManaSheetsDialog riêng) |
| - [ ] `ViewHubDialog.py` | `ViewHub.xaml` | Không ai import (ManaViews dùng AdvancedViewManagerDialog) |
| - [ ] `ViewTemplateDialog.py` | `ViewTemplateManager.xaml`, `ViewTemplateBatchRename.xaml` | Không ai import |

- [ ] Trước khi chuyển từng file: `grep -rn "<tên module>" --include="*.py" T3Lab.extension` xác nhận 0 tham chiếu ngoài chính nó
- [ ] Chuyển cả cặp .py + .xaml vào `_archive/` giữ nguyên cây thư mục con

### 2c. Cập nhật audit script

- [ ] Nếu quyết định GIỮ file nào tại chỗ → thêm tên vào `KNOWN_ORPHAN_XAML` trong `dev/audit_tools.py` kèm comment lý do
- [ ] `python3 dev/audit_tools.py --quiet` → **exit 0, AUDIT: clean**
- [ ] `python3 dev/sync_wpf_styles.py --check` → vẫn 0 lệch (số file giảm là bình thường)

### Chốt ngày 2
- [ ] Commit: `chore: archive orphan XAML files and dead GUI dialogs`
- [ ] Cập nhật bảng tiến độ trong `dev/plan/README.md` (Ngày 1–2 → ✅)
- [ ] Ghi lại danh sách file đã archive vào mục Phát sinh bên dưới (để trace)

---

## Phát sinh trong giai đoạn 1

_(ghi tại đây các vấn đề mới phát hiện khi thực hiện)_
