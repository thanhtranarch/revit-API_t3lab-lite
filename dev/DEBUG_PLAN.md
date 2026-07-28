# T3Lab — Review & Kế hoạch Debug từng Tool

> Cập nhật: 2026-07-05 (chốt GĐ4) · Phạm vi: toàn bộ 41 pushbutton trong `T3Lab.extension/T3Lab.tab/`
> Công cụ đi kèm: `python3 dev/audit_tools.py` (audit tĩnh) · `python3 dev/audit_ui.py` (audit UI Lumina) · `python3 dev/sync_wpf_styles.py --check` (đồng bộ style)
>
> **📅 Kế hoạch thực hiện theo ngày (14 ngày, chia theo panel): xem [`dev/plan/README.md`](plan/README.md)** — file này là bản ghi review gốc; checklist thực hiện hằng ngày nằm trong `dev/plan/`.
>
> **✅ TRẠNG THÁI CUỐI — 2026-07-05 (GĐ4)**: Kế hoạch 12 ngày hoàn tất — GĐ1/GĐ2 xong từ
> 2026-07-03, GĐ3 (41/41 tool) đóng theo xác nhận chung của user 2026-07-05 *"các tool hiện
> tại đều đã hoạt động tốt"*, GĐ4 chốt cùng ngày. 4 roadmap độc lập đã đóng & xóa file (kết
> cục ghi ở `dev/plan/README.md`). Regression cuối: `audit_tools` clean · `audit_ui` 0/54 ·
> `sync_wpf_styles --check` 53/53 0 lệch. **Issue còn lại & review GĐ4: xem mục 7.**

---

## 1. Tổng quan kiến trúc

Codebase có **2 kiểu tool**:

| Kiểu | Mô tả | Số lượng |
|------|-------|----------|
| **Launcher** | `script.py` mỏng (15–56 dòng) → gọi `lib/GUI/<Tool>Dialog.py` (logic + WPF) → load XAML trong `lib/GUI/Tools/` | ~20 tool |
| **Self-contained** | Toàn bộ logic nằm trong `script.py` của pushbutton (122–3764 dòng), XAML vẫn ở `lib/GUI/Tools/` | ~18 tool |
| **Không UI / urlbutton** | PDF import, CloudLinks... | 3 |

### Kết quả audit tĩnh đã chạy (2026-07-02)

| Check | Kết quả |
|-------|---------|
| Syntax IronPython 2.7 (f-string, annotation, nonlocal) trên toàn bộ lib + tab | ✅ Sạch (trừ `lib/core/bridge.py` — chủ đích chạy CPython 3) |
| Chuỗi launcher → Dialog module → entry function → XAML | ✅ 19/19 launcher resolve đúng |
| Style block chia sẻ (`sync_wpf_styles.py --check`) | ✅ 76 file, 0 lệch |
| Transaction: `Start()` có `Commit()`/`Assimilate()` tương ứng | ✅ Không phát hiện transaction treo |
| Handler XAML ↔ method Python | ⚠️ 2 file lỗi thật (mục 2) |
| XAML mồ côi (không Python nào tham chiếu) | ⚠️ 12 file (mục 2) |

---

## 2. Phát hiện (Findings)

### ✅ F1 — `FindReplace.xaml` thiếu handler `button_run` (ĐÃ SỬA — GĐ1 Ngày 1)
- `lib/GUI/Tools/FindReplace.xaml:1181` khai báo `Click="button_run"` nhưng `lib/GUI/FindReplace.py` không có method này (WPF_Base cũng không).
- Vì class dùng `wpf.LoadComponent` → **XamlParseException ngay khi mở cửa sổ** nếu được gọi.
- Hiện không tool nào khởi tạo `FindReplace`, nhưng `lib/GUI/forms.py` vẫn import class này — bất kỳ tool mới nào dùng lại sẽ crash.
- **Hành động**: thêm `def button_run(self, sender, e)` hoặc xoá module nếu xác nhận không dùng.

### ✅ F2 — `CreateFromRooms.xaml` thiếu 4 handler (ĐÃ XỬ LÝ — GĐ1 Ngày 1)
- Thiếu: `NumberValidationTextBox`, `UIe_ItemChecked`, `button_run`, `text_filter_updated` trong `lib/GUI/CreateFromRooms.py`.
- Không nơi nào import `CreateFromRooms` → dead code, nhưng sẽ crash khi mở nếu được dùng lại.
- **Hành động**: xoá hoặc hoàn thiện handler trước khi tái sử dụng.

### ✅ F3 — UIShowcase phụ thuộc cấu trúc repo (ĐÃ SỬA — GĐ1 Ngày 1)
- `lib/GUI/UIShowcaseDialog.py:33` load XAML từ `[repo]/.claude/standard/UIStandardShowcase.xaml` — **ngoài** thư mục `T3Lab.extension`.
- Nếu deploy extension rời khỏi repo (copy `T3Lab.extension` vào `%APPDATA%\pyRevit\Extensions`), tool hỏng.
- **Hành động**: copy XAML vào `lib/GUI/Tools/` hoặc fallback path.

### ✅ F4 — 12 XAML mồ côi trong `lib/GUI/Tools/` (ĐÃ DỌN — GĐ1 Ngày 2, cùng F9)
`AutoClick`, `BulkFamilyExport`, `ConvertGridline`, `ConvertLevel`, `ExportManagerTest`(*), `InPlaceModelMain`, `JSONtoFamily`, `LocationManager_backup`, `ModelChecker`, `ParameterLoaderCatPicker/Main/Mapper`, `TransferPara` — không file Python nào tham chiếu.
- (*) `ExportManagerTest.xaml` đang được CLAUDE.md dẫn làm ví dụ wizard nav — giữ lại nhưng đánh dấu.
- `DatumManagerDialog.py` (dead code) còn trỏ tới các pushbutton `Datum.pulldown/...` **không tồn tại**.
- **Hành động**: dọn dẹp hoặc chuyển vào thư mục `_archive/`.

### 🟡 F5 — Bare `except:` nuốt lỗi hàng loạt (GIỮ NGUYÊN — tech debt, xem mục 7)
| File | Số bare except |
|------|----------------|
| `Wall Cut Profile/script.py` | 37 |
| `BatchOut/script.py` | 30 |
| `Wall_Adjust Base/script.py` | 4 |
| `DoorThreshold/script.py` | 3 |
| `ManaTabs/script.py` | 2 |

Khi debug các tool này, lỗi thật sẽ bị nuốt im lặng → phải tạm thêm logger vào các nhánh except trước khi test.

### 🟡 F6 — BatchOut: module "Intelligence" degrade im lặng (GIỮ NGUYÊN — tech debt, xem mục 7)
- `BatchOut/script.py:62-67`: import `api_learner`/`api_updater` bọc trong bare except → nếu fail, `HAS_API_LEARNER=False` không có log. Cần in cảnh báo khi debug.

### ✅ F7 — Tài liệu ui-design-standard lỗi thời về vị trí copyright (ĐÃ SỬA DOC 2026-07-02)
- Kiểm tra thực tế: **74/76 XAML và cả file chuẩn `UIStandardShowcase.xaml` đã đặt copyright ở footer (status bar) BÊN TRÁI** từ trước (copyright → divider 1px `#DCDCE0` → status text). Tài liệu `ui-design-standard.md` cũ bắt overlay góc phải-dưới (`Grid.RowSpan="99"`) mới là thứ sai so với codebase.
- **Đã cập nhật** `.claude/rules/ui-design-standard.md`: chuẩn chính thức = copyright ở footer, left-most element, theo snippet của `UIStandardShowcase.xaml`; Variant B không có footer dùng fallback overlay **trái**-dưới; item-template XAML (`CadtoFloorLayerItem.xaml`) được miễn trừ. `dev/audit_ui.py` đã sửa theo rule mới.
- Việc còn lại (2 outlier, đều thuộc T3LabAssistant — gộp xử lý với F8 trong GĐ2 Ngày 3):
  - `T3LabAssistant.xaml:1766` — copyright (số dòng cũ ~2109 đã lệch sau khi file co lại còn 1982 dòng; khối copyright vẫn còn, không bị mất)
  - `AssistantPane.xaml:~1357` — copyright theo chuẩn cũ (overlay góc **phải**-dưới); đây là file duy nhất từng theo doc cũ
- Các hạng mục UI khác đều đạt: palette Lumina (trừ F8), font, WindowChrome, glyph, shared styles, không dot-notation. KHÔNG sửa 2 file UI-locked (`DWGManagement.xaml`, `ExportManager.xaml`).

### ✅ F8 — T3LabAssistant.xaml sót palette Terra cũ (ĐÃ SỬA — GĐ2 Ngày 3)
- `#15803D` (Terra success) ×1, `#166534` (Terra success hover) ×1 → migrate sang `#10B981` / `#059669`.

### ✅ F9 — Mở rộng F4: thêm dead code phát hiện sau (ĐÃ DỌN — GĐ1 Ngày 2)
- XAML mồ côi thêm: `InPlaceModelRename.xaml`, `InPlaceModelSettings.xaml`, `ParaSync.xaml` (tổng 15 file).
- 6 dialog trung gian **không pushbutton nào gọi tới**: `DatumManagerDialog.py` (+`DatumManager.xaml`), `FamilyBatchCreatorDialog.py`, `OpeningAssignValuesDialog.py`, `SheetHubDialog.py`, `ViewHubDialog.py`, `ViewTemplateDialog.py` (+2 XAML).
- Danh sách archive đầy đủ: `dev/plan/phase-1-cleanup-fixes.md`.

### ✅ F11 — PDF import: mở tool không hiện thông tin + crash file (PHÁT HIỆN & SỬA — 2026-07-05, user báo)
- Bug functional thật (không phải UI): mở tool thì grid "Target Views" kẹt overlay "Loading
  views…" (không view nào hiện) và crash/đóng file Revit. Phơi ra rằng xác nhận chung "các
  tool hoạt động tốt" (2026-07-05) chưa thực sự test functional PDF import.
- **Nguyên nhân kép** trong `PDFImportDialog.py`:
  1. *Không hiện thông tin* — refactor `0a2a886` hoãn nạp grid sang event `ContentRendered`,
     mà event này **không chạy đáng tin dưới modal `ShowDialog()`** trong Revit → `_load_views()`
     không chạy → `pnl_loading` không collapse → kẹt "Loading views…".
  2. *Crash* — bind `ItemsSource` + mutate `ObservableCollection` đặt trong `Loaded`/`ContentRendered`
     (SAU khi cửa sổ render); mutate collection bound DataGrid non-virtualized giữa layout có thể
     **hard-crash host** (lỗi tầng WPF/native, không phải exception Python bắt được).
- **Fix (vòng 1)**: dời tạo OC + bind + `_load_views()` về `__init__`, chạy **đồng bộ trước `ShowDialog()`**
  (cửa sổ chưa render) — đúng pattern kiểm chứng ở `ManaViews`/`ManaContains`/`ManaSheets`. Xóa
  `_on_loaded`/`_on_content_rendered`/`_views_loaded`; guard `mode_changed` giữ nguyên.
- **Fix (vòng 2, 2026-07-05) — nguyên nhân chậm-mở + crash thật**: user báo lại tool vẫn **mở chậm +
  crash file**. Root cause còn lại: DataGrid `grid_views` đặt `EnableRowVirtualization="False"` →
  với model nhiều view, WPF dựng TOÀN BỘ row cùng lúc (mỗi row có template checkbox + nhiều Path/Border)
  → chậm + hard-crash host Revit. ManaViews (chạy tốt với nhiều view) **không** tắt virtualization và
  checkbox **không** có event — chỉ bind `IsSelected`. PDFImport tắt virtualization vì checkbox có
  handler `view_selection_changed` gọi `Items.Refresh()` mỗi lần tick (không tương thích virtualization).
  Sửa theo đúng pattern chuẩn: (a) bật lại virtualization (`VirtualizingPanel.IsVirtualizing=True` +
  `VirtualizationMode=Recycling`, bỏ `EnableRowVirtualization="False"`); (b) `ViewItem` implement
  `INotifyPropertyChanged` (giống `RenumberItem` của ManaSheets) — PAGE pill + checkbox cập nhật phản
  ứng, KHÔNG cần `Items.Refresh()`; (c) bỏ event `Checked/Unchecked` XAML, chuyển tính-lại-page +
  single-select sang **setter của `IsSelected`** (chỉ chạy khi user thực sự tick, KHÔNG chạy khi cuộn
  realize row → virtualization-safe). Static clean (`ast.parse` + XML well-formed + 3 audit).
  **Chưa smoke test Revit** — cần user reload pyRevit + mở lại (kỳ vọng mở nhanh, không crash kể cả
  model nhiều view; tick/single-select/page-number vẫn đúng).
- **Bài học**: (1) xác nhận "tất cả tool OK" theo lô KHÔNG thay được smoke test functional từng tool.
  (2) `EnableRowVirtualization="False"` trên DataGrid là bẫy chậm-mở + crash với danh sách lớn — không
  bao giờ tắt virtualization để né vấn đề checkbox; thay vào đó dùng `INotifyPropertyChanged` + setter-
  callback (không dùng `Items.Refresh()` + không dùng XAML `Checked` event vốn bắn cả khi cuộn).

### ✅ F10 — 2 điểm ghi JSON `ensure_ascii=False` trên file bytes-mode (PHÁT HIỆN & SỬA — review GĐ4, 2026-07-05)
- `lib/Intelligence/t3lab_assistant.py::save_learned_patterns` và `lib/config/user_profile.py::save`
  còn dùng `open(path, 'w')` + `json.dump(..., ensure_ascii=False)` — đúng failure class đã biết
  (từng làm hỏng chat history / tool_registry.json): IronPython 2.7 ném `UnicodeEncodeError` giữa
  chừng khi dữ liệu có tiếng Việt → file bị truncate 0 byte, lỗi bị `except: pass` / `return False`
  nuốt im lặng. `user_profile.py` **chắc chắn dính** vì mặc định `_DEFAULT_NAME = u"Thạnh"`.
- Đã sửa cả 2 theo pattern chuẩn của codebase: `json.dumps(..., ensure_ascii=True)` → decode →
  `io.open(path, 'w', encoding='utf-8')` ghi một lần. Các điểm ghi khác đã an toàn từ trước
  (history/registry trong `T3LabAssistant/script.py`, `tool_discovery.py` dùng cùng pattern;
  `IFCSGDialog.py` dùng `codecs.open`; body HTTP trong `llm_provider.py` encode UTF-8 đúng).

---

## 3. Quy trình smoke-test chuẩn (áp dụng cho mọi tool)

Chạy trong Revit (2022+ khuyến nghị, có model test nhỏ):

1. **Mở tool** từ ribbon → cửa sổ hiện, không XamlParseException (bắt lỗi handler/XAML).
2. **Chrome window**: kéo title bar, Minimize / Maximize / Close hoạt động; resize grip ở góc.
3. **Happy path**: chạy chức năng chính với input hợp lệ → kiểm tra kết quả trong model.
4. **Undo test**: 1 lần Ctrl+Z = revert đúng 1 thao tác tool (transaction/TransactionGroup gom đúng).
5. **Empty/edge case**: không chọn gì, model rỗng, huỷ giữa chừng → phải có thông báo thân thiện, không stacktrace đỏ.
6. **Log check**: bật pyRevit debug (`CTRL+click` nút hoặc `pyrevit env`) xem exception bị nuốt.
7. **Doc khác nhau**: chạy trên file workshared (nếu tool đụng workset/sheet) và file thường.

Ghi kết quả vào bảng mục 4 (cột Trạng thái: ⬜ chưa test / ✅ pass / ❌ fail + issue).

---

## 4. Kế hoạch debug từng tool (theo panel)

### 4.1 Annotation & Select.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 1 | AutoDimension | launcher → `AutoDimensionDialog.py` | 1721 | Logic dim theo reference dễ vỡ giữa các version Revit | Tạo dim trên wall/grid ở plan + section; test view không hỗ trợ | ✅ |
| 2 | ManaDWG | self + `DWGManagement.xaml` (UI-locked) | 432 | Xoá/relink DWG import vs link | List đúng DWG import & link; xoá 1 import; Undo | ✅ |
| 3 | ManaSelect | launcher → `ManaSelectDialog.py` | 660 | Filter selection theo category/param | Chọn hỗn hợp element rồi lọc; selection rỗng | ✅ |
| 4 | ManaAnno | launcher → `ManaAnnoDialog.py` | 1416 | Thao tác hàng loạt trên text/tag/dim | Batch đổi type annotation; view template ảnh hưởng | ✅ |

### 4.2 Modeling & Datum.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 5 | CADToElements | launcher → `CADToElementsDialog.py` (5 XAML con) | 3143 | Đọc geometry DWG; 3 flow (Wall/Floor/Beam) | Import DWG mẫu → convert từng loại; DWG không có layer khớp | ✅ |
| 6 | DoorThreshold | self | 356 | 3 bare except; family threshold phải có sẵn | Cửa trên wall thường + curtain wall; thiếu family | ✅ |
| 7 | ImageToDrafting | self | 1585 | **Compile C# runtime (Vectorizer)** — phụ thuộc .NET compiler trong Revit | Ảnh PNG đen trắng nhỏ; test cả 2 mode (centerline/outline); máy không compile được | ✅ |
| 8 | Text to Element | launcher → `TextToElementDialog.py` | 466 | Parse text note thành element | Text hợp lệ + text rác | ✅ |
| 9 | PointCloud | self | 1502 | 6 transaction; file scan lớn | File điểm nhỏ; format sai; cancel giữa chừng | ✅ |
| 10 | RoomToFloor | self | 383 | Boundary phức tạp (room có island/arc) | Room chữ nhật, room có cột giữa (opening), room chưa placed | ✅ |
| 11 | Tile Layout | self | 2450 | Không dùng Transaction trực tiếp trong script (kiểm tra wrapper) — logic layout tính toán nhiều | Layout trên floor nhỏ; đổi origin/góc; Undo | ✅ |
| 12 | PropertyLine | launcher → `PropertyLineDialog.py` | 1727 | Tạo property line từ tọa độ/bảng | Nhập bảng tọa độ hợp lệ + sai đơn vị | ✅ |
| 13 | FamiGen | launcher → `FamiGenDialog.py` | 2360 | Tạo family qua Document khác (family doc) — dễ leak doc | Gen 1 family, load vào project; huỷ giữa chừng xem family doc có đóng | ✅ |
| 14 | ManaFami | launcher → `ManaFamiDialog.py` | 1428 | Rename/purge family hàng loạt | Batch rename; family in-use | ✅ |
| 15 | Wall_Adjust Base | self | 311 | 4 bare except | Wall thường + wall bị attach top/base | ✅ |
| 16 | SplitElements | launcher → `SplitElementsDialog.py` | 61 | Dialog rất mỏng — logic có thể ở Snippets | Split wall/line theo level | ✅ |
| 17 | AutoJoin | self | 708 | JoinGeometry hàng loạt → chậm/fail từng cặp | Model nhỏ nhiều wall-floor giao nhau; phần tử không join được | ✅ |
| 18 | Wall Cut Profile | self | 1624 | **37 bare except**; 8 transaction; edit sketch profile | Wall thẳng + wall cong; profile đè lên nhau; Undo group | ✅ |

### 4.3 Views & Sheets.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 19 | ManaSheets | launcher → `ManaSheetsDialog.py` | 583 | Batch sửa sheet | Rename/renumber; sheet đang mở | ✅ |
| 20 | SheetGen | self | 1262 | 7 transaction, 0 rollback — lỗi giữa chừng để lại state dở | Gen sheet từ template; titleblock thiếu; fail giữa batch | ✅ |
| 21 | ManaViews | launcher → `ManaViewsDialog.py` | 689 | Xoá/đổi view hàng loạt | View đang active; view trên sheet | ✅ |
| 22 | BatchOut | self + `ExportManager.xaml` (UI-locked) | 3764 | **30 bare except; F6 (Intelligence degrade im lặng)**; export PDF/DWG/NWC phụ thuộc máy | Export 2 sheet PDF + DWG; máy không có PDF printer; đường dẫn output không tồn tại; kiểm tra `HAS_API_LEARNER` | ✅ |

### 4.4 Data & IFC-SG.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 23 | IFC-SG | launcher → `IFCSGDialog.py` + `configs/` | 2236 | Mapping param IFC-SG từ config JSON | Config mặc định; config thiếu field; apply lên model mẫu | ✅ |
| 24 | Foundation Volume | self | 410 | Tính volume qua geometry boolean | Foundation giao nhau; element không có solid | ✅ |
| 25 | ManaContains | launcher → `ManaContainsDialog.py` | 2204 | Logic spatial containment nặng | Element trong/ngoài room; link model | ✅ |
| 26 | ManaSched | launcher → `ManaSchedDialog.py` | 1601 | Sửa schedule field/filter qua API | Schedule thường + key schedule | ✅ |
| 27 | ManaPara | launcher → `ManaParaDialog.py` | 2548 | Shared param file; binding category | Add/remove shared param; file .txt bị lock | ✅ |
| 28 | BCF Reader | self | 3041 | Parse ZIP/XML BCF; 4 transaction; viewpoint camera | BCF 2.1 mẫu; BCF hỏng; apply viewpoint | ✅ |

### 4.5 Standards & Settings.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 29 | ManaLoca | self + `ManaLoca.xaml` | 871 | Site/shared coordinates — thao tác nguy hiểm với model thật | **Chỉ test trên file copy**; đổi project base point; Undo | ✅ |
| 30 | ManaStyles | launcher → `ManaStylesDialog.py` | 2091 | Object styles/line weights toàn model | Đổi 1 style; template lock | ✅ |
| 31 | ManaWorkset | self | 609 | Chỉ chạy được trên file workshared | File thường (phải báo lỗi đẹp); tạo/rename workset | ✅ |
| 32 | ModelAuditor | launcher → `ModelAuditorDialog.py` | 1828 | Quét toàn model → chậm trên model lớn | Model nhỏ ra report đúng; model lớn đo thời gian | ✅ |

### 4.6 Support.panel & Standard.panel

| # | Tool | Chain | LOC | Rủi ro chính | Test trọng tâm | Trạng thái |
|---|------|-------|-----|--------------|----------------|-----------|
| 33 | MCPControl | launcher → `MCPControlDialog.py` | 348 | Start/stop server `core/server.py`; port 48884 | Bật/tắt server; port bị chiếm | ✅ |
| 34 | Feedback | self | 232 | Gửi feedback ra ngoài (network) | Không có mạng; nội dung unicode tiếng Việt | ✅ |
| 35 | PDF import | self, không UI | 33 | Phụ thuộc API import PDF theo version Revit | Revit 2022+ import 1 PDF | ✅ |
| 36 | T3LabAssistant | self | 8115 | **Tool lớn nhất**; import động (`imp`/`importlib` có guard); gọi `core.server`; UI chat | Mở UI; server chưa chạy; gửi 1 lệnh; kiểm tra fallback import trên IronPython (`importlib.util` không tồn tại → phải rơi vào nhánh `imp`) | ✅ |
| 37 | ManaTabs | self | 147 | Thao tác UI Revit qua Windows API — dễ vỡ theo version | Revit nhiều tab; 2 bare except | ✅ |
| 38 | Ribbon Names | self | 154 | Sửa text ribbon qua AdWindows.dll | Đổi tên; restart Revit xem persist | ✅ |
| 39 | BG Theme | self | 122 | Đổi background/theme qua UIApplication settings | Toggle theme; Revit dark mode | ✅ |
| 40 | UIShowcase | launcher → `UIShowcaseDialog.py` | 129 | **F3: XAML nằm ngoài extension** | Mở từ repo checkout OK; mở từ extension deploy rời → fail | ✅ |
| 41 | AutoWork | self + `AutoWork.xaml` | 353 | Không có Transaction trong script — kiểm tra có ghi model không | Chạy happy path; xem có cần transaction | ✅ |

---

## 5. Thứ tự ưu tiên (ĐÃ HOÀN THÀNH TOÀN BỘ — giữ làm tài liệu quy trình)

**Đợt 1 — Sửa lỗi đã xác nhận (không cần Revit):**
1. F1: thêm `button_run` vào `FindReplace.py` (hoặc xoá module + bỏ import trong `forms.py`).
2. F2: xử lý `CreateFromRooms` (xoá hoặc hoàn thiện).
3. F3: đưa `UIStandardShowcase.xaml` vào trong extension.
4. F4: dọn 12 XAML mồ côi + `DatumManagerDialog.py`.

**Đợt 2 — Smoke test tool rủi ro cao (cần Revit):**
BatchOut (22) → Wall Cut Profile (18) → T3LabAssistant (36) → BCF Reader (28) → CADToElements (5) → ImageToDrafting (7) → SheetGen (20) → ManaLoca (29, trên file copy).

**Đợt 3 — Smoke test tool trung bình:**
Tile Layout, PointCloud, FamiGen, ManaPara, ManaContains, IFC-SG, ModelAuditor, ManaStyles, AutoDimension, ManaAnno.

**Đợt 4 — Tool mỏng/ít rủi ro:**
Các launcher còn lại + UI.stack + PDF import + Feedback.

**Sau mỗi đợt:** chạy lại `python3 dev/audit_tools.py --quiet` và `python3 dev/sync_wpf_styles.py --check` để bảo đảm không tạo regression tĩnh.

---

## 6. Ghi chú debug trong Revit

- Bật debug log: giữ `CTRL` khi bấm pushbutton → chạy kèm output window; hoặc `pyrevit env` / `pyrevit logs`.
- Với các file nhiều bare `except:` (F5), tạm chèn `import traceback; print(traceback.format_exc())` vào nhánh except khi truy lỗi — nhớ gỡ trước khi commit.
- Tool ghi model luôn test kèm **Ctrl+Z** để xác nhận transaction gom đúng 1 bước.
- Tool phụ thuộc máy (BatchOut in PDF, ImageToDrafting compile C#, MCPControl mở port) cần test trên đúng máy triển khai, không chỉ máy dev.

---

## 7. Tổng kết cuối — GĐ4 (2026-07-05)

### Kết quả
- **Regression cuối**: `audit_tools --quiet` clean · `audit_ui --quiet` 0/54 file có vấn đề
  (lỗi IFCSG copyright trùng ×2 từng ghi ở GĐ2 Ngày 4 nay đã hết) · `sync_wpf_styles --check`
  53/53, 0 lệch.
- **41 tool đóng theo xác nhận chung của user 2026-07-05**; bảng 6 panel (`dev/plan/panel-*.md`)
  không còn ô ⬜. **PDF import từng mở lại (🔄) do user báo crash + không hiện thông tin + mở chậm khi
  thực sự dùng — đã sửa 2 vòng (F11), user xác nhận đã debug hết 2026-07-05 → ✅.** Bài học vẫn giữ:
  xác nhận theo lô không thay được smoke test functional; tool chỉ mới sửa UI có thể còn bug functional.
- **Review code GĐ4** trên commit mới nhất `5c4c3b1` (Assistant native agent + providers, ~2.1k
  dòng): kiến trúc ExternalEvent đăng ký trên main thread + worker thread block chờ — đúng chuẩn
  modeless; UI thread chỉ dùng `get_status_instant()` (probe đầy đủ chạy nền) — đúng bài học cũ;
  ghi history/registry đã dùng pattern dumps→io.open an toàn; SheetGen bổ sung commit/rollback
  hợp lý; AutoDimension helper mới có guard đầy đủ. Phát hiện + sửa ngay **F10** (2 điểm ghi
  JSON không an toàn còn sót — xem mục 2).
- Screenshot bộ UI: ⏭️ skip có lý do — user QA trực tiếp trong Revit (tiền lệ GĐ2 Ngày 4).
- 4 roadmap độc lập: đóng toàn bộ 2026-07-05, file đã xóa — bảng kết cục ở `dev/plan/README.md`.

### Issue còn lại (xếp ưu tiên — nguồn: tổng hợp mục "Phát sinh" 6 file panel + review GĐ4)

> **Đợt fix 2026-07-05 — ĐÃ ĐÓNG**: issue 0/1 + F6 & DoorThreshold-message đã sửa (6 commit riêng
> lượt: PDF import F11 ×2 vòng · ManaSheets Excel · BatchOut F6 · DoorThreshold message + doc).
> Tất cả static-clean và **user xác nhận đã debug hết 2026-07-05** ("các issue đã debug hết").
> Không còn issue mở nào ở mức chặn chức năng; chỉ còn tech debt (F5) và backlog đóng-không-triển-khai.

| # | Trạng thái | Issue | Ghi chú |
|---|-----------|-------|---------|
| 0 | ✅ Xong (user xác nhận) | **PDF import — crash + không hiện thông tin + mở chậm (F11)**: v1 nạp grid đồng bộ trong `__init__`; v2 bật lại virtualization + `INotifyPropertyChanged` + selection qua setter. | Commit `91159c7` + `3eb5a9c`; user xác nhận OK. `panel-6` Phát sinh |
| 1 | ✅ Xong (user xác nhận) | **ManaSheets Excel Import/Export chết** (lệch API service ↔ dialog): thêm adapter `export_sheets`/`import_sheets` (dict-based, cột `ID` để round-trip khớp lại theo ElementId; giữ nguyên 2 method gốc cho `import_excel_dialog`), dialog khớp theo id trước rồi fallback sheet_number. | Commit `05a5628`; user xác nhận OK. `panel-3` Phát sinh |
| 2 | ✅ F6 xong · 🟡 F5 giữ | **F6** (BatchOut degrade Intelligence im lặng): đã log khi import `api_learner`/`api_updater` fail — commit `e535764`. **F5** (bare `except:` — Wall Cut Profile 37, BatchOut 30, Wall_Adjust 4, DoorThreshold 3, ManaTabs 2): **giữ nguyên làm tech debt** — không chặn chức năng, sửa hàng loạt trên tool đang chạy tốt là churn rủi ro không tương xứng; chỉ chèn log tạm khi thực sự cần debug tool đó. | mục 2 (F5/F6) |
| 3 | ✅ Message xong · 🟡 geometry giữ | **DoorThreshold**: đã thêm thông báo rõ khi model không có FloorType nào (commit `09cee9b`). Phần **curtain wall / wall nghiêng** vẫn ngoài scope — là feature geometry cần phát triển + test trong Revit, không phải bug; mở việc riêng nếu cần. | `panel-2` |
| 4 | 🟢 Thấp — đóng không triển khai | **Backlog** (mở roadmap mới nếu cần): AutoDimension GĐ B/C/D ~16 mục; FamiGen WS4/WS5; 3 probe FamiGen prompt v2 (chỉ mới pass theo xác nhận chung). | `dev/plan/README.md` |

**Kết luận GĐ4 (2026-07-05):** toàn bộ 12 ngày + 4 roadmap độc lập hoàn tất; mọi issue functional
đã đóng (user xác nhận đã debug hết). Còn lại chỉ là tech debt F5 và backlog cố ý hoãn — không có
việc bắt buộc nào đang mở.

---

## 8. Đợt debug T3Lab Assistant + LLMs Setting (2026-07-27)

> Hai tool AI là phần chưa từng đi qua chương trình debug GĐ1–GĐ4 (đóng 2026-07-05,
> phủ 41 tool còn lại). Phạm vi user chốt: **sửa lỗi + hoàn thiện UX, không thêm
> tính năng mới**. Icon ribbon dark theme: user từ chối, ngoài scope.

### Vì sao audit tĩnh không bắt được

Trước đợt này cả 3 audit + 2 test suite đều xanh; quét thêm AST, import chéo,
cú pháp py3-only, pyflakes, đối chiếu XAML handler ↔ Python (2 chiều) cũng sạch.
**Toàn bộ 33 lỗi dưới đây tìm bằng đọc code**, không phải bằng công cụ.

### Nhóm lỗi đã sửa

| # | Lỗi | Sửa |
|---|-----|-----|
| A1 | **Stop không dừng**: `_cancel_requested` chỉ được đọc ở spellcheck + native loop; vòng tool cũ (5 vòng) và knowledge/comment agent không kiểm tra → bấm Stop xong agent vẫn chạy tiếp tool **ghi** vào model, nút Stop bị disable nên không bấm lại được | 3 checkpoint trong vòng lặp + sau mỗi lần chạy tool, `_finish_cancelled` báo rõ "Stopped. N step(s) had already run", watchdog 20s giải phóng UI kể cả khi worker treo. Cancel exit của knowledge/comment agent **phải `return True`** — trả `False` sẽ khởi động lại request mới ngay sau khi user bấm Stop |
| A2 | **5 handler nuốt lỗi im lặng**: xoá typing dots + mở khoá busy nhưng không in gì → user gửi tin, xoáy quay biến mất, assistant không bao giờ trả lời | `_report_error()` dùng chung: log `logger.error`, hiện bubble cảnh báo, `_set_busy(False)` trong `finally` |
| A3 | `str(execute_err)` raise `UnicodeEncodeError` trên IronPython 2.7 khi lỗi Revit có ký tự non-ASCII (tên family tiếng Việt) — **nguyên nhân gốc của A2** | `_exc_text()` không bao giờ raise |
| A4 | Hết 5 vòng lặp → `result=None` → báo *"I didn't understand this request"* **sau khi 5 tool đã chạy thành công và sửa model** | Báo đúng giới hạn + số bước đã chạy, kèm chip "Continue" |
| A5 | `_stream_llm_turn(**kwargs)` nhận rồi vứt kwargs → vòng 1 stream không có JSON mode, câu trả lời đúng user vừa xem stream bị ghi đè bằng *"Could not read data from the model"* | Forward `**kwargs` (phải cùng commit với D4) |
| A6–A10 | Thread discovery `Dispatcher.Invoke` không guard; `reset_chat` không chặn khi đang bận + để lại ref bubble mồ côi; `_typing_timer` không stop khi đóng window (leak cả cây window); test `q in u'memory'` **ngược** (`/e`, `/r`, `/` đều hiện Memory); `undo_clicked` gọi `revit.doc.CanUndo()/.Undo()` **không tồn tại** trên Document | Guard + busy guard + stop timer + `u'memory'.startswith(q)` + `PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.Undo))` (async → copy là "Undo sent to Revit.") |
| B11–B17 | Dialog LLMs Setting: `Dispatcher.Invoke` cuối thread không guard (đóng window giữa lúc probe → crash); probe nền **ghi đè API key user đang dán**; nút Save Key **im lặng** khi ô đang hiện mask; đổi provider trong ~8s đầu không nạp model; cờ busy latch vĩnh viễn nếu thread không start; `-=`/`+=` handler không guard | `_ui_invoke`/`_start_worker` dùng chung; dirty-tracking (counter, không phải bool) + `force_fields`; Save Key re-validate key đã lưu; `_probe_pending` xếp hàng; un-latch khi start fail; `_prov_guard` thay cho detach/attach |
| C18–C22 | Ollama hỏng đầu-cuối khi không dùng host mặc định: `set_host` chỉ ghi RAM; `get_active_model` đi qua `OLLAMA_HOST` module-level nên hỏi localhost (dot xanh "Ready" nhưng Test báo "No model selected"); `_get_text`/`_post_json` **bỏ qua tham số timeout** (WebClient không có timeout → ~100s .NET mặc định); `format:"json"` **hardcode cho MỌI call**; `get_status_instant` gọi HTTP **trên UI thread** | Lưu `Ollama_Host` (đối xứng LM Studio); `pick_best()` thuần + `_probe_tags` đúng host; `HttpWebRequest` có `.Timeout`; `format` **opt-in** theo `response_format`; `get_status_instant` đọc settings, không I/O |
| D23–D29 | Router/settings: `probe_provider` không set `_status_ts`; `get_status` báo remote "available" chỉ vì **có key** (mâu thuẫn với `check_health` trong cùng dialog); 5 thread ghi dict chung không lock (IronPython không có GIL); `set_model` không refresh cache → chip model hiện model cũ; 4 setter ghi đè cả file từ dict cũ → **2 phiên Revit xoá key của nhau**; `_load_settings` fallback defaults với **mọi** lỗi đọc → 1 lần bật toggle trên file settings.json bị cụt là **mất sạch API key**; project override ghi đè provider mặc định toàn cục | TTL chỉ arm khi cache đủ provider; `check_health` thật + cờ `probed` (dot "Checking…"); lock cục bộ; refresh cache; `_update(mutator)` reload→patch→save cho mọi setter; phân biệt *thiếu file* / *parse lỗi* (quarantine `settings.corrupt-<stamp>.json`, giữ nguyên bytes) / *không đọc được* (từ chối ghi); `switch_provider(persist=False)` |
| E30–E33 | `AssistantPaneController` (~350 dòng) + `AssistantPane.xaml` (~1370 dòng) là **code chết** — `SetupDockablePane` nạp thẳng full window nên `get_pane_controller()` luôn `None`; `show_assistant_pane` vẫn trả `{'success': True}` nên agent tưởng đã gửi message; Command Palette (~170 dòng XAML + bảng 50 lệnh) **không có nút nào mở được**; log context-menu ghi **mỗi lần right-click**, đồng bộ trên UI thread, không giới hạn dung lượng | Xoá code chết (user chốt); `message_injected: False` + note trung thực; **mở** Command Palette (nút composer + Ctrl+K); log tắt mặc định (`T3LAB_CTXMENU_DEBUG=1`) + cap 256KB |

### KHÔNG phải lỗi (đã kiểm tra — đừng "sửa")

- ~~`_is_viet_text` trả `False` vô điều kiện — **cố ý**, có ghi chú ("UI + replies locked to English, 2026-07-18").~~ **Đã đảo lại ở §9 (2026-07-28)**: khoá này không thực sự làm UI thành tiếng Anh — các chuỗi tiếng Việt viết KHÔNG có nhánh `if viet:` vẫn hiện, nên cửa sổ lẫn hai ngôn ngữ.
- ~~`_FAST_CONTEXT_ENABLED = False` — **cố ý** (tắt 2026-07-06 vì keyword match cướp query không liên quan).~~ **Đã xoá hẳn ở §9 (2026-07-28)**: lý do tắt vẫn đúng, nhưng giữ ~220 dòng không bao giờ chạy chỉ tạo ảo giác còn đường tắt; tầng agent đã trả lời các câu hỏi này bằng tool thật.
- `_update_knowledge_status` là no-op hook — **cố ý** sau khi UI knowledge chuyển sang LLMs Setting.
- `get_status(use_cache=True)` **không có caller nào** trong repo → claim "TTL chết làm probe chậm" là phóng đại; vẫn arm TTL nhưng chỉ khi cache đủ provider (arm sau probe 1 provider sẽ phục vụ snapshot thiếu 4 provider trong 30s).
- 2 hàm knowledge/comment agent trả `False` ở except ngoài cùng là **đường degrade có chủ đích** (user vẫn nhận trả lời từ legacy path) — chỉ nâng log `debug` → `error`, không thêm bubble.
- Theo tiền lệ **F5**: chỉ sửa đúng 5 handler kết thúc lượt chat, KHÔNG quét ~310 handler còn lại.

### Regression

Thêm `dev/test_assistant_llm.py` (44 check, CPython3) — khoá lại 4 lỗi mất dữ liệu/UI nói dối
đã repro được, cộng host Ollama và định tuyến `response_format`. Chạy cùng
`test_knowledge.py`, `test_llm_config.py` và 3 audit; tất cả xanh.

**Còn lại cần QA trong Revit** (không tự bịa kết quả — chờ user xác nhận): Stop giữa chừng ·
lỗi hiện bubble · giới hạn 5 bước · ô API key không bị ghi đè · đổi provider lúc mới mở ·
Ollama host LAN + Test Connection trả câu văn · 2 phiên Revit · nút Undo · popup `/` · palette Ctrl+K ·
log right-click ngừng phình.

---

## 9. Đợt debug logic T3Lab Assistant (2026-07-28)

> Phạm vi user chốt: **tính logic của tool assistant và các vấn đề liên quan** — catalog tool,
> launcher, định tuyến, ngôn ngữ, side effect lúc khởi động. Không đụng XAML.

### Vì sao audit tĩnh không bắt được (lần nữa)

3 audit + 5 test suite đều xanh trước đợt này. Cả 4 nhóm lỗi dưới đây đều là **dữ liệu
tĩnh nói dối về thực tế trên đĩa** — không có lỗi cú pháp, không có import hỏng, không có
handler thiếu. Chỉ đối chiếu từng đường dẫn với `os.path.exists()` mới lộ ra.

### Nhóm lỗi đã sửa

| # | Lỗi | Sửa |
|---|-----|-----|
| G1 | **Catalog tool nói dối**: 10 intent mở tool được quảng cáo cho LLM; 7 trỏ tới pushbutton **không còn tồn tại** (ParaSync, ProjectName, Workset, DimText, UpperAll, Reset Overrides, Beam), `open_grids` không có launcher nào, `open_loadfamily_cloud` trùng khít `open_loadfamily` (`show_family_manager` chỉ có tab 0/1, không có tab cloud). Catalog bị nhân bản ở **10 bảng trong 5 module**, không bảng nào đối chiếu với đĩa | `TOOL_LAUNCHERS` giữ BatchOut + Family Loader, phần còn lại sinh từ registry `tool_discovery`, loại mọi intent không có script trên đĩa và ghi vào activity log. `_BUILTIN_TOOLS`, bảng trigger/threshold/label/message của `nlu_engine`, `AVAILABLE_INTENTS` của `t3lab_agent`, prompt local model — đều rút về 2 launcher đặc biệt + đọc registry |
| G2 | **`make_generic_launcher` báo thành công dối** — đường đi của **mọi** tool auto-discovered: nạp script bằng `imp.load_source` tên `_auto_<title>`, nên khối `if __name__ == '__main__':` mà **36/42 pushbutton** dùng không bao giờ chạy; vẫn `return True`. Kèm 4 launcher hardcode chỉ `import` rồi `return mod is not None` | `run_tool_script()` exec source với `__name__='__main__'` (đọc **bytes** — Python 2 từ chối coding declaration trong unicode source), trả `(ok, error_text)`; nhánh `imp` chỉ tính thành công khi **thật sự** `ShowDialog()` được; `SystemExit` = thành công |
| G3 | `_SKIP_BUTTONS` loại `PropertyLine.pushbutton` với lý do "đã hardcode trong TOOL_LAUNCHERS" — nhưng nó **tồn tại thật** và **chưa bao giờ** có trong `TOOL_LAUNCHERS` → assistant hoàn toàn không với tới được. `discover_new_tools` chỉ thêm, không bao giờ xoá → pushbutton bị đổi tên/xoá nằm lại registry vĩnh viễn | Skip list còn đúng hạ tầng + 2 launcher đặc biệt; `discover_new_tools` prune entry không còn script |
| G4 | **Song ngữ chết một nửa**: `_is_viet_text` khoá `return False` (2026-07-18) khiến 57 nhánh `if viet:` chỉ chạy vế Anh — nhưng chuỗi tiếng Việt viết KHÔNG có nhánh vẫn hiện (`/memory`, thông báo tải model, toàn bộ `AssistantCards`, xác nhận spellcheck). Một nhánh `if/else` có **hai vế giống hệt nhau**. Prompt cũng mâu thuẫn: `agent_loop` ép tiếng Anh trong khi ví dụ của `t3lab_assistant` phát ra tiếng Việt; prompt knowledge agent viết bằng tiếng Việt kèm dòng "always answer in English" | Setting `agents.reply_language` (`auto`/`vi`/`en`, mặc định auto-detect); `_note_user_language` ghi nhớ ngôn ngữ mỗi lượt để chuỗi không thuộc lượt nào (card, thông báo nền) đi theo; `AssistantCards` có đủ 2 biến thể; `build_agent_system_prompt` / `build_specialist_prompt` / `KnowledgeAgent.answer` / prompt pane nhận cùng tham số ngôn ngữ |
| G5 | **Khởi động chạy ngầm không xin phép**: `_bg_startup_probe` spawn `ollama serve` bằng `subprocess.Popen` và POST `/api/pull` tải `qwen2.5:1.5b` (timeout 600s) — tiến trình nền + ~1 GB tải về, không hỏi, không huỷ được, mọi lỗi nuốt trong `except Exception: logger.debug`. Ngay bên dưới, nhánh knowledge embeddings đã ghi đúng chuẩn ("the ~270 MB pull is never triggered silently") | `_offer_local_engine` chỉ **mời**, hành động nằm sau cú click; từ chối được nhớ (`agents.local_engine_declined`); không mời khi đã có cloud provider hoặc chưa cài Ollama. `_append_action_chips` gọi thẳng handler (khác `_append_quick_replies` vốn đẩy text qua LLM). Lỗi của các bước còn lại ra activity log |
| G6 | `_exec_tool` closure lên `doc_guard` được gán **sau 50 dòng** — chạy được chỉ nhờ thứ tự gọi; đổi thứ tự là `NameError` ngay tool call đầu tiên | Khai báo `doc_guard` trước `_exec_tool` |
| G7 | **Carryover quá lỏng**: `'?' in last_bot[-250:] and len(raw) <= 80` — dấu hỏi ở **bất kỳ đâu** trong 250 ký tự cuối đều tính, và 80 ký tự đủ chứa một lệnh mới. "tô vàng sàn" (13 ký tự) qua cả hai, chỉ còn escape hatch keyword chặn — một lớp duy nhất | `routing.is_continuation()`: tin cuối của bot phải **kết thúc** bằng câu hỏi, và câu trả lời phải **có dạng câu trả lời** (số thứ tự, chọn phương án, từ đồng ý, hoặc cụm danh từ ngắn không mang động từ hành động — lấy bảng động từ từ chính dispatcher) |
| G8 | Learned pattern (Jaccard ≥ 0.8) chạy **trước** NLU và cướp lượt vô điều kiện → mapping cũ thắng cả khi NLU đọc đúng catalog hiện tại; mapping học cho tool đã đổi tên vẫn thắng | Tính NLU trước, `routing.learned_pattern_wins()` phân xử: pattern chỉ thắng khi NLU không có ý kiến hoặc cả hai đồng ý |

### Regression

Thêm `dev/test_assistant_routing.py` (99 check, CPython3): launcher chạy thật, catalog khớp
đĩa, phát hiện ngôn ngữ, thang định tuyến, và **các cặp tranh chấp precedence của dispatcher**
(`đổi model` → multi_doc chứ không phải action; `bản vẽ` không tính là động từ `vẽ`; `tô đỏ`
chỉ khớp khi có dấu) — trước đây hoàn toàn không có test, sửa bảng từ khoá là âm thầm phá tuning.
Kèm một **drift check** fail ngay khi bất kỳ prompt tĩnh nào nhắc một intent `open_*` không phải
launcher đặc biệt, không có trong registry, và cũng không phải tool MCP.

`dev/test_knowledge.py` khẳng định prompt knowledge chỉ-tiếng-Việt cũ → đã cập nhật cho cả 2 ngôn ngữ.
Toàn bộ 5 test suite + 3 audit xanh.

### Còn lại cần QA trong Revit (không tự bịa kết quả — chờ user xác nhận)

1. `mở workset` → mở **ManaWorkset** thật, không phải lỗi console.
2. `mở property line` → mở PropertyLine (trước đợt này hoàn toàn không với tới được).
3. `mở parasync` → trả lời trung thực "không có tool này" + gợi ý tool gần nhất.
4. **Mọi tool auto-discovered giờ CHẠY THẬT** thay vì no-op im lặng — cần smoke test vài tool;
   một số có thể lộ lỗi vốn bị che.
5. Gõ tiếng Việt → toàn bộ reply (kể cả tool card, confirm card) tiếng Việt; gõ tiếng Anh → toàn Anh.
6. Máy chưa cài/chưa chạy Ollama → thấy lời mời có nút, **không** tự tải model.
7. Assistant hỏi lại → trả lời ngắn ("1", "toàn bộ project") → tiếp đúng mạch;
   gõ lệnh mới ngay sau câu hỏi → route mới, không lặp lại việc cũ.
8. Undo hoàn tác trọn TransactionGroup sau một lệnh sửa model.
